import json
import os
import io
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
from pathlib import Path

import boto3
from botocore.exceptions import ClientError

"""
Lambda: BigQuery -> CSV(in-memory) -> S3
デバッグしやすく、Step Functions 連携を正しくするための完全版。

主な特徴:
- 例外は握りつぶさず **raise** → Step Functions 上で "失敗(赤)" になる
- 返却 JSON は API や人間が読みやすい `statusCode/body` も含むが、**Step Functions 用にトップレベル `result` も返す**
- gRPC(=BigQuery Storage API) が使えない場合は **REST に自動フォールバック**
- S3 の SSE-KMS / Tagging / ACL を環境変数で指定可
- DRY_RUN / SAMPLE_LIMIT / FORCE_S3_FAIL のデバッグフラグ
- 1 行 1 JSON の構造化ログ (`jlog`)

Step Functions 側の取り出し例:
  "group.$": "$.Payload.result"
API Gateway 経由でもそのまま利用可。
"""

import os, json
import os.path as _p



def load_group_companies_config_from_s3(bucket: str, key: str) -> Dict[str, Any]:
    """
    S3からグループ会社設定JSONファイルを読み込む
    """
    try:
        s3 = boto3.client('s3', region_name=AWS_REGION)
        response = s3.get_object(Bucket=bucket, Key=key)
        config_data = json.loads(response['Body'].read().decode('utf-8'))
        group_count = len(config_data.get("group_companies", {}))
        jlog("info", "Config loaded from S3", bucket=bucket, key=key, group_count=group_count)
        return config_data
    except Exception as e:
        jlog("error", "Failed to load config from S3", error=str(e), bucket=bucket, key=key)
        return {"group_companies": {}}

def load_group_companies_config(config_path: str = None, s3_bucket: str = None, s3_key: str = None) -> Dict[str, Any]:
    """
    グループ会社設定を読み込む（S3優先、ローカルファイルはフォールバック）
    """
    # S3からの読み込みを優先
    if s3_bucket and s3_key:
        jlog("debug", "Loading config from S3", s3_bucket=s3_bucket, s3_key=s3_key)
        return load_group_companies_config_from_s3(s3_bucket, s3_key)
    
    # ローカルファイルからの読み込み（開発・テスト用）
    jlog("debug", "Checking config file", config_path=config_path, exists=os.path.exists(config_path) if config_path else False)
    if config_path and os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                group_count = len(config_data.get("group_companies", {}))
                jlog("info", "Config loaded from local file", path=config_path, group_count=group_count)
                return config_data
        except Exception as e:
            jlog("error", "Failed to load local config file", error=str(e), path=config_path)
    
    # デフォルト設定を使用（現在はコメントアウト）
    jlog("error", "No config file found", 
         config_path=config_path, 
         s3_bucket=s3_bucket, 
         s3_key=s3_key,
         local_file_exists=os.path.exists(config_path) if config_path else False)
    return {"group_companies": {}}

# def get_default_group_companies() -> Dict[str, Any]:
#     """
#     デフォルトのグループ会社設定を返す（あおば様とリーフ様）
#     """
#     return {
#         "group_companies": {
#             "aoba": {
#                 "name": "株式会社あおば",
#                 "property_numbers": [
#                     "E10325967000001","E10325968000001","E10325969000001",
#                     "E10325970000001","E10325971000001","E10325972000001",
#                     "E10328768000001","E10329713000001","E10331622000001",
#                     "E10333612000001","E10333613000001","E10334125000001",
#                     "E10334306000001","E10334307000001","E10335046000001",
#                     "E10335047000001","E10335430000001","E10335438000001",
#                     "E10335991000001","E10336750000001","E10336882000001",
#                     "E10336883000001","E10336884000001","E10337382000001",
#                     "E10337579000001","E10338079000001","E10338080000001",
#                     "E10338081000001","E10338728000001","E10339072000001",
#                     "E10339069000001","E10339071000001","E10339070000001"
#                 ]
#             },
#             "leaf": {
#                 "name": "株式会社リーフ",
#                 "property_numbers": [
#                     "E10331193000001","E10331197000001","E10331198000001",
#                     "E10331200000001","E10331201000001","E10337069000001",
#                     "E10337070000001"
#                 ]
#             }
#         }
#     }

def process_all_group_companies(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    s3_prefix: str = "bigquery-data/group_year_data/",
    config_path: str = None,
    s3_bucket: str = None,
    s3_key: str = None
) -> Dict[str, Any]:
    """
    全27グループ会社を順次処理する
    """
    trace_id = str(uuid.uuid4())
    t_start = time.perf_counter()
    
    jlog("info", "Starting batch processing for all group companies", trace_id=trace_id)
    
    # 設定ファイルを読み込み（ローカルファイル優先）
    jlog("debug", "Loading config", config_path=config_path, s3_bucket=s3_bucket, s3_key=s3_key)
    config = load_group_companies_config(config_path, s3_bucket, s3_key)
    group_companies = config.get("group_companies", {})
    
    # デバッグ: 読み込まれた設定の詳細ログ
    jlog("info", "Loaded group companies config", 
         total_groups=len(group_companies),
         group_ids=list(group_companies.keys())[:10],  # 最初の10個のみ表示
         config_source="S3" if s3_bucket and s3_key else "local" if config_path else "default")
    
    if len(group_companies) == 0:
        jlog("error", "No group companies found in config", config=config)
        return {
            "trace_id": trace_id,
            "success": False,
            "error": "No group companies found in config",
            "total_groups": 0,
            "success_count": 0,
            "failed_count": 0,
            "total_rows": 0,
            "results": []
        }
    
    results = []
    errors = []
    total_rows = 0
    
    for group_id, group_info in group_companies.items():
        try:
            # グループIDから会社の正式名称を取得
            company_name = group_info.get("name", group_id)
            
            jlog("info", "Processing company", 
                 company_name=company_name, 
                 property_count=len(group_info.get("property_numbers", [])))
            
            # デバッグ: 物件番号の詳細ログ
            property_numbers = group_info.get("property_numbers", [])
            jlog("debug", "Property numbers for company", 
                 company_name=company_name,
                 property_numbers=property_numbers[:5],  # 最初の5個のみ
                 total_count=len(property_numbers))
            
            # 個別グループ会社を処理
            result = process(
                property_numbers=property_numbers,
                start_date=start_date,
                end_date=end_date,
                s3_prefix=f"{s3_prefix.rstrip('/')}/group_year_data/",
                trace_id=f"{trace_id}_{group_id}",
                company_name=company_name
            )
            
            # 結果を記録
            group_result = {
                "company_name": company_name,
                "group_id": group_id,  # デバッグ用に残す
                "success": True,
                "rows": result.get("rows", 0),
                "s3_bucket": result.get("bucket"),
                "s3_key": result.get("key"),
                "s3_url": result.get("url")
            }
            results.append(group_result)
            total_rows += result.get("rows", 0)
            
            jlog("info", "Company processing completed", 
                 company_name=company_name, 
                 rows=result.get("rows", 0))
            
        except Exception as e:
            error_msg = f"Company {company_name} processing failed: {str(e)}"
            jlog("error", "Company processing failed", 
                 company_name=company_name,
                 error=str(e))
            
            group_result = {
                "company_name": company_name,
                "group_id": group_id,  # デバッグ用に残す
                "success": False,
                "error": str(e)
            }
            results.append(group_result)
            errors.append(error_msg)
    
    t_ms = round((time.perf_counter() - t_start) * 1000)
    success_count = len([r for r in results if r.get("success", False)])
    failed_count = len([r for r in results if not r.get("success", False)])
    
    jlog("info", "Batch processing completed", 
         trace_id=trace_id,
         total_companies=len(group_companies),
         success_count=success_count,
         failed_count=failed_count,
         total_rows=total_rows,
         ms=t_ms)
    
    return {
        "trace_id": trace_id,
        "total_companies": len(group_companies),
        "success_count": success_count,
        "failed_count": failed_count,
        "total_rows": total_rows,
        "results": results,
        "errors": errors,
        "processing_time_ms": t_ms
    }

def lambda_handler(event, context):
    cred_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    print(json.dumps({
        "check": "ADC path",
        "cred_path": cred_path,
        "exists": _p.exists(cred_path) if cred_path else False
    }))
    ...

# --- BigQuery 認証（Secrets Manager 経由） ---
from google.oauth2 import service_account

def get_bq_client():
    """
    Secrets Manager のシークレット（環境変数 SA_SERCRET_NAME）から
    GCP サービスアカウントJSONを取得して BigQuery Client を返す。
    """
    secret_name = os.environ.get("SA_SERCRET_NAME")  # ← ご指定の名前
    if not secret_name:
        raise RuntimeError("環境変数 SA_SERCRET_NAME が未設定です。")

    # プロジェクトIDは既存の BIGQUERY_PROJECT_ID を使用
    project_id = os.environ.get("BIGQUERY_PROJECT_ID", "epark-refre-users")
    if not project_id:
        raise RuntimeError("環境変数 BIGQUERY_PROJECT_ID が未設定です。")

    # Secrets Manager から鍵JSONを取得
    sm = boto3.client("secretsmanager", region_name=AWS_REGION)
    sec = sm.get_secret_value(SecretId=secret_name)["SecretString"]
    keyinfo = json.loads(sec)

    # 資格情報を生成して BigQuery クライアント作成
    creds = service_account.Credentials.from_service_account_info(keyinfo)
    from google.cloud import bigquery
    return bigquery.Client(project=project_id, credentials=creds)

# ---- 環境フラグ ----
DEBUG = os.getenv("DEBUG", "true").lower() == "true"
DRY_RUN = os.getenv("DRY_RUN", "false").lower() == "true"
FORCE_S3_FAIL = os.getenv("FORCE_S3_FAIL", "false").lower() == "true"
SAMPLE_LIMIT = int(os.getenv("SAMPLE_LIMIT", "0"))  # 0 = 無制限

# S3 追加設定 (必要に応じて)
S3_CONTENT_TYPE = os.getenv("S3_CONTENT_TYPE", "text/csv")
S3_SSE = os.getenv("S3_SSE")  # 例: "aws:kms"
S3_KMS_KEY_ID = os.getenv("S3_KMS_KEY_ID")  # 例: arn:aws:kms:ap-northeast-1:1234:key/xxxx
S3_TAGGING = os.getenv("S3_TAGGING")  # 例: "project=refre&owner=data"
S3_ACL = os.getenv("S3_ACL")  # 例: "bucket-owner-full-control"

# BigQuery 設定
BQ_PROJECT_ID = os.getenv("BIGQUERY_PROJECT_ID", "epark-refre-users")
BQ_CREDENTIALS_PATH = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
# 追加：S3 既定バケット（環境変数）
S3_BUCKET_NAME = os.getenv("S3_BUCKET_NAME")
if not S3_BUCKET_NAME:
    raise RuntimeError("環境変数 S3_BUCKET_NAME が未設定です")

AWS_REGION = os.getenv("AWS_REGION", os.getenv("AWS_DEFAULT_REGION", "ap-northeast-1"))


def jlog(level: str, msg: str, **fields: Any) -> None:
    payload = {
        "ts": datetime.utcnow().isoformat() + "Z",
        "level": level,
        "msg": msg,
        "fields": fields,
    }
    print(json.dumps(payload, ensure_ascii=False))


cred_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
from os import path as _p
jlog("info", "ADC check",
     has_env=bool(cred_path),
     path=cred_path if cred_path else None,
     exists=_p.exists(cred_path) if cred_path else False,
     has_json_env=bool(os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")),
     has_secret=bool(os.getenv("GCP_SA_JSON_SECRET")))

def get_s3_client():
    try:
        c = boto3.client("s3", region_name=AWS_REGION)
        return c
    except Exception as e:
        jlog("error", "S3 client creation failed", error=str(e))
        return None


@dataclass
class S3Result:
    success: bool
    bucket: Optional[str] = None
    key: Optional[str] = None
    url: Optional[str] = None
    error: Optional[str] = None


def upload_bytes_to_s3(data: bytes, bucket: str, key: str) -> S3Result:
    if FORCE_S3_FAIL:
        return S3Result(success=False, error="FORCE_S3_FAIL=true")

    s3 = get_s3_client()
    if not s3:
        return S3Result(success=False, error="S3 client not available")

    extra_args: Dict[str, Any] = {"ContentType": S3_CONTENT_TYPE}
    if S3_SSE:
        extra_args["ServerSideEncryption"] = S3_SSE
    if S3_KMS_KEY_ID:
        extra_args["SSEKMSKeyId"] = S3_KMS_KEY_ID
    if S3_TAGGING:
        extra_args["Tagging"] = S3_TAGGING
    if S3_ACL:
        extra_args["ACL"] = S3_ACL

    try:
        s3.put_object(Bucket=bucket, Key=key, Body=data, **extra_args)
        url = f"https://{bucket}.s3.{AWS_REGION}.amazonaws.com/{key}"
        return S3Result(success=True, bucket=bucket, key=key, url=url)
    except ClientError as e:
        return S3Result(success=False, error=f"ClientError: {e.response.get('Error', {}).get('Message', str(e))}")
    except Exception as e:
        return S3Result(success=False, error=f"Exception: {str(e)}")


# ---------- BigQuery ----------

def query_bigquery(property_numbers: List[str], start_date: str, end_date: str):
    """BigQuery から DataFrame を取得。BQS が無ければ REST にフォールバック。"""
    from google.cloud import bigquery

    # BQ 資格情報
    if BQ_CREDENTIALS_PATH:
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = BQ_CREDENTIALS_PATH

    client = get_bq_client()

    # Storage API (gRPC) はオプション
    bqstorage_client = None
    try:
        from google.cloud import bigquery_storage
        bqstorage_client = bigquery_storage.BigQueryReadClient()
        jlog("info", "BigQuery Storage API enabled")
    except Exception as e:
        jlog("warn", "BigQuery Storage API unavailable; fallback to REST", error=str(e))

    prop_cond = ",".join([f"'{x}'" for x in property_numbers])

    query = f"""
    WITH w_web AS (
        SELECT 
            concat(s.shopowner_id,"_",r.time_appoint_id) as petano_log_no,
            f.facility_id,
            f.property_number,
            r.shopowner_no,
            s.shopowner_id,
            CASE
                WHEN r.SYSTEM_USAGE_KIND = "1" THEN "初回受付"
                WHEN r.SYSTEM_USAGE_KIND = "2" THEN "再受付"
                WHEN r.SYSTEM_USAGE_KIND = "3" THEN "初回キャンセル"
            END AS SYSTEM_USAGE_NAME,
            m.MENU_NAME,
            m.PRICE,
            m.TIME_REQUIRED,
            SITE_CODE,
            FORMAT_TIMESTAMP('%Y/%m/%d %H:%M', PARSE_TIMESTAMP('%Y%m%d%H%M', r.APPOINT_START_DATETIME)) AS modify_start_date,
            PARSE_TIMESTAMP('%Y%m%d%H%M', r.APPOINT_START_DATETIME) AS appoint_timestamp,
            FORMAT_DATETIME('%Y-%m-%d %H:%M:%S', DATETIME(PARSE_TIMESTAMP('%Y%m%d%H%M', r.APPOINT_START_DATETIME))) AS appoint_datetime_clean,
            substr(r.create_date,0,10) AS create_date
        FROM `epark-refre-rook.rook.T_MR_T_TIME_APPOINT_20250831` r
        LEFT JOIN `epark-refre-rook.common.T_AD_M_SHOPOWNER_20250831` s ON s.shopowner_no = r.shopowner_no
        LEFT JOIN `epark-refre-rook.rook.T_MR_M_TIME_APPOINT_MENU_20250831` m ON m.shopowner_no = r.shopowner_no AND m.MENU_ID = r.MENU_ID
        LEFT JOIN `epark-refre-rook.rook.T_MR_M_TIME_APPOINT_STAFF_20250831` ss ON ss.STAFF_ID = r.STAFF_ID AND ss.SHOPOWNER_NO = r.SHOPOWNER_NO
        LEFT JOIN `epark-refre-users.public.T_BI_M_FACILITY_20250831` f ON f.RESERVE_SHOPOWNER_ID = s.SHOPOWNER_ID
        WHERE r.inside_outside_flg = "2"
          AND r.SYSTEM_USAGE_KIND != "0"
          AND f.status = "1"
          AND f.property_number IN ({prop_cond})
    )
    SELECT w_web.*
    FROM w_web
    WHERE appoint_timestamp BETWEEN TIMESTAMP("{start_date} 00:00:00") AND TIMESTAMP("{end_date} 23:59:59")
    ORDER BY appoint_timestamp DESC
    """

    jlog("info", "BQ query start", project=BQ_PROJECT_ID, start_date=start_date, end_date=end_date)
    t0 = time.perf_counter()
    job = client.query(query)

    # Storage API が有効なら DataFrame 生成で使う
    if bqstorage_client:
        df = job.to_dataframe(bqstorage_client=bqstorage_client)
    else:
        df = job.to_dataframe()

    if SAMPLE_LIMIT and len(df) > SAMPLE_LIMIT:
        df = df.head(SAMPLE_LIMIT)

    dur = round((time.perf_counter() - t0) * 1000)
    jlog("info", "BQ query done", rows=len(df), ms=dur)
    return df


def df_to_csv_bytes(df) -> bytes:
    import pandas as pd  # レイヤーに含まれている想定
    buf = io.StringIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig", na_rep="")
    return buf.getvalue().encode("utf-8-sig")


# ---------- Core processing ----------

def process(
    property_numbers: Optional[List[str]],
    start_date: Optional[str],
    end_date: Optional[str],
    s3_prefix: str,
    trace_id: str,
    company_name: str = None
) -> Dict[str, Any]:
    # 期間はlambda_handlerで設定済み
    jlog("debug", "process関数で受信した期間", start_date=start_date, end_date=end_date)

    if not property_numbers:
        return {"success": False, "error": "property_numbersが指定されていません"}
    jlog("info", "parameters",
         trace_id=trace_id,
         company_name=company_name,
         property_numbers_count=len(property_numbers),
         property_numbers_sample=property_numbers[:3] if property_numbers else [],
         start_date=start_date, end_date=end_date,
         s3_bucket=S3_BUCKET_NAME,   # ← ここは常に環境変数
         s3_prefix=s3_prefix,
         region=AWS_REGION, dry_run=DRY_RUN)

    # BigQuery
    df = query_bigquery(property_numbers, start_date, end_date)
    
    # デバッグ: BigQuery結果のログ
    jlog("info", "BigQuery result", 
         company_name=company_name,
         rows_count=len(df) if df is not None else 0,
         columns=df.columns.tolist() if df is not None and not df.empty else [])

    # CSV へ
    csv_bytes = df_to_csv_bytes(df)
    # 会社名を使用してファイル名を生成
    company_name_for_file = company_name or "unknown"
    filename = f"web_reservation_{company_name_for_file}_{datetime.today().replace(day=1).strftime('%Y%m%d')}.csv"
    
    # S3キーに会社名フォルダを含める
    year_month = f"{datetime.today().year}_{datetime.today().month:02d}"
    key = f"{s3_prefix.rstrip('/')}/{year_month}/{company_name_for_file}/{filename}"
    
    # デバッグ: ファイル名とS3キーのログ
    jlog("info", "Generated file info", 
         company_name=company_name,
         filename=filename,
         s3_key=key,
         csv_size_bytes=len(csv_bytes))

    if DRY_RUN:
        jlog("warn", "DRY_RUN: skip S3 upload", key=key, bytes=len(csv_bytes), bucket=S3_BUCKET_NAME)
        return {"rows": len(df), "dry_run": True, "bucket": S3_BUCKET_NAME, "key": key, "url": None}

    # S3 へ（常に S3_BUCKET_NAME を使用）
    t0 = time.perf_counter()
    s3res = upload_bytes_to_s3(csv_bytes, S3_BUCKET_NAME, key)
    dur = round((time.perf_counter() - t0) * 1000)
    if not s3res.success:
        jlog("error", "S3 upload failed", error=s3res.error, ms=dur)
        raise RuntimeError(s3res.error or "S3 upload failed")
    jlog("info", "S3 upload ok", bucket=s3res.bucket, key=s3res.key, url=s3res.url, ms=dur)

    return {"rows": len(df), "bucket": s3res.bucket, "key": s3res.key, "url": s3res.url}

# ---------- Lambda handler ----------

def lambda_handler(event, context):
    trace_id = str(uuid.uuid4())
    t_start = time.perf_counter()
    jlog("info", "lambda start", trace_id=trace_id, event=event)

    try:
        # 全27グループ会社を順次処理
        jlog("info", "Starting batch processing for all group companies", trace_id=trace_id)
        
        # イベントからPayloadを取得
        payload = event.get("Payload", {}) if isinstance(event, dict) else {}
        jlog("debug", "Event payload", payload=payload)
        
        # 期間を明示的に設定（常に13ヶ月前から前月まで）
        start_date = payload.get("start_date") if isinstance(payload, dict) else None
        end_date = payload.get("end_date") if isinstance(payload, dict) else None
        
        # 期間を常に13ヶ月前から前月までに設定（強制実行）
        if True:  # 常に実行
            from dateutil.relativedelta import relativedelta
            now = datetime.now()
            # 過去13ヶ月前の1日を開始日とする
            thirteen_months_ago = now - relativedelta(months=13)
            thirteen_months_ago = thirteen_months_ago.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            
            # 前月の最終日を終了日とする
            last_month_end = (now.replace(day=1) - timedelta(days=1))
            last_month_end = last_month_end.replace(hour=23, minute=59, second=59, microsecond=999999)
            
            start_date = thirteen_months_ago.strftime("%Y-%m-%d")
            end_date = last_month_end.strftime("%Y-%m-%d")
            
            jlog("info", "Lambda期間設定（強制実行）", 
                 start_date=start_date, 
                 end_date=end_date,
                 thirteen_months_ago=thirteen_months_ago.strftime("%Y-%m-%d %H:%M:%S"),
                 last_month_end=last_month_end.strftime("%Y-%m-%d %H:%M:%S"),
                 now=now.strftime("%Y-%m-%d %H:%M:%S"),
                 months_diff=13)
        s3_prefix = payload.get("s3_prefix", "bigquery-data/") if isinstance(payload, dict) else "bigquery-data/"
        
        # 設定ファイルの場所を指定（S3優先）
        config_path = payload.get("config_path") if isinstance(payload, dict) else None
        s3_bucket = payload.get("s3_bucket") if isinstance(payload, dict) else None
        s3_key = payload.get("s3_key", "bigquery-data/group_companies_config.json") if isinstance(payload, dict) else "bigquery-data/group_companies_config.json"
        
        jlog("debug", "Config parameters", s3_bucket=s3_bucket, s3_key=s3_key, config_path=config_path)
        
        result = process_all_group_companies(start_date, end_date, s3_prefix, config_path, s3_bucket, s3_key)
        
        t_ms = round((time.perf_counter() - t_start) * 1000)
        jlog("info", "batch processing done", trace_id=trace_id, ms=t_ms)
        
        return {
            "statusCode": 200,
            "result": result,
            "body": json.dumps({"trace_id": trace_id, "result": result}, ensure_ascii=False),
        }

    except Exception as e:
        t_ms = round((time.perf_counter() - t_start) * 1000)
        jlog("error", "lambda failed", trace_id=trace_id, ms=t_ms, error=str(e))
        # Step Functions を「失敗（赤）」にしたいので raise
        # ただし API 経由のデバッグもしやすいように一応戻り値も構築してから投げる
        body = {"trace_id": trace_id, "error": str(e)}
        # API Gateway 専用に 500 を返したいなら、ここで return に切り替える運用も可
        raise


# -------- Local test --------
if __name__ == "__main__":
    # ローカルテスト用（バッチ処理）
    # 期間を指定しない場合のデフォルト動作をテスト
    test_event = {
        "Payload": {
            "s3_prefix": "bigquery-data/",
            "s3_bucket": "monthly-report-output-group",
            "s3_key": "bigquery-data/group_companies_config.json"
            # start_date, end_date を指定しない（デフォルト期間を使用）
        }
    }
    print(json.dumps(lambda_handler(test_event, None), ensure_ascii=False, indent=2))
