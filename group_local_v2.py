import json
import os
import io
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

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
    trace_id: str
) -> Dict[str, Any]:
    # デフォルト期間（既存ロジックをそのまま）
    if not start_date or not end_date:
        now = datetime.now()
        last_month_end = (now.replace(day=1) - timedelta(days=1))
        last_month_start = last_month_end.replace(day=1)
        one_year_ago = last_month_start - timedelta(days=365)
        start_date = one_year_ago.strftime("%Y-%m-%d")
        end_date = last_month_end.strftime("%Y-%m-%d")

    if not property_numbers:
        property_numbers = [
            "E10325967000001","E10325968000001","E10325969000001",
            "E10325970000001","E10325971000001","E10325972000001",
            "E10328768000001","E10329713000001","E10331622000001",
            "E10333612000001","E10333613000001","E10334125000001",
            "E10334306000001","E10334307000001","E10335046000001",
            "E10335047000001","E10335430000001","E10335438000001",
            "E10335991000001","E10336750000001","E10336882000001",
            "E10336883000001","E10336884000001","E10337382000001",
            "E10337579000001","E10338079000001","E10338080000001",
            "E10338081000001","E10338728000001","E10339072000001",
            "E10339069000001","E10339071000001","E10339070000001"
        ]

    jlog("info", "parameters",
         trace_id=trace_id,
         property_numbers_count=len(property_numbers),
         start_date=start_date, end_date=end_date,
         s3_bucket=S3_BUCKET_NAME,   # ← ここは常に環境変数
         s3_prefix=s3_prefix,
         region=AWS_REGION, dry_run=DRY_RUN)

    # BigQuery
    df = query_bigquery(property_numbers, start_date, end_date)

    # CSV へ
    csv_bytes = df_to_csv_bytes(df)
    filename = f"web_reservation_leaf_{datetime.today().replace(day=1).strftime('%Y%m%d')}.csv"
    key = f"{s3_prefix.rstrip('/')}/{filename}"

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
        # 入力取得
        property_numbers = event.get("property_numbers") if isinstance(event, dict) else None
        start_date = event.get("start_date") if isinstance(event, dict) else None
        end_date = event.get("end_date") if isinstance(event, dict) else None
        output_to_s3 = bool(event.get("output_to_s3", False)) if isinstance(event, dict) else False
        s3_prefix = event.get("s3_prefix", "bigquery-data/") if isinstance(event, dict) else "bigquery-data/"

        if not output_to_s3 and not DRY_RUN:
            jlog("warn", "output_to_s3=false and DRY_RUN=false: S3 へは出力されません", trace_id=trace_id)

        result = process(property_numbers, start_date, end_date, s3_prefix, trace_id)

        t_ms = round((time.perf_counter() - t_start) * 1000)
        jlog("info", "lambda done", trace_id=trace_id, ms=t_ms)
        # Step Functions 用にトップレベル result を返しつつ、API 互換の body も返す
        body = {"trace_id": trace_id, "result": result}
        return {
            "statusCode": 200,
            "result": result,               # ← これで $.Payload.result をそのまま参照可能
            "body": json.dumps(body, ensure_ascii=False),
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
    # ローカルテスト用
    test_event = {
        "property_numbers": ["E10331191000001", "E10331193000001"],
        "start_date": "2024-01-01",
        "end_date": "2024-12-31",
        "output_to_s3": True,
        "s3_bucket": "your-bucket-name",
        "s3_prefix": "bigquery-data/",
    }
    print(json.dumps(lambda_handler(test_event, None), ensure_ascii=False, indent=2))
