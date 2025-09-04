from pathlib import Path
import os
import json
import csv
from typing import Optional, Tuple, List, Dict
from datetime import datetime
import io


def _get_boto3_client_or_none(service: str):
    """boto3 クライアントを取得（ローカル環境では None を返す）"""
    try:
        import boto3  # type: ignore
        return boto3.client(service)
    except Exception:
        return None


def find_latest_leaf_csv_in_s3(bucket: str, prefix: str) -> Optional[Tuple[str, str]]:
    """S3 上の prefix 配下から最新の web_reservation_leaf_*.csv を返す"""
    s3 = _get_boto3_client_or_none("s3")
    if s3 is None:
        return None
    try:
        paginator = s3.get_paginator("list_objects_v2")
        latest: Optional[Dict] = None
        for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
            for obj in page.get("Contents", []):
                key = obj.get("Key", "")
                if key.endswith(".csv") and "web_reservation_leaf_" in key:
                    if latest is None or obj["LastModified"] > latest["LastModified"]:
                        latest = obj
        if latest is None:
            return None
        return bucket, latest["Key"]
    except Exception:
        return None


def load_csv_from_s3(bucket: str, key: str) -> Optional[List[Dict]]:
    """S3 から CSV を読み込んで辞書リストとして返す"""
    s3 = _get_boto3_client_or_none("s3")
    if s3 is None:
        return None
    try:
        obj = s3.get_object(Bucket=bucket, Key=key)
        body = obj["Body"].read().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(body))
        return list(reader)
    except Exception:
        try:
            obj = s3.get_object(Bucket=bucket, Key=key)
            body = obj["Body"].read().decode("utf-8")
            reader = csv.DictReader(io.StringIO(body))
            return list(reader)
        except Exception:
            return None


def parse_datetime(date_str: str) -> Optional[datetime]:
    """日時文字列を datetime に変換（ISO8601/TZ対応：“2025-08-08 16:30:00+00:00”等もOK）"""
    if not date_str:
        return None
    s = str(date_str).strip()
    # ISO8601（+00:00 / Z / T区切り）優先
    try:
        return datetime.fromisoformat(s.replace("T", " ").replace("Z", "+00:00"))
    except Exception:
        pass
    # 既存フォーマット群
    for fmt in (
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def get_datetime_series(data: List[Dict]) -> List[Optional[datetime]]:
    """データの日時列を datetime(JST) のリストに変換"""
    from zoneinfo import ZoneInfo
    JST = ZoneInfo("Asia/Tokyo")

    # 使う列を決める（実データが入っている最初の列）
    for col in ["appoint_timestamp", "appoint_datetime_clean", "modify_start_date"]:
        if data and col in data[0]:
            parsed: List[Optional[datetime]] = []
            for row in data:
                dt = parse_datetime(row.get(col))
                if dt is None:
                    parsed.append(None)
                    continue
                # tz を JST に揃える
                if dt.tzinfo:
                    dt = dt.astimezone(JST)
                else:
                    dt = dt.replace(tzinfo=JST)
                parsed.append(dt)
            # 1件でも有効ならその列を採用
            if any(p is not None for p in parsed):
                return parsed
    raise ValueError("日時列が見つからない")


def _create_excel_file(data: List[Dict], output_path: Path, sheet_title: str = "データ") -> None:
    """Excelファイルを作成する共通関数"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    
    if data:
        headers = list(data[0].keys())
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        for row_idx, row_data in enumerate(data, 2):
            for col_idx, header in enumerate(headers, 1):
                ws.cell(row=row_idx, column=col_idx, value=row_data.get(header))
    
    wb.save(str(output_path))


def _get_target_months() -> List[Tuple[int, int]]:
    """過去2か月の対象月を取得する共通関数"""
    from zoneinfo import ZoneInfo
    current_date = datetime.now(ZoneInfo("Asia/Tokyo"))
    target_months = []
    
    for i in range(2):
        if current_date.month - (i + 1) <= 0:
            year = current_date.year - 1
            month = 12 + (current_date.month - (i + 1))
        else:
            year = current_date.year
            month = current_date.month - (i + 1)
        target_months.append((year, month))
    
    return target_months


def _generate_outputs(data: List[Dict], results_dir: Path) -> List[Path]:
    # outputs を最初に初期化
    outputs: List[Path] = []

    # 対象月を取得
    target_months = _get_target_months()

    """フィルタしたデータを月別・カテゴリ別に Excel 出力"""
    if not data:
        # データが空の場合はフォールバックファイルのみ生成
        fb = results_dir / "fallback_all_rows.xlsx"
        _create_excel_file([], fb, "ALL")
        outputs.append(fb)
        print({"metric":"fallback_emitted","path":str(fb)})
        return outputs
    
    for col in ["SYSTEM_USAGE_NAME", "SITE_CODE"]:
        if col not in data[0]:
            raise ValueError(f"必須列が見つかりません: {col}")

    dt_list = get_datetime_series(data)
    # デバッグ情報を追加
    print(f"データ件数: {len(data)}")
    print(f"日時解析結果: {len([d for d in dt_list if d is not None])}件の有効な日時")
    
    # サンプルデータの情報を出力
    if data:
        sample_row = data[0]
        print(f"利用可能な列: {list(sample_row.keys())}")
        print(f"SYSTEM_USAGE_NAMEの値: {sample_row.get('SYSTEM_USAGE_NAME', 'N/A')}")
        print(f"SITE_CODEの値: {sample_row.get('SITE_CODE', 'N/A')}")
        # 日時列のサンプル
        for col in ["appoint_timestamp", "appoint_datetime_clean", "modify_start_date"]:
            if col in sample_row:
                print(f"{col}のサンプル: {sample_row.get(col, 'N/A')}")
                break

    print(f"対象月: {target_months}")

    def is_hp(site_code: str) -> bool:
        return str(site_code).strip().lower() == "hp"
    for target_year, target_month in target_months:
        month_name = f"{target_month}月"
        filtered_data = []
        month_matches = 0
        for i, row in enumerate(data):
            dt = dt_list[i]
            if dt and dt.month == target_month and dt.year == target_year:
                month_matches += 1
                site_code = row.get("SITE_CODE", "")
                system_usage = row.get("SYSTEM_USAGE_NAME", "")
                site_type = "hp" if is_hp(site_code) else "その他"
                if system_usage == "再受付" and site_type == "hp":
                    filtered_data.append(("再受付_hp", row))
                elif system_usage == "初回受付" and site_type == "hp":
                    filtered_data.append(("初回受付_hp", row))
                elif system_usage == "再受付" and site_type == "その他":
                    filtered_data.append(("再受付_その他", row))
                elif system_usage == "初回受付" and site_type == "その他":
                    filtered_data.append(("初回受付_その他", row))
        
        print(f"{month_name}: 月マッチ{month_matches}件, カテゴリマッチ{len(filtered_data)}件")

        grouped_data: Dict[str, List[Dict]] = {}
        for category, row in filtered_data:
            grouped_data.setdefault(category, []).append(row)

        results_dir.mkdir(parents=True, exist_ok=True)
        for category, rows in grouped_data.items():
            filename = f"{month_name}_{category}.xlsx"
            out_path = results_dir / filename
            _create_excel_file(rows, out_path, "データ")
            outputs.append(out_path)

    # メイン処理でファイルが生成されなかった場合のフォールバック
    if not outputs:
        fb = results_dir / "fallback_all_rows.xlsx"
        _create_excel_file(data, fb, "ALL")
        outputs.append(fb)
        print({"metric":"fallback_emitted","path":str(fb)})

    return outputs


def process_monthly_usage(
    *,
    s3_bucket: str,
    s3_key: Optional[str] = None,
    s3_input_prefix: Optional[str] = None,
    output_to_s3: bool = False,
) -> dict:
    """S3 から CSV を読み込み、Excel を bigquery-data/ に出力する処理"""

    data: Optional[List[Dict]] = None
    input_desc = ""

    # 入力
    key_to_use: Optional[str] = s3_key
    if key_to_use is None:
        # デフォルトで直下を探索（prefix 未指定なら ""）
        latest = find_latest_leaf_csv_in_s3(s3_bucket, s3_input_prefix or "")
        if latest is None:
            return {"success": False, "error": "S3に対象CSVが見つかりません"}
        _, key_to_use = latest

    data = load_csv_from_s3(s3_bucket, key_to_use)
    if data is None:
        return {"success": False, "error": "S3からのCSV読み込みに失敗しました"}
    input_desc = f"s3://{s3_bucket}/{key_to_use}"

    # 出力先: Lambda なら /tmp、ローカルなら bigquery-data/
    out_dir = Path("/tmp" if os.environ.get("AWS_LAMBDA_FUNCTION_NAME") else "bigquery-data/")
    try:
        outputs = _generate_outputs(data, out_dir)
    except Exception as e:
        return {"success": False, "error": f"出力生成でエラー: {e}"}

    # S3 にアップロード（必ず bigquery-data/ 配下に置く）
    uploaded: List[str] = []
    if output_to_s3:
        s3 = _get_boto3_client_or_none("s3")
        if s3 is None:
            return {"success": False, "error": "boto3が利用できません"}
        prefix = "bigquery-data/"
        for p in outputs:
            key = f"{prefix}{p.name}"
            try:
                s3.upload_file(str(p), s3_bucket, key)
                uploaded.append(f"s3://{s3_bucket}/{key}")
            except Exception as e:
                return {"success": False, "error": f"S3アップロード失敗: {e}"}

    return {
        "success": True,
        "input": input_desc,
        "outputs": [str(p) for p in outputs],
        "uploaded": uploaded,
        "count": len(outputs),
    }


def lambda_handler(event, context):
    try:
        s3_bucket = os.getenv("S3_BUCKET_NAME")  # 環境変数からバケット名
        s3_key = event.get("s3_key")             # イベントのキー名から取得
        s3_input_prefix = event.get("s3_input_prefix")
        output_to_s3 = bool(event.get("output_to_s3", True))

        if not s3_bucket:
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "環境変数 S3_BUCKET_NAME が未設定です"}, ensure_ascii=False)}

        result = process_monthly_usage(
            s3_bucket=s3_bucket,
            s3_key=s3_key,
            s3_input_prefix=s3_input_prefix,
            output_to_s3=output_to_s3,
        )

        status = 200 if result.get("success") else 400
        return {"statusCode": status, "body": json.dumps(result, ensure_ascii=False)}

    except Exception as e:
        return {"statusCode": 500, "body": json.dumps({"success": False, "error": str(e)}, ensure_ascii=False)}
