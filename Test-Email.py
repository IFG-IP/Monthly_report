import boto3
import os
import csv
import io
import re
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# SES と S3 のクライアント
ses = boto3.client("ses", region_name="ap-northeast-1")
s3 = boto3.client("s3")


def parse_datetime(date_str: str) -> Optional[datetime]:
    """日時文字列を datetime に変換（ISO8601/TZ対応）"""
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


def load_csv_from_s3(bucket: str, key: str) -> Optional[List[Dict]]:
    """S3 から CSV を読み込んで辞書リストとして返す"""
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


def get_target_months_from_data(data: List[Dict]) -> List[Tuple[int, int]]:
    """データから直近1か月の対象月を動的に取得する関数"""
    if not data:
        return []
    
    try:
        from zoneinfo import ZoneInfo
        JST = ZoneInfo("Asia/Tokyo")
    except:
        from datetime import timezone, timedelta
        JST = timezone(timedelta(hours=9))
    
    # 現在時刻をJSTで取得
    now = datetime.now(JST)
    
    # 直近1か月（現在の月）の開始日を計算
    current_year = now.year
    current_month = now.month
    
    # 現在の月の1日を開始日とする
    current_month_start = datetime(current_year, current_month, 1, tzinfo=JST)
    
    print(f"🔍 対象期間: {current_month_start.strftime('%Y年%m月')} 以降")
    
    # データから実際に存在する年月を収集
    months_found = set()
    for row in data:
        timestamp_str = row.get('appoint_timestamp', '')
        if not timestamp_str:
            continue
            
        dt = parse_datetime(timestamp_str)
        if dt is None:
            continue
            
        # JSTに変換
        if dt.tzinfo:
            dt = dt.astimezone(JST)
        else:
            dt = dt.replace(tzinfo=JST)
        
        # 現在の月以降のデータのみを対象
        if dt >= current_month_start:
            months_found.add((dt.year, dt.month))
    
    # 年月でソートして返す
    return sorted(list(months_found))


def find_latest_files_for_email(bucket: str, prefix: str) -> List[str]:
    """メール添付用のファイルを検索する関数（直近1か月の統合データファイル + 月別集計ファイル）"""
    try:
        # 最新のCSVファイルを取得
        paginator = s3.get_paginator("list_objects_v2")
        latest_csv = None
        for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
            for obj in page.get("Contents", []):
                key = obj.get("Key", "")
                if key.endswith(".csv") and "web_reservation_leaf_" in key:
                    if latest_csv is None or obj["LastModified"] > latest_csv["LastModified"]:
                        latest_csv = obj
        
        if latest_csv is None:
            print("⚠️ CSVファイルが見つかりませんでした")
            return []
        
        # CSVファイルを読み込んで対象月を取得
        print(f"🔍 CSVファイルから対象月を取得中: {latest_csv['Key']}")
        csv_data = load_csv_from_s3(bucket, latest_csv['Key'])
        if not csv_data:
            print("⚠️ CSVファイルの読み込みに失敗しました")
            return []
        
        # 直近1か月の対象月を取得
        target_months = get_target_months_from_data(csv_data)
        if not target_months:
            print("⚠️ 直近1か月のデータが見つかりませんでした")
            return []
        
        print(f"🔍 対象月: {target_months}")
        
    except Exception as e:
        print(f"⚠️ CSVファイルからの対象月取得に失敗: {e}")
        # フォールバック: 固定で現在の月を使用
        now = datetime.now()
        target_months = [(now.year, now.month)]
        print(f"🔍 フォールバック: 現在月を使用 {target_months}")
    
    # 統合データファイルを検索
    pattern = re.compile(r"(\d+)月_統合データ\.xlsx$")
    paginator = s3.get_paginator("list_objects_v2")
    by_month: Dict[Tuple[int, int], List[str]] = {}
    
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            name = key.split("/")[-1]
            m = pattern.search(name)
            if m:
                month = int(m.group(1))
                # 対象月に含まれるかチェック
                for year, target_month in target_months:
                    if month == target_month:
                        by_month.setdefault((year, month), []).append(key)
                        break
    
    if not by_month:
        print("⚠️ 直近1か月の統合データファイルが見つかりませんでした")
        return []
    
    # 直近1か月の統合データファイルを取得
    result: List[str] = []
    for year, month in target_months:
        if (year, month) in by_month:
            files = by_month[(year, month)]
            result.extend(files)
            print(f"✅ {year}年{month}月の統合データファイル: {len(files)}件")
        else:
            print(f"⚠️ {year}年{month}月の統合データファイルが見つかりませんでした")
    
    # 月別集計ファイルを検索（最新の1つを取得）
    monthly_summary_pattern = re.compile(r"月別集計_一年間_\d{6}\.xlsx$")
    latest_monthly_summary = None
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            name = key.split("/")[-1]
            if monthly_summary_pattern.search(name):
                if latest_monthly_summary is None or obj["LastModified"] > latest_monthly_summary["LastModified"]:
                    latest_monthly_summary = obj
    
    if latest_monthly_summary:
        result.append(latest_monthly_summary["Key"])
        print(f"✅ 月別集計ファイル発見: {latest_monthly_summary['Key'].split('/')[-1]}")
    else:
        print("⚠️ 月別集計ファイルが見つかりませんでした")
    
    return result

def lambda_handler(event, context):
    # ========================
    # 設定値（環境変数にしてもOK）
    # ========================
    SENDER = "IT_Solutions_Department@epark.co.jp"     # 送信元（Verified 済み）
    RECIPIENT = "kenta.murakami@iflag.co.jp"  # 宛先（Verified 済み）
    SUBJECT = "毎月レポート送信（直近1か月分 + 月別集計）"

    BUCKET_NAME = "monthly-report-output-group"
    PREFIX = "bigquery-data/"
    
    # メール添付用のファイルを動的に取得
    print("🔍 メール添付用ファイルを検索中...")
    object_keys = find_latest_files_for_email(BUCKET_NAME, PREFIX)
    
    if not object_keys:
        print("❌ 対象ファイルが見つかりませんでした")
        return {"statusCode": 404, "body": "対象ファイルが見つかりませんでした"}
    
    print(f"✅ {len(object_keys)}個のファイルが見つかりました: {object_keys}")

    # ========================
    # メール構築
    # ========================
    msg = MIMEMultipart()
    msg["Subject"] = SUBJECT
    msg["From"] = SENDER
    msg["To"] = RECIPIENT

    # ========================
    # 複数ファイルを添付
    # ========================
    attached_files = []  # 添付されたファイルのリスト
    
    for object_key in object_keys:
        try:
            # S3からファイル取得
            s3_obj = s3.get_object(Bucket=BUCKET_NAME, Key=object_key)
            file_content = s3_obj["Body"].read()
            
            # 添付ファイル
            attachment = MIMEApplication(file_content)
            attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(object_key)
            )
            msg.attach(attachment)
            attached_files.append(os.path.basename(object_key))
            print(f"添付ファイル追加: {os.path.basename(object_key)}")
            
        except Exception as e:
            print(f"ファイル取得エラー: {object_key} - {str(e)}")
            continue

    # ========================
    # 本文を動的に生成
    # ========================
    if attached_files:
        body_text = f"""S3に保存された最新レポートをお送りします。

添付ファイル（{len(attached_files)}件）：
"""
        for i, filename in enumerate(attached_files, 1):
            body_text += f"{i}. {filename}\n"
        
        body_text += "\nご確認ください。"
    else:
        body_text = "申し訳ございませんが、添付ファイルの取得に失敗しました。"
    
    # 本文を添付
    msg.attach(MIMEText(body_text, "plain"))

    # ========================
    # SES送信
    # ========================
    response = ses.send_raw_email(
        Source=SENDER,
        Destinations=[RECIPIENT],
        RawMessage={"Data": msg.as_string()}
    )

    print("Email sent! Message ID:", response["MessageId"])
    return {"statusCode": 200, "body": "Email sent successfully"}
