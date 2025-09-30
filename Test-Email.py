import boto3
import os
import zipfile
import tempfile
from datetime import datetime
from typing import List, Dict, Optional
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# SES と S3 のクライアント
ses = boto3.client("ses", region_name="ap-northeast-1")
s3 = boto3.client("s3")








def get_current_year_month_folder(bucket: str) -> str:
    """現在の年_月フォルダ名を取得"""
    now = datetime.now()
    year_month = f"{now.year}_{now.month:02d}"
    print(f"🔍 対象フォルダ: {year_month}")
    return year_month

def find_files_in_year_month_folder(bucket: str, year_month_folder: str) -> List[Dict[str, str]]:
    """指定された年_月フォルダ内のすべてのファイルを取得"""
    try:
        paginator = s3.get_paginator("list_objects_v2")
        files = []
        
        for page in paginator.paginate(Bucket=bucket, Prefix=year_month_folder):
            for obj in page.get("Contents", []):
                key = obj["Key"]
                # フォルダ内のファイルのみを対象（フォルダ自体は除外）
                if not key.endswith("/") and key != year_month_folder:
                    files.append({
                        "key": key,
                        "filename": key.split("/")[-1],
                        "size": obj.get("Size", 0),
                        "last_modified": obj.get("LastModified")
                    })
        
        print(f"✅ {year_month_folder}フォルダ内のファイル: {len(files)}件")
        for file_info in files:
            print(f"  - {file_info['filename']} ({file_info['size']} bytes)")
        
        return files
        
    except Exception as e:
        print(f"❌ フォルダ内ファイル取得エラー: {e}")
        return []

def create_zip_from_s3_files(bucket: str, files: List[Dict[str, str]], year_month_folder: str) -> Optional[bytes]:
    """S3ファイルからZIPファイルを作成"""
    try:
        # 一時ファイルでZIPを作成
        with tempfile.NamedTemporaryFile() as temp_zip:
            with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_info in files:
                    try:
                        # S3からファイルをダウンロード
                        s3_obj = s3.get_object(Bucket=bucket, Key=file_info["key"])
                        file_content = s3_obj["Body"].read()
                        
                        # ZIPに追加（フォルダ構造を保持）
                        zip_path = file_info["key"].replace(f"{year_month_folder}/", "")
                        zipf.writestr(zip_path, file_content)
                        print(f"✅ ZIPに追加: {zip_path}")
                        
                    except Exception as e:
                        print(f"❌ ファイル追加エラー {file_info['filename']}: {e}")
                        continue
            
            # ZIPファイルの内容を読み込み
            with open(temp_zip.name, 'rb') as f:
                zip_content = f.read()
            
            print(f"✅ ZIPファイル作成完了: {len(zip_content)} bytes")
            return zip_content
            
    except Exception as e:
        print(f"❌ ZIPファイル作成エラー: {e}")
        return None

def lambda_handler(event, context):
    # ========================
    # 設定値（環境変数にしてもOK）
    # ========================
    SENDER = "IT_Solutions_Department@epark.co.jp"     # 送信元（Verified 済み）
    RECIPIENT = "kenta.murakami@iflag.co.jp"  # 宛先（Verified 済み）
    
    BUCKET_NAME = "monthly-report-output-group"
    
    # 現在の年_月フォルダを取得
    print("🔍 現在の年_月フォルダを取得中...")
    year_month_folder = get_current_year_month_folder(BUCKET_NAME)
    SUBJECT = f"毎月レポート送信（{year_month_folder}フォルダ）"
    
    # 年_月フォルダ内のファイルを取得
    print(f"🔍 {year_month_folder}フォルダ内のファイルを検索中...")
    files = find_files_in_year_month_folder(BUCKET_NAME, year_month_folder)
    
    if not files:
        print(f"❌ {year_month_folder}フォルダ内にファイルが見つかりませんでした")
        return {"statusCode": 404, "body": f"{year_month_folder}フォルダ内にファイルが見つかりませんでした"}
    
    print(f"✅ {len(files)}個のファイルが見つかりました")

    # ========================
    # ZIPファイル作成
    # ========================
    print("📦 ZIPファイルを作成中...")
    zip_content = create_zip_from_s3_files(BUCKET_NAME, files, year_month_folder)
    
    if not zip_content:
        print("❌ ZIPファイルの作成に失敗しました")
        return {"statusCode": 500, "body": "ZIPファイルの作成に失敗しました"}

    # ========================
    # メール構築
    # ========================
    msg = MIMEMultipart()
    msg["Subject"] = SUBJECT
    msg["From"] = SENDER
    msg["To"] = RECIPIENT

    # ========================
    # ZIPファイルを添付
    # ========================
    zip_filename = f"{year_month_folder}_レポート.zip"
    
    try:
        # ZIPファイルを添付
        attachment = MIMEApplication(zip_content)
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename=zip_filename
        )
        msg.attach(attachment)
        print(f"✅ ZIPファイル添付完了: {zip_filename}")
        
    except Exception as e:
        print(f"❌ ZIPファイル添付エラー: {e}")
        return {"statusCode": 500, "body": f"ZIPファイル添付エラー: {e}"}

    # ========================
    # 本文を動的に生成
    # ========================
    body_text = f"""S3に保存された{year_month_folder}フォルダのレポートをお送りします。

添付ファイル：
- {zip_filename} ({len(zip_content):,} bytes)

ZIPファイル内のファイル（{len(files)}件）：
"""
    for i, file_info in enumerate(files, 1):
        body_text += f"{i}. {file_info['filename']} ({file_info['size']:,} bytes)\n"
    
    body_text += "\nご確認ください。"
    
    # 本文を添付
    msg.attach(MIMEText(body_text, "plain"))

    # ========================
    # SES送信
    # ========================
    try:
        response = ses.send_raw_email(
            Source=SENDER,
            Destinations=[RECIPIENT],
            RawMessage={"Data": msg.as_string()}
        )

        print("✅ Email sent! Message ID:", response["MessageId"])
        return {
            "statusCode": 200, 
            "body": f"Email sent successfully. ZIP file: {zip_filename} ({len(zip_content):,} bytes)"
        }
        
    except Exception as e:
        print(f"❌ メール送信エラー: {e}")
        return {"statusCode": 500, "body": f"メール送信エラー: {e}"}
