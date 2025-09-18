import boto3
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# SES と S3 のクライアント
ses = boto3.client("ses", region_name="ap-northeast-1")
s3 = boto3.client("s3")

def lambda_handler(event, context):
    # ========================
    # 設定値（環境変数にしてもOK）
    # ========================
    SENDER = "IT_Solutions_Department@epark.co.jp"     # 送信元（Verified 済み）
    RECIPIENT = "kenta.murakami@iflag.co.jp"  # 宛先（Verified 済み）
    SUBJECT = "毎月レポート送信（複数ファイル添付）"

    BUCKET_NAME = "monthly-report-output-group"
    OBJECT_KEYS = [
        "bigquery-data/5月_統合データ.xlsx",
        "bigquery-data/月別集計_一年間_20250918061304.xlsx"  # 追加したいファイル
    ]

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
    
    for object_key in OBJECT_KEYS:
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
