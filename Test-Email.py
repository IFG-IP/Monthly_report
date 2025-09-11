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
    SUBJECT = "毎月レポート送信"
    BODY_TEXT = "S3に保存された最新レポートをお送りします。"

    BUCKET_NAME = "monthly-report-output-group"
    OBJECT_KEY = "bigquery-data/8月_初回受付_その他.xlsx"

    # ========================
    # S3からファイル取得
    # ========================
    s3_obj = s3.get_object(Bucket=BUCKET_NAME, Key=OBJECT_KEY)
    file_content = s3_obj["Body"].read()

    # ========================
    # メール構築
    # ========================
    msg = MIMEMultipart()
    msg["Subject"] = SUBJECT
    msg["From"] = SENDER
    msg["To"] = RECIPIENT

    # 本文
    msg.attach(MIMEText(BODY_TEXT, "plain"))

    # 添付ファイル
    attachment = MIMEApplication(file_content)
    attachment.add_header(
        "Content-Disposition",
        "attachment",
        filename=os.path.basename(OBJECT_KEY)
    )
    msg.attach(attachment)

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
