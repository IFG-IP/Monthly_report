import os
from google.cloud import bigquery
from google.cloud import bigquery_storage
import pandas as pd
from datetime import datetime, timedelta

# ✅ ワーニング対策（quota project を指定）
os.environ["GOOGLE_CLOUD_PROJECT"] = "account-key.json"

client = bigquery.Client(project="epark-refre-users")
bqstorage_client = bigquery_storage.BigQueryReadClient()

# 特定の物件番号を指定
property_numbers = ["E10331191000001", "E10331193000001","E10331194000001","E10331197000001","E10331198000001","E10331200000001","E10331201000001","E10337069000001","E10337070000001"]  # ここに物件番号を設定してください
property_condition = ",".join([f"'{x}'" for x in property_numbers])

# 日付範囲を動的に計算（先月から過去1年）
current_date = datetime.now()
# 先月の1日
last_month_start = (current_date.replace(day=1) - timedelta(days=1)).replace(day=1)
# 先月の最終日（月末）
last_month_end = current_date.replace(day=1) - timedelta(days=1)
# 過去1年前
one_year_ago = last_month_start - timedelta(days=365)

# 日付を文字列形式に変換
start_date = one_year_ago.strftime("%Y-%m-%d")
end_date = last_month_end.strftime("%Y-%m-%d")

# テーブル名用の日付（現在の日付）
table_date = current_date.strftime("%Y%m%d")

print(f"データ取得期間: {start_date} から {end_date} まで")
print(f"使用テーブル日付: {table_date}")

# BigQueryクエリを定義
query = f"""
WITH w_web AS (
    SELECT 
        concat(s.shopowner_id,"_",r.time_appoint_id) as petano_log_no,
        f.facility_id, --施設ID
        f.property_number, --物件番号
        r.shopowner_no, --契約者番号
        s.shopowner_id, --契約者ID
        CASE
            WHEN r.SYSTEM_USAGE_KIND = "1" THEN "初回受付"
            WHEN r.SYSTEM_USAGE_KIND = "2" THEN "再受付"
            WHEN r.SYSTEM_USAGE_KIND = "3" THEN "初回キャンセル"
        END AS SYSTEM_USAGE_NAME,
        m.MENU_NAME, --メニュー名
        m.PRICE, --メニュー料金(税抜)
        m.TIME_REQUIRED, --所要時間
        SITE_CODE, --サイトコード
        -- APPOINT_START_DATETIMEを適切な日時形式に変換
        FORMAT_TIMESTAMP('%Y/%m/%d %H:%M', 
            PARSE_TIMESTAMP('%Y%m%d%H%M', r.APPOINT_START_DATETIME)
        ) AS modify_start_date,
        -- 元のTIMESTAMP型も保持（日時比較用）
        PARSE_TIMESTAMP('%Y%m%d%H%M', r.APPOINT_START_DATETIME) AS appoint_timestamp,
        -- タイムゾーン情報なしの読みやすい形式
        FORMAT_DATETIME('%Y-%m-%d %H:%M:%S', 
            DATETIME(PARSE_TIMESTAMP('%Y%m%d%H%M', r.APPOINT_START_DATETIME))
        ) AS appoint_datetime_clean,
        substr(r.create_date,0,10) AS create_date --予約登録日
    FROM `epark-refre-rook.rook.T_MR_T_TIME_APPOINT_{table_date}` r
    LEFT JOIN `epark-refre-rook.common.T_AD_M_SHOPOWNER_{table_date}` s 
        ON s.shopowner_no = r.shopowner_no
    LEFT JOIN `epark-refre-rook.rook.T_MR_M_TIME_APPOINT_MENU_{table_date}` m 
        ON m.shopowner_no = r.shopowner_no AND m.MENU_ID = r.MENU_ID
    LEFT JOIN `epark-refre-rook.rook.T_MR_M_TIME_APPOINT_STAFF_{table_date}` ss 
        ON ss.STAFF_ID = r.STAFF_ID AND ss.SHOPOWNER_NO = r.SHOPOWNER_NO
    LEFT JOIN `epark-refre-users.public.T_BI_M_FACILITY_{table_date}` f
        ON f.RESERVE_SHOPOWNER_ID = s.SHOPOWNER_ID
    WHERE r.inside_outside_flg = "2" --WEB予約のみ
        AND r.SYSTEM_USAGE_KIND != "0" --WEB予約のみ※念の為
        AND f.status = "1"
        AND f.property_number IN ({property_condition}) --物件番号で絞り込み
)

SELECT 
    w_web.*
FROM w_web
WHERE appoint_timestamp BETWEEN TIMESTAMP("{start_date} 00:00:00") AND TIMESTAMP("{end_date} 23:59:59") --先月から過去1年の日時範囲で絞り込み
ORDER BY appoint_timestamp DESC
"""

# クエリを実行
try:
    df = client.query(query).to_dataframe(bqstorage_client=bqstorage_client)
    
    # 結果を保存
    base_date = datetime.today().replace(day=1)
    output_dir = "C:/Users/8961083/.cursor/workspace_petlife/results"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"web_reservation_leaf_{base_date.strftime('%Y%m%d')}.csv")
    
    df.to_csv(output_path, index=False, encoding="utf-8-sig", na_rep="")
    
    print(f"クエリが正常に実行されました。結果は {output_path} に保存されました。")
    print(f"対象物件番号: {property_numbers}")
    print(f"取得件数: {len(df)}件")
    print("\nデータの最初の5行:")
    print(df.head())
    
except Exception as e:
    print(f"エラーが発生しました: {e}")

