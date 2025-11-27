# 定例会資料作成（グループ店）　定義書

## BigQuery

## 構成

- 以下の３つのプロジェクトからなる
    - epark-refre-users
        - このプロジェクトでサービスアカウントキーを発行している
    - epark-refre-admin
    - epark-refre-rook

### データの抽出場所

- 施設情報
    
    参照先：epark-refre-users.0023635_toshima.V_SHISETHUMAI_6MONTH`
    
    ```python
    query = f"""
          SELECT * 
          FROM `epark-refre-users.0023635_toshima.V_SHISETHUMAI_6MONTH`
          WHERE facility_id IN ({id_sql})
          ORDER BY facility_id
        """
    ```
    
- 予約情報
    
    参照先：epark-refre-rook.rook.T_MR_T_TIME_APPOINT_20250603
    
    ```python
    q_app = f"""
          SELECT
            *
          FROM {tbl_app}
          WHERE SHOPOWNER_NO IN ({nos_sql})
            AND INSIDE_OUTSIDE_FLG = '2'
            AND SITE_CODE IS NOT NULL
            AND CREATE_DATE BETWEEN '{start_date}' AND '{end_date}'
        """
    ```
    
- アクセス数
    
    参照先：epark-refre-users.public.T_BI_T_GA_SESSIONS_PER_HIT_SMR_20250501
    
    ```python
    q_pv = f"""
                SELECT
                  facility_id,
                  COUNT(1) AS pv,
                  "{label}" AS month
                FROM {table}
                WHERE facility_id IN ({id_sql})
                GROUP BY facility_id
            """
    ```
    
    

## 権限設定

- 付与するべきロール
    - Bigqueryデータ閲覧者
    - Bigqueryユーザー
    - Project IAM 管理者
    - Service Usage 管理者
    - サービスアカウントキー　管理者
    - サービスアカウント 管理者
    - リソース設定閲覧者
    - 環境とストレージオブジェクト閲覧者

## AWS
 -IAMロール
    - AmazonS3FullAccess
    - S3フルアクセス
    - AWSLambdaBasicExecutionRole-c49481e1-0a0b-41f8-9103-59ab711e082a
    - AWSLambdaVPCAccessExecutionRole-d8db4320-7d87-46b2-a0a8-eb7d80416a21
    - LambdaS3UploadPolicy
    - service-account

  -環境変数
    - S3_BUCKET_NAME（キー名）
       - S3バケット名
       - monthly-report-output-group（値）
    - SA_SERCRET_NAME（キー名）
       - bigquery
       - gcp/bigquery/service-key（値）
 
  -レイヤー
    - db-types
    - bigquery-layer
    - excel-layer
    - pandas-dependencies2
       - pandasの既存レイヤーが見つからなかったため自作


  


### Step Functions


<img width="153" height="352" alt="image" src="https://github.com/user-attachments/assets/862a5417-62dd-4efd-b30a-a4454e1d0fbb" />

---

###S3構成
**monthly-report-output-group（バゲット名）**

S3 Bucket: monthly-report-output-group
│
├── bigquery-data/
│   
│   └── group_year_data/
│       ├── web_reservation_aoba_20250901.csv
│       └── web_reservation_leaf_20250901.csv
│
└── 2025_09/                    ← 現在の月（9月）
├── あおばグループ/
│   ├── aoba_7月_統合データ.xlsx          ← 10シート（データ+円グラフ）
│   ├── aoba_8月_統合データ.xlsx          ← 10シート（データ+円グラフ）
│   └── aoba_2025_09_月別集計_一年間.xlsx
└── リーフグループ/
├── leaf_7月_統合データ.xlsx          ← 10シート（データ+円グラフ）
├── leaf_8月_統合データ.xlsx          ← 10シート（データ+円グラフ）
└── leaf_2025_09_月別集計_一年間.xlsx

**monthly_report_group_output_json**
group_companies_config.json

### Lambda関数

**monthly_report_group**
-group_local_v2.py
  S3からグループ会社の物件番号が記載されているファイル（JSON）を読み取り、その物件番号をもとにGCPから直近1年分の予約情報を取得。
-monthly_6_usage_group.py
　group_local_v2.pyによって作成されたweb_reservation_*.csvから最新の一か月の予約情報を取得。（現在が10月3日だったら9月の1日から末尾までのデータを取得）。以下の4つのシートに分類される。
 初回受付か再受付か。流入経路はホームページ(hp)経由かそれ以外（その他）か。
 　　-再受付その他
    -初回受付その他
    -再受付hp
    -初回受付hp
-monthly_line_chart_group.py
 monthly_6_usage_group.pyで作成されたcsvファイルをもとに直近一年分の予約情報を線グラフで表示。（現在が2025年10月3日だったら2024年10月から2025年9月まで)
-menu_pie.py
 monthly_6_usage_group.pyで作成したcsvファイルの4つのデータシートごとにメニュー別の円グラフを作成。（現状計8シートになる）
-site_pie.py
 monthly_6_usage_group.pyで作成したcsvファイルの2つのその他のシートにそれぞれ流入経路別の円グラフを作成。（最終的に1月のcsvファイルで10シートになる）
-test_email.py
 作成した2つのcsvファイルをグループ店舗様分zipファイルにして、メールで社員さんに送信（毎月3日）  
 

**monthly_report_group_check**
上記のmonthly_report_groupのtest_email.pyの代わりにtest_email_check.py
-test_email_check.py
 test_emil.pyと同じプログラム。送信先が社員さんではなく管理者宛てになっている（毎月1日）
 動作の確認用である

### EventBridge Scheduler
- 毎月3日にstepfunction(monthly_report_group)を作動させる
- 毎月1日にstepfunction(monthly_report_group_check)を作動させる
- バゲット内部（monthly-report-output-group）は自動で作成から60日後に削除される

### JSONファイル（物件番号を含む）
このファイル（group_companies_config.json）だけ異なるS3のバケットに入っている。（monthly-report-group-json ）
### Secrets Manager

- BigQueryにアクセスするために必要なサービスアカウントの鍵を保管



