from pathlib import Path
from datetime import datetime, timedelta
import sys
import boto3
from botocore.exceptions import ClientError
import os
import json
import csv
from collections import defaultdict

# Python 3.13互換性のための修正
if not hasattr(os, 'add_dll_directory'):
    def add_dll_directory(path):
        pass
    os.add_dll_directory = add_dll_directory

def get_s3_client():
    """S3クライアントを取得"""
    try:
        return boto3.client('s3')
    except Exception as e:
        print(f"S3クライアントの作成に失敗: {e}")
        return None

def download_from_s3(s3_path: str) -> str:
    """S3からファイルをダウンロード"""
    try:
        print(f"download_from_s3開始: '{s3_path}'")
        
        # スペースを除去
        s3_path = s3_path.replace(' ', '')
        
        # S3パスを解析 (s3://bucket/key)
        if not s3_path.startswith('s3://'):
            print(f"S3パスではありません: {s3_path}")
            return None
        
        path_parts = s3_path[5:].split('/', 1)  # s3://を除去して分割
        if len(path_parts) != 2:
            print(f"パス分割に失敗: {path_parts}")
            return None
        
        bucket_name, s3_key = path_parts
        bucket_name = bucket_name.strip()
        s3_key = s3_key.lstrip('/')
        
        # ローカルファイル名を生成
        filename = os.path.basename(s3_key)
        local_path = os.path.join('/tmp', filename)
        
        # S3クライアントを取得
        s3_client = get_s3_client()
        if not s3_client:
            return None
        
        # ファイルの存在確認
        try:
            s3_client.head_object(Bucket=bucket_name, Key=s3_key)
            print(f"✅ S3ファイル存在確認: {s3_path}")
        except ClientError as e:
            error_code = e.response['Error']['Code']
            if error_code == '404':
                print(f"❌ S3ファイルが見つかりません: {s3_path}")
                return None
            else:
                print(f"❌ S3ファイル確認エラー: {e}")
                return None
        
        # ファイルをダウンロード
        s3_client.download_file(bucket_name, s3_key, local_path)
        print(f"S3からダウンロード完了: {s3_path} -> {local_path}")
        
        return local_path
        
    except Exception as e:
        print(f"S3ダウンロードエラー: {e}")
        return None

def upload_to_s3(file_path: str, bucket_name: str, s3_key: str) -> dict:
    """ファイルをS3にアップロード"""
    try:
        print(f"🔍 デバッグ: S3アップロード開始")
        print(f"🔍 デバッグ: ファイルパス: '{file_path}'")
        print(f"🔍 デバッグ: バケット名: '{bucket_name}'")
        print(f"🔍 デバッグ: S3キー: '{s3_key}'")
        
        # バケット名の正規化（スペースを除去）
        bucket_name = bucket_name.strip() if bucket_name else None
        s3_key = s3_key.strip()
        
        # ファイルの存在確認
        if not os.path.exists(file_path):
            print(f"❌ ファイルが存在しません: {file_path}")
            return {"success": False, "error": f"ファイルが存在しません: {file_path}"}
        
        s3_client = get_s3_client()
        if not s3_client:
            print("❌ S3クライアントの作成に失敗")
            return {"success": False, "error": "S3クライアントの作成に失敗"}
        
        # ファイルアップロード
        s3_client.upload_file(file_path, bucket_name, s3_key)
        print(f"✅ S3アップロード成功")
        
        # S3パスを生成
        s3_path = f"s3://{bucket_name}/{s3_key}"
        
        return {
            "success": True,
            "s3_key": s3_key,
            "s3_path": s3_path,
            "bucket": bucket_name
        }
        
    except Exception as e:
        print(f"❌ S3アップロード失敗: {str(e)}")
        return {"success": False, "error": f"S3アップロードに失敗: {str(e)}"}

def load_csv_data(csv_path: Path):
    """CSVファイルを読み込む（標準ライブラリ使用）"""
    try:
        data = []
        
        # ファイルの詳細を確認
        print(f"🔍 ファイル詳細確認: {csv_path}")
        if not os.path.exists(csv_path):
            print(f"❌ ファイルが存在しません: {csv_path}")
            return None
        
        file_size = os.path.getsize(csv_path)
        print(f"📁 ファイルサイズ: {file_size} bytes")
        
        # 複数のエンコーディングを試す
        encodings = [
            'utf-8-sig', 'utf-8', 'utf-16', 'utf-32',
            'shift_jis', 'cp932', 'euc-jp', 'iso-2022-jp',
            'latin1', 'cp1252', 'ascii', 'big5', 'gb2312'
        ]
        
        for encoding in encodings:
            try:
                print(f"🔍 エンコーディング試行: {encoding}")
                with open(csv_path, 'r', encoding=encoding) as file:
                    reader = csv.DictReader(file)
                    for row in reader:
                        data.append(row)
                
                print(f"✅ エンコーディング成功: {encoding}")
                print(f"📊 CSVファイル読み込み完了: {len(data)}行")
                return data
                
            except UnicodeDecodeError as e:
                print(f"❌ エンコーディング失敗: {encoding} - {str(e)[:100]}")
                data = []  # リセット
                continue
            except csv.Error as e:
                print(f"❌ CSV解析エラー: {encoding} - {str(e)[:100]}")
                data = []  # リセット
                continue
            except Exception as e:
                print(f"❌ その他のエラー: {encoding} - {str(e)[:100]}")
                data = []  # リセット
                continue
        
        # すべてのエンコーディングが失敗した場合
        print(f"❌ すべてのエンコーディングで読み込み失敗")
        return None
        
    except Exception as e:
        print(f"❌ CSVファイル読み込みエラー: {e}")
        return None

def parse_date_columns(data):
    """日付列を解析してdatetime型に変換"""
    try:
        for row in data:
            # appoint_datetime_clean列を日付型に変換
            date_str = row.get('appoint_datetime_clean', '')
            try:
                if date_str:
                    # 日付文字列を解析（複数の形式に対応）
                    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d']:
                        try:
                            row['appoint_date'] = datetime.strptime(date_str, fmt)
                            break
                        except ValueError:
                            continue
                    else:
                        row['appoint_date'] = None
                else:
                    row['appoint_date'] = None
            except Exception:
                row['appoint_date'] = None
        
        # 日付が解析できない行を除外
        before_count = len(data)
        data = [row for row in data if row.get('appoint_date') is not None]
        after_count = len(data)
        
        print(f"📅 日付解析完了: {before_count}行 -> {after_count}行")
        return data
        
    except Exception as e:
        print(f"❌ 日付解析エラー: {e}")
        return None

def create_monthly_summary(data, site_type: str = "all"):
    """月別の初回受付・再受付の集計を作成（HP/その他別）"""
    try:
        # 年月でグループ化
        monthly_data = defaultdict(lambda: {'初回受付': 0, '再受付': 0, '合計': 0})
        
        for row in data:
            # サイトタイプでフィルタリング
            site_code = row.get('SITE_CODE', '').strip().lower()
            
            if site_type == "hp" and site_code != 'hp':
                continue
            elif site_type == "other" and site_code == 'hp':
                continue
            
            # 年月を取得
            appoint_date = row.get('appoint_date')
            if not appoint_date:
                continue
                
            year_month = appoint_date.strftime('%Y-%m')
            
            # 初回受付・再受付を判定（例：appoint_type列がある場合）
            appoint_type = row.get('appoint_type', '').strip()
            
            if '初回' in appoint_type or 'first' in appoint_type.lower():
                monthly_data[year_month]['初回受付'] += 1
            elif '再' in appoint_type or 'repeat' in appoint_type.lower():
                monthly_data[year_month]['再受付'] += 1
            else:
                # デフォルトで初回受付として扱う
                monthly_data[year_month]['初回受付'] += 1
            
            monthly_data[year_month]['合計'] += 1
        
        # 結果をリストに変換してソート
        result = []
        for year_month in sorted(monthly_data.keys()):
            data_point = monthly_data[year_month]
            data_point['月'] = year_month
            result.append(data_point)
        
        print(f"📊 {site_type}月別集計完了: {len(result)}ヶ月")
        return result
        
    except Exception as e:
        print(f"❌ {site_type}月別集計エラー: {e}")
        return None


def save_to_csv(hp_data, other_data, output_path: Path) -> bool:
    """CSVファイルに結果を保存"""
    try:
        print(f"🔍 デバッグ: CSVファイル保存開始: {output_path}")
        
        # HPデータをCSVに保存
        hp_csv_path = output_path.with_suffix('_HP.csv')
        with open(hp_csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            if hp_data:
                fieldnames = hp_data[0].keys()
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(hp_data)
        
        # その他データをCSVに保存
        other_csv_path = output_path.with_suffix('_その他.csv')
        with open(other_csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            if other_data:
                fieldnames = other_data[0].keys()
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(other_data)
        
        print(f"✅ CSVファイル保存完了: {hp_csv_path}, {other_csv_path}")
        return True
        
    except Exception as e:
        print(f"❌ CSVファイル保存エラー: {e}")
        return False

def save_to_excel_with_chart(hp_data, other_data, output_path: Path) -> bool:
    """Excelファイルに結果とグラフを保存（openpyxl使用）"""
    try:
        print(f"🔍 デバッグ: openpyxlのインポートを試行中...")
        print(f"🔍 デバッグ: Pythonパス: {sys.path}")
        try:
            from openpyxl import Workbook
            from openpyxl.chart import LineChart, Reference
            from openpyxl.chart.axis import DateAxis
            print(f"✅ デバッグ: openpyxlのインポート成功")
        except ImportError as e:
            print(f"❌ デバッグ: openpyxlのインポート失敗: {e}")
            raise
        
        # ワークブックを作成
        print(f"🔍 デバッグ: ワークブック作成開始")
        wb = Workbook()
        print(f"✅ デバッグ: ワークブック作成完了")
        
        # HP用のシート
        print(f"🔍 デバッグ: HPシート作成開始")
        ws_hp = wb.active
        ws_hp.title = 'HP経由'
        print(f"✅ デバッグ: HPシート作成完了")
        
        # HPデータのヘッダーを設定
        headers = ['月', '初回受付', '再受付', '合計']
        for col, header in enumerate(headers, 1):
            ws_hp.cell(row=1, column=col, value=header)
        
        # HPデータを書き込み
        for row, data_point in enumerate(hp_data, 2):
            ws_hp.cell(row=row, column=1, value=data_point['月'])
            ws_hp.cell(row=row, column=2, value=data_point['初回受付'])
            ws_hp.cell(row=row, column=3, value=data_point['再受付'])
            ws_hp.cell(row=row, column=4, value=data_point['合計'])
        
        # HP用の折れ線グラフを作成
        print(f"🔍 デバッグ: HPグラフ作成開始")
        chart_hp = LineChart()
        chart_hp.title = "月別予約受付数推移 (HP経由)"
        chart_hp.style = 13
        chart_hp.y_axis.title = '件数'
        chart_hp.x_axis.title = '月'
        print(f"✅ デバッグ: HPグラフオブジェクト作成完了")
        
        # HPデータ範囲を指定
        data_hp = Reference(ws_hp, min_col=2, min_row=1, max_row=len(hp_data) + 1, max_col=3)
        chart_hp.add_data(data_hp, titles_from_data=True)
        
        # HPカテゴリ軸を設定
        cats_hp = Reference(ws_hp, min_col=1, min_row=2, max_row=len(hp_data) + 1)
        chart_hp.set_categories(cats_hp)
        
        # HPグラフをワークシートに追加
        ws_hp.add_chart(chart_hp, "F2")
        
        # その他用のシート
        ws_other = wb.create_sheet("その他経由")
        
        # その他データのヘッダーを設定
        for col, header in enumerate(headers, 1):
            ws_other.cell(row=1, column=col, value=header)
        
        # その他データを書き込み
        for row, data_point in enumerate(other_data, 2):
            ws_other.cell(row=row, column=1, value=data_point['月'])
            ws_other.cell(row=row, column=2, value=data_point['初回受付'])
            ws_other.cell(row=row, column=3, value=data_point['再受付'])
            ws_other.cell(row=row, column=4, value=data_point['合計'])
        
        # その他用の折れ線グラフを作成
        chart_other = LineChart()
        chart_other.title = "月別予約受付数推移 (その他経由)"
        chart_other.style = 13
        chart_other.y_axis.title = '件数'
        chart_other.x_axis.title = '月'
        
        # その他データ範囲を指定
        data_other = Reference(ws_other, min_col=2, min_row=1, max_row=len(other_data) + 1, max_col=3)
        chart_other.add_data(data_other, titles_from_data=True)
        
        # その他カテゴリ軸を設定
        cats_other = Reference(ws_other, min_col=1, min_row=2, max_row=len(other_data) + 1)
        chart_other.set_categories(cats_other)
        
        # その他グラフをワークシートに追加
        ws_other.add_chart(chart_other, "F2")
        
        # ファイルを保存
        wb.save(output_path)
        print(f"📊 Excelファイル保存完了: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Excelファイル保存エラー: {e}")
        return False

def process_csv_file(csv_path: str, upload_to_s3_flag: bool = False, s3_bucket: str = None) -> dict:
    """CSVファイルを処理して月別グラフを作成"""
    try:
        print(f"🎯 CSVファイル処理開始: {csv_path}")
        
        # CSVファイルを読み込み
        data = load_csv_data(Path(csv_path))
        if data is None:
            return {"success": False, "error": "CSVファイルの読み込みに失敗"}
        
        # 日付解析
        data = parse_date_columns(data)
        if data is None:
            return {"success": False, "error": "日付解析に失敗"}
        
        # HPとその他で月別集計
        hp_data = create_monthly_summary(data, "hp")
        other_data = create_monthly_summary(data, "other")
        
        if hp_data is None or other_data is None:
            return {"success": False, "error": "月別集計に失敗"}
        
        # 出力ファイル名を生成（一つにまとめる）
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        base_filename = f"月別集計_一年間_{timestamp}"
        
        output_path = Path("/tmp") / base_filename
        
        # Excelファイルにグラフを作成（HPとその他を一つのファイルに）
        try:
            excel_success = save_to_excel_with_chart(hp_data, other_data, output_path.with_suffix('.xlsx'))
            
            if not excel_success:
                return {"success": False, "error": "Excelファイルの作成に失敗"}
        except Exception as e:
            print(f"❌ Excelファイル作成エラー: {str(e)}")
            return {"success": False, "error": f"Excelファイルの作成に失敗: {str(e)}"}
        
        # S3にアップロード（フラグがTrueの場合）
        s3_results = []
        if upload_to_s3_flag and s3_bucket:
            upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
            
            # Excelファイルをアップロード
            excel_filename = f"{base_filename}.xlsx"
            excel_s3_key = f"{upload_prefix}{excel_filename}"
            excel_result = upload_to_s3(str(output_path.with_suffix('.xlsx')), s3_bucket, excel_s3_key)
            s3_results.append(excel_result)
        
        # 成功したアップロードのみを結果に含める
        successful_uploads = [result for result in s3_results if result.get('success', False)]
        
        return {
            "success": True,
            "message": f"処理完了: {base_filename}",
            "hp_months_count": len(hp_data),
            "other_months_count": len(other_data),
            "excel_created": excel_success,
            "s3_uploads": successful_uploads,
            "local_path": str(output_path)
        }
        
    except Exception as e:
        print(f"❌ CSVファイル処理エラー: {e}")
        return {"success": False, "error": f"処理失敗: {str(e)}"}

def lambda_handler(event, context):
    """AWS Lambda関数のエントリーポイント"""
    try:
        print("Lambda関数が呼び出されました。")
        print(f"イベント: {json.dumps(event, ensure_ascii=False)}")
        
        # 環境変数から設定を取得（デフォルト値を設定）
        s3_bucket = os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
        upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
        
        # デバッグ用: 環境変数が設定されていない場合はハードコード
        if s3_bucket == 'S3_BUCKET_NAME' or s3_bucket == 'NOT_SET':
            s3_bucket = 'monthly-report-output-group'
            print("⚠️ デバッグ: 環境変数S3_BUCKET_NAMEが設定されていません。デフォルト値を使用します。")
        
        if upload_prefix == 'S3_UPLOAD_PREFIX' or upload_prefix == 'NOT_SET':
            upload_prefix = 'bigquery-data/'
            print("⚠️ デバッグ: 環境変数S3_UPLOAD_PREFIXが設定されていません。デフォルト値を使用します。")
        
        print(f"🔍 デバッグ: 環境変数S3_BUCKET_NAME: '{os.environ.get('S3_BUCKET_NAME', 'NOT_SET')}'")
        print(f"🔍 デバッグ: 環境変数S3_UPLOAD_PREFIX: '{os.environ.get('S3_UPLOAD_PREFIX', 'NOT_SET')}'")
        print(f"🔍 デバッグ: すべての環境変数: {dict(os.environ)}")
        
        print(f"🔍 デバッグ: 使用するS3バケット: '{s3_bucket}'")
        print(f"🔍 デバッグ: アップロードプレフィックス: '{upload_prefix}'")
        
        print(f"🔍 デバッグ: 環境変数 - バケット: '{s3_bucket}', プレフィックス: '{upload_prefix}'")
        
        if not s3_bucket:
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "s3_bucketが未指定"}, ensure_ascii=False)}
        
        # Step Functions用の入力形式に対応（monthly_6_usage形式）
        # 引数: {"s3_bucket": "...", "s3_key": "...", "output_to_s3": true, ...}
        print(f"🔍 デバッグ: 受信したイベント全体: {json.dumps(event, ensure_ascii=False, indent=2)}")
        print(f"🔍 デバッグ: イベントの型: {type(event)}")
        print(f"🔍 デバッグ: イベントのキー: {list(event.keys()) if isinstance(event, dict) else 'Not a dict'}")
        
        # monthly_6_usage形式の引数からS3情報を取得
        input_s3_bucket = event.get('s3_bucket', 'monthly-report-output-group')
        input_s3_key = event.get('s3_key', 'bigquery-data/web_reservation_leaf_20250901.csv')
        
        print(f"🔍 デバッグ: 入力S3バケット: {input_s3_bucket}")
        print(f"🔍 デバッグ: 入力S3キー: {input_s3_key}")
        
        if not input_s3_bucket or not input_s3_key:
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "s3_bucketまたはs3_keyが指定されていません"}, ensure_ascii=False)}
        
        # S3パスを構築
        s3_path = f"s3://{input_s3_bucket}/{input_s3_key}"
        result_data = {"bucket": input_s3_bucket, "key": input_s3_key}
        
        print(f"🔍 デバッグ: 構築されたS3パス: {s3_path}")
        print(f"🔍 デバッグ: result_data: {result_data}")
        
        processed_files = []
        uploaded_files = []
        errors = []
        
        # 1つのファイルを処理
        try:
            print(f"🔍 デバッグ: ファイル処理開始")
            
            # S3パスを構築
            bucket = result_data.get('bucket', '')
            key = result_data.get('key', '')
            if not bucket or not key:
                return {"statusCode": 400, "body": json.dumps({"success": False, "error": "bucketまたはkeyが指定されていません"}, ensure_ascii=False)}
            
            s3_path = f"s3://{bucket}/{key}"
            
            # 入力ファイル名をログ出力
            print(f"📁 入力ファイルS3パス: {s3_path}")
            if s3_path:
                # S3パスからファイル名を抽出
                filename = s3_path.split('/')[-1] if '/' in s3_path else s3_path
                print(f"📄 入力ファイル名: {filename}")
            else:
                print("⚠️ ファイルパスが空です")
            
            # S3からファイルをダウンロード
            print(f"⬇️ S3からファイルをダウンロード開始: {s3_path}")
            print(f"🔍 デバッグ: バケット名: {input_s3_bucket}")
            print(f"🔍 デバッグ: オブジェクトキー: {input_s3_key}")
            local_path = download_from_s3(s3_path)
            if not local_path:
                error_msg = f"S3ダウンロード失敗: {s3_path}"
                print(f"❌ {error_msg}")
                return {"statusCode": 400, "body": json.dumps({"success": False, "error": error_msg}, ensure_ascii=False)}
            
            # ダウンロード完了ログ
            downloaded_filename = local_path.split('/')[-1] if '/' in local_path else local_path
            print(f"✅ S3ダウンロード完了: {downloaded_filename}")
            
            # 一つのファイルでHPとその他を処理
            print(f"🎯 データ処理開始（HPとその他を一つのファイルに）")
            
            result = process_csv_file(
                csv_path=local_path,
                upload_to_s3_flag=True,
                s3_bucket=s3_bucket
            )
            
            if result.get('success'):
                processed_files.append(result)
                uploaded_files.extend(result.get('s3_uploads', []))
                print(f"✅ 処理完了: {result.get('message')}")
            else:
                error_msg = f"処理失敗: {result.get('error')}"
                print(f"❌ {error_msg}")
                return {"statusCode": 500, "body": json.dumps({"success": False, "error": error_msg}, ensure_ascii=False)}
            
        except Exception as e:
            error_msg = f"ファイル処理エラー: {str(e)}"
            print(f"❌ {error_msg}")
            return {"statusCode": 500, "body": json.dumps({"success": False, "error": error_msg}, ensure_ascii=False)}
        
        print(f"\n=== 処理完了サマリー ===")
        print(f"✅ 成功: {len(processed_files)} ファイル")
        print(f"📊 合計: 1 ファイル")
        
        # 結果を返す
        result = {
            'statusCode': 200,
            'body': json.dumps({
                'success': True,
                'message': '月別集計ファイルが正常に作成されました',
                'processed': processed_files,
                'uploaded': uploaded_files,
                'count': len(processed_files)
            }, ensure_ascii=False)
        }
        
        print(f"🔍 デバッグ: 最終結果 - 処理済み: {len(processed_files)}, アップロード済み: {len(uploaded_files)}")
        
        return result
            
    except Exception as e:
        print(f"Lambda関数でエラーが発生: {str(e)}")
        return {
            'statusCode': 500,
            'body': json.dumps({
                'error': f'予期しないエラーが発生しました: {str(e)}'
            }, ensure_ascii=False)
        }

if __name__ == "__main__":
    # ローカルテスト用（monthly_6_usage形式をシミュレート）
    test_event = {
        "s3_bucket": "monthly-report-output-group",
        "s3_key": "bigquery-data/web_reservation_leaf_20250901.csv",
        "output_to_s3": True,
        "s3_output_bucket": "monthly-report-output-group",
        "s3_output_prefix": "bigquery-data/"
    }
    
    result = lambda_handler(test_event, None)
    print(f"テスト結果: {json.dumps(result, ensure_ascii=False, indent=2)}")
