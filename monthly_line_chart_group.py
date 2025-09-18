from pathlib import Path
from datetime import datetime, timedelta
import sys
import boto3
from botocore.exceptions import ClientError
import os
import json
import csv
from collections import defaultdict

# Excel操作用のライブラリをインポート
try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
    print("✅ xlsxwriterのインポート成功")
except ImportError as e:
    print(f"❌ xlsxwriterのインポート失敗: {e}")
    XLSXWRITER_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.chart import LineChart, Reference
    OPENPYXL_AVAILABLE = True
    print("✅ openpyxlのインポート成功")
except ImportError as e:
    print(f"❌ openpyxlのインポート失敗: {e}")
    OPENPYXL_AVAILABLE = False

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
        print(f" デバッグ: ファイルパス: '{file_path}'")
        print(f" デバッグ: バケット名: '{bucket_name}'")
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
        print(f" ファイルサイズ: {file_size} bytes")
        
        # 複数のエンコーディングを試す
        encodings = [
            'utf-8-sig', 'utf-8', 'utf-16', 'utf-32',
            'shift_jis', 'cp932', 'euc-jp', 'iso-2022-jp',
            'latin1', 'cp1252', 'ascii', 'big5', 'gb2312'
        ]
        
        for encoding in encodings:
            try:
                print(f" エンコーディング試行: {encoding}")
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
        monthly_data = defaultdict(lambda: {'初回受付': 0, '再受付': 0, '初回キャンセル': 0, '合計': 0})
        
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
            
            # SYSTEM_USAGE_NAMEによる受付種別の判定
            system_usage = row.get('SYSTEM_USAGE_NAME', '').strip()
            
            if '初回受付' in system_usage:
                monthly_data[year_month]['初回受付'] += 1
            elif '再受付' in system_usage:
                monthly_data[year_month]['再受付'] += 1
            elif '初回キャンセル' in system_usage:
                monthly_data[year_month]['初回キャンセル'] += 1
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

def create_combined_data(hp_data, other_data):
    """HPとその他のデータを統合してグラフ用のデータを作成"""
    try:
        # 全月のリストを作成
        all_months = set()
        if hp_data:
            all_months.update([d['月'] for d in hp_data])
        if other_data:
            all_months.update([d['月'] for d in other_data])
        
        all_months = sorted(list(all_months))
        
        # 統合データを作成
        combined_data = []
        for month in all_months:
            # HPデータを取得
            hp_month_data = next((d for d in hp_data if d['月'] == month), {
                '月': month, '初回受付': 0, '再受付': 0, '初回キャンセル': 0, '合計': 0
            })
            
            # その他データを取得
            other_month_data = next((d for d in other_data if d['月'] == month), {
                '月': month, '初回受付': 0, '再受付': 0, '初回キャンセル': 0, '合計': 0
            })
            
            # 統合データを作成
            combined_data.append({
                '月': month,
                'HP初回受付': hp_month_data['初回受付'],
                'HP再受付': hp_month_data['再受付'],
                'その他初回受付': other_month_data['初回受付'],
                'その他再受付': other_month_data['再受付'],
                'HP合計': hp_month_data['合計'],
                'その他合計': other_month_data['合計'],
                '総合計': hp_month_data['合計'] + other_month_data['合計']
            })
        
        return combined_data
        
    except Exception as e:
        print(f"❌ 統合データ作成エラー: {e}")
        return None

def create_excel_with_enhanced_chart(hp_data, other_data, output_path: Path) -> bool:
    """xlsxwriterで新しいExcelファイルを作成（AWS_monthly_line_chart.pyと同じ方式）"""
    try:
        if not XLSXWRITER_AVAILABLE:
            print("❌ xlsxwriterが利用できないため、Excelファイルは作成されません")
            return False
            
        print(f"🔍 デバッグ: xlsxwriterで新しいExcelファイル作成開始")
        
        # 新しいExcelファイルを作成
        workbook = xlsxwriter.Workbook(str(output_path))
        
        # フォーマットを定義
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1
        })
        
        data_format = workbook.add_format({
            'border': 1
        })
        
        # 1. HP経由シートを作成
        if hp_data:
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
            ws_hp = workbook.add_worksheet(f'HP月別集計_{timestamp}')
            
            # ヘッダーを書き込み
            hp_headers = ['月', '初回受付', '再受付', '初回キャンセル', '合計']
            for col, header in enumerate(hp_headers):
                ws_hp.write(0, col, header, header_format)
            
            # データを書き込み
            for row, data_point in enumerate(hp_data, 1):
                ws_hp.write(row, 0, data_point['月'], data_format)
                ws_hp.write(row, 1, data_point['初回受付'], data_format)
                ws_hp.write(row, 2, data_point['再受付'], data_format)
                ws_hp.write(row, 3, data_point['初回キャンセル'], data_format)
                ws_hp.write(row, 4, data_point['合計'], data_format)
            
            # 列幅を自動調整
            for col in range(len(hp_headers)):
                ws_hp.set_column(col, col, 15)
            
            # HP用のグラフを作成（xlsxwriter版）
            create_hp_chart_xlsxwriter(workbook, ws_hp, hp_data, f'HP月別集計_{timestamp}')
        
        # 2. その他経由シートを作成
        if other_data:
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
            ws_other = workbook.add_worksheet(f'その他月別集計_{timestamp}')
            
            # ヘッダーを書き込み
            other_headers = ['月', '初回受付', '再受付', '初回キャンセル', '合計']
            for col, header in enumerate(other_headers):
                ws_other.write(0, col, header, header_format)
            
            # データを書き込み
            for row, data_point in enumerate(other_data, 1):
                ws_other.write(row, 0, data_point['月'], data_format)
                ws_other.write(row, 1, data_point['初回受付'], data_format)
                ws_other.write(row, 2, data_point['再受付'], data_format)
                ws_other.write(row, 3, data_point['初回キャンセル'], data_format)
                ws_other.write(row, 4, data_point['合計'], data_format)
            
            # 列幅を自動調整
            for col in range(len(other_headers)):
                ws_other.set_column(col, col, 15)
            
            # その他用のグラフを作成（xlsxwriter版）
            create_other_chart_xlsxwriter(workbook, ws_other, other_data, f'その他月別集計_{timestamp}')
        
        # ワークブックを閉じる
        workbook.close()
        
        print(f"✅ xlsxwriterで新しいExcelファイル作成完了: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ 拡張Excelファイル作成エラー: {e}")
        return False


def create_hp_chart_xlsxwriter(workbook, worksheet, hp_data, sheet_name):
    """HP用のグラフを作成（AWS_monthly_line_chart.pyと完全に同じスタイル）"""
    try:
        if not XLSXWRITER_AVAILABLE:
            print("❌ xlsxwriterが利用できないため、グラフは作成されません")
            return False
            
        # グラフを作成
        chart = workbook.add_chart({'type': 'line'})
        
        # グラフのタイトルとスタイル（AWS_monthly_line_chart.pyと完全に同じ）
        chart.set_title({
            'name': f'月別予約受付数推移 ({sheet_name})',
            'name_font': {'size': 16, 'bold': True}
        })
        
        # データ系列を追加（HP用の色 - AWS_monthly_line_chart.pyと完全に同じ）
        series_colors = [
            {'name': '初回受付', 'color': '#2E8B57'},      # 緑色
            {'name': '再受付', 'color': '#4169E1'}        # 青色
        ]
        
        for i, series_info in enumerate(series_colors, 1):
            chart.add_series({
                'name': series_info['name'],
                'categories': [sheet_name, 1, 0, len(hp_data), 0],  # A2:A(last_row)
                'values': [sheet_name, 1, i, len(hp_data), i],      # B2:C(last_row)
                'line': {'color': series_info['color'], 'width': 2},
                'marker': {
                    'type': 'circle',
                    'size': 8,
                    'line': {'color': series_info['color']},
                    'fill': {'color': series_info['color']}
                },
                'data_labels': {
                    'value': True,
                    'font': {'size': 13, 'bold': True}
                }
            })
        
        # グラフのレイアウト設定（AWS_monthly_line_chart.pyと完全に同じ）
        chart.set_size({'width': 800, 'height': 400})
        
        # グラフエリア全体の背景色を設定（余白部分も含む）- 薄い緑色
        chart.set_chartarea({
            'fill': {'color': '#E8F5E8'}  # 薄い緑色
        })
        
        # プロットエリアの背景色を設定
        chart.set_plotarea({
            'fill': {'color': '#E8F5E8'},  # 薄い緑色
            'layout': {
                'x': 0.1,
                'y': 0.1,
                'width': 0.8,
                'height': 0.8
            }
        })
        
        # X軸の設定（ラベルなし）
        chart.set_x_axis({
            'num_font': {'size': 10},
            'major_gridlines': {'visible': False}
        })
        
        # Y軸の設定（ラベルなし、数値なし）
        chart.set_y_axis({
            'num_font': {'color': '#E8F5E8'},  # 数値を背景色と同じにして非表示
            'major_gridlines': {'visible': False},
            'minor_gridlines': {'visible': False},
            'line': {'color': '#E8F5E8'}  # 軸線も背景色と同じにして非表示
        })
        
        # 凡例の設定
        chart.set_legend({
            'position': 'top',
            'font': {'size': 10, 'bold': True}
        })
        
        # グラフをワークシートに挿入
        worksheet.insert_chart('F2', chart)
        
        print(f"✅ {sheet_name}グラフ作成完了")
        return True
        
    except Exception as e:
        print(f"❌ {sheet_name}グラフ作成エラー: {e}")
        return False

def create_other_chart_xlsxwriter(workbook, worksheet, other_data, sheet_name):
    """その他用のグラフを作成（AWS_monthly_line_chart.pyと完全に同じスタイル）"""
    try:
        if not XLSXWRITER_AVAILABLE:
            print("❌ xlsxwriterが利用できないため、グラフは作成されません")
            return False
            
        # グラフを作成
        chart = workbook.add_chart({'type': 'line'})
        
        # グラフのタイトルとスタイル（AWS_monthly_line_chart.pyと完全に同じ）
        chart.set_title({
            'name': f'月別予約受付数推移 ({sheet_name})',
            'name_font': {'size': 16, 'bold': True}
        })
        
        # データ系列を追加（その他用の色 - AWS_monthly_line_chart.pyと完全に同じ）
        series_colors = [
            {'name': '初回受付', 'color': '#2E8B57'},      # 緑色
            {'name': '再受付', 'color': '#4169E1'}        # 青色
        ]
        
        for i, series_info in enumerate(series_colors, 1):
            chart.add_series({
                'name': series_info['name'],
                'categories': [sheet_name, 1, 0, len(other_data), 0],  # A2:A(last_row)
                'values': [sheet_name, 1, i, len(other_data), i],      # B2:C(last_row)
                'line': {'color': series_info['color'], 'width': 2},
                'marker': {
                    'type': 'circle',
                    'size': 8,
                    'line': {'color': series_info['color']},
                    'fill': {'color': series_info['color']}
                },
                'data_labels': {
                    'value': True,
                    'font': {'size': 13, 'bold': True}
                }
            })
        
        # グラフのレイアウト設定（AWS_monthly_line_chart.pyと完全に同じ）
        chart.set_size({'width': 800, 'height': 400})
        
        # グラフエリア全体の背景色を設定（余白部分も含む）
        chart.set_chartarea({
            'fill': {'color': '#8EB4E3'}  # グラフ全体の背景色
        })
        
        # プロットエリアの背景色を設定
        chart.set_plotarea({
            'fill': {'color': '#8EB4E3'},  # プロットエリアの背景色
            'layout': {
                'x': 0.1,
                'y': 0.1,
                'width': 0.8,
                'height': 0.8
            }
        })
        
        # X軸の設定（ラベルなし）
        chart.set_x_axis({
            'num_font': {'size': 10},
            'major_gridlines': {'visible': False}
        })
        
        # Y軸の設定（ラベルなし、数値なし）
        chart.set_y_axis({
            'num_font': {'color': '#8EB4E3'},  # 数値を背景色と同じにして非表示
            'major_gridlines': {'visible': False},
            'minor_gridlines': {'visible': False},
            'line': {'color': '#8EB4E3'}  # 軸線も背景色と同じにして非表示
        })
        
        # 凡例の設定
        chart.set_legend({
            'position': 'top',
            'font': {'size': 10, 'bold': True}
        })
        
        # グラフをワークシートに挿入
        worksheet.insert_chart('F2', chart)
        
        print(f"✅ {sheet_name}グラフ作成完了")
        return True
        
    except Exception as e:
        print(f"❌ {sheet_name}グラフ作成エラー: {e}")
        return False

def create_hp_chart_openpyxl(workbook, worksheet, hp_data, sheet_name):
    """HP用のグラフを作成（シンプルで確実に動作するバージョン）"""
    try:
        from openpyxl.chart import LineChart, Reference
        
        print(f"🔍 デバッグ: HP用グラフ作成開始 - シート: {sheet_name}")
        print(f"🔍 デバッグ: データ行数: {len(hp_data)}")
        
        # グラフを作成
        chart = LineChart()
        chart.title = f'月別予約受付数推移 ({sheet_name})'
        chart.style = 13
        
        # データ範囲を設定（初回受付と再受付のみ、初回キャンセルは除外）
        data = Reference(worksheet, min_col=2, min_row=1, max_row=len(hp_data)+1, max_col=3)
        chart.add_data(data, titles_from_data=True)
        
        # カテゴリ（月）を設定
        categories = Reference(worksheet, min_col=1, min_row=2, max_row=len(hp_data)+1)
        chart.set_categories(categories)
        
        # グラフのサイズ設定
        chart.width = 20
        chart.height = 10
        
        # 軸の設定
        chart.x_axis.majorGridlines = None
        chart.y_axis.majorGridlines = None
        chart.y_axis.minorGridlines = None
        chart.y_axis.scaling.min = 0
        
        # 凡例の設定
        chart.legend.position = 't'  # top
        
        # グラフをシートに追加
        worksheet.add_chart(chart, "F2")
        
        print(f"✅ HP用グラフ作成完了: {sheet_name}")
        
    except Exception as e:
        print(f"❌ HP用グラフ作成エラー: {e}")
        import traceback
        traceback.print_exc()

def create_other_chart_openpyxl(workbook, worksheet, other_data, sheet_name):
    """その他用のグラフを作成（シンプルで確実に動作するバージョン）"""
    try:
        from openpyxl.chart import LineChart, Reference
        
        print(f"🔍 デバッグ: その他用グラフ作成開始 - シート: {sheet_name}")
        print(f"🔍 デバッグ: データ行数: {len(other_data)}")
        
        # グラフを作成
        chart = LineChart()
        chart.title = f'月別予約受付数推移 ({sheet_name})'
        chart.style = 13
        
        # データ範囲を設定（初回受付と再受付のみ、初回キャンセルは除外）
        data = Reference(worksheet, min_col=2, min_row=1, max_row=len(other_data)+1, max_col=3)
        chart.add_data(data, titles_from_data=True)
        
        # カテゴリ（月）を設定
        categories = Reference(worksheet, min_col=1, min_row=2, max_row=len(other_data)+1)
        chart.set_categories(categories)
        
        # グラフのサイズ設定
        chart.width = 20
        chart.height = 10
        
        # 軸の設定
        chart.x_axis.majorGridlines = None
        chart.y_axis.majorGridlines = None
        chart.y_axis.minorGridlines = None
        chart.y_axis.scaling.min = 0
        
        # 凡例の設定
        chart.legend.position = 't'  # top
        
        # グラフをシートに追加
        worksheet.add_chart(chart, "F2")
        
        print(f"✅ その他用グラフ作成完了: {sheet_name}")
        
    except Exception as e:
        print(f"❌ その他用グラフ作成エラー: {e}")
        import traceback
        traceback.print_exc()

def create_hp_chart(workbook, worksheet, hp_data, sheet_name):
    """HP用のグラフを作成（monthly_line_chart.pyと同じスタイル）"""
    try:
        # グラフを作成
        chart = workbook.add_chart({'type': 'line'})
        
        # グラフのタイトルとスタイル
        chart.set_title({
            'name': f'月別予約受付数推移 ({sheet_name})',
            'name_font': {'size': 16, 'bold': True}
        })
        
        # データ系列を追加（HP用の色）
        series_colors = [
            {'name': '初回受付', 'color': '#2E8B57'},      # 緑色
            {'name': '再受付', 'color': '#4169E1'}        # 青色
        ]
        
        for i, series_info in enumerate(series_colors, 1):
            chart.add_series({
                'name': series_info['name'],
                'categories': [sheet_name, 1, 0, len(hp_data), 0],  # A2:A(last_row)
                'values': [sheet_name, 1, i, len(hp_data), i],      # B2:C(last_row)
                'line': {'color': series_info['color'], 'width': 2},
                'marker': {
                    'type': 'circle',
                    'size': 8,
                    'line': {'color': series_info['color']},
                    'fill': {'color': series_info['color']}
                },
                'data_labels': {
                    'value': True,
                    'font': {'size': 13, 'bold': True}
                }
            })
        
        # グラフのレイアウト設定
        chart.set_size({'width': 800, 'height': 400})
        
        # グラフエリア全体の背景色を設定（余白部分も含む）- 薄い緑色
        chart.set_chartarea({
            'fill': {'color': '#E8F5E8'}  # 薄い緑色
        })
        
        # プロットエリアの背景色を設定
        chart.set_plotarea({
            'fill': {'color': '#E8F5E8'},  # 薄い緑色
            'layout': {
                'x': 0.1,
                'y': 0.1,
                'width': 0.8,
                'height': 0.8
            }
        })
        
        # X軸の設定（ラベルなし）
        chart.set_x_axis({
            'num_font': {'size': 10},
            'major_gridlines': {'visible': False}
        })
        
        # Y軸の設定（ラベルなし、数値なし）
        chart.set_y_axis({
            'num_font': {'color': '#E8F5E8'},  # 数値を背景色と同じにして非表示
            'major_gridlines': {'visible': False},
            'minor_gridlines': {'visible': False},
            'line': {'color': '#E8F5E8'}  # 軸線も背景色と同じにして非表示
        })
        
        # 凡例の設定
        chart.set_legend({
            'position': 'top',
            'font': {'size': 10, 'bold': True}
        })
        
        # グラフをワークシートに挿入
        worksheet.insert_chart('F2', chart)
        
        print(f"✅ {sheet_name}グラフ作成完了")
        
    except Exception as e:
        print(f"❌ {sheet_name}グラフ作成エラー: {e}")

def create_other_chart(workbook, worksheet, other_data, sheet_name):
    """その他用のグラフを作成（monthly_line_chart.pyと同じスタイル）"""
    try:
        # グラフを作成
        chart = workbook.add_chart({'type': 'line'})
        
        # グラフのタイトルとスタイル
        chart.set_title({
            'name': f'月別予約受付数推移 ({sheet_name})',
            'name_font': {'size': 16, 'bold': True}
        })
        
        # データ系列を追加（その他用の色）
        series_colors = [
            {'name': '初回受付', 'color': '#2E8B57'},      # 緑色
            {'name': '再受付', 'color': '#4169E1'}        # 青色
        ]
        
        for i, series_info in enumerate(series_colors, 1):
            chart.add_series({
                'name': series_info['name'],
                'categories': [sheet_name, 1, 0, len(other_data), 0],  # A2:A(last_row)
                'values': [sheet_name, 1, i, len(other_data), i],      # B2:C(last_row)
                'line': {'color': series_info['color'], 'width': 2},
                'marker': {
                    'type': 'circle',
                    'size': 8,
                    'line': {'color': series_info['color']},
                    'fill': {'color': series_info['color']}
                },
                'data_labels': {
                    'value': True,
                    'font': {'size': 13, 'bold': True}
                }
            })
        
        # グラフのレイアウト設定
        chart.set_size({'width': 800, 'height': 400})
        
        # グラフエリア全体の背景色を設定（余白部分も含む）
        chart.set_chartarea({
            'fill': {'color': '#8EB4E3'}  # グラフ全体の背景色
        })
        
        # プロットエリアの背景色を設定
        chart.set_plotarea({
            'fill': {'color': '#8EB4E3'},  # プロットエリアの背景色
            'layout': {
                'x': 0.1,
                'y': 0.1,
                'width': 0.8,
                'height': 0.8
            }
        })
        
        # X軸の設定（ラベルなし）
        chart.set_x_axis({
            'num_font': {'size': 10},
            'major_gridlines': {'visible': False}
        })
        
        # Y軸の設定（ラベルなし、数値なし）
        chart.set_y_axis({
            'num_font': {'color': '#8EB4E3'},  # 数値を背景色と同じにして非表示
            'major_gridlines': {'visible': False},
            'minor_gridlines': {'visible': False},
            'line': {'color': '#8EB4E3'}  # 軸線も背景色と同じにして非表示
        })
        
        # 凡例の設定
        chart.set_legend({
            'position': 'top',
            'font': {'size': 10, 'bold': True}
        })
        
        # グラフをワークシートに挿入
        worksheet.insert_chart('F2', chart)
        
        print(f"✅ {sheet_name}グラフ作成完了")
        
    except Exception as e:
        print(f"❌ {sheet_name}グラフ作成エラー: {e}")

def create_combined_chart(workbook, worksheet, combined_data, sheet_name):
    """統合グラフを作成（monthly_line_chart.pyと同じスタイル）"""
    try:
        # グラフを作成
        chart = workbook.add_chart({'type': 'line'})
        
        # グラフのタイトルとスタイル
        chart.set_title({
            'name': '月別予約受付数推移',
            'name_font': {'size': 16, 'bold': True}
        })
        
        # データ系列を追加（monthly_line_chart.pyと同じ色）
        series_colors = [
            {'name': 'HP新規', 'color': '#CD8433'},      # HP初回受付
            {'name': 'HP再受付', 'color': '#808080'},    # HP再受付
            {'name': 'WEB新規', 'color': '#2E8B57'},     # その他初回受付
            {'name': 'WEB再受付', 'color': '#4169E1'}    # その他再受付
        ]
        
        for i, series_info in enumerate(series_colors, 1):
            chart.add_series({
                'name': series_info['name'],
                'categories': [sheet_name, 1, 0, len(combined_data), 0],  # A2:A(last_row)
                'values': [sheet_name, 1, i, len(combined_data), i],      # B2:E(last_row)
                'line': {'color': series_info['color'], 'width': 2},
                'marker': {
                    'type': 'circle',
                    'size': 8,
                    'line': {'color': series_info['color']},
                    'fill': {'color': series_info['color']}
                },
                'data_labels': {
                    'value': True,
                    'font': {'size': 13, 'bold': True}
                }
            })
        
        # グラフのレイアウト設定
        chart.set_size({'width': 800, 'height': 400})
        
        # グラフエリア全体の背景色を設定（余白部分も含む）
        chart.set_chartarea({
            'fill': {'color': '#8EB4E3'}  # グラフ全体の背景色
        })
        
        # プロットエリアの背景色を設定
        chart.set_plotarea({
            'fill': {'color': '#8EB4E3'},  # プロットエリアの背景色
            'layout': {
                'x': 0.1,
                'y': 0.1,
                'width': 0.8,
                'height': 0.8
            }
        })
        
        # X軸の設定（ラベルなし）
        chart.set_x_axis({
            'num_font': {'size': 10},
            'major_gridlines': {'visible': False}
        })
        
        # Y軸の設定（ラベルなし、数値なし）
        chart.set_y_axis({
            'num_font': {'color': '#8EB4E3'},  # 数値を背景色と同じにして非表示
            'major_gridlines': {'visible': False},
            'minor_gridlines': {'visible': False},
            'line': {'color': '#8EB4E3'}  # 軸線も背景色と同じにして非表示
        })
        
        # 凡例の設定
        chart.set_legend({
            'position': 'top',
            'font': {'size': 10, 'bold': True}
        })
        
        # グラフをワークシートに挿入
        worksheet.insert_chart('G2', chart)
        
        print(f"✅ {sheet_name}グラフ作成完了")
        
    except Exception as e:
        print(f"❌ {sheet_name}グラフ作成エラー: {e}")

def process_csv_file_enhanced(csv_path: str, upload_to_s3_flag: bool = False, s3_bucket: str = None) -> dict:
    """CSVファイルを処理して新しい月別集計Excelファイルを作成"""
    try:
        print(f" 拡張CSVファイル処理開始: {csv_path}")
        
        # S3クライアントを取得（アップロード用）
        s3_client = get_s3_client()
        if not s3_client:
            return {"success": False, "error": "S3クライアントの作成に失敗"}
        
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
        
        # 新しいExcelファイルを作成
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        output_filename = f"月別集計_一年間_{timestamp}.xlsx"
        output_path = Path(f"/tmp/{output_filename}")
        
        try:
            excel_success = create_excel_with_enhanced_chart(hp_data, other_data, output_path)
            
            if not excel_success:
                return {"success": False, "error": "拡張Excelファイルの作成に失敗"}
                
        except Exception as e:
            print(f"❌ ファイル作成エラー: {str(e)}")
            return {"success": False, "error": f"ファイルの作成に失敗: {str(e)}"}
        
        # S3にアップロード（フラグがTrueの場合）
        s3_results = []
        if upload_to_s3_flag and s3_bucket:
            upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
            
            # 新しい月別集計ファイルをアップロード
            if output_path.exists():
                file_s3_key = f"{upload_prefix}{output_filename}"
                file_result = upload_to_s3(str(output_path), s3_bucket, file_s3_key)
                s3_results.append(file_result)
        
        # 成功したアップロードのみを結果に含める
        successful_uploads = [result for result in s3_results if result.get('success', False)]
        
        return {
            "success": True,
            "message": f"月別集計ファイル作成完了: {output_filename}",
            "hp_months_count": len(hp_data),
            "other_months_count": len(other_data),
            "excel_created": excel_success,
            "s3_uploads": successful_uploads,
            "local_path": str(output_path)
        }
        
    except Exception as e:
        print(f"❌ 拡張CSVファイル処理エラー: {e}")
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
        
        print(f" デバッグ: 環境変数S3_BUCKET_NAME: '{os.environ.get('S3_BUCKET_NAME', 'NOT_SET')}'")
        print(f"🔍 デバッグ: 環境変数S3_UPLOAD_PREFIX: '{os.environ.get('S3_UPLOAD_PREFIX', 'NOT_SET')}'")
        print(f"🔍 デバッグ: すべての環境変数: {dict(os.environ)}")
        
        print(f" デバッグ: 使用するS3バケット: '{s3_bucket}'")
        print(f"🔍 デバッグ: アップロードプレフィックス: '{upload_prefix}'")
        
        print(f" デバッグ: 環境変数 - バケット: '{s3_bucket}', プレフィックス: '{upload_prefix}'")
        
        if not s3_bucket:
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "s3_bucketが未指定"}, ensure_ascii=False)}
        
        # Step Functions用の入力形式に対応（monthly_6_usage形式）
        # 引数: {"s3_bucket": "...", "s3_key": "...", "output_to_s3": true, ...}
        print(f"🔍 デバッグ: 受信したイベント全体: {json.dumps(event, ensure_ascii=False, indent=2)}")
        print(f" デバッグ: イベントの型: {type(event)}")
        print(f" デバッグ: イベントのキー: {list(event.keys()) if isinstance(event, dict) else 'Not a dict'}")
        
        # monthly_6_usage形式の引数からS3情報を取得
        input_s3_bucket = event.get('s3_bucket', 'monthly-report-output-group')
        input_s3_key = event.get('s3_key', 'bigquery-data/web_reservation_leaf_20250901.csv')
        
        print(f" デバッグ: 入力S3バケット: {input_s3_bucket}")
        print(f" デバッグ: 入力S3キー: {input_s3_key}")
        
        if not input_s3_bucket or not input_s3_key:
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "s3_bucketまたはs3_keyが指定されていません"}, ensure_ascii=False)}
        
        # S3パスを構築
        s3_path = f"s3://{input_s3_bucket}/{input_s3_key}"
        result_data = {"bucket": input_s3_bucket, "key": input_s3_key}
        
        print(f" デバッグ: 構築されたS3パス: {s3_path}")
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
            print(f" デバッグ: バケット名: {input_s3_bucket}")
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
            print(f" データ処理開始（HPとその他を一つのファイルに）")
            
            result = process_csv_file_enhanced(
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
