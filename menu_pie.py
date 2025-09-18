import json
import os
import tempfile
import shutil
from pathlib import Path
import sys
import re
try:
    import boto3
    from botocore.exceptions import ClientError
except ImportError:
    # ローカル環境ではboto3が不要
    boto3 = None
    ClientError = Exception
from datetime import datetime
from typing import List, Dict, Optional, Tuple



def get_s3_client():
    """S3クライアントを取得"""
    try:
        if boto3 is None:
            return None
        return boto3.client('s3')
    except Exception as e:
        print(f"S3クライアントの作成に失敗: {e}")
        return None


def download_from_s3(s3_path: str) -> str:
    """S3からファイルをダウンロード"""
    try:
        print(f"download_from_s3開始: '{s3_path}' (長さ: {len(s3_path)})")
        
        # スペースを除去
        original_path = s3_path
        s3_path = s3_path.replace(' ', '')
        print(f"スペース除去後: '{s3_path}' (長さ: {len(s3_path)})")
        
        # S3パスを解析 (s3://bucket/key)
        if not s3_path.startswith('s3://'):
            print(f"S3パスではありません: {s3_path}")
            return None
        
        path_parts = s3_path[5:].split('/', 1)  # s3://を除去して分割
        print(f"パス分割結果: {path_parts}")
        if len(path_parts) != 2:
            print(f"パス分割に失敗: {path_parts}")
            return None
        
        bucket_name, s3_key = path_parts
        bucket_name = bucket_name.strip()
        s3_key = s3_key.lstrip('/')
        print(f"バケット名: '{bucket_name}', S3キー: '{s3_key}'")
        
        # ローカルファイル名を生成
        filename = os.path.basename(s3_key)
        local_path = os.path.join('/tmp', filename)
        
        # S3クライアントを取得
        s3_client = get_s3_client()
        if not s3_client:
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
        print(f"🔍 デバッグ: バケット名: '{bucket_name}' (長さ: {len(bucket_name)})")
        print(f"🔍 デバッグ: S3キー: '{s3_key}' (長さ: {len(s3_key)})")
        print(f"🔍 デバッグ: ファイル存在: {os.path.exists(file_path)}")
        
        # バケット名の検証
        if not bucket_name or bucket_name.strip() == '':
            print("❌ バケット名が空です")
            return {"success": False, "error": "バケット名が空です"}
        
        # バケット名の正規化（スペースを除去）
        bucket_name = bucket_name.strip()
        print(f"🔍 デバッグ: 正規化後のバケット名: '{bucket_name}'")
        
        # バケット名の形式チェック
        import re
        if not re.match(r'^[a-zA-Z0-9.\-_]{1,255}$', bucket_name):
            print(f"❌ 無効なバケット名: '{bucket_name}'")
            return {"success": False, "error": f"無効なバケット名: '{bucket_name}'"}
        
        # S3キーの正規化（スペースを除去）
        s3_key = s3_key.strip()
        print(f"🔍 デバッグ: 正規化後のS3キー: '{s3_key}'")
        
        # ファイルの存在確認
        if not os.path.exists(file_path):
            print(f"❌ ファイルが存在しません: {file_path}")
            return {"success": False, "error": f"ファイルが存在しません: {file_path}"}
        
        # ファイルサイズの確認
        file_size = os.path.getsize(file_path)
        print(f"🔍 デバッグ: ファイルサイズ: {file_size} bytes")
        
        s3_client = get_s3_client()
        if not s3_client:
            print("❌ S3クライアントの作成に失敗")
            return {"success": False, "error": "S3クライアントの作成に失敗"}
        
        print(f"🔍 デバッグ: S3クライアント取得成功")
        
        # バケットの存在確認
        try:
            s3_client.head_bucket(Bucket=bucket_name)
            print(f"🔍 デバッグ: バケット '{bucket_name}' の存在確認成功")
        except Exception as e:
            print(f"❌ バケット '{bucket_name}' の存在確認失敗: {str(e)}")
            return {"success": False, "error": f"バケット '{bucket_name}' にアクセスできません: {str(e)}"}
        
        # ファイルアップロード
        print(f"🔍 デバッグ: ファイルアップロード開始...")
        s3_client.upload_file(file_path, bucket_name, s3_key)
        print(f"✅ S3アップロード成功")
        
        # S3 URLを生成
        region = os.environ.get('AWS_REGION', 'ap-northeast-1')
        s3_url = f"https://{bucket_name}.s3.{region}.amazonaws.com/{s3_key}"
        
        # S3パスを生成
        s3_path = f"s3://{bucket_name}/{s3_key}"
        
        print(f"🔍 デバッグ: 生成されたS3パス: {s3_path}")
        
        return {
            "success": True,
            "s3_key": s3_key,
            "s3_path": s3_path,
            "s3_url": s3_url,
            "bucket": bucket_name
        }
        
    except Exception as e:
        print(f"❌ S3アップロード失敗: {str(e)}")
        import traceback
        print(f"🔍 デバッグ: エラーの詳細: {traceback.format_exc()}")
        return {"success": False, "error": f"S3アップロードに失敗: {str(e)}"}


def extract_bracket_content(menu_name: str) -> str:
    """
    メニュー名から括弧内の内容を抽出する
    例: "メニューA（カテゴリ1）" → "カテゴリ1"
    """
    if not menu_name or not isinstance(menu_name, str):
        return "健康保険適用の施術を相談したい方"
    
    # 文字列を正規化（前後の空白を除去、連続する空白を単一の空白に統一）
    menu_name = re.sub(r'\s+', ' ', menu_name.strip())
    
    # 括弧のパターンを検索（全角・半角両方対応）
    patterns = [
        r'（(.+?)）',  # 全角括弧
        r'\((.+?)\)',  # 半角括弧
        r'【(.+?)】',  # 全角角括弧
        r'\[(.+?)\]',  # 半角角括弧
    ]
    
    for pattern in patterns:
        match = re.search(pattern, menu_name)
        if match:
            content = match.group(1).strip()
            # 空文字や空白文字のみの場合は「健康保険適用の施術を相談したい方」を返す
            if content and not content.isspace():
                return content
            else:
                # 空の括弧の場合は「健康保険適用の施術を相談したい方」を返す
                return "健康保険適用の施術を相談したい方"
    
    # 括弧が見つからない場合は「健康保険適用の施術を相談したい方」を返す
    return "健康保険適用の施術を相談したい方"


def add_pivot_pie(target_path: Path, upload_to_s3_flag: bool = False, s3_bucket: str = None, original_s3_key: str = None) -> dict:
    """
    メニュー分類別の円グラフを作成する関数
    Lambda環境用に修正（win32comの代わりにopenpyxlを使用）
    """
    print(f"🎯 グラフ作成開始: {target_path}")
    print(f"🔍 デバッグ: target_pathの型: {type(target_path)}")
    
    # target_pathがPathオブジェクトであることを確認
    if not isinstance(target_path, Path):
        print(f"❌ エラー: target_pathがPathオブジェクトではありません: {type(target_path)}")
        print(f"❌ エラー: target_pathの内容: {target_path}")
        return {
            "success": False,
            "error": f"target_pathの型エラー: {type(target_path)}"
        }
    
    print(f"🔍 デバッグ: target_pathが存在するか: {target_path.exists()}")
    
    try:
        from openpyxl import load_workbook
        from openpyxl.chart import PieChart, Reference
        from openpyxl.chart.label import DataLabelList
    except ImportError as e:
        print(f"❌ openpyxlのインポート失敗: {e}")
        return {
            "success": False,
            "error": "openpyxl が未インストールです。`pip install openpyxl` を実行してください。"
        }

    try:
        # ファイルを読み込み
        wb = load_workbook(str(target_path))
        
        # 統合されたExcelファイルの場合、各データシートごとに円グラフを作成
        print(f"    🔍 メニューグラフ作成開始: {target_path}")
        print(f"    📋 全シート一覧: {wb.sheetnames}")
        
        data_sheets = []
        for sheet_name in wb.sheetnames:
            # データシートを特定（グラフシートやピボットシートを除外）
            if not any(keyword in sheet_name.lower() for keyword in ['ピボット', 'グラフ', 'chart', 'pivot']):
                data_sheets.append(wb[sheet_name])
                print(f"    📊 データシートとして認識: {sheet_name}")
        
        print(f"    📊 データシート数: {len(data_sheets)}")
        if not data_sheets:
            return {
                "success": False,
                "error": "データシートが見つかりません"
            }

        # 各シートごとに円グラフを作成
        created_charts = 0
        for i, data_ws in enumerate(data_sheets):
            print(f"    📊 シート {i+1} 処理中: {data_ws.title}")
            max_row = data_ws.max_row
            max_col = data_ws.max_column
            print(f"    📊 シートサイズ: {max_row}行 x {max_col}列")
            
            if max_row <= 1 or max_col <= 0:
                print(f"    ⚠️ シート {i+1} はデータが不足")
                continue

            # MENU_NAME列の検索
            menu_name_col = None
            for col in range(1, max_col + 1):
                header_val = data_ws.cell(row=1, column=col).value
                if header_val == "MENU_NAME":
                    menu_name_col = col
                    break
            
            if menu_name_col is None:
                print(f"    ⚠️ シート {i+1} にMENU_NAME列が見つからない")
                continue
            
            print(f"    📊 MENU_NAME列: {menu_name_col}列目")

            # このシートからメニュー名と件数を収集
            menu_counts = {}
            total_count = 0
            
            for row in range(2, max_row + 1):  # ヘッダー行を除く
                menu_name = data_ws.cell(row=row, column=menu_name_col).value
                if menu_name and menu_name != "MENU_NAME":  # ヘッダー行を除外
                    # 括弧内の内容で分類
                    category = extract_bracket_content(str(menu_name))
                    
                    if category in menu_counts:
                        menu_counts[category] += 1
                    else:
                        menu_counts[category] = 1
                    total_count += 1
            
            
            print(f"    📊 シート {i+1} から収集: {total_count}件, カテゴリ数: {len(menu_counts)}")

            if not menu_counts:
                print(f"    ⚠️ シート {i+1} に有効なデータがありません")
                continue

            # このシート用の円グラフを作成
            try:
                ts = datetime.now().strftime("%Y%m%d%H%M%S")
                pivot_name = f"ピボット_メニュー分類円_{data_ws.title}_{ts}"
                # Excelのシート名制限（31文字）に合わせる
                if len(pivot_name) > 31:
                    pivot_name = f"メニュー分類円_{data_ws.title}_{ts}"[:31]
                print(f"    📊 新しいシート作成: '{pivot_name}'")
                pivot_ws = wb.create_sheet(title=pivot_name)

                # ヘッダーを設定
                pivot_ws['A1'] = "メニュー分類"
                pivot_ws['B1'] = "件数"
                pivot_ws['C1'] = "割合(%)"

                # データを書き出し（件数順にソート）
                row = 2
                for category, count in sorted(menu_counts.items(), key=lambda x: x[1], reverse=True):
                    percentage = (count / total_count * 100) if total_count > 0 else 0
                    pivot_ws[f'A{row}'] = category
                    pivot_ws[f'B{row}'] = count
                    pivot_ws[f'C{row}'] = round(percentage, 1)
                    row += 1

                # 円グラフを作成
                chart = PieChart()
                chart.title = f"メニュー分類別割合"
                
                # データ範囲を正しく設定
                if row > 2:  # データが存在する場合
                    data = Reference(pivot_ws, min_col=2, min_row=1, max_row=row-1)
                    labels = Reference(pivot_ws, min_col=1, min_row=2, max_row=row-1)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(labels)
                else:
                    # データが少ない場合はサンプルデータを追加
                    pivot_ws['A2'] = "データなし"
                    pivot_ws['B2'] = 1
                    pivot_ws['C2'] = 100.0
                    data = Reference(pivot_ws, min_col=2, min_row=1, max_row=2)
                    labels = Reference(pivot_ws, min_col=1, min_row=2, max_row=2)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(labels)
                
                # データラベルを設定
                chart.dataLabels = DataLabelList()
                chart.dataLabels.showPercent = True
                chart.dataLabels.showCatName = True
                
                # 凡例を非表示にする
                chart.legend = None
                
                # グラフのサイズを設定（大きめに調整）
                chart.width = 20
                chart.height = 15
                
                # グラフをシートに追加
                pivot_ws.add_chart(chart, "A5")
                
                created_charts += 1
                print(f"    ✅ シート {i+1} のメニュー分類円グラフ作成完了")
                
            except Exception as e:
                print(f"    ❌ シート {i+1} の円グラフ作成エラー: {e}")
                continue

        print(f"📊 メニュー分類円グラフ作成完了: {created_charts}個")

        if created_charts == 0:
            return {
                "success": False,
                "error": "有効なデータが見つかりません"
            }

        # ファイルを保存
        print(f"🔍 デバッグ: ファイル保存開始: {target_path}")
        wb.save(str(target_path))
        print(f"🔍 デバッグ: ファイル保存完了")
        
        # S3にアップロード（フラグがTrueの場合）
        s3_result = None
        print(f"🔍 デバッグ: S3アップロードフラグ: {upload_to_s3_flag}")
        if upload_to_s3_flag:
            # 引数で渡されたバケット名を使用、なければ環境変数から取得
            bucket_name = s3_bucket or os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
            # バケット名のスペースを除去
            bucket_name = bucket_name.strip() if bucket_name else None
            upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
            print(f"🔍 デバッグ: バケット名: '{bucket_name}', プレフィックス: '{upload_prefix}'")
            
            if bucket_name:
                # 元のファイルと同じキーで上書き保存
                s3_key = original_s3_key if original_s3_key else f"{upload_prefix}{target_path.name}"
                print(f"🔍 デバッグ: S3キー: '{s3_key}'")
                
                s3_result = upload_to_s3(str(target_path), bucket_name, s3_key)
            else:
                s3_result = {"success": False, "error": "S3_BUCKET_NAME環境変数が設定されていません"}
        else:
            print("🔍 デバッグ: S3アップロードが無効化されています")
        
        return {
            "success": True,
            "message": f"更新完了: {target_path.name}",
            "created_charts": created_charts,
            "timestamp": datetime.now().strftime("%Y%m%d%H%M%S"),
            "s3_upload": s3_result
        }
        
    except Exception as e:
        print(f"❌ 処理失敗: {str(e)}")
        return {
            "success": False,
            "error": f"処理失敗: {str(e)}"
        }


def add_site_pivot_pie(target_path: Path, upload_to_s3_flag: bool = False, s3_bucket: str = None) -> dict:
    """
    サイト別の円グラフを作成する関数
    Lambda環境用に修正（win32comの代わりにopenpyxlを使用）
    """
    try:
        from openpyxl import load_workbook
        from openpyxl.chart import PieChart, Reference
        from openpyxl.chart.label import DataLabelList
    except ImportError:
        return {
            "success": False,
            "error": "openpyxl が未インストールです。`pip install openpyxl` を実行してください。"
        }

    try:
        # ファイルを読み込み
        wb = load_workbook(str(target_path))
        
        # データシートの取得
        data_ws = None
        try:
            data_ws = wb["データ"]
        except KeyError:
            data_ws = wb.worksheets[0]  # 最初のシート

        # データ範囲の取得
        max_row = data_ws.max_row
        max_col = data_ws.max_column
        
        # SITE_CODE列の検索
        site_code_col = None
        for col in range(1, max_col + 1):
            header_val = data_ws.cell(row=1, column=col).value
            if header_val == "SITE_CODE":
                site_code_col = col
                break
        
        if site_code_col is None:
            return {
                "success": False,
                "error": "SITE_CODE列が見つかりません"
            }

        # サイト別件数を収集
        site_counts = {}
        total_count = 0
        
        for row in range(2, max_row + 1):  # ヘッダー行を除く
            site_code = data_ws.cell(row=row, column=site_code_col).value
            if site_code:
                site_name = str(site_code)
                if site_name in site_counts:
                    site_counts[site_name] += 1
                else:
                    site_counts[site_name] = 1
                total_count += 1

        if not site_counts:
            return {
                "success": False,
                "error": "有効なデータが見つかりません"
            }

        # 新しいシートを作成
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        pivot_sheet_name = f"ピボット_サイト円_{timestamp}"
        pivot_ws = wb.create_sheet(title=pivot_sheet_name)

        # ヘッダーを設定
        pivot_ws['A1'] = "サイト名"
        pivot_ws['B1'] = "件数"
        pivot_ws['C1'] = "割合(%)"

        # データを書き出し（件数順にソート）
        row = 2
        for site_name, count in sorted(site_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_count * 100) if total_count > 0 else 0
            pivot_ws[f'A{row}'] = site_name
            pivot_ws[f'B{row}'] = count
            pivot_ws[f'C{row}'] = round(percentage, 1)
            row += 1

        # 円グラフを作成
        chart = PieChart()
        chart.title = "サイト別割合"
        
        # データ範囲を設定
        data = Reference(pivot_ws, min_col=2, min_row=1, max_row=row-1)
        labels = Reference(pivot_ws, min_col=1, min_row=2, max_row=row-1)
        
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(labels)
        
        # データラベルを設定
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showCatName = True
        
        # グラフのサイズを設定（メニュー名と被らないよう大きく調整）
        chart.width = 25
        chart.height = 18
        
        # グラフをシートに追加（右側に配置してメニュー名と被らないように）
        pivot_ws.add_chart(chart, "F3")

        # ファイルを保存
        wb.save(str(target_path))
        
        # S3にアップロード（フラグがTrueの場合）
        s3_result = None
        if upload_to_s3_flag:
            # 引数で渡されたバケット名を使用、なければ環境変数から取得
            bucket_name = s3_bucket or os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
            upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
            
            if bucket_name:
                # S3キーを生成
                file_name = target_path.name
                s3_key = f"{upload_prefix}site_pivot_{timestamp}_{file_name}"
                
                s3_result = upload_to_s3(str(target_path), bucket_name, s3_key)
            else:
                s3_result = {"success": False, "error": "S3_BUCKET_NAME環境変数が設定されていません"}
        
        return {
            "success": True,
            "message": f"更新完了: {target_path.name}",
            "sites": list(site_counts.keys()),
            "total_count": total_count,
            "timestamp": timestamp,
            "s3_upload": s3_result
        }
        
    except Exception as e:
        return {
            "success": False,
            "error": f"処理失敗: {str(e)}"
        }


def process_excel_file(file_path: str, process_type: str = "menu", upload_to_s3: bool = False, s3_bucket: str = None) -> dict:
    """
    Excelファイルを処理する関数
    """
    try:
        # file_pathが文字列であることを確認
        if not isinstance(file_path, str):
            print(f"❌ エラー: file_pathが文字列ではありません: {type(file_path)}")
            return {
                "success": False,
                "error": f"file_pathの型エラー: {type(file_path)}"
            }
        
        # file_pathが文字列であることを確認
        if not isinstance(file_path, str):
            print(f"❌ エラー: file_pathが文字列ではありません: {type(file_path)}")
            return {
                "success": False,
                "error": f"file_pathの型エラー: {type(file_path)}"
            }
        
        target_path = Path(file_path)
        
        if not target_path.exists():
            return {
                "success": False,
                "error": f"ファイルが見つかりません: {file_path}"
            }
        
        if process_type == "menu":
            return add_pivot_pie(target_path, upload_to_s3, s3_bucket)
        elif process_type == "site":
            return add_site_pivot_pie(target_path, upload_to_s3, s3_bucket)
        else:
            return {
                "success": False,
                "error": f"不明な処理タイプ: {process_type}"
            }
            
    except Exception as e:
        return {
            "success": False,
            "error": f"ファイル処理でエラー: {str(e)}"
        }


def _list_last_two_months_files(s3, bucket: str, prefix: str, process_type: str = "menu") -> List[Dict[str, str]]:
    """5月・6月分の統合データファイルを検索する関数"""
    # 5月・6月の統合データファイルを検索
    target_months = [5, 6]  # 5月と6月
    pattern = re.compile(r"(\d+)月_統合データ\.xlsx$")
    paginator = s3.get_paginator("list_objects_v2")
    by_month: Dict[int, List[Dict[str, str]]] = {}
    
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            name = key.split("/")[-1]
            m = pattern.search(name)
            if m:
                month = int(m.group(1))
                # 5月・6月のみを対象とする
                if month in target_months:
                    by_month.setdefault(month, []).append({"bucket": bucket, "key": key, "name": name})
    
    if not by_month:
        print("⚠️ 5月・6月の統合データファイルが見つかりませんでした")
        return []
    
    # 5月・6月の順で結果を返す
    result: List[Dict[str, str]] = []
    for m in target_months:
        if m in by_month:
            files = by_month[m]
            result.extend(files)
            print(f"✅ {m}月の統合データファイル: {len(files)}件")
        else:
            print(f"⚠️ {m}月の統合データファイルが見つかりませんでした")
    
    return result


def lambda_handler(event, context):
    """S3上の直近2か月分のExcelファイルを自動検索して処理し、結果をS3に保存"""
    try:
        print(f"Lambda関数が呼び出されました。イベント: {json.dumps(event, ensure_ascii=False)}")
        # オプション: 明示的にmain()を呼びたい場合（ローカル処理フロー）
        if event and isinstance(event, dict) and event.get("invoke_main"):
            print("🔧 invoke_main フラグを検知。main() を実行します。")
            exit_code = main()
            status = 200 if exit_code == 0 else 500
            return {
                'statusCode': status,
                'body': json.dumps({
                    'success': exit_code == 0,
                    'mode': 'main',
                    'exit_code': exit_code,
                }, ensure_ascii=False)
            }
        
        # 環境変数から取得
        s3_bucket = os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
        # バケット名のスペースを除去
        s3_bucket = s3_bucket.strip() if s3_bucket else None
        s3_prefix = os.environ.get('S3_OUTPUT_PREFIX', 'bigquery-data/')
        upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
        process_type = os.environ.get('PROCESS_TYPE', 'menu')
        
        print(f"🔍 デバッグ: 環境変数 - バケット: '{s3_bucket}', プレフィックス: '{s3_prefix}', アップロードプレフィックス: '{upload_prefix}'")
        
        if not s3_bucket:
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "s3_bucketが未指定"}, ensure_ascii=False)}
        
        s3_client = get_s3_client()
        if not s3_client:
            return {"statusCode": 500, "body": json.dumps({"success": False, "error": "S3クライアントの作成に失敗"}, ensure_ascii=False)}
        
        # Step Functions用の入力形式に対応
        input_files = event.get('input_files', event.get('input', event.get('uploaded', [])))
        
        # 入力ファイルが指定されていない場合は、直近2か月分のファイルを自動検索
        if not input_files:
            print(f"🔍 直近2か月分のファイルを検索中...")
            print(f"📁 Lambda関数: S3バケット '{s3_bucket}' からファイルを検索中...")
            target_files = _list_last_two_months_files(s3_client, s3_bucket, s3_prefix, process_type)
            
            if not target_files:
                print("❌ 対象ファイルが見つかりません")
                return {"statusCode": 404, "body": json.dumps({"success": False, "error": "対象ファイルが見つかりません"}, ensure_ascii=False)}
            
            print(f"✅ {len(target_files)}個のファイルが見つかりました")
            # 自動検索結果をinput_files形式に変換
            input_files = [f"s3://{s3_bucket}/{t['key']}" for t in target_files]
        
        print(f"🔍 デバッグ: 入力ファイル数: {len(input_files)}")
        
        # 処理対象ファイルを準備検索する
        files_to_process = []
        for file_path in input_files:
            if isinstance(file_path, str) and file_path.startswith('s3://'):
                # S3パスの正規化
                if 'your-bucket' in file_path:
                    file_path = file_path.replace('your-bucket', s3_bucket)
                file_path = file_path.strip()
                
                # S3パスを解析
                if not file_path.startswith('s3://'):
                    print(f"⚠️ 無効なS3パス: {file_path}")
                    continue
                    
                # s3://bucket/key を解析
                path_parts = file_path[5:].split('/', 1)
                if len(path_parts) != 2:
                    print(f"⚠️ 無効なS3パス形式: {file_path}")
                    continue
                    
                bucket_name = path_parts[0].strip()
                s3_key = path_parts[1].strip()
                
                # バケット名の検証
                if not re.match(r'^[a-zA-Z0-9.\-_]{1,255}$', bucket_name):
                    print(f"⚠️ 無効なバケット名: {bucket_name}")
                    continue
                
                filename = s3_key.split('/')[-1]
                print(f"🔍 デバッグ: ファイル名抽出 - s3_key: {s3_key}, filename: {filename}, 型: {type(filename)}")
                
                file_info = {
                    's3_path': file_path,
                    'bucket': bucket_name,
                    'key': s3_key,
                    'filename': filename
                }
                print(f"🔍 デバッグ: 追加するfile_info: {file_info}")
                print(f"🔍 デバッグ: file_infoの型: {type(file_info)}")
                files_to_process.append(file_info)
        
        if not files_to_process:
            print("❌ 処理対象ファイルがありません")
            return {"statusCode": 400, "body": json.dumps({"success": False, "error": "処理対象ファイルがありません"}, ensure_ascii=False)}
        
        print(f"🎯 合計 {len(files_to_process)} ファイルを処理します")
        print(f"🔍 デバッグ: files_to_processの型: {type(files_to_process)}")
        print(f"🔍 デバッグ: files_to_processの長さ: {len(files_to_process)}")
        
        for i, file_info in enumerate(files_to_process):
            print(f"  [{i+1}] file_info: {file_info}")
            print(f"  [{i+1}] file_infoの型: {type(file_info)}")
            if isinstance(file_info, dict) and 'filename' in file_info:
                print(f"  [{i+1}] filename: {file_info['filename']}")
            else:
                print(f"  [{i+1}] filenameが見つかりません")
        
        # files_to_processの詳細を確認
        print(f"🔍 デバッグ: files_to_processの内容:")
        for i, file_info in enumerate(files_to_process):
            print(f"  [{i+1}] file_info: {file_info}")
            print(f"  [{i+1}] file_infoの型: {type(file_info)}")
            if isinstance(file_info, dict):
                print(f"  [{i+1}] file_infoのキー: {list(file_info.keys())}")
                if 'filename' in file_info:
                    print(f"  [{i+1}] filename: {file_info['filename']}, 型: {type(file_info['filename'])}")
            print()
        
        processed_files = []
        uploaded_files = []
        errors = []
        
        for i, file_info in enumerate(files_to_process, 1):
            print(f"\n--- 処理 {i}/{len(files_to_process)} ---")
            print(f"🔍 デバッグ: ループ開始 - file_info: {file_info}")
            print(f"🔍 デバッグ: file_infoの型: {type(file_info)}")
            
            # file_infoが辞書でない場合はエラー
            if not isinstance(file_info, dict):
                print(f"❌ エラー: file_infoが辞書ではありません: {type(file_info)}")
                errors.append(f"file_infoの型エラー: {type(file_info)}")
                continue
            
            print(f"🎯 処理開始: {file_info['filename']}")
            print(f"処理中: {file_info['s3_path']}")
            
            # ローカルファイル名を生成
            print(f"🔍 デバッグ: file_info: {file_info}")
            print(f"🔍 デバッグ: file_infoの型: {type(file_info)}")
            print(f"🔍 デバッグ: file_infoのキー: {list(file_info.keys()) if isinstance(file_info, dict) else 'N/A'}")
            print(f"🔍 デバッグ: file_info['filename']: {file_info['filename']}")
            print(f"🔍 デバッグ: file_info['filename']の型: {type(file_info['filename'])}")
            print(f"🔍 デバッグ: file_info['filename']の内容: {repr(file_info['filename'])}")
            
            # filenameが文字列でない場合はエラー
            if not isinstance(file_info['filename'], str):
                print(f"❌ エラー: filenameが文字列ではありません: {type(file_info['filename'])}")
                errors.append(f"filenameの型エラー: {type(file_info['filename'])}")
                continue
            
            local_path = Path("/tmp") / file_info['filename']
            print(f"🔍 デバッグ: local_path作成後: {local_path}, 型: {type(local_path)}")
            
            try:
                # 直接S3からダウンロード
                s3_client.download_file(file_info['bucket'], file_info['key'], str(local_path))
                print(f"ダウンロード成功: {local_path}")
                
                # ファイル処理
                print(f"🔍 デバッグ: add_pivot_pie呼び出し - ファイル: {local_path}, 型: {type(local_path)}, S3アップロード: True, バケット: {s3_bucket}")
                print(f"🔍 デバッグ: local_pathが存在するか: {local_path.exists()}")
                # local_pathがPathオブジェクトであることを確認
                if not isinstance(local_path, Path):
                    print(f"❌ エラー: local_pathがPathオブジェクトではありません: {type(local_path)}")
                    errors.append(f"ファイルパスの型エラー: {type(local_path)}")
                    continue
                
                print(f"🔍 デバッグ: add_pivot_pie呼び出し前 - local_path: {local_path}, 型: {type(local_path)}")
                result = add_pivot_pie(local_path, True, s3_bucket, file_info['key'])
                print(f"🔍 デバッグ: 処理結果: {result}")
                
                if result.get('success', False):
                    processed_files.append(file_info['s3_path'])
                    print(f"✅ 処理完了: {file_info['filename']}")
                    print(f"   カテゴリ数: {len(result.get('categories', []))}")
                    print(f"   総件数: {result.get('total_count', 0)}")
                    if 's3_upload' in result and result['s3_upload'] and result['s3_upload'].get('success', False):
                        uploaded_files.append(result['s3_upload']['s3_path'])
                        print(f"✅ S3アップロード成功: {result['s3_upload']['s3_path']}")
                    else:
                        s3_error = result.get('s3_upload', {}).get('error', 'S3アップロードが実行されませんでした')
                        print(f"⚠️ S3アップロード失敗: {s3_error}")
                else:
                    print(f"❌ 処理失敗: {file_info['filename']} - {result.get('error', '不明なエラー')}")
                    errors.append(f"ファイル処理に失敗: {file_info['s3_path']} - {result.get('error', '不明なエラー')}")
                    
            except Exception as e:
                errors.append(f"S3からのダウンロードに失敗: {file_info['s3_path']} - {str(e)}")
        
        # 処理結果のサマリー
        print(f"\n=== 処理完了サマリー ===")
        print(f"✅ 成功: {len(processed_files)} ファイル")
        print(f"❌ 失敗: {len(errors)} ファイル")
        print(f"📊 合計: {len(files_to_process)} ファイル")
        
        # 結果を返す
        status = 207 if errors and processed_files else (200 if processed_files else 500)
        result = {
            'statusCode': status,
            'body': json.dumps({
                'success': bool(processed_files),
                'processed': processed_files,
                'uploaded': uploaded_files,
                'count': len(processed_files),
                'errors': errors
            }, ensure_ascii=False)
        }
        
        print(f"🔍 デバッグ: 最終結果 - 処理済み: {len(processed_files)}, アップロード済み: {len(uploaded_files)}, エラー: {len(errors)}")
        print(f"🔍 デバッグ: アップロード済みファイル: {uploaded_files}")
        
        return result
            
    except Exception as e:
        print(f"Lambda関数でエラーが発生: {str(e)}")
        return {
            'statusCode': 500,
            'body': json.dumps({
                'error': f'予期しないエラーが発生しました: {str(e)}'
            }, ensure_ascii=False)
        }


def main() -> int:
    """メニュー分類別の円グラフを作成するメイン関数"""
    try:
        # bigquery-dataディレクトリ内のファイルを自動検索
        bigquery_data_dir = Path("bigquery-data")
        
        if not bigquery_data_dir.exists():
            print("❌ bigquery-dataディレクトリが見つかりません")
            return 1
        
        # Excelファイルを検索
        excel_files = list(bigquery_data_dir.glob("*.xlsx"))
        
        if not excel_files:
            print("❌ bigquery-dataディレクトリにExcelファイルが見つかりません")
            return 1
        
        # その他ファイルを優先的に処理
        target_files = []
        for file in excel_files:
            if "その他" in file.name:
                target_files.insert(0, file)  # 先頭に追加
            else:
                target_files.append(file)
        
        print(f"📁 処理対象ファイル: {[f.name for f in target_files]}")
        print(f"🎯 合計 {len(target_files)} ファイルを処理します")
        
        # 全てのファイルを処理
        success_count = 0
        error_count = 0
        
        for i, target in enumerate(target_files, 1):
            print(f"\n--- 処理 {i}/{len(target_files)} ---")
            print(f"🎯 処理開始: {target.name}")
            
            try:
                result = process_excel_file(str(target), "menu", False, None)
                if result.get('success', False):
                    print(f"✅ 処理完了: {target.name}")
                    print(f"   カテゴリ数: {len(result.get('categories', []))}")
                    print(f"   総件数: {result.get('total_count', 0)}")
                    success_count += 1
                else:
                    print(f"❌ 処理失敗: {target.name} - {result.get('error', '不明なエラー')}")
                    error_count += 1
            except Exception as e:
                print(f"❌ 処理中にエラー: {target.name} - {e}")
                error_count += 1
        
        # 処理結果のサマリー
        print(f"\n=== 処理完了サマリー ===")
        print(f"✅ 成功: {success_count} ファイル")
        print(f"❌ 失敗: {error_count} ファイル")
        print(f"📊 合計: {len(target_files)} ファイル")
        
        return 0 if error_count == 0 else 1
        
    except Exception as e:
        print(f"❌ メイン処理でエラー: {e}")
        return 1


def show_help():
    """使用方法を表示"""
    print("""
📊 メニュー分類別円グラフ作成ツール (openpyxl版)

使用方法:
  python menu_copy.py                    # bigquery-dataディレクトリ内の全ファイルを処理
  python menu_copy.py ファイル名.xlsx     # 特定のファイルを処理
  python menu_copy.py --help             # このヘルプを表示

例:
  python menu_copy.py 7月_初回受付_その他.xlsx
  python menu_copy.py "6月_再受付_hp.xlsx"
  python menu_copy.py bigquery-data/7月_初回受付_hp.xlsx

注意:
  - ファイル名のみ指定した場合、bigquery-dataディレクトリ内を検索します
  - 絶対パスまたは相対パスで指定することも可能です
  - ファイル名にスペースが含まれる場合は引用符で囲んでください
  - openpyxlライブラリが必要です: pip install openpyxl
""")


if __name__ == "__main__":
    try:
        # ヘルプ表示
        if len(sys.argv) > 1 and sys.argv[1] in ["--help", "-h", "-?"]:
            show_help()
            raise SystemExit(0)
        
        # コマンドライン引数でファイル名が指定されているかチェック
        if len(sys.argv) > 1:
            # 指定されたファイルを処理
            specified_file = sys.argv[1]
            target = Path(specified_file)
            
            # 相対パスの場合はbigquery-dataディレクトリ内を検索
            if not target.is_absolute():
                bigquery_data_dir = Path("bigquery-data")
                if bigquery_data_dir.exists():
                    target = bigquery_data_dir / specified_file
                else:
                    print("❌ bigquery-dataディレクトリが見つかりません")
                    raise SystemExit(1)
            
            if not target.exists():
                print(f"❌ 指定されたファイルが見つかりません: {target}")
                print("💡 ヒント: ファイル名が正しいか確認してください")
                print("💡 ヒント: python menu_copy.py --help で使用方法を確認")
                raise SystemExit(1)
            
            print(f"🎯 指定されたファイルを処理: {target.name}")
            result = process_excel_file(str(target), "menu", False, None)
            if result.get('success', False):
                print(f"✅ 処理完了: {target.name}")
                print(f"   カテゴリ数: {len(result.get('categories', []))}")
                print(f"   総件数: {result.get('total_count', 0)}")
                raise SystemExit(0)
            else:
                print(f"❌ 処理失敗: {target.name} - {result.get('error', '不明なエラー')}")
                raise SystemExit(1)
        else:
            # ファイル名が指定されていない場合は、bigquery-dataディレクトリ内のファイルを自動検索
            exit_code = main()
            if exit_code != 0:
                print(f"⚠️ 処理が正常に完了しませんでした（終了コード: {exit_code}）")
            raise SystemExit(exit_code)
    except KeyboardInterrupt:
        print("\n⚠️ 処理が中断されました")
        raise SystemExit(1)
    except Exception as e:
        print(f"❌ 予期しないエラー: {e}")
        raise SystemExit(1)
