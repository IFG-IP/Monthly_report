from pathlib import Path
import os
import json
import csv
from typing import Optional, Tuple, List, Dict
from datetime import datetime
import io
import re


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
    try:
        from zoneinfo import ZoneInfo
        JST = ZoneInfo("Asia/Tokyo")
    except:
        # フォールバック: UTC+9を使用
        from datetime import timezone, timedelta
        JST = timezone(timedelta(hours=9))

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


def filter_last_two_months_data(data: List[Dict]) -> List[Dict]:
    """appoint_timestampを基準に直近2か月のデータをフィルタリング"""
    if not data:
        return data
    
    try:
        from zoneinfo import ZoneInfo
        JST = ZoneInfo("Asia/Tokyo")
    except:
        from datetime import timezone, timedelta
        JST = timezone(timedelta(hours=9))
    
    # 現在時刻をJSTで取得
    now = datetime.now(JST)
    
    # 直近2か月の開始日を計算（月単位で正確に）
    current_year = now.year
    current_month = now.month
    
    # 2か月前の月を計算
    if current_month <= 2:
        two_months_ago_year = current_year - 1
        two_months_ago_month = current_month + 10  # 12 + (month - 2)
    else:
        two_months_ago_year = current_year
        two_months_ago_month = current_month - 2
    
    # 2か月前の月の1日を開始日とする
    two_months_ago = datetime(two_months_ago_year, two_months_ago_month, 1, tzinfo=JST)
    
    print(f"🔍 フィルタリング条件: {two_months_ago.strftime('%Y年%m月')} 以降")
    
    filtered_data = []
    for row in data:
        # appoint_timestampを取得
        timestamp_str = row.get('appoint_timestamp', '')
        if not timestamp_str:
            continue
            
        # 日時をパース
        dt = parse_datetime(timestamp_str)
        if dt is None:
            continue
            
        # JSTに変換
        if dt.tzinfo:
            dt = dt.astimezone(JST)
        else:
            dt = dt.replace(tzinfo=JST)
        
        # 2か月以内かチェック
        if dt >= two_months_ago:
            filtered_data.append(row)
    
    print(f"📊 フィルタリング結果: {len(data)}件 → {len(filtered_data)}件")
    return filtered_data


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

def add_charts_to_excel(excel_path: Path) -> Dict:
    """
    統合されたExcelファイルにグラフを追加する関数
    site_copy.pyとmenu_copy.pyの機能を統合
    """
    try:
        print(f"📊 グラフ追加開始: {excel_path}")
        print(f"🔍 ファイル存在確認: {excel_path.exists()}")
        
        # ファイルの内容を確認
        if excel_path.exists():
            from openpyxl import load_workbook
            wb = load_workbook(str(excel_path))
            print(f"📋 既存シート一覧: {wb.sheetnames}")
            wb.close()
        
        # 修正された機能をインポート
        print("🔍 site_copy.pyをインポート中...")
        try:
            from site_copy import _add_site_pie_inplace, _classify_site_code
            print("✅ site_copy.pyインポート成功")
        except Exception as e:
            print(f"❌ site_copy.pyインポート失敗: {e}")
            return {"success": False, "error": f"site_copy.pyインポート失敗: {e}"}
        
        print("🔍 menu_copy.pyをインポート中...")
        try:
            from menu_copy import add_pivot_pie, extract_bracket_content
            print("✅ menu_copy.pyインポート成功")
        except Exception as e:
            print(f"❌ menu_copy.pyインポート失敗: {e}")
            return {"success": False, "error": f"menu_copy.pyインポート失敗: {e}"}
        
        results = {
            "site_chart": {"success": False, "error": None},
            "menu_chart": {"success": False, "error": None}
        }
        
        # サイト分類別円グラフを追加
        try:
            print("  🥧 サイト分類別円グラフを追加中...")
            site_result = _add_site_pie_inplace(excel_path, "サイト分類別割合")
            print(f"  📊 サイトグラフ結果: {site_result}")
            results["site_chart"] = site_result
            if site_result["success"]:
                print("  ✅ サイト分類別円グラフ追加完了")
            else:
                print(f"  ❌ サイト分類別円グラフ追加失敗: {site_result['error']}")
        except Exception as e:
            results["site_chart"] = {"success": False, "error": str(e)}
            print(f"  ❌ サイト分類別円グラフ追加エラー: {e}")
            import traceback
            traceback.print_exc()
        
        # メニュー分類別円グラフを追加
        try:
            print("  🥧 メニュー分類別円グラフを追加中...")
            menu_result = add_pivot_pie(excel_path, False, None)
            print(f"  📊 メニューグラフ結果: {menu_result}")
            results["menu_chart"] = menu_result
            if menu_result["success"]:
                print("  ✅ メニュー分類別円グラフ追加完了")
            else:
                print(f"  ❌ メニュー分類別円グラフ追加失敗: {menu_result['error']}")
        except Exception as e:
            results["menu_chart"] = {"success": False, "error": str(e)}
            print(f"  ❌ メニュー分類別円グラフ追加エラー: {e}")
            import traceback
            traceback.print_exc()
        
        # 最終的なシート一覧を確認
        print("🔍 最終的なシート一覧を確認中...")
        try:
            from openpyxl import load_workbook
            wb = load_workbook(str(excel_path))
            print(f"📋 最終シート一覧: {wb.sheetnames}")
            wb.close()
        except Exception as e:
            print(f"⚠️ シート一覧確認エラー: {e}")
        
        # 結果をまとめる
        success_count = sum(1 for result in results.values() if result["success"])
        total_count = len(results)
        
        return {
            "success": success_count > 0,
            "message": f"{success_count}/{total_count} 個のグラフを追加しました",
            "results": results,
            "excel_path": str(excel_path)
        }
        
    except Exception as e:
        print(f"❌ グラフ追加処理エラー: {e}")
        import traceback
        traceback.print_exc()
        return {
            "success": False,
            "error": str(e),
            "excel_path": str(excel_path)
        }


def _get_target_months(data: List[Dict]) -> List[Tuple[int, int]]:
    """直近２か月の対象月を取得する共通関数"""
    # 直近2か月の取得ロジックをコメントアウト
    # from zoneinfo import ZoneInfo
    # current_date = datetime.now(ZoneInfo("Asia/Tokyo"))
    # target_months = []
    # 
    # for i in range(2):
    #     if current_date.month - (i + 1) <= 0:
    #         year = current_date.year - 1
    #         month = 12 + (current_date.month - (i + 1))
    #     else:
    #         year = current_date.year
    #         month = current_date.month - (i + 1)
    #     target_months.append((year, month))
    
    # データから直近2か月の対象月を動的に取得
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
    
    # 直近2か月の開始日を計算（月単位で正確に）
    current_year = now.year
    current_month = now.month
    
    # 2か月前の月を計算
    if current_month <= 2:
        two_months_ago_year = current_year - 1
        two_months_ago_month = current_month + 10  # 12 + (month - 2)
    else:
        two_months_ago_year = current_year
        two_months_ago_month = current_month - 2
    
    # 2か月前の月の1日を開始日とする
    two_months_ago = datetime(two_months_ago_year, two_months_ago_month, 1, tzinfo=JST)
    
    print(f"🔍 対象期間: {two_months_ago.strftime('%Y年%m月')} 以降")
    
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
        
        # 2か月前の月以降のデータのみを対象
        if dt >= two_months_ago:
            months_found.add((dt.year, dt.month))
    
    # 年月でソートして返す
    return sorted(list(months_found))


def _generate_outputs(data: List[Dict], results_dir: Path) -> List[Path]:
    # outputs を最初に初期化
    outputs: List[Path] = []

    # 対象月を取得
    target_months = _get_target_months(data)

    """フィルタしたデータを月別・カテゴリ別に Excel 出力（月ごとに1つのファイル）"""
    if not data:
        print("⚠️ データが空です")
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
    
    results_dir.mkdir(parents=True, exist_ok=True)
    
    for target_year, target_month in target_months:
        month_name = f"{target_month}月"
        filtered_data = []
        month_matches = 0
        
        # その月のデータをフィルタ
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

        # その月のデータがない場合はスキップ
        if not filtered_data:
            print(f"⚠️ {month_name}: データがありません")
            continue

        # カテゴリごとにグループ化
        grouped_data: Dict[str, List[Dict]] = {}
        for category, row in filtered_data:
            grouped_data.setdefault(category, []).append(row)

        # 月ごとに1つのExcelファイルを作成
        filename = f"{month_name}_統合データ.xlsx"
        out_path = results_dir / filename
        
        # 新しいワークブックを作成
        from openpyxl import Workbook
        wb = Workbook()
        
        # デフォルトシートを削除
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        # 各カテゴリを別々のシートに追加
        for category, rows in grouped_data.items():
            if rows:  # データがある場合のみシートを作成
                # シート名を生成（Excelの制限に合わせて調整）
                sheet_name = f"{month_name}_{category}"
                if len(sheet_name) > 31:  # Excelのシート名制限
                    sheet_name = sheet_name[:31]
                
                ws = wb.create_sheet(title=sheet_name)
                
                # ヘッダーを設定
                if rows:
                    headers = list(rows[0].keys())
                    for col, header in enumerate(headers, 1):
                        ws.cell(row=1, column=col, value=header)
                    
                    # データを書き込み
                    for row_idx, row_data in enumerate(rows, 2):
                        for col_idx, header in enumerate(headers, 1):
                            ws.cell(row=row_idx, column=col_idx, value=row_data.get(header))
                
                print(f"  ✅ シート作成: {sheet_name} ({len(rows)}件)")
        
        # ファイルを保存
        wb.save(str(out_path))
        
        # グラフを追加
        print(f"📊 {month_name}にグラフを追加中...")
        chart_result = add_charts_to_excel(out_path)
        if chart_result["success"]:
            print(f"✅ {month_name}グラフ追加完了: {chart_result['message']}")
        else:
            print(f"⚠️ {month_name}グラフ追加に問題: {chart_result.get('error', '不明なエラー')}")
        
        outputs.append(out_path)
        print(f"✅ {month_name}統合ファイル作成完了: {filename} ({len(grouped_data)}シート + グラフ)")

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

    # 直近2か月のデータをフィルタリング
    print(f"📊 全データ件数: {len(data)}件")
    filtered_data = filter_last_two_months_data(data)
    print(f"📊 フィルタリング後: {len(filtered_data)}件")
    
    if not filtered_data:
        return {"success": False, "error": "直近2か月のデータが見つかりませんでした"}

    # 出力先: Lambda なら /tmp、ローカルなら bigquery-data/
    out_dir = Path("/tmp" if os.environ.get("AWS_LAMBDA_FUNCTION_NAME") else "bigquery-data/")
    try:
        outputs = _generate_outputs(filtered_data, out_dir)
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
