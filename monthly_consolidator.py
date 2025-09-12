from pathlib import Path
import pandas as pd
import win32com.client as win32
from datetime import datetime
import sys
import re

def extract_month_from_filename(filename: str):
    """ファイル名から月を抽出"""
    # 月名のパターンを検索
    month_patterns = [
        r'(\d+)月',  # 数字+月
        r'(\d{1,2})月',  # 1-2桁の数字+月
    ]
    
    for pattern in month_patterns:
        match = re.search(pattern, filename)
        if match:
            return int(match.group(1))
    
    return None

def get_file_type_from_filename(filename: str):
    """ファイル名からタイプを判定"""
    if "hp" in filename.lower():
        return "HP"
    elif "その他" in filename:
        return "その他"
    else:
        return "不明"

def get_visit_type_from_filename(filename: str):
    """ファイル名から受付タイプを判定"""
    if "初回受付" in filename:
        return "初回受付"
    elif "再受付" in filename:
        return "再受付"
    else:
        return "不明"

def consolidate_monthly_data(results_dir: Path = Path("results")):
    """月ごとにデータを統合"""
    try:
        print("🔍 resultsフォルダ内のExcelファイルを分析中...")
        
        # Excelファイルを取得
        excel_files = list(results_dir.glob("*.xlsx"))
        if not excel_files:
            print("❌ Excelファイルが見つかりません")
            return
        
        print(f"📁 見つかったファイル数: {len(excel_files)}")
        
        # 月ごとにファイルをグループ化
        monthly_groups = {}
        
        for file_path in excel_files:
            month = extract_month_from_filename(file_path.name)
            if month is None:
                print(f"⚠️ 月を抽出できませんでした: {file_path.name}")
                continue
            
            if month not in monthly_groups:
                monthly_groups[month] = []
            
            monthly_groups[month].append(file_path)
        
        print(f"📅 対象月: {sorted(monthly_groups.keys())}")
        
        # 各月のデータを統合
        for month in sorted(monthly_groups.keys()):
            print(f"\n🔄 {month}月のデータを統合中...")
            
            files = monthly_groups[month]
            print(f"  ファイル数: {len(files)}")
            
            # 月別統合ファイルを作成
            create_monthly_consolidated_file(month, files)
        
        print(f"\n🎉 月別統合完了!")
        
    except Exception as e:
        print(f"❌ エラー: {e}")
        import traceback
        traceback.print_exc()

def create_monthly_consolidated_file(month: int, files: list):
    """月別統合ファイルを作成"""
    try:
        # 統合データを格納する辞書
        consolidated_data = {
            'HP_初回受付': [],
            'HP_再受付': [],
            'その他_初回受付': [],
            'その他_再受付': []
        }
        
        # 円グラフデータを格納する辞書
        pie_chart_data = {
            'メニュー分類': [],
            'サイト分類': []
        }
        
        # 各ファイルからデータを読み込み
        for file_path in files:
            print(f"  📊 読み込み中: {file_path.name}")
            
            try:
                # 全シートを読み込み
                all_sheets = pd.read_excel(file_path, sheet_name=None)
                
                # データシートを処理
                if 'データ' in all_sheets:
                    df = all_sheets['データ']
                    
                    if not df.empty:
                        # ファイルタイプと受付タイプを判定
                        file_type = get_file_type_from_filename(file_path.name)
                        visit_type = get_visit_type_from_filename(file_path.name)
                        
                        # データに列を追加
                        df['ファイル名'] = file_path.name
                        df['ファイルタイプ'] = file_type
                        df['受付タイプ'] = visit_type
                        df['月'] = month
                        
                        # カテゴリに分類
                        category = f"{file_type}_{visit_type}"
                        if category in consolidated_data:
                            consolidated_data[category].append(df)
                            print(f"    ✅ データ追加: {category}")
                        else:
                            print(f"    ⚠️ 未知のカテゴリ: {category}")
                
                # 円グラフシートを処理
                for sheet_name, sheet_df in all_sheets.items():
                    if "円" in sheet_name and not sheet_df.empty:
                        print(f"    📊 円グラフシート処理: {sheet_name}")
                        
                        # ファイルタイプと受付タイプを判定
                        file_type = get_file_type_from_filename(file_path.name)
                        visit_type = get_visit_type_from_filename(file_path.name)
                        
                        # 列名をクリーンアップ（Unnamed列を削除）
                        sheet_df = sheet_df.loc[:, ~sheet_df.columns.str.contains('^Unnamed')].copy()
                        
                        # シートデータに列を追加
                        sheet_df.loc[:, 'ファイル名'] = file_path.name
                        sheet_df.loc[:, 'ファイルタイプ'] = file_type
                        sheet_df.loc[:, '受付タイプ'] = visit_type
                        sheet_df.loc[:, '月'] = month
                        
                        # メニュー分類かサイト分類かを判定
                        if "メニュー" in sheet_name:
                            pie_chart_data['メニュー分類'].append(sheet_df)
                            print(f"    ✅ メニュー分類円グラフデータ追加")
                        elif "サイト" in sheet_name:
                            pie_chart_data['サイト分類'].append(sheet_df)
                            print(f"    ✅ サイト分類円グラフデータ追加")
                
            except Exception as e:
                print(f"    ❌ 読み込みエラー: {e}")
                continue
        
        # 統合ファイルを作成
        output_file = Path(f"results/{month}月_統合データ.xlsx")
        create_consolidated_excel(month, consolidated_data, pie_chart_data, output_file)
        
        # 円グラフを作成
        create_monthly_pie_charts(month, output_file)
        
        print(f"  ✅ {month}月統合ファイル作成完了: {output_file.name}")
        
    except Exception as e:
        print(f"❌ {month}月統合でエラー: {e}")

def create_consolidated_excel(month: int, consolidated_data: dict, pie_chart_data: dict, output_file: Path):
    """統合Excelファイルを作成"""
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 各カテゴリのデータを統合してシートに出力
            for category, dataframes in consolidated_data.items():
                if not dataframes:
                    continue
                
                # データフレームを結合
                combined_df = pd.concat(dataframes, ignore_index=True)
                
                # シート名を設定
                sheet_name = f"{category}_{month}月"
                
                # データを書き込み
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"    📊 シート作成: {sheet_name} ({len(combined_df)}行)")
            
            # 円グラフデータを統合してシートに出力
            for chart_type, dataframes in pie_chart_data.items():
                if not dataframes:
                    continue
                
                # データフレームを結合
                combined_df = pd.concat(dataframes, ignore_index=True)
                
                # データを整理・フォーマット
                if chart_type == "メニュー分類":
                    # メニュー分類データの整理
                    if 'MENU_NAME' in combined_df.columns:
                        # メニュー名でグループ化して件数を集計
                        menu_summary = combined_df.groupby(['MENU_NAME', 'ファイルタイプ', '受付タイプ']).size().reset_index(name='件数')
                        menu_summary = menu_summary.sort_values('件数', ascending=False)
                        combined_df = menu_summary
                elif chart_type == "サイト分類":
                    # サイト分類データの整理
                    if 'SITE_CODE' in combined_df.columns:
                        # サイトコードでグループ化して件数を集計
                        site_summary = combined_df.groupby(['SITE_CODE', 'ファイルタイプ', '受付タイプ']).size().reset_index(name='件数')
                        site_summary = site_summary.sort_values('件数', ascending=False)
                        combined_df = site_summary
                
                # シート名を設定
                sheet_name = f"{chart_type}_統合_{month}月"
                
                # データを書き込み
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"    📊 円グラフシート作成: {sheet_name} ({len(combined_df)}行)")
            
            # サマリーシートを作成
            create_summary_sheet(writer, month, consolidated_data, pie_chart_data)
        
        print(f"  📁 統合ファイル保存完了: {output_file}")
        
    except Exception as e:
        print(f"❌ Excelファイル作成エラー: {e}")

def create_summary_sheet(writer, month: int, consolidated_data: dict, pie_chart_data: dict):
    """サマリーシートを作成"""
    try:
        summary_data = []
        
        # 基本データのサマリー
        for category, dataframes in consolidated_data.items():
            if not dataframes:
                continue
            
            # データを結合
            combined_df = pd.concat(dataframes, ignore_index=True)
            
            # サマリー情報を計算
            total_records = len(combined_df)
            unique_menus = combined_df['MENU_NAME'].nunique() if 'MENU_NAME' in combined_df.columns else 0
            unique_sites = combined_df['SITE_CODE'].nunique() if 'SITE_CODE' in combined_df.columns else 0
            
            summary_data.append({
                'カテゴリ': category,
                'レコード数': total_records,
                'ユニークメニュー数': unique_menus,
                'ユニークサイト数': unique_sites,
                'データタイプ': '基本データ'
            })
        
        # 円グラフデータのサマリー
        for chart_type, dataframes in pie_chart_data.items():
            if not dataframes:
                continue
            
            # データを結合
            combined_df = pd.concat(dataframes, ignore_index=True)
            
            # サマリー情報を計算
            total_records = len(combined_df)
            unique_menus = combined_df['MENU_NAME'].nunique() if 'MENU_NAME' in combined_df.columns else 0
            unique_sites = combined_df['SITE_CODE'].nunique() if 'SITE_CODE' in combined_df.columns else 0
            
            summary_data.append({
                'カテゴリ': f"{chart_type}_統合",
                'レコード数': total_records,
                'ユニークメニュー数': unique_menus,
                'ユニークサイト数': unique_sites,
                'データタイプ': '円グラフデータ'
            })
        
        # サマリーデータフレームを作成
        summary_df = pd.DataFrame(summary_data)
        
        # サマリーシートに書き込み
        summary_df.to_excel(writer, sheet_name=f'サマリー_{month}月', index=False)
        
        print(f"    📊 サマリーシート作成: サマリー_{month}月")
        
    except Exception as e:
        print(f"❌ サマリーシート作成エラー: {e}")

def create_monthly_pie_charts(month: int, output_file: Path):
    """月別統合ファイルに既存の円グラフをコピー"""
    try:
        print(f"🔄 {month}月統合ファイルに既存の円グラフをコピー中...")
        
        # Excelアプリケーションを起動
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = None
        try:
            wb = excel.Workbooks.Open(str(output_file.resolve()))
            
            # 各シートに既存の円グラフをコピー
            for sheet in wb.Worksheets:
                if sheet.Name.startswith('HP_') or sheet.Name.startswith('その他_'):
                    copy_existing_pie_charts_to_sheet(sheet, month)
            
            try:
                wb.Close(SaveChanges=True)
                print(f"  ✅ {month}月統合ファイルに円グラフをコピーしました")
            except:
                print(f"  ⚠️ 統合ファイルのクローズでエラー（無視）")
            
        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=True)
                except:
                    pass
            try:
                excel.Quit()
            except:
                pass
            
    except Exception as e:
        print(f"❌ 円グラフコピーエラー: {e}")
        import traceback
        traceback.print_exc()

def copy_existing_pie_charts_to_sheet(sheet, month: int):
    """既存のExcelファイルから円グラフをコピーしてシートに配置"""
    try:
        print(f"    📊 既存の円グラフをコピー中: {sheet.Name}")
        
        # シートタイプを判定
        is_hp_sheet = "HP" in sheet.Name
        
        # 対応する元ファイルを探す
        results_dir = Path("results")
        source_files = []
        
        if is_hp_sheet:
            # HPシートの場合
            if "初回受付" in sheet.Name:
                search_pattern = f"{month}月_初回受付_hp.xlsx"
                source_files = list(results_dir.glob(search_pattern))
                print(f"    🔍 HP初回受付検索: {search_pattern}")
            elif "再受付" in sheet.Name:
                search_pattern = f"{month}月_再受付_hp.xlsx"
                source_files = list(results_dir.glob(search_pattern))
                print(f"    🔍 HP再受付検索: {search_pattern}")
                print(f"    📊 検索結果: {len(source_files)}個のファイルが見つかりました")
        else:
            # その他シートの場合
            if "初回受付" in sheet.Name:
                source_files = list(results_dir.glob(f"{month}月_初回受付_その他.xlsx"))
            elif "再受付" in sheet.Name:
                source_files = list(results_dir.glob(f"{month}月_再受付_その他.xlsx"))
        
        if not source_files:
            print(f"    ⚠️ 対応する元ファイルが見つかりません: {sheet.Name}")
            print(f"    🔍 検索パターン: {month}月_*_hp.xlsx または {month}月_*_その他.xlsx")
            print(f"    🔍 実際の検索パターン: {month}月_再受付_hp.xlsx")
            # resultsディレクトリのファイル一覧を表示
            all_files = list(results_dir.glob("*.xlsx"))
            print(f"    📁 resultsディレクトリ内のファイル:")
            for f in all_files:
                print(f"      - {f.name}")
            return
        
        source_file = source_files[0]
        print(f"    📁 元ファイル: {source_file.name}")
        
        # 元ファイルを開く
        source_wb = None
        try:
            source_wb = sheet.Parent.Application.Workbooks.Open(str(source_file.resolve()))
            
            # 元ファイルから円グラフを探してコピー
            chart_found = False
            for source_sheet in source_wb.Worksheets:
                print(f"    🔍 シート確認: {source_sheet.Name}")
                if "円" in source_sheet.Name or "ピボット" in source_sheet.Name:
                    print(f"    📊 円グラフシート発見: {source_sheet.Name}")
                    
                    # 円グラフをコピー
                    for shape in source_sheet.Shapes:
                        if shape.Type == 3:  # msoChart
                            print(f"    📈 円グラフをコピー中: {shape.Name}")
                            
                            # HPシートの場合はサイト別円グラフをスキップ
                            if "サイト" in source_sheet.Name and is_hp_sheet:
                                print(f"    ⚠️ HPシートのためサイト別円グラフはスキップ: {sheet.Name}")
                                continue
                            
                            # 円グラフをコピー
                            shape.Copy()
                            
                            # 統合シートに貼り付け
                            target_sheet = sheet
                            target_sheet.Paste()
                            
                            # 貼り付けたグラフの位置を調整
                            pasted_shape = target_sheet.Shapes(target_sheet.Shapes.Count)
                            
                            if "メニュー" in source_sheet.Name:
                                # メニュー別円グラフの位置（データの右側）
                                pasted_shape.Left = target_sheet.Range("T1").Left
                                pasted_shape.Top = target_sheet.Range("D1").Top
                                print(f"    ✅ メニュー別円グラフをコピー: {sheet.Name}")
                                chart_found = True
                            elif "サイト" in source_sheet.Name and not is_hp_sheet:
                                # サイト別円グラフの位置（メニューグラフの下に配置、HPシート以外のみ）
                                pasted_shape.Left = target_sheet.Range("T1").Left
                                pasted_shape.Top = target_sheet.Range("D1").Top + 650  # メニューグラフの下に配置
                                print(f"    ✅ サイト別円グラフをコピー: {sheet.Name}")
                                chart_found = True
                            elif "ピボット_メニュー" in source_sheet.Name:
                                # ピボットメニューシートの場合
                                pasted_shape.Left = target_sheet.Range("T1").Left
                                pasted_shape.Top = target_sheet.Range("D1").Top
                                print(f"    ✅ ピボットメニュー円グラフをコピー: {sheet.Name}")
                                chart_found = True
                            elif "ピボット_サイト" in source_sheet.Name and not is_hp_sheet:
                                # ピボットサイトシートの場合（HPシート以外のみ）
                                pasted_shape.Left = target_sheet.Range("T1").Left
                                pasted_shape.Top = target_sheet.Range("D1").Top + 650  # メニューグラフの下に配置
                                print(f"    ✅ ピボットサイト円グラフをコピー: {sheet.Name}")
                                chart_found = True
                            elif "ピボット_サイト" in source_sheet.Name and is_hp_sheet:
                                # HPシートにはピボットサイト円グラフをコピーしない
                                print(f"    ⚠️ HPシートのためピボットサイト円グラフはスキップ: {sheet.Name}")
                                # 貼り付けたグラフを削除
                                pasted_shape.Delete()
                                continue
                            
                            # サイズ調整
                            pasted_shape.Width = 700
                            pasted_shape.Height = 500
                            
                            print(f"    ✅ 円グラフコピー完了: {shape.Name}")
                            break
            
            if not chart_found:
                print(f"    ⚠️ 円グラフが見つかりませんでした: {sheet.Name}")
                print(f"    🔍 元ファイル内のシート一覧:")
                for source_sheet in source_wb.Worksheets:
                    print(f"      - {source_sheet.Name}")
            
            try:
                source_wb.Close(SaveChanges=False)
                print(f"    ✅ 既存の円グラフコピー完了: {sheet.Name}")
            except:
                print(f"    ⚠️ 元ファイルのクローズでエラー（無視）: {sheet.Name}")
            
        finally:
            if source_wb:
                try:
                    source_wb.Close(SaveChanges=False)
                except:
                    pass
        
    except Exception as e:
        print(f"    ❌ 円グラフコピーエラー ({sheet.Name}): {e}")
        print(f"    🔍 シート名: {sheet.Name}")
        print(f"    🔍 月: {month}")
        print(f"    🔍 HPシート判定: {is_hp_sheet}")
        import traceback
        traceback.print_exc()

def create_simple_pie_chart_for_sheet(sheet, month: int):
    """シートにメニュー別とサイト別の円グラフを作成（HPはメニュー別のみ）"""
    try:
        print(f"    📊 円グラフ作成中: {sheet.Name}")
        
        # データ範囲を取得
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
        last_col = sheet.Cells(1, sheet.Columns.Count).End(-4159).Column
        
        if last_row <= 1:
            print(f"    ⚠️ データが不足: {sheet.Name}")
            return
        
        # シートタイプを判定
        is_hp_sheet = "HP" in sheet.Name
        
        # 列インデックスを探す
        menu_col = None
        site_col = None
        
        for col in range(1, last_col + 1):
            cell_value = sheet.Cells(1, col).Value
            if cell_value:
                cell_str = str(cell_value)
                if "MENU_NAME" in cell_str:
                    menu_col = col
                elif "SITE_CODE" in cell_str or "SITE_NAME" in cell_str:
                    site_col = col
        
        # メニュー別円グラフを作成（全シート共通）
        if menu_col is not None:
            create_pie_chart_by_column(sheet, month, menu_col, "メニュー", "MENU_NAME", last_row, is_hp_sheet)
        else:
            print(f"    ⚠️ MENU_NAME列が見つかりません: {sheet.Name}")
        
        # サイト別円グラフを作成（HPシート以外のみ）
        if not is_hp_sheet and site_col is not None:
            create_pie_chart_by_column(sheet, month, site_col, "サイト", "SITE_CODE", last_row, is_hp_sheet)
        elif not is_hp_sheet:
            print(f"    ⚠️ SITE_CODE列が見つかりません: {sheet.Name}")
        else:
            print(f"    ℹ️ HPシートのためサイト別円グラフは作成しません: {sheet.Name}")
        
        print(f"    ✅ 円グラフ作成完了: {sheet.Name}")
        
    except Exception as e:
        print(f"    ❌ 円グラフ作成エラー ({sheet.Name}): {e}")
        import traceback
        traceback.print_exc()

def create_pie_chart_by_column(sheet, month: int, col_index: int, chart_type: str, column_name: str, last_row: int, is_hp_sheet: bool = False):
    """指定された列で円グラフを作成"""
    try:
        # データを集計
        counts = {}
        for row in range(2, last_row + 1):
            cell_value = sheet.Cells(row, col_index).Value
            if cell_value:
                value_str = str(cell_value).strip()
                if value_str:
                    counts[value_str] = counts.get(value_str, 0) + 1
        
        if not counts:
            print(f"    ⚠️ {chart_type}データが見つかりません: {sheet.Name}")
            return
        
        # 既存のデータの下に集計データを書き込み
        chart_data_start_row = last_row + 3
        if chart_type == "メニュー":
            chart_data_start_row = last_row + 3
        else:  # サイト
            # メニュー円グラフの下に配置（HPシート以外のみ）
            if not is_hp_sheet:
                chart_data_start_row = last_row + 3 + len(counts) + 5
        
        sheet.Cells(chart_data_start_row, 1).Value = f"{chart_type}名"
        sheet.Cells(chart_data_start_row, 2).Value = "件数"
        
        row_idx = chart_data_start_row + 1
        for name, count in sorted(counts.items(), key=lambda x: x[1], reverse=True):
            sheet.Cells(row_idx, 1).Value = name
            sheet.Cells(row_idx, 2).Value = count
            row_idx += 1
        
        # 円グラフを作成
        chart_data_end_row = row_idx - 1
        chart_range = sheet.Range(sheet.Cells(chart_data_start_row, 1), sheet.Cells(chart_data_end_row, 2))
        
        # 円グラフを挿入
        chart_shape = sheet.Shapes.AddChart2(227, 5)  # xlPie
        chart = chart_shape.Chart
        chart.SetSourceData(chart_range)
        
        # チャートタイトル設定
        chart.HasTitle = True
        chart.ChartTitle.Text = f"{month}月_{sheet.Name}_{chart_type}別件数"
        
        # データラベル設定
        chart.ApplyDataLabels()
        chart.FullSeriesCollection(1).ApplyDataLabels()
        data_labels = chart.FullSeriesCollection(1).DataLabels()
        data_labels.ShowPercentage = True
        data_labels.ShowCategoryName = True
        data_labels.ShowValue = False
        chart.HasLegend = True
        
        # チャート位置調整（データの右側に配置）
        if chart_type == "メニュー":
            # メニュー別円グラフはデータの右側（列T）に配置
            chart_shape.Left = sheet.Range("T1").Left
            chart_shape.Top = sheet.Range("D1").Top
        else:  # サイト（HPシート以外のみ）
            if not is_hp_sheet:
                # サイト別円グラフはさらに右側（列AA）に配置
                chart_shape.Left = sheet.Range("AA1").Left
                chart_shape.Top = sheet.Range("D1").Top
        
        chart_shape.Width = 800
        chart_shape.Height = 600
        
        print(f"    ✅ {chart_type}別円グラフ作成完了: {sheet.Name}")
        
    except Exception as e:
        print(f"    ❌ {chart_type}別円グラフ作成エラー ({sheet.Name}): {e}")
        import traceback
        traceback.print_exc()


def main():
    """メイン関数"""
    try:
        print("📊 月別データ統合ツール")
        print("=" * 50)
        
        # resultsディレクトリを確認
        results_dir = Path("results")
        if not results_dir.exists():
            print("❌ resultsディレクトリが見つかりません")
            return 1
        
        # 月別統合を実行
        consolidate_monthly_data(results_dir)
        
        print("\n🎉 月別統合処理が完了しました!")
        print("💡 各月の統合ファイルが results フォルダに作成されました")
        
        return 0
        
    except Exception as e:
        print(f"❌ メイン処理でエラー: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
