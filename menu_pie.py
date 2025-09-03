from pathlib import Path
import sys
import re
from datetime import datetime, timedelta
import pandas as pd


def extract_bracket_content(menu_name: str) -> str:
    """
    メニュー名から括弧内の内容を抽出する
    例: "メニューA（カテゴリ1）" → "カテゴリ1"
    """
    if not menu_name or not isinstance(menu_name, str):
        return "その他"
    
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
            # 空文字や空白文字のみの場合は「その他」を返す
            if content and not content.isspace():
                return content
            else:
                # 空の括弧の場合は「その他」を返す
                return "その他"
    
    # 括弧が見つからない場合は元のメニュー名を返す
    return menu_name.strip() if menu_name.strip() else "その他"


def add_pivot_pie(target_path: Path) -> int:
    try:
        import win32com.client as win32
    except Exception:
        print("pywin32 が未インストールです。`pip install pywin32` を実行してください。")
        return 1

    # 絶対パスに変換（Excel は相対パスだと別の作業ディレクトリを見ることがある）
    abs_path = target_path.resolve()
    if not abs_path.exists():
        print(f"ファイルが見つかりません: {abs_path}")
        return 1

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Excel 定数
    xlDatabase = 1
    xlRowField = 1
    xlPie = 5
    xlColumns = 2

    wb = None
    try:
        wb = excel.Workbooks.Open(str(abs_path))

        # データシートの推定（既存スクリプトでは "データ"）
        data_ws = None
        try:
            data_ws = wb.Worksheets("データ")
        except Exception:
            try:
                data_ws = wb.Worksheets(1)
            except Exception as e:
                print(f"❌ データシートの取得に失敗: {e}")
                wb.Close(SaveChanges=False)
                return 1

        # 最終行・最終列
        try:
            last_row = data_ws.Cells(data_ws.Rows.Count, 1).End(-4162).Row  # xlUp=-4162
            last_col = data_ws.Cells(1, data_ws.Columns.Count).End(-4159).Column  # xlToLeft=-4159
            if last_row <= 1 or last_col <= 0:
                print("❌ データが不足しています")
                wb.Close(SaveChanges=False)
                return 1
        except Exception as e:
            print(f"❌ データ範囲の取得に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        src_range = data_ws.Range(data_ws.Cells(1, 1), data_ws.Cells(last_row, last_col))

        # MENU_NAME列の存在確認
        header_names = []
        menu_name_col = None
        try:
            for col in range(1, last_col + 1):
                header_val = data_ws.Cells(1, col).Value
                header_names.append(header_val)
                if header_val == "MENU_NAME":
                    menu_name_col = col
            
            if menu_name_col is None:
                print("❌ MENU_NAME列が見つかりません！")
                wb.Close(SaveChanges=False)
                return 1
        except Exception as e:
            print(f"❌ ヘッダーの読み取りに失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # 既存は残し、タイムスタンプ付きの新規シートを作成
        from datetime import datetime
        pivot_sheet_name = f"ピボット_メニュー分類円_{datetime.now():%Y%m%d%H%M%S}"
        pivot_ws = wb.Worksheets.Add()
        pivot_ws.Name = pivot_sheet_name

        # ピボットキャッシュ作成
        try:
            pc = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=src_range)
            # ピボットテーブル作成（A3 起点）
            pt = pc.CreatePivotTable(TableDestination=pivot_ws.Range("A3"), TableName=f"PivotMenu_{pivot_sheet_name}")
        except Exception as e:
            print(f"❌ ピボットテーブルの作成に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # 行フィールド: MENU_NAME（カテゴリ）
        try:
            pf = pt.PivotFields("MENU_NAME")
            pf.Orientation = xlRowField
            pf.Position = 1
        except Exception as e:
            print(f"❌ MENU_NAMEフィールドの設定に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # データフィールド: petano_log_no列をカウント
        try:
            data_field = pt.AddDataField(pt.PivotFields("petano_log_no"), "件数", -4112)
            data_field.Function = -4112
            data_field.NumberFormat = "0"
        except Exception as e:
            print(f"❌ データフィールドの設定に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # ピボットを更新
        try:
            pt.RefreshTable()
        except Exception as e:
            print(f"⚠️ ピボット更新でエラー（処理は続行）: {e}")

        # 円グラフ作成（ピボットの結果を固定値として別領域に展開してから参照）
        header_row = 3
        try:
            used_last_row = pivot_ws.Cells(pivot_ws.Rows.Count, 1).End(-4162).Row  # A列の最終行
            if used_last_row <= header_row:
                print("ピボットのデータが空です。")
                wb.Close(SaveChanges=True)
                return 0
        except Exception as e:
            print(f"❌ ピボットデータの読み取りに失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # 固定範囲ヘッダ
        out_col_start = 8  # H列
        pivot_ws.Cells(header_row, out_col_start).Value = "メニュー分類"
        pivot_ws.Cells(header_row, out_col_start + 1).Value = "件数"
        pivot_ws.Cells(header_row, out_col_start + 2).Value = "割合(%)"

        # ピボット表のラベル(A)と値(B)を固定範囲(H:I)へ書き出し
        src_first = header_row + 1
        row_count = used_last_row - header_row
        out_write_row = header_row + 1
        total_count = 0
        menu_data = []
        
        # まずデータを収集
        try:
            for i in range(row_count):
                # A列: メニュー名, B列: 件数（ラベル行を除く）
                label_val = pivot_ws.Cells(src_first + i, 1).Value
                value_val = pivot_ws.Cells(src_first + i, 2).Value
                
                # 空行や総計/Grand Total を除外
                label_str = str(label_val).strip() if label_val is not None else ""
                if not label_str or label_str in ("総計", "Grand Total", "行ラベル"):
                    continue
                if value_val is not None and str(value_val).replace('.0', '').isdigit():
                    # 括弧内の内容で分類
                    category = extract_bracket_content(label_str)
                    menu_data.append((category, int(float(value_val))))
                    total_count += int(float(value_val))
        except Exception as e:
            print(f"❌ データ収集でエラー: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # カテゴリ別に集計
        category_counts = {}
        for category, count in menu_data:
            if category in category_counts:
                category_counts[category] += count
            else:
                category_counts[category] = count

        # データを書き出し（パーセンテージ計算付き）
        try:
            for category, count in sorted(category_counts.items(), key=lambda x: x[1], reverse=True):
                percentage = (count / total_count * 100) if total_count > 0 else 0
                pivot_ws.Cells(out_write_row, out_col_start).Value = category
                pivot_ws.Cells(out_write_row, out_col_start + 1).Value = count
                pivot_ws.Cells(out_write_row, out_col_start + 2).Value = round(percentage, 1)
                out_write_row += 1
        except Exception as e:
            print(f"❌ データ書き出しでエラー: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # 実データの最終行（ヘッダ含む）
        used_out_last = out_write_row - 1
        if used_out_last <= header_row:
            print("ピボットのカテゴリがありません（総計のみか空です）。")
            wb.Close(SaveChanges=True)
            return 0

        # 円グラフ作成
        try:
            chart_range = pivot_ws.Range(
                pivot_ws.Cells(header_row, out_col_start),
                pivot_ws.Cells(used_out_last, out_col_start + 1),
            )
            shape = pivot_ws.Shapes.AddChart2(227, xlPie)
            chart = shape.Chart
            chart.SetSourceData(chart_range)
            chart.HasTitle = True
            chart.ChartTitle.Text = "メニュー分類別割合"
            
            # データラベル（割合とメニュー分類名）の設定を強化
            try:
                chart.ApplyDataLabels()
                chart.FullSeriesCollection(1).ApplyDataLabels()
                
                # データラベルの表示設定
                data_labels = chart.FullSeriesCollection(1).DataLabels()
                data_labels.ShowPercentage = True      # パーセンテージを表示
                data_labels.ShowCategoryName = True    # メニュー分類名を表示
                data_labels.ShowValue = False          # 件数は非表示
                data_labels.Position = -4142          # xlLabelPositionOutsideEnd（外側）
                
                # 吹き出し形式の設定
                try:
                    chart.ApplyDataLabels()
                    chart.FullSeriesCollection(1).ApplyDataLabels()
                    
                    # データラベルの表示設定
                    data_labels = chart.FullSeriesCollection(1).DataLabels()
                    data_labels.ShowPercentage = True      # パーセンテージを表示
                    data_labels.ShowCategoryName = True    # メニュー分類名を表示
                    data_labels.ShowValue = False          # 件数は非表示
                    data_labels.Position = -4142          # xlLabelPositionOutsideEnd（外側）
                    
                    # 吹き出し形式の設定
                    try:
                        data_labels.Format.TextFrame2.TextRange.Font.Size = 9
                        # 吹き出し線の設定
                        data_labels.Format.Line.Visible = True
                        data_labels.Format.Line.ForeColor.RGB = 0x000000  # 黒色の線
                    except Exception:
                        pass
                    
                    # データラベルのフォント設定（見やすくするため）
                    try:
                        data_labels.Font.Size = 9
                        data_labels.Font.Name = "Arial"
                    except Exception:
                        pass
                    
                    # 引き出し線を有効にする（長いメニュー分類名の場合）
                    try:
                        data_labels.ShowLeaderLines = True
                    except Exception:
                        pass
                    
                    # 凡例を非表示にする（データラベルにメニュー分類名が表示されるため）
                    chart.HasLegend = False
                    
                except Exception as e:
                    print(f"⚠️ データラベル設定でエラー（基本的な設定のみ適用）: {e}")
                    # フォールバック：基本的な設定のみ
                    try:
                        chart.FullSeriesCollection(1).DataLabels().ShowPercentage = True
                        chart.FullSeriesCollection(1).DataLabels().ShowCategoryName = True
                        chart.HasLegend = False  # 凡例を確実に非表示
                    except Exception:
                        pass
                
                # 念のため、凡例を確実に非表示にする
                try:
                    chart.HasLegend = False
                except Exception:
                    pass

            except Exception as e:
                print(f"⚠️ データラベル設定でエラー（基本的な設定のみ適用）: {e}")
                # フォールバック：基本的な設定のみ
                try:
                    chart.FullSeriesCollection(1).DataLabels().ShowPercentage = True
                    chart.FullSeriesCollection(1).DataLabels().ShowCategoryName = True
                    chart.HasLegend = False  # 凡例を確実に非表示
                except Exception:
                    pass
            
            # 念のため、凡例を確実に非表示にする
            try:
                chart.HasLegend = False
            except Exception:
                pass

            # レイアウト調整（サイズを大きくしてメニュー分類名を見やすく）
            shape.Left = pivot_ws.Range("D3").Left
            shape.Top = pivot_ws.Range("D3").Top
            shape.Width = 600   # 幅を広く
            shape.Height = 450  # 高さも大きく

        except Exception as e:
            print(f"❌ 円グラフ作成でエラー: {e}")
            wb.Close(SaveChanges=False)
            return 1

        wb.Close(SaveChanges=True)
        print(f"✅ 更新完了: {target_path}")
        print(f"📊 作成された分類: {list(category_counts.keys())}")
        return 0
        
    except Exception as e:
        print(f"❌ 処理失敗: {e}")
        # 開かれて保存できない場合は別名で保存
        try:
            if wb is not None:
                alt = target_path.with_name(target_path.stem + "_pivot" + target_path.suffix)
                wb.SaveCopyAs(str(alt))
                wb.Close(SaveChanges=False)
                print(f"⚠️ 警告: 元ファイルへ保存できないため複製を保存しました → {alt}")
                return 0
        except Exception as save_error:
            print(f"❌ 複製保存も失敗: {save_error}")
        return 1
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def add_site_pivot_pie(target_path: Path) -> int:
    try:
        import win32com.client as win32
    except Exception:
        print("pywin32 が未インストールです。`pip install pywin32` を実行してください。")
        return 1

    # 絶対パスに変換（Excel は相対パスだと別の作業ディレクトリを見ることがある）
    abs_path = target_path.resolve()
    if not abs_path.exists():
        print(f"ファイルが見つかりません: {abs_path}")
        return 1

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Excel 定数
    xlDatabase = 1
    xlRowField = 1
    xlPie = 5
    xlColumns = 2

    wb = None
    try:
        wb = excel.Workbooks.Open(str(abs_path))

        # データシートの推定（既存スクリプトでは "データ"）
        data_ws = None
        try:
            data_ws = wb.Worksheets("データ")
        except Exception:
            data_ws = wb.Worksheets(1)

        # 最終行・最終列
        last_row = data_ws.Cells(data_ws.Rows.Count, 1).End(-4162).Row  # xlUp=-4162
        last_col = data_ws.Cells(1, data_ws.Columns.Count).End(-4159).Column  # xlToLeft=-4159
        src_range = data_ws.Range(data_ws.Cells(1, 1), data_ws.Cells(last_row, last_col))

        # SITE_CODE列の存在確認
        header_names = []
        for col in range(1, last_col + 1):
            header_val = data_ws.Cells(1, col).Value
            header_names.append(header_val)
        
        if "SITE_CODE" not in header_names:
            print("❌ SITE_CODE列が見つかりません！")
            wb.Close(SaveChanges=False)
            return 1

        # 既存は残し、タイムスタンプ付きの新規シートを作成
        from datetime import datetime
        pivot_sheet_name = f"ピボット_サイト円_{datetime.now():%Y%m%d%H%M%S}"
        pivot_ws = wb.Worksheets.Add()
        pivot_ws.Name = pivot_sheet_name

        # ピボットキャッシュ作成
        pc = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=src_range)
        # ピボットテーブル作成（A3 起点）
        pt = pc.CreatePivotTable(TableDestination=pivot_ws.Range("A3"), TableName=f"PivotSITE_{pivot_sheet_name}")

        # 行フィールド: SITE_CODE（カテゴリ）
        try:
            pf = pt.PivotFields("SITE_CODE")
            pf.Orientation = xlRowField
            pf.Position = 1
        except Exception as e:
            print(f"❌ SITE_CODEフィールドの設定に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # データフィールド: petano_log_no列をカウントして、SITE_CODEでグループ化
        try:
            data_field = pt.AddDataField(pt.PivotFields("petano_log_no"), "件数", -4112)
            data_field.Function = -4112
            data_field.NumberFormat = "0"
        except Exception as e:
            print(f"❌ データフィールドの設定に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # ピボットを更新
        try:
            pt.RefreshTable()
        except Exception as e:
            pass

        # 円グラフ作成（ピボットの結果を固定値として別領域に展開してから参照）
        header_row = 3
        used_last_row = pivot_ws.Cells(pivot_ws.Rows.Count, 1).End(-4162).Row  # A列の最終行
        if used_last_row <= header_row:
            print("ピボットのデータが空です。")
            wb.Close(SaveChanges=True)
            return 0

        # 固定範囲ヘッダ
        out_col_start = 8  # H列
        pivot_ws.Cells(header_row, out_col_start).Value = "サイト名"
        pivot_ws.Cells(header_row, out_col_start + 1).Value = "件数"
        pivot_ws.Cells(header_row, out_col_start + 2).Value = "割合(%)"

        # ピボット表のラベル(A)と値(B)を固定範囲(H:I)へ書き出し
        src_first = header_row + 1
        row_count = used_last_row - header_row
        out_write_row = header_row + 1
        total_count = 0
        site_data = []
        
        # まずデータを収集
        for i in range(row_count):
            # A列: サイト名, B列: 件数（ラベル行を除く）
            label_val = pivot_ws.Cells(src_first + i, 1).Value
            value_val = pivot_ws.Cells(src_first + i, 2).Value
            
            # 空行や総計/Grand Total を除外
            label_str = str(label_val).strip() if label_val is not None else ""
            if not label_str or label_str in ("総計", "Grand Total", "行ラベル"):
                continue
            if value_val is not None and str(value_val).replace('.0', '').isdigit():
                site_data.append((label_val, int(float(value_val))))
                total_count += int(float(value_val))
        
        # データを書き出し（パーセンテージ計算付き）
        for label_val, value_val in site_data:
            percentage = (value_val / total_count * 100) if total_count > 0 else 0
            pivot_ws.Cells(out_write_row, out_col_start).Value = label_val
            pivot_ws.Cells(out_write_row, out_col_start + 1).Value = value_val
            pivot_ws.Cells(out_write_row, out_col_start + 2).Value = round(percentage, 1)
            out_write_row += 1

        # 実データの最終行（ヘッダ含む）
        used_out_last = out_write_row - 1
        if used_out_last <= header_row:
            print("ピボットのカテゴリがありません（総計のみか空です）。")
            wb.Close(SaveChanges=True)
            return 0

        chart_range = pivot_ws.Range(
            pivot_ws.Cells(header_row, out_col_start),
            pivot_ws.Cells(used_out_last, out_col_start + 1),
        )
        shape = pivot_ws.Shapes.AddChart2(227, xlPie)
        chart = shape.Chart
        chart.SetSourceData(chart_range)
        chart.HasTitle = True
        chart.ChartTitle.Text = "サイト別割合"
        
        # データラベル（割合とサイト名）の設定を強化
        try:
            chart.ApplyDataLabels()
            chart.FullSeriesCollection(1).ApplyDataLabels()
            
            # データラベルの表示設定
            data_labels = chart.FullSeriesCollection(1).DataLabels()
            data_labels.ShowPercentage = True      # パーセンテージを表示
            data_labels.ShowCategoryName = True    # サイト名を表示
            data_labels.ShowValue = False          # 件数は非表示
            data_labels.Position = -4142          # xlLabelPositionOutsideEnd（外側）
            
            # 吹き出し形式の設定
            try:
                # xlCallout = 2 （吹き出し形式）
                data_labels.Format.TextFrame2.TextRange.Font.Size = 9
                # 吹き出し線の設定
                data_labels.Format.Line.Visible = True
                data_labels.Format.Line.ForeColor.RGB = 0x000000  # 黒色の線
            except:
                pass
            
            # データラベルのフォント設定（見やすくするため）
            try:
                data_labels.Font.Size = 9
                data_labels.Font.Name = "Arial"
            except:
                pass
            
            # 引き出し線を有効にする（長いサイト名の場合）
            try:
                data_labels.ShowLeaderLines = True
            except:
                pass
            
            # 凡例を非表示にする（データラベルにサイト名が表示されるため）
            chart.HasLegend = False
            
        except Exception as e:
            print(f"データラベル設定でエラー: {e}")
            # フォールバック：基本的な設定のみ
            try:
                chart.FullSeriesCollection(1).DataLabels().ShowPercentage = True
                chart.FullSeriesCollection(1).DataLabels().ShowCategoryName = True
                chart.HasLegend = False  # 凡例を確実に非表示
            except:
                pass
        
        # 念のため、凡例を確実に非表示にする
        try:
            chart.HasLegend = False
        except:
            pass

        # レイアウト調整（サイズを大きくしてサイト名を見やすく）
        shape.Left = pivot_ws.Range("D3").Left
        shape.Top = pivot_ws.Range("D3").Top
        shape.Width = 600   # 幅を広く
        shape.Height = 450  # 高さも大きく

        wb.Close(SaveChanges=True)
        print(f"更新完了: {target_path}")
        return 0
    except Exception as e:
        # 開かれて保存できない場合は別名で保存
        try:
            if wb is not None:
                alt = target_path.with_name(target_path.stem + "_pivot" + target_path.suffix)
                wb.SaveCopyAs(str(alt))
                wb.Close(SaveChanges=False)
                print(f"警告: 元ファイルへ保存できないため複製を保存しました → {alt}")
                return 0
        except Exception:
            pass
        print(f"処理失敗: {e}")
        return 1
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def show_help():
    """使用方法を表示"""
    print("""
📊 メニュー分類別円グラフ作成ツール

使用方法:
  python menu_pie.py                                    # 過去2か月のデータを自動処理
  python menu_pie.py ファイル名.xlsx                     # 特定のファイルを処理
  python menu_pie.py --months 2024-08,2024-09           # 特定の2か月を指定して処理
  python menu_pie.py --all                              # resultsディレクトリ内の全ファイルを処理
  python menu_pie.py --test                             # 括弧内内容抽出テストを実行
  python menu_pie.py --help                             # このヘルプを表示

例:
  python menu_pie.py                                    # 過去2か月を自動処理（推奨）
  python menu_pie.py 7月_初回受付_その他.xlsx
  python menu_pie.py "6月_再受付_hp.xlsx"
  python menu_pie.py results/7月_初回受付_hp.xlsx
  python menu_pie.py --months 2024-08,2024-09          # 8月と9月のデータを処理
  python menu_pie.py --all                              # 全ファイルを処理

月指定の形式:
  --months YYYY-MM,YYYY-MM  (例: 2024-08,2024-09)
  - カンマ区切りで2つの月を指定
  - 年-月の形式で指定（例: 2024-08）

注意:
  - 引数なしで実行すると、現在の日付から過去2か月のデータを自動処理します
  - ファイル名のみ指定した場合、resultsディレクトリ内を検索します
  - 絶対パスまたは相対パスで指定することも可能です
  - ファイル名にスペースが含まれる場合は引用符で囲んでください
  - 月指定時は、該当月のデータを含むファイルを自動検索します
""")


def parse_months_argument(months_str: str) -> list:
    """月指定の引数をパースする"""
    try:
        months = months_str.split(',')
        if len(months) != 2:
            raise ValueError("2つの月を指定してください")
        
        parsed_months = []
        for month in months:
            month = month.strip()
            # YYYY-MM形式をチェック
            if len(month) != 7 or month[4] != '-':
                raise ValueError(f"月の形式が正しくありません: {month}")
            
            year, month_num = month.split('-')
            if not year.isdigit() or not month_num.isdigit():
                raise ValueError(f"年月が数値ではありません: {month}")
            
            month_num = int(month_num)
            if month_num < 1 or month_num > 12:
                raise ValueError(f"月の値が範囲外です: {month_num}")
            
            parsed_months.append((int(year), month_num))
        
        return parsed_months
    except Exception as e:
        print(f"❌ 月指定の解析エラー: {e}")
        print("💡 正しい形式: --months 2024-08,2024-09")
        return None


def get_past_two_months() -> list:
    """現在の日付から過去2か月を計算する"""
    current_date = datetime.now()
    
    # 今月の1日
    current_month = current_date.replace(day=1)
    
    # 先月の1日
    last_month = (current_month - timedelta(days=1)).replace(day=1)
    
    # 先々月の1日
    two_months_ago = (last_month - timedelta(days=1)).replace(day=1)
    
    months = [
        (last_month.year, last_month.month),      # 先月
        (two_months_ago.year, two_months_ago.month)  # 先々月
    ]
    
    return months


def find_files_for_months(months: list) -> list:
    """指定された月に対応するファイルを検索する"""
    results_dir = Path("results")
    if not results_dir.exists():
        print("❌ resultsディレクトリが見つかりません")
        return []
    
    # 月名のマッピング
    month_names = {
        1: "1月", 2: "2月", 3: "3月", 4: "4月", 5: "5月", 6: "6月",
        7: "7月", 8: "8月", 9: "9月", 10: "10月", 11: "11月", 12: "12月"
    }
    
    target_files = []
    for year, month in months:
        month_name = month_names[month]
        # 該当月を含むファイルを検索
        pattern_files = list(results_dir.glob(f"*{month_name}*.xlsx"))
        if pattern_files:
            target_files.extend(pattern_files)
            print(f"📅 {year}年{month_name}: {len(pattern_files)}ファイル見つかりました")
        else:
            print(f"⚠️ {year}年{month_name}: 該当ファイルが見つかりません")
    
    return target_files


def main() -> int:
    """メニュー分類別の円グラフを作成するメイン関数"""
    try:
        # ヘルプ表示
        if len(sys.argv) > 1 and sys.argv[1] in ["--help", "-h", "-?"]:
            show_help()
            return 0
        
        # 月指定の処理
        if len(sys.argv) > 1 and sys.argv[1].startswith("--months"):
            if len(sys.argv) < 3:
                print("❌ --months の後に月を指定してください")
                print("💡 例: python menu_pie.py --months 2024-08,2024-09")
                return 1
            
            months = parse_months_argument(sys.argv[2])
            if months is None:
                return 1
            
            print(f"🎯 指定された月: {months[0][0]}年{months[0][1]}月, {months[1][0]}年{months[1][1]}月")
            target_files = find_files_for_months(months)
            
            if not target_files:
                print("❌ 指定された月のファイルが見つかりません")
                return 1
            
            # 重複を除去
            target_files = list(set(target_files))
            print(f"📁 処理対象ファイル: {[f.name for f in target_files]}")
            
            # 全てのファイルを処理
            success_count = 0
            error_count = 0
            
            for i, target in enumerate(target_files, 1):
                print(f"\n--- 処理 {i}/{len(target_files)} ---")
                print(f"🎯 処理開始: {target.name}")
                
                try:
                    result = add_pivot_pie(target)
                    if result == 0:
                        print(f"✅ 処理完了: {target.name}")
                        success_count += 1
                    else:
                        print(f"❌ 処理失敗: {target.name}")
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
        
        # コマンドライン引数でファイル名が指定されているかチェック
        if len(sys.argv) > 1 and sys.argv[1] != "--test":
            # 指定されたファイルを処理
            specified_file = sys.argv[1]
            target = Path(specified_file)
            
            # 相対パスの場合はresultsディレクトリ内を検索
            if not target.is_absolute():
                results_dir = Path("results")
                if results_dir.exists():
                    target = results_dir / specified_file
                else:
                    print("❌ resultsディレクトリが見つかりません")
                    return 1
            
            if not target.exists():
                print(f"❌ 指定されたファイルが見つかりません: {target}")
                print("💡 ヒント: ファイル名が正しいか確認してください")
                print("💡 ヒント: python menu_pie.py --help で使用方法を確認")
                return 1
            
            print(f"🎯 指定されたファイルを処理: {target.name}")
            result = add_pivot_pie(target)
            if result == 0:
                print(f"✅ 処理完了: {target.name}")
            else:
                print(f"❌ 処理失敗: {target.name}")
            return result
        
        # --all オプションの処理
        if len(sys.argv) > 1 and sys.argv[1] == "--all":
            # resultsディレクトリ内の全ファイルを処理
            results_dir = Path("results")
            if not results_dir.exists():
                print("❌ resultsディレクトリが見つかりません")
                return 1
            
            # Excelファイルを検索（その他ファイルを優先）
            excel_files = list(results_dir.glob("*.xlsx"))
            if not excel_files:
                print("❌ resultsディレクトリにExcelファイルが見つかりません")
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
            print("💡 特定のファイルのみ処理したい場合は: python menu_pie.py ファイル名.xlsx")
            
            # 全てのファイルを処理
            success_count = 0
            error_count = 0
            
            for i, target in enumerate(target_files, 1):
                print(f"\n--- 処理 {i}/{len(target_files)} ---")
                print(f"🎯 処理開始: {target.name}")
                
                try:
                    result = add_pivot_pie(target)
                    if result == 0:
                        print(f"✅ 処理完了: {target.name}")
                        success_count += 1
                    else:
                        print(f"❌ 処理失敗: {target.name}")
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
        
        # 引数なしの場合は過去2か月を自動処理
        print("🔄 過去2か月のデータを自動処理します...")
        months = get_past_two_months()
        print(f"🎯 対象期間: {months[0][0]}年{months[0][1]}月, {months[1][0]}年{months[1][1]}月")
        
        target_files = find_files_for_months(months)
        
        if not target_files:
            print("❌ 過去2か月のファイルが見つかりません")
            print("💡 ヒント: python menu_pie.py --all で全ファイルを処理")
            print("💡 ヒント: python menu_pie.py --help で使用方法を確認")
            return 1
        
        # 重複を除去
        target_files = list(set(target_files))
        print(f"📁 処理対象ファイル: {[f.name for f in target_files]}")
        
        # 全てのファイルを処理
        success_count = 0
        error_count = 0
        
        for i, target in enumerate(target_files, 1):
            print(f"\n--- 処理 {i}/{len(target_files)} ---")
            print(f"🎯 処理開始: {target.name}")
            
            try:
                result = add_pivot_pie(target)
                if result == 0:
                    print(f"✅ 処理完了: {target.name}")
                    success_count += 1
                else:
                    print(f"❌ 処理失敗: {target.name}")
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


def main_site() -> int:
    """サイト別の円グラフを作成する関数（既存のまま）"""
    try:
        target = Path("results") / "web_reservation_leaf_20250801.csv"
        if not target.exists():
            print(f"❌ ファイルが見つかりません: {target}")
            return 1
        return add_site_pivot_pie(target)
    except Exception as e:
        print(f"❌ サイト処理でエラー: {e}")
        return 1


def test_extract_bracket_content():
    """括弧内の内容を抽出する関数のテスト"""
    test_cases = [
        ("メニューA（カテゴリ1）", "カテゴリ1"),
        ("メニューB(カテゴリ2)", "カテゴリ2"),
        ("メニューC【カテゴリ3】", "カテゴリ3"),
        ("メニューD[カテゴリ4]", "カテゴリ4"),
        ("メニューE（　）", "その他"),  # 空の括弧
        ("メニューF", "メニューF"),  # 括弧なし
        ("", "その他"),  # 空文字
        (None, "その他"),  # None
        ("メニューG（カテゴリ5）詳細", "カテゴリ5"),  # 括弧後に文字あり
    ]
    
    print("🧪 括弧内内容抽出テスト:")
    for input_text, expected in test_cases:
        result = extract_bracket_content(input_text)
        status = "✅" if result == expected else "❌"
        print(f"  {status} 入力: '{input_text}' → 出力: '{result}' (期待値: '{expected}')")


if __name__ == "__main__":
    try:
        # テストモードの場合はテストを実行
        if len(sys.argv) > 1 and sys.argv[1] == "--test":
            test_extract_bracket_content()
            raise SystemExit(0)
        
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


