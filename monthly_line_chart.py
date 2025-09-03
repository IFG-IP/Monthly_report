from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import sys
import numpy as np

# 日本語フォントの設定
plt.rcParams['font.family'] = ['DejaVu Sans', 'Hiragino Sans', 'Yu Gothic', 'Meiryo', 'Takao', 'IPAexGothic', 'IPAPGothic', 'VL PGothic', 'Noto Sans CJK JP']

def load_csv_data(csv_path: Path) -> pd.DataFrame:
    """CSVファイルを読み込む"""
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        print(f"📊 CSVファイル読み込み完了: {len(df)}行")
        return df
    except Exception as e:
        print(f"❌ CSVファイル読み込みエラー: {e}")
        return None

def parse_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """日付列を解析してdatetime型に変換"""
    try:
        # appoint_datetime_clean列を日付型に変換
        df['appoint_date'] = pd.to_datetime(df['appoint_datetime_clean'], errors='coerce')
        
        # 日付が解析できない行を除外
        before_count = len(df)
        df = df.dropna(subset=['appoint_date'])
        after_count = len(df)
        
        if before_count != after_count:
            print(f"⚠️ 日付解析エラーで{before_count - after_count}行を除外")
        
        print(f"📅 日付解析完了: {after_count}行")
        return df
    except Exception as e:
        print(f"❌ 日付解析エラー: {e}")
        return None

def create_monthly_summary(df: pd.DataFrame, site_type: str = "all") -> pd.DataFrame:
    """月別の初回受付・再受付の集計を作成（HP/その他別）"""
    try:
        # 年月を抽出
        df['year_month'] = df['appoint_date'].dt.to_period('M')
        
        # HPとその他で分ける
        if site_type == "hp":
            df_filtered = df[df['SITE_CODE'].str.strip().str.lower() == 'hp']
            print(f"🏠 HP経由データ: {len(df_filtered)}行")
        elif site_type == "other":
            df_filtered = df[df['SITE_CODE'].str.strip().str.lower() != 'hp']
            print(f"🌐 その他経由データ: {len(df_filtered)}行")
        else:
            df_filtered = df
            print(f"📊 全データ: {len(df_filtered)}行")
        
        # 月別・受付種別で集計
        monthly_summary = df_filtered.groupby(['year_month', 'SYSTEM_USAGE_NAME']).size().unstack(fill_value=0)
        
        # 列名を整理
        if '初回受付' in monthly_summary.columns:
            monthly_summary = monthly_summary.rename(columns={'初回受付': '初回受付'})
        if '再受付' in monthly_summary.columns:
            monthly_summary = monthly_summary.rename(columns={'再受付': '再受付'})
        if '初回キャンセル' in monthly_summary.columns:
            monthly_summary = monthly_summary.rename(columns={'初回キャンセル': '初回キャンセル'})
        
        # 存在しない列を追加（0で初期化）
        for col in ['初回受付', '再受付', '初回キャンセル']:
            if col not in monthly_summary.columns:
                monthly_summary[col] = 0
        
        # 合計列を追加
        monthly_summary['合計'] = monthly_summary['初回受付'] + monthly_summary['再受付'] + monthly_summary['初回キャンセル']
        
        print(f"📊 {site_type}月別集計完了: {len(monthly_summary)}か月分")
        return monthly_summary
    except Exception as e:
        print(f"❌ {site_type}月別集計エラー: {e}")
        return None

def create_line_chart(monthly_summary: pd.DataFrame, output_path: Path) -> bool:
    """折れ線グラフを作成"""
    try:
        # データを準備
        months = monthly_summary.index.to_timestamp()
        first_visit = monthly_summary['初回受付']
        repeat_visit = monthly_summary['再受付']
        total = monthly_summary['合計']
        
        # グラフのサイズを設定
        plt.figure(figsize=(14, 8))
        
        # 背景色を水色に設定
        plt.gca().set_facecolor('#E0F6FF')
        
        # 折れ線グラフを描画
        plt.plot(months, first_visit, marker='o', linewidth=2, markersize=6, label='初回受付', color='#2E8B57')
        plt.plot(months, repeat_visit, marker='s', linewidth=2, markersize=6, label='再受付', color='#4169E1')
        
        # 各データポイントに数値を表示
        for i, (month, first_val, repeat_val) in enumerate(zip(months, first_visit, repeat_visit)):
            # 初回受付の数値を表示（マーカーの上側）
            plt.annotate(f'{first_val}', 
                        xy=(month, first_val), 
                        xytext=(0, 10), 
                        textcoords='offset points',
                        ha='center', va='bottom',
                        fontsize=17, fontweight='bold',
                        color='#2E8B57')
            
            # 再受付の数値を表示（マーカーの下側）
            plt.annotate(f'{repeat_val}', 
                        xy=(month, repeat_val), 
                        xytext=(0, -15), 
                        textcoords='offset points',
                        ha='center', va='top',
                        fontsize=17, fontweight='bold',
                        color='#4169E1')
        
        # グラフの設定
        plt.title('月別予約受付数推移', fontsize=16, fontweight='bold', pad=20)
        plt.xlabel('月', fontsize=12)
        plt.ylabel('件数', fontsize=12)
        plt.legend(fontsize=11, loc='upper left')
        plt.grid(True, alpha=0.3)
        
        # X軸の設定（月表示）
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.gca().xaxis.set_major_locator(mdates.MonthLocator(interval=1))
        plt.xticks(rotation=45)
        
        # レイアウト調整
        plt.tight_layout()
        
        # グラフを保存
        chart_path = output_path.with_suffix('.png')
        plt.savefig(chart_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"📈 折れ線グラフ保存完了: {chart_path}")
        return True
    except Exception as e:
        print(f"❌ グラフ作成エラー: {e}")
        return False

def save_to_excel_with_chart(monthly_summary: pd.DataFrame, output_path: Path, site_type: str = "all") -> bool:
    """Excelファイルに結果とグラフを保存（Excelのグラフ機能を使用）"""
    try:
        import win32com.client as win32
        
        # Excelファイルを作成（データのみ）
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 初回受付と再受付のみのデータを準備
            chart_data = monthly_summary[['初回受付', '再受付']].copy()
            chart_data.to_excel(writer, sheet_name='月別集計', index=True)
            
            # ワークブックを取得してフォーマット調整
            workbook = writer.book
            worksheet = writer.sheets['月別集計']
            
            # 列幅を自動調整
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 20)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Excelアプリケーションを起動してグラフを作成
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            # ワークブックを開く
            wb = excel.Workbooks.Open(str(output_path.resolve()))
            ws = wb.Worksheets("月別集計")
            
            # データ範囲を取得
            last_row = len(chart_data) + 1  # ヘッダー行を含む
            last_col = len(chart_data.columns) + 1  # インデックス列を含む
            
            # グラフ用のデータ範囲（初回受付と再受付のみ）
            # A列: 年月, B列: 初回受付, C列: 再受付
            chart_data_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, 3))  # A1:C(last_row)
            
            # 折れ線グラフを作成
            chart = ws.Shapes.AddChart2(227, 4)  # xlLine = 4
            chart_obj = chart.Chart
            
            # グラフのデータソースを設定
            chart_obj.SetSourceData(chart_data_range)
            
            # 系列の順序を確認して色を設定
            # 初回受付を青色、再受付をグレー色に設定
            if chart_obj.SeriesCollection().Count >= 1:
                # 初回受付（1番目の系列）
                series1 = chart_obj.SeriesCollection(1)
                series1.Format.Line.ForeColor.RGB = 0xCD8433  # 青色（BGR形式）
                series1.MarkerBackgroundColor = 0xCD8433
                series1.MarkerForegroundColor = 0xCD8433
                
            if chart_obj.SeriesCollection().Count >= 2:
                # 再受付（2番目の系列）
                series2 = chart_obj.SeriesCollection(2)
                series2.Format.Line.ForeColor.RGB = 0x808080  # グレー色
                series2.MarkerBackgroundColor = 0x808080
                series2.MarkerForegroundColor = 0x808080
            
            # グラフのタイトルを設定（サイトタイプ別）
            chart_obj.HasTitle = True
            if site_type == "hp":
                chart_obj.ChartTitle.Text = "HP予約受付数推移"
            elif site_type == "other":
                chart_obj.ChartTitle.Text = "WEB予約受付数推移"
            else:
                chart_obj.ChartTitle.Text = "予約受付数推移"
            chart_obj.ChartTitle.Font.Size = 16
            chart_obj.ChartTitle.Font.Bold = True
            
            # タイトルをプロットエリア内に配置
            try:
                plot_left = chart_obj.PlotArea.Left
                plot_top = chart_obj.PlotArea.Top
                plot_width = chart_obj.PlotArea.Width
                
                # タイトルをプロットエリア内の上部中央に配置
                title_left = plot_left + (plot_width / 2) - 30  # 中央寄せ（タイトル幅を考慮）
                title_top = plot_top  # プロットエリア内の上部
                
                chart_obj.ChartTitle.Left = title_left
                chart_obj.ChartTitle.Top = title_top
                print("✅ タイトルをプロットエリア内に配置しました")
            except Exception as e:
                print(f"⚠️ タイトル位置の設定に失敗: {e}")
            
            # X軸の設定（年月表示）
            chart_obj.Axes(1).CategoryNames = chart_data.index.strftime('%Y年%m月').tolist()
            chart_obj.Axes(1).TickLabels.Font.Size = 10
            chart_obj.Axes(1).TickLabels.Orientation = 0  # 水平表示（斜めにしない）
            
            # グラフの位置とサイズを設定（右側凡例用に幅を広げる）
            chart.Left = ws.Range("E2").Left
            chart.Top = ws.Range("E2").Top
            chart.Width = 800  # 右側凡例用に幅を広げる
            chart.Height = 400
            
            # データラベルと色を設定
            for i in range(1, chart_obj.SeriesCollection().Count + 1):
                series = chart_obj.SeriesCollection(i)
                series.HasDataLabels = True
                series.DataLabels().ShowValue = True
                series.DataLabels().Font.Size = 13
                series.DataLabels().Font.Bold = True
                
                # 系列の色と名前を設定
                if i == 1:  # 初回受付
                    series.Format.Line.ForeColor.RGB = 0xCD8433  # 青色（BGR形式）
                    series.Format.Line.Weight = 1  # 線の太さ
                    series.MarkerBackgroundColor = 0xCD8433
                    series.MarkerForegroundColor = 0xCD8433
                    series.MarkerStyle = 8  # 円形マーカー
                    
                    # 凡例名をサイトタイプに応じて変更
                    if site_type == "hp":
                        series.Name = "HP新規"
                    elif site_type == "other":
                        series.Name = "WEB新規"
                    else:
                        series.Name = "初回受付"
                        
                elif i == 2:  # 再受付
                    series.Format.Line.ForeColor.RGB = 0x808080  # グレー色
                    series.Format.Line.Weight = 1  # 線の太さ
                    series.MarkerBackgroundColor = 0x808080
                    series.MarkerForegroundColor = 0x808080
                    series.MarkerStyle = 8  # 円形マーカー
                    
                    # 凡例名をサイトタイプに応じて変更
                    if site_type == "hp":
                        series.Name = "HP再受付"
                    elif site_type == "other":
                        series.Name = "WEB再受付"
                    else:
                        series.Name = "再受付"
            
            # 凡例の設定（プロットエリア内に手動配置）
            chart_obj.HasLegend = True
            try:
                # まず自動配置を無効にする
                chart_obj.Legend.Position = -4107  # xlLegendPositionCustom
                
                # プロットエリア内の左上に手動で配置
                plot_left = chart_obj.PlotArea.Left
                plot_top = chart_obj.PlotArea.Top
                plot_width = chart_obj.PlotArea.Width
                plot_height = chart_obj.PlotArea.Height
                
                # 凡例をプロットエリア内の左上に配置（マージンを考慮）
                legend_left = plot_left + 10  # 左から10ピクセル
                legend_top = plot_top + 10    # 上から10ピクセル
                
                chart_obj.Legend.Left = legend_left
                chart_obj.Legend.Top = legend_top
                
                print("✅ 凡例をプロットエリア内の左上に配置しました")
            except Exception as e:
                print(f"⚠️ プロットエリア内凡例の設定に失敗、上部に設定します: {e}")
                chart_obj.Legend.Position = 2  # xlLegendPositionTop（上部）
            
            # グリッド線の設定（非表示）
            chart_obj.Axes(1).HasMajorGridlines = False
            chart_obj.Axes(2).HasMajorGridlines = False
            
            # 縦軸（Y軸）のラベルを完全に非表示
            chart_obj.Axes(2).TickLabels.Font.Color = 0xF0D4B6  # 背景色と同じ色で見えなくする
            chart_obj.Axes(2).TickLabels.NumberFormat = ""
            chart_obj.Axes(2).HasMajorGridlines = False  # グリッド線も非表示
            chart_obj.Axes(2).HasMinorGridlines = False
            
            # Y軸の線の色も背景色と同じに設定
            try:
                chart_obj.Axes(2).Border.Color = 0xF0D4B6  # 背景色と同じ薄緑色
                chart_obj.Axes(2).Border.LineStyle = 1  # 実線
                print("✅ Y軸の線の色を背景色と同じに設定しました")
            except Exception as e:
                print(f"⚠️ Y軸の線の色設定に失敗: {e}")
            
            # 縦軸の最大値を動的に設定（最高値＋30）
            try:
                # 初回受付と再受付の最大値を取得
                max_first_visit = chart_data['初回受付'].max() if '初回受付' in chart_data.columns else 0
                max_repeat_visit = chart_data['再受付'].max() if '再受付' in chart_data.columns else 0
                max_value = max(max_first_visit, max_repeat_visit)
                
                # 最高値＋30を設定（最低でも50は確保）
                y_max = max(max_value + 30, 50)
                chart_obj.Axes(2).MaximumScale = y_max
                print(f"📊 縦軸最大値を{max_value}→{y_max}に設定")
            except Exception as e:
                print(f"⚠️ 縦軸最大値の設定に失敗、デフォルト100を使用: {e}")
                chart_obj.Axes(2).MaximumScale = 100
            
            # 背景色を設定
            chart_obj.PlotArea.Format.Fill.ForeColor.RGB = 0xF0D4B6
            
            # グラフエリア（外側）も同じ背景色で塗りつぶし
            try:
                # グラフエリアの背景を薄緑色に設定
                chart_obj.ChartArea.Format.Fill.Visible = True
                chart_obj.ChartArea.Format.Fill.ForeColor.RGB = 0xF0D4B6  # 薄緑色
                chart_obj.ChartArea.Format.Line.Visible = False
                
                # プロットエリアのサイズをグラフ全体に合わせる
                chart_obj.PlotArea.Left = 0
                chart_obj.PlotArea.Top = 0
                chart_obj.PlotArea.Width = chart_obj.ChartArea.Width
                chart_obj.PlotArea.Height = chart_obj.ChartArea.Height
                
                print("✅ グラフエリアとプロットエリアを同じ背景色で塗りつぶしました")
            except Exception as e:
                print(f"⚠️ グラフエリア背景色設定に失敗: {e}") 
            
            # 保存
            wb.Save()
            wb.Close()
            
        finally:
            excel.Quit()
        
        print(f"📊 Excelファイル保存完了: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Excel保存エラー: {e}")
        return False

def show_help():
    """使用方法を表示"""
    print("""
📈 月別予約受付数推移グラフ作成ツール

使用方法:
  python monthly_line_chart.py                                    # デフォルトCSVファイルを処理
  python monthly_line_chart.py ファイル名.csv                      # 特定のCSVファイルを処理
  python monthly_line_chart.py --help                             # このヘルプを表示

例:
  python monthly_line_chart.py
  python monthly_line_chart.py results/web_reservation_leaf_20250901.csv
  python monthly_line_chart.py "results/予約データ.csv"

出力ファイル:
  - 月別集計.xlsx: 月別の集計データ + Excelグラフ

注意:
  - CSVファイルには以下の列が必要です:
    * appoint_datetime_clean: 予約日時
    * SYSTEM_USAGE_NAME: 受付種別（初回受付、再受付、初回キャンセル）
  - ファイル名にスペースが含まれる場合は引用符で囲んでください
""")

def main() -> int:
    """メイン処理"""
    try:
        # ヘルプ表示
        if len(sys.argv) > 1 and sys.argv[1] in ["--help", "-h", "-?"]:
            show_help()
            return 0
        
        # CSVファイルのパスを決定
        if len(sys.argv) > 1:
            csv_path = Path(sys.argv[1])
        else:
            # デフォルトファイルを検索
            results_dir = Path("results")
            csv_files = list(results_dir.glob("web_reservation_leaf_*.csv"))
            if not csv_files:
                print("❌ CSVファイルが見つかりません")
                print("💡 ヒント: python monthly_line_chart.py --help で使用方法を確認")
                return 1
            csv_path = csv_files[0]  # 最初に見つかったファイルを使用
        
        print(f"🎯 処理対象ファイル: {csv_path}")
        
        if not csv_path.exists():
            print(f"❌ ファイルが見つかりません: {csv_path}")
            return 1
        
        # データ読み込み
        df = load_csv_data(csv_path)
        if df is None:
            return 1
        
        # 日付解析
        df = parse_date_columns(df)
        if df is None:
            return 1
        
        # HPとその他で分けて処理
        site_types = ["hp", "other"]
        success_count = 0
        
        for site_type in site_types:
            print(f"\n{'='*50}")
            print(f"🔍 {site_type.upper()}経由データの処理開始")
            print(f"{'='*50}")
            
            # 月別集計
            monthly_summary = create_monthly_summary(df, site_type)
            if monthly_summary is None:
                print(f"⚠️ {site_type}経由データの集計に失敗")
                continue
            
            # 結果表示
            print(f"\n📊 {site_type}経由月別集計結果:")
            print(monthly_summary)
            
            # 出力ファイル名を生成
            if site_type == "hp":
                output_path = Path("月別集計_HP")
            else:
                output_path = Path("月別集計_その他")
            
            # Excelファイルに保存（Excelのグラフ機能を使用）
            excel_success = save_to_excel_with_chart(monthly_summary, output_path.with_suffix('.xlsx'), site_type)
            
            if excel_success:
                print(f"✅ {site_type}経由データの処理完了!")
                print(f"📊 Excelファイル: {output_path.with_suffix('.xlsx')}")
                print(f"📈 Excel内にグラフが作成されました")
                success_count += 1
            else:
                print(f"❌ {site_type}経由データの処理が失敗しました")
        
        if success_count > 0:
            print(f"\n🎉 処理完了! {success_count}/2 ファイルが正常に作成されました")
            return 0
        else:
            print(f"\n❌ 全ての処理が失敗しました")
            return 1
        
    except Exception as e:
        print(f"❌ メイン処理でエラー: {e}")
        return 1

if __name__ == "__main__":
    try:
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
