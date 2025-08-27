from pathlib import Path


def get_site_code_classification():
    """
    サイトコードマスタの分類情報を返す
    提携サイトコード → パワポ用分類名 のマッピング
    """
    return {
        # EPARK 接骨・鍼灸
        "epark": "EPARK 接骨・鍼灸",
        "refre": "EPARK 接骨・鍼灸",
        "taiken": "EPARK 接骨・鍼灸",
        "n.refre": "EPARK 接骨・鍼灸",
        "refre.jp": "EPARK 接骨・鍼灸",
        "r-ogc01": "EPARK 接骨・鍼灸",
        "r-ogj01": "EPARK 接骨・鍼灸",
        "r-leg": "EPARK 接骨・鍼灸",
        "r-ley": "EPARK 接骨・鍼灸",
        "k-t": "EPARK 接骨・鍼灸",
        "r_ogc01": "EPARK 接骨・鍼灸",
        "r_ogj01": "EPARK 接骨・鍼灸",
        "r_leg": "EPARK 接骨・鍼灸",
        "r_ley": "EPARK 接骨・鍼灸",
        "k_t": "EPARK 接骨・鍼灸",
        "gmb": "EPARK 接骨・鍼灸",  # 修正: EPARK接骨・鍼灸に分類
        
        # EPARK 連携メディア
        "hp": "EPARK 連携メディア",
        "innai": "EPARK 連携メディア",
        "CL": "EPARK 連携メディア",
        "refhp": "EPARK 連携メディア",
        "k-k": "EPARK 連携メディア",
        "k-se": "EPARK 連携メディア",
        "k-c": "EPARK 連携メディア",
        "k-m": "EPARK 連携メディア",
        "shinbus": "EPARK 連携メディア",
        "karadam": "EPARK 連携メディア",
        "lshinbus": "EPARK 連携メディア",
        "lkaradam": "EPARK 連携メディア",
        "refcp20": "EPARK 連携メディア",
        "k_k": "EPARK 連携メディア",
        "k_se": "EPARK 連携メディア",
        "k_c": "EPARK 連携メディア",
        "k_m": "EPARK 連携メディア",
        "faspa": "EPARK 連携メディア",  # 修正: EPARK連携メディアに分類
        "l.yahoo": "EPARK 連携メディア",  # 修正: EPARK連携メディアに分類
        "l_yahoo": "EPARK 連携メディア",  # 正規化後のパターン
        
        # EPARK アプリ
        "epark_app": "EPARK アプリ",
        "eparkapp": "EPARK アプリ",
        "app-fasp": "EPARK アプリ",  # 実際のデータに存在
        "app_fasp": "EPARK アプリ",  # 正規化後のパターン
        
        # 店頭割引クーポン
        "coupon": "店頭割引クーポン",
        "birthday": "店頭割引クーポン",
        "birthday_coupon": "店頭割引クーポン",
        "store_coupon": "店頭割引クーポン",
        "kuchikom": "店頭割引クーポン",  # 修正: 店頭割引クーポンに分類
        
        # その他の分類（ポイント系、モニター系など）
        "f-com": "その他",
        "f_com": "その他",
        "mm_01": "その他",
        "mm_02": "その他",
        "mm_03": "その他",
        "mm_04": "その他",
        "mm_05": "その他",
        "mm_06": "その他",
        "mm_07": "その他",
        "mm_08": "その他",
        "mm_09": "その他",
        "mm_10": "その他",
        "mm_11": "その他",
        "mm_12": "その他",
        "r_f_moni": "その他",
        "a_c_poin": "その他",
        "f_c_poin": "その他",
        "i_c_poin": "その他",
        "c_c_poin": "その他",
        "r_p_adwa": "その他",
        "r_p_ecna": "その他",
        "r_p_a8ne": "その他",
        "r_p_at": "その他",
        "r_epmail": "その他",
        "mm01": "その他",
        "mm02": "その他",
        "mm03": "その他",
        "mm04": "その他",
        "mm05": "その他",
        "mm06": "その他",
        "mm07": "その他",
        "mm08": "その他",
        "mm09": "その他",
        "mm10": "その他",
        "mm11": "その他",
        "mm12": "その他",
    }


def classify_site_code(site_code: str) -> str:
    """
    提携サイトコードを分類名に変換する
    マスタに存在しないコードはNoneを返す（出力しない）
    """
    if not site_code or not isinstance(site_code, str):
        return None
    
    # サイトコードを正規化（小文字化、アンダースコア化）
    normalized_code = site_code.lower().replace("-", "_").replace(".", "_")
    
    # 分類マスタから検索
    classification = get_site_code_classification()
    
    # 完全一致を試行
    if normalized_code in classification:
        return classification[normalized_code]
    
    # 部分一致を試行（より柔軟なマッチング）
    for code, category in classification.items():
        if code in normalized_code or normalized_code in code:
            return category
    
    # デフォルト分類（サイトコードの特徴から推測）
    if "epark" in normalized_code or "refre" in normalized_code:
        return "EPARK 接骨・鍼灸"
    elif "hp" in normalized_code or "innai" in normalized_code:
        return "EPARK 連携メディア"
    elif "mm_" in normalized_code or "poin" in normalized_code:
        return "その他"
    elif "coupon" in normalized_code or "birthday" in normalized_code:
        return "店頭割引クーポン"
    
    # マスタに存在しない場合はNoneを返す（出力しない）
    return None


def add_reception_pivot_pie(target_path: Path, chart_title: str = "サイト分類別割合") -> int:
    """
    既存のExcelファイル（.xlsx）を開いて、SITE_CODE×petano_log_no のピボット＆円グラフを
    タイムスタンプ付きの新規シートに作成します。
    提携サイトコードを分類して円グラフで出力します。
    """
    try:
        import win32com.client as win32
    except Exception:
        print("pywin32 が未インストールです。`pip install pywin32` を実行してください。")
        return 1

    abs_path = target_path.resolve()
    if not abs_path.exists():
        print(f"ファイルが見つかりません: {abs_path}")
        return 1

    # Excel 起動（非表示・警告非表示）
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Excel 定数
    xlDatabase = 1
    xlRowField = 1
    xlPie = 5

    wb = None
    try:
        wb = excel.Workbooks.Open(str(abs_path))

        # データシートの推定
        try:
            data_ws = wb.Worksheets("データ")
        except Exception:
            data_ws = wb.Worksheets(1)

        # 最終行・最終列を取得
        last_row = data_ws.Cells(data_ws.Rows.Count, 1).End(-4162).Row  # xlUp=-4162
        last_col = data_ws.Cells(1, data_ws.Columns.Count).End(-4159).Column  # xlToLeft=-4159
        src_range = data_ws.Range(data_ws.Cells(1, 1), data_ws.Cells(last_row, last_col))

        # SITE_CODE列の存在確認（1行目ヘッダを走査）
        header_names = []
        site_code_col = None
        for col in range(1, last_col + 1):
            header_val = data_ws.Cells(1, col).Value
            header_names.append(header_val)
            if header_val == "SITE_CODE":
                site_code_col = col

        if site_code_col is None:
            print("❌ SITE_CODE列が見つかりません！")
            wb.Close(SaveChanges=False)
            return 1

        print(f"📊 データ範囲: {last_row}行 × {last_col}列")
        print(f"📍 SITE_CODE列: {site_code_col}列目")

        # タイムスタンプ付きの新規シート作成
        from datetime import datetime
        pivot_sheet_name = f"ピボット_サイト分類円_{datetime.now():%Y%m%d%H%M%S}"
        pivot_ws = wb.Worksheets.Add()
        pivot_ws.Name = pivot_sheet_name

        # ピボットキャッシュ → ピボットテーブル
        pc = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=src_range)
        pt = pc.CreatePivotTable(TableDestination=pivot_ws.Range("A3"),
                                 TableName=f"PivotSITE_{pivot_sheet_name}")

        # 行フィールド: SITE_CODE
        try:
            pf = pt.PivotFields("SITE_CODE")
            pf.Orientation = xlRowField
            pf.Position = 1
        except Exception as e:
            print(f"❌ SITE_CODEフィールドの設定に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # データフィールド: petano_log_no をカウント
        try:
            # -4112 = xlCount
            data_field = pt.AddDataField(pt.PivotFields("petano_log_no"), "件数", -4112)
            data_field.Function = -4112
            data_field.NumberFormat = "0"
        except Exception as e:
            print(f"❌ データフィールドの設定に失敗: {e}")
            wb.Close(SaveChanges=False)
            return 1

        # ピボット更新（失敗しても続行）
        try:
            pt.RefreshTable()
        except Exception:
            pass

        # ピボット結果を固定範囲に展開して円グラフのデータソースにする
        header_row = 3
        used_last_row = pivot_ws.Cells(pivot_ws.Rows.Count, 1).End(-4162).Row  # A列の最終行
        if used_last_row <= header_row:
            print("ピボットのデータが空です。")
            wb.Close(SaveChanges=True)
            return 0

        print(f"📊 ピボット結果: {used_last_row - header_row}行のデータ")

        # 固定範囲ヘッダ
        out_col_start = 8  # H列
        pivot_ws.Cells(header_row, out_col_start).Value = "サイト分類"
        pivot_ws.Cells(header_row, out_col_start + 1).Value = "件数"
        pivot_ws.Cells(header_row, out_col_start + 2).Value = "割合(%)"

        # A列: ラベル, B列: 値 を読み取り
        src_first = header_row + 1
        row_count = used_last_row - header_row
        out_write_row = header_row + 1
        total_count = 0
        site_data = []
        raw_site_codes = []  # デバッグ用

        print("\n🔍 サイトコード分類処理:")
        for i in range(row_count):
            label_val = pivot_ws.Cells(src_first + i, 1).Value
            value_val = pivot_ws.Cells(src_first + i, 2).Value

            label_str = str(label_val).strip() if label_val is not None else ""
            if not label_str or label_str in ("総計", "Grand Total", "行ラベル"):
                continue
            # 値が数値なら採用
            if value_val is not None:
                try:
                    iv = int(float(value_val))
                    raw_site_codes.append(label_str)
                    
                    # サイトコードを分類名に変換
                    classification = classify_site_code(label_str)
                    if classification is not None:  # マスタに存在するコードのみ追加
                        site_data.append((classification, iv))
                        total_count += iv
                        print(f"  ✅ {label_str} → {classification} ({iv}件)")
                    else:
                        print(f"  ❌ {label_str} → マスタに存在しません（除外）")
                        # 除外されたコードの詳細分析
                        normalized = label_str.lower().replace("-", "_").replace(".", "_")
                        print(f"      💡 正規化後: '{normalized}'")
                        if "hp" in normalized or "innai" in normalized or "cl" in normalized:
                            print(f"      💡 推測: EPARK 連携メディア の可能性")
                        elif "coupon" in normalized or "birthday" in normalized or "store" in normalized:
                            print(f"      💡 推測: 店頭割引クーポン の可能性")
                        elif "app" in normalized or "faspa" in normalized:
                            print(f"      💡 推測: EPARK アプリ の可能性")
                        elif "epark" in normalized or "refre" in normalized:
                            print(f"      💡 推測: EPARK 接骨・鍼灸 の可能性")
                        else:
                            print(f"      💡 推測: その他 の可能性")
                except Exception as e:
                    print(f"  ⚠️ {label_str}: 値の変換エラー - {e}")

        # 分類別に集計
        classification_counts = {}
        classification_details = {}  # 各分類に含まれるサイトコードの詳細
        
        for classification, count in site_data:
            if classification in classification_counts:
                classification_counts[classification] += count
            else:
                classification_counts[classification] = count
                classification_details[classification] = []
        
        # 各分類に含まれるサイトコードを記録
        for i, (label_val, value_val) in enumerate(zip(raw_site_codes, [data[1] for data in site_data])):
            if label_val and value_val:
                classification = classify_site_code(label_val)
                if classification and classification in classification_details:
                    if label_val not in classification_details[classification]:
                        classification_details[classification].append(label_val)

        print(f"\n📊 分類結果サマリー:")
        print(f"  元のサイトコード数: {len(raw_site_codes)}")
        print(f"  分類済みデータ数: {len(site_data)}")
        print(f"  総件数: {total_count}")
        print(f"  分類カテゴリ数: {len(classification_counts)}")
        
        for category, count in classification_counts.items():
            percentage = (count / total_count * 100) if total_count > 0 else 0
            print(f"    {category}: {count}件 ({percentage:.1f}%)")
            if category in classification_details and classification_details[category]:
                print(f"      📋 含まれるサイトコード: {', '.join(classification_details[category])}")
            else:
                print(f"      📋 含まれるサイトコード: 詳細なし")

        if not classification_counts:
            print("❌ 分類可能なデータがありません")
            wb.Close(SaveChanges=True)
            return 0

        # データを書き出し（パーセンテージ計算付き）
        for classification, count in sorted(classification_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_count * 100) if total_count > 0 else 0
            pivot_ws.Cells(out_write_row, out_col_start).Value = classification
            pivot_ws.Cells(out_write_row, out_col_start + 1).Value = count
            pivot_ws.Cells(out_write_row, out_col_start + 2).Value = round(percentage, 1)
            out_write_row += 1

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
        chart.ChartTitle.Text = chart_title

        # データラベル設定（割合＋カテゴリ名、凡例は非表示）
        try:
            chart.ApplyDataLabels()
            chart.FullSeriesCollection(1).ApplyDataLabels()
            data_labels = chart.FullSeriesCollection(1).DataLabels()
            data_labels.ShowPercentage = True
            data_labels.ShowCategoryName = True
            data_labels.ShowValue = False
            data_labels.Position = -4142  # xlLabelPositionOutsideEnd
            try:
                data_labels.Format.TextFrame2.TextRange.Font.Size = 9
                data_labels.Format.Line.Visible = True
                data_labels.Format.Line.ForeColor.RGB = 0x000000
            except Exception:
                pass
            try:
                data_labels.Font.Size = 9
                data_labels.Font.Name = "Arial"
            except Exception:
                pass
            try:
                data_labels.ShowLeaderLines = True
            except Exception:
                pass
            chart.HasLegend = False
        except Exception as e:
            print(f"データラベル設定でエラー: {e}")
            try:
                chart.FullSeriesCollection(1).DataLabels().ShowPercentage = True
                chart.FullSeriesCollection(1).DataLabels().ShowCategoryName = True
                chart.HasLegend = False
            except Exception:
                pass

        # レイアウト調整
        shape.Left = pivot_ws.Range("D3").Left
        shape.Top = pivot_ws.Range("D3").Top
        shape.Width = 600
        shape.Height = 450

        wb.Close(SaveChanges=True)
        print(f"✅ 更新完了: {target_path}")
        print(f"📊 作成された分類: {list(classification_counts.keys())}")
        return 0

    except Exception as e:
        # 開かれて保存できない場合は別名保存
        try:
            if wb is not None:
                alt = target_path.with_name(target_path.stem + "_pivot" + target_path.suffix)
                wb.SaveCopyAs(str(alt))
                wb.Close(SaveChanges=False)
                print(f"⚠️ 警告: 元ファイルへ保存できないため複製を保存しました → {alt}")
                return 0
        except Exception:
            pass
        print(f"❌ 処理失敗: {e}")
        return 1
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def test_classification():
    """サイトコード分類のテスト"""
    test_cases = [
        ("epark", "EPARK 接骨・鍼灸"),
        ("refre", "EPARK 接骨・鍼灸"),
        ("hp", "EPARK 連携メディア"),
        ("innai", "EPARK 連携メディア"),
        ("mm_01", "その他"),
        ("coupon", "店頭割引クーポン"),
        ("unknown", None),  # マスタに存在しないコード
        ("", None),         # 空文字
        (None, None),       # None
    ]
    
    print("🧪 サイトコード分類テスト:")
    for input_code, expected in test_cases:
        result = classify_site_code(input_code)
        status = "✅" if result == expected else "❌"
        expected_str = str(expected) if expected is not None else "None（出力しない）"
        result_str = str(result) if result is not None else "None（出力しない）"
        print(f"  {status} 入力: '{input_code}' → 出力: '{result_str}' (期待値: '{expected_str}')")
    
    print("\n📊 分類マスタの内容:")
    classification = get_site_code_classification()
    for code, category in classification.items():
        print(f"  {code} → {category}")


def main() -> int:
    # ここは .xlsx を直接指定（CSVではありません）
    target = Path("results") / "6月_再受付_その他.xlsx"  # 6月初回受付その他ファイルに戻す
    print(f"🎯 処理対象ファイル: {target}")
    print(f"📁 ファイル存在確認: {target.exists()}")
    
    # タイトルは任意で差し替え可
    result = add_reception_pivot_pie(target, chart_title="サイト分類別割合")
    print(f"🏁 処理完了: 終了コード {result}")
    return result


if __name__ == "__main__":
    import sys
    
    # テストモードの場合はテストを実行
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        test_classification()
        raise SystemExit(0)
    
    raise SystemExit(main())
