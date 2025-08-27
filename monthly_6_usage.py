from pathlib import Path
import pandas as pd


def find_latest_leaf_csv(results_dir: Path) -> Path | None:
    if not results_dir.exists():
        return None
    # 最新の web_reservation_leaf_*.csv を探索
    files = sorted(results_dir.glob("web_reservation_leaf_*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def load_csv(csv_path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(csv_path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(csv_path)


def get_datetime_series(df: pd.DataFrame) -> pd.Series:
    # 優先: appoint_timestamp、次点: appoint_datetime_clean、modify_start_date
    for col in ["appoint_timestamp", "appoint_datetime_clean", "modify_start_date"]:
        if col in df.columns:
            s = pd.to_datetime(df[col], errors="coerce")
            if s.notna().any():
                return s
    raise ValueError("日時列が見つからないか解析できません (appoint_timestamp)")


# ... existing code ...

def main() -> int:
    results_dir = Path("results")
    csv_path = find_latest_leaf_csv(results_dir)
    if csv_path is None:
        print("results/web_reservation_leaf_*.csv が見つかりません。先にデータ出力を実行してください。")
        return 1

    print(f"入力CSV: {csv_path}")
    df = load_csv(csv_path)

    # 必須列チェック
    for col in ["SYSTEM_USAGE_NAME", "SITE_CODE"]:
        if col not in df.columns:
            print(f"必須列が見つかりません: {col}")
            return 1

    dt = get_datetime_series(df)
    df = df.copy()
    df["__dt"] = dt

    # 現在の年月を取得
    current_date = pd.Timestamp.now()
    # 過去2か月分の年月を生成
    target_months = []
    for i in range(2):
        target_date = current_date - pd.DateOffset(months=i+1)
        target_months.append((target_date.year, target_date.month))
    
    print(f"処理対象月: {[f'{year}年{month}月' for year, month in target_months]}")

    def is_hp(x: object) -> bool:
        return str(x).strip().lower() == "hp"

    # サイト分類
    df["__site"] = df["SITE_CODE"].apply(lambda x: "hp" if is_hp(x) else "その他")

    # 各月ごとに処理
    for target_year, target_month in target_months:
        month_name = f"{target_month}月"
        
        # 該当月のフィルタ条件
        base_cond = (df["__dt"].dt.month == target_month) & (df["__dt"].dt.year == target_year)
        
        filters = {
            f"{month_name}_再受付_hp.xlsx": (df["SYSTEM_USAGE_NAME"] == "再受付") & (df["__site"] == "hp"),
            f"{month_name}_初回受付_hp.xlsx": (df["SYSTEM_USAGE_NAME"] == "初回受付") & (df["__site"] == "hp"),
            f"{month_name}_再受付_その他.xlsx": (df["SYSTEM_USAGE_NAME"] == "再受付") & (df["__site"] == "その他"),
            f"{month_name}_初回受付_その他.xlsx": (df["SYSTEM_USAGE_NAME"] == "初回受付") & (df["__site"] == "その他"),
        }

        results_dir.mkdir(parents=True, exist_ok=True)
        for filename, cond in filters.items():
            sub = df[base_cond & cond].drop(columns=["__dt", "__site"], errors="ignore")
            out_path = results_dir / filename
            try:
                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    sub.to_excel(writer, index=False, sheet_name="データ")
                print(f"出力: {out_path} (行数: {len(sub)})")
            except PermissionError:
                # ファイルが開かれていて書き込めない場合は別名で保存
                alt_name = f"{out_path.stem}_{pd.Timestamp.now():%Y%m%d%H%M%S}{out_path.suffix}"
                alt_path = results_dir / alt_name
                with pd.ExcelWriter(alt_path, engine="openpyxl") as writer:
                    sub.to_excel(writer, index=False, sheet_name="データ")
                print(f"警告: {out_path.name} が開かれているため、別名で出力しました → {alt_path} (行数: {len(sub)})")

    return 0

# ... existing code ...

if __name__ == "__main__":
    raise SystemExit(main())


