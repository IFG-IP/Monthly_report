import json
import os
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple

import boto3
from botocore.exceptions import ClientError


def _get_s3_client():
	try:
		return boto3.client("s3")
	except Exception as e:
		print(f"S3クライアント初期化失敗: {e}")
		return None




def _classify_site_code(site_code: Optional[str]) -> Optional[str]:
	if not site_code:
		return None
	code = str(site_code).replace('-', '_').lower().strip()
	m = {
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
		"gmb": "EPARK 接骨・鍼灸",
		"hp": "EPARK 連携メディア",
		"innai": "EPARK 連携メディア",
		"cl": "EPARK 連携メディア",
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
		"faspa": "EPARK 連携メディア",
		"l.yahoo": "EPARK 連携メディア",
		"l_yahoo": "EPARK 連携メディア",
		"epark_app": "EPARK アプリ",
		"eparkapp": "EPARK アプリ",
		"app-fasp": "EPARK アプリ",
		"app_fasp": "EPARK アプリ",
		"coupon": "店頭割引クーポン",
		"birthday": "店頭割引クーポン",
		"birthday_coupon": "店頭割引クーポン",
		"store_coupon": "店頭割引クーポン",
		"kuchikom": "店頭割引クーポン",
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
	return m.get(code)


def _add_site_pie_inplace(target_path: Path, chart_title: str = "サイト分類別割合") -> Dict:
	try:
		from openpyxl import load_workbook
		from openpyxl.chart import PieChart, Reference
		from openpyxl.chart.label import DataLabelList
	except ImportError:
		return {"success": False, "error": "openpyxl が未インストールです（Layerを確認）"}

	try:
		wb = load_workbook(str(target_path))
		try:
			data_ws = wb["データ"]
		except KeyError:
			data_ws = wb.worksheets[0]

		max_row = data_ws.max_row
		max_col = data_ws.max_column
		if max_row <= 1 or max_col <= 0:
			return {"success": False, "error": "データが不足"}

		site_code_col = None
		for col in range(1, max_col + 1):
			h = data_ws.cell(row=1, column=col).value
			if h == "SITE_CODE":
				site_code_col = col
				break
		if site_code_col is None:
			return {"success": False, "error": "SITE_CODE列が見つからない"}

		counts: Dict[str, int] = {}
		total = 0
		for r in range(2, max_row + 1):
			sc = data_ws.cell(row=r, column=site_code_col).value
			cat = _classify_site_code(sc)
			if cat:
				counts[cat] = counts.get(cat, 0) + 1
				total += 1
		if not counts:
			return {"success": False, "error": "有効なデータなし"}

		ts = datetime.now().strftime("%Y%m%d%H%M%S")
		pivot_name = f"ピボット_サイト分類円_{ts}"
		pivot_ws = wb.create_sheet(title=pivot_name)

		pivot_ws['A1'] = "サイト分類"
		pivot_ws['B1'] = "件数"
		pivot_ws['C1'] = "割合(%)"

		row = 2
		for name, cnt in sorted(counts.items(), key=lambda x: x[1], reverse=True):
			pct = (cnt / total * 100) if total else 0
			pivot_ws[f'A{row}'] = name
			pivot_ws[f'B{row}'] = cnt
			pivot_ws[f'C{row}'] = round(pct, 1)
			row += 1

		chart = PieChart()
		chart.title = chart_title
		data = Reference(pivot_ws, min_col=2, min_row=1, max_row=row-1)
		labels = Reference(pivot_ws, min_col=1, min_row=2, max_row=row-1)
		chart.add_data(data, titles_from_data=True)
		chart.set_categories(labels)
		chart.dataLabels = DataLabelList()
		chart.dataLabels.showPercent = True
		chart.dataLabels.showCatName = True
		
		# グラフのサイズを設定
		chart.width = 15
		chart.height = 10
		
		
		
		# グラフをより見やすい位置に配置
		pivot_ws.add_chart(chart, "A5")

		wb.save(str(target_path))
		return {"success": True, "timestamp": ts, "total_count": total, "categories": list(counts.keys())}
	except Exception as e:
		return {"success": False, "error": f"処理失敗: {e}"}


def _list_last_two_months_others(s3, bucket: str, prefix: str) -> List[Dict[str, str]]:
	pattern = re.compile(r"(\d+)月_(再受付|初回受付)_その他\.xlsx$")
	paginator = s3.get_paginator("list_objects_v2")
	by_month: Dict[int, List[Dict[str, str]]] = {}
	for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
		for obj in page.get("Contents", []):
			key = obj["Key"]
			name = key.split("/")[-1]
			m = pattern.search(name)
			if m:
				month = int(m.group(1))
				by_month.setdefault(month, []).append({"bucket": bucket, "key": key, "name": name})
	if not by_month:
		return []
	months = sorted(by_month.keys(), reverse=True)[:2]
	result: List[Dict[str, str]] = []
	for m in months:
		# 再受付→初回受付の順で安定化
		files = sorted(by_month[m], key=lambda x: ("初回受付" in x["name"], x["name"]))
		result.extend(files)
	return result


def lambda_handler(event, context):
	"""S3上の直近2か月分の『その他』Excelをまとめて処理し、結果を同一バケットへ保存。"""
	try:
		# 環境変数から取得、なければデフォルト値を設定
		s3_bucket = os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
		s3_prefix = os.environ.get('S3_OUTPUT_PREFIX', 'bigquery-data/')
		upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'results/')
		chart_title = (event or {}).get("chart_title", "サイト分類別割合")
		if not s3_bucket:
			return {"statusCode": 400, "body": json.dumps({"success": False, "error": "s3_bucketが未指定"}, ensure_ascii=False)}

		s3 = _get_s3_client()
		if s3 is None:
			return {"statusCode": 500, "body": json.dumps({"success": False, "error": "boto3初期化失敗"}, ensure_ascii=False)}

		# 直近2か月分の"その他"ファイルを抽出
		targets = _list_last_two_months_others(s3, s3_bucket, s3_prefix)
		if not targets:
			return {"statusCode": 404, "body": json.dumps({"success": False, "error": "対象Excelが見つかりません"}, ensure_ascii=False)}

		processed: List[Dict[str, str]] = []
		errors: List[str] = []
		for t in targets:
			key = t["key"]
			name = t["name"]
			local = Path("/tmp") / name
			try:
				# DL
				s3.download_file(s3_bucket, key, str(local))
				# 加工
				res = _add_site_pie_inplace(local, chart_title)
				if not res.get("success"):
					errors.append(f"{name}: {res.get('error')}")
					continue
				# UL（別名で保存）
				ts = res.get("timestamp")
				out_key = f"{upload_prefix.rstrip('/')}/site_pivot_{ts}_{name}"
				s3.upload_file(str(local), s3_bucket, out_key)
				processed.append({"bucket": s3_bucket, "key": out_key, "source_key": key})
			except Exception as e:
				errors.append(f"{name}: {e}")

		status = 207 if errors and processed else (200 if processed else 500)
		return {
			"statusCode": status,
			"body": json.dumps({
				"success": bool(processed),
				"processed": processed,
				"errors": errors,
			}, ensure_ascii=False),
		}
	except Exception as e:
		return {"statusCode": 500, "body": json.dumps({"success": False, "error": f"予期しないエラー: {e}"}, ensure_ascii=False)}
