import json
import os
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple

try:
    import boto3
    from botocore.exceptions import ClientError
except ImportError:
    # ローカル環境ではboto3が不要
    boto3 = None
    ClientError = Exception


def _get_s3_client():
	try:
		if boto3 is None:
			return None
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
		print(f"    🔍 サイトグラフ作成開始: {target_path}")
		wb = load_workbook(str(target_path))
		print(f"    📋 全シート一覧: {wb.sheetnames}")
		
		# 統合されたExcelファイルの場合、各データシートごとに個別の円グラフを作成
		data_sheets = []
		for sheet_name in wb.sheetnames:
			# データシートを特定（グラフシートやピボットシートを除外）
			if not any(keyword in sheet_name.lower() for keyword in ['ピボット', 'グラフ', 'chart', 'pivot']):
				data_sheets.append(wb[sheet_name])
				print(f"    📊 データシートとして認識: {sheet_name}")
		
		print(f"    📊 データシート数: {len(data_sheets)}")
		if not data_sheets:
			return {"success": False, "error": "データシートが見つかりません"}

		# 各データシートごとに個別の円グラフを作成
		created_charts = 0
		for i, data_ws in enumerate(data_sheets):
			# その他シートのみを対象とする
			if "その他" not in data_ws.title:
				print(f"    ⚠️ シート {i+1} はその他シートではないためスキップ: {data_ws.title}")
				continue
				
			print(f"    📊 シート {i+1} 処理中: {data_ws.title}")
			max_row = data_ws.max_row
			max_col = data_ws.max_column
			print(f"    📊 シートサイズ: {max_row}行 x {max_col}列")
			
			if max_row <= 1 or max_col <= 0:
				print(f"    ⚠️ シート {i+1} はデータが不足")
				continue

			site_code_col = None
			for col in range(1, max_col + 1):
				h = data_ws.cell(row=1, column=col).value
				if h == "SITE_CODE":
					site_code_col = col
					break
			
			if site_code_col is None:
				print(f"    ⚠️ シート {i+1} にSITE_CODE列が見つからない")
				continue
			
			print(f"    📊 SITE_CODE列: {site_code_col}列目")

			# このシートからデータを収集
			counts: Dict[str, int] = {}
			total = 0
			sheet_count = 0
			for r in range(2, max_row + 1):
				sc = data_ws.cell(row=r, column=site_code_col).value
				cat = _classify_site_code(sc)
				if cat:
					counts[cat] = counts.get(cat, 0) + 1
					total += 1
					sheet_count += 1
			
			print(f"    📊 シート {i+1} から収集: {sheet_count}件")
			
			if not counts:
				print(f"    ⚠️ シート {i+1} に有効なデータなし")
				continue

			# このシート用のピボットシートを作成
			ts = datetime.now().strftime("%Y%m%d%H%M%S")
			pivot_name = f"ピボット_サイト分類円_{data_ws.title}_{ts}"
			# Excelのシート名制限（31文字）に合わせる
			if len(pivot_name) > 31:
				pivot_name = f"サイト分類円_{data_ws.title}_{ts}"[:31]
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

			# このシート用の円グラフを作成
			chart = PieChart()
			chart.title = f"{data_ws.title} - {chart_title}"
			
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
			
			chart.dataLabels = DataLabelList()
			chart.dataLabels.showPercent = True
			chart.dataLabels.showCatName = True
			
			# 凡例を非表示にする
			chart.legend = None
			
			# グラフのサイズを設定（大きめに調整）
			chart.width = 20
			chart.height = 15
			
			# グラフをより見やすい位置に配置
			pivot_ws.add_chart(chart, "A5")
			
			created_charts += 1
			print(f"    ✅ シート {i+1} の円グラフ作成完了: {pivot_name}")

		print(f"    📊 作成された円グラフ数: {created_charts}")
		if created_charts == 0:
			return {"success": False, "error": "その他シートが見つからないか、有効なデータなし"}

		wb.save(str(target_path))
		return {"success": True, "timestamp": ts, "created_charts": created_charts}
	except Exception as e:
		return {"success": False, "error": f"処理失敗: {e}"}


def _list_last_two_months_others(s3, bucket: str, prefix: str) -> List[Dict[str, str]]:
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
	"""Step Functionsから呼び出されるLambda関数 - サイト分類別円グラフ作成"""
	try:
		print(f"🔍 デバッグ: イベント受信: {event}")
		
		# 環境変数から取得、なければデフォルト値を設定
		s3_bucket = os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
		s3_prefix = os.environ.get('S3_OUTPUT_PREFIX', 'bigquery-data/')
		upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', 'bigquery-data/')
		chart_title = (event or {}).get("chart_title", "サイト分類別割合")
		chart_style = (event or {}).get("chart_style", "vibrant")
		
		if not s3_bucket:
			return {
				"success": False,
				"error": "s3_bucketが未指定"
			}

		s3 = _get_s3_client()
		if s3 is None:
			return {
				"success": False,
				"error": "boto3初期化失敗"
			}

		# Step Functions用の入力形式に対応
		input_files = event.get('input_files', event.get('input', event.get('uploaded', [])))
		
		# 入力ファイルが指定されていない場合は、5月・6月分の"その他"ファイルを自動検索
		if not input_files:
			print("🔍 デバッグ: 入力ファイルが指定されていないため、5月・6月分のファイルを自動検索")
			targets = _list_last_two_months_others(s3, s3_bucket, s3_prefix)
			if not targets:
				return {
					"success": False,
					"error": "対象Excelが見つかりません"
				}
			# 自動検索結果をinput_files形式に変換
			input_files = [f"s3://{s3_bucket}/{t['key']}" for t in targets]

		print(f"🔍 デバッグ: 入力ファイル数: {len(input_files)}")
		
		# 処理対象ファイルを準備
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
				
				files_to_process.append({
					's3_path': file_path,
					'bucket': bucket_name,
					'key': s3_key,
					'filename': s3_key.split('/')[-1]
				})

		if not files_to_process:
			print("❌ 処理対象ファイルがありません")
			return {
				"success": False,
				"error": "処理対象ファイルがありません"
			}

		print(f"🔍 デバッグ: 処理対象ファイル数: {len(files_to_process)}")
		
		# 各ファイルを処理
		processed_files = []
		output_files = []
		errors = []
		
		for file_info in files_to_process:
			try:
				print(f"🔍 デバッグ: 処理中: {file_info['s3_path']}")
				
				# ファイルをダウンロード
				local_path = Path("/tmp") / file_info['filename']
				s3.download_file(file_info['bucket'], file_info['key'], str(local_path))
				print(f"🔍 デバッグ: ダウンロード成功: {local_path}")
				
				# グラフを追加
				print(f"🎯 グラフ作成開始: {file_info['filename']}")
				result = _add_site_pie_inplace(local_path, chart_title)
				
				if result.get('success'):
					print(f"✅ グラフ作成完了: {result.get('message', '')}")
					processed_files.append(file_info['s3_path'])
					
					# 同じファイルをS3に上書き保存
					s3.upload_file(str(local_path), s3_bucket, file_info['key'])
					
					output_path = f"s3://{s3_bucket}/{file_info['key']}"
					output_files.append(output_path)
					print(f"✅ S3上書き保存成功: {output_path}")
				else:
					error_msg = f"グラフ作成失敗: {result.get('error', '')}"
					print(f"❌ {error_msg}")
					errors.append(error_msg)
					
			except Exception as e:
				error_msg = f"ファイル処理に失敗: {file_info['s3_path']} - {str(e)}"
				print(f"❌ {error_msg}")
				errors.append(error_msg)
				import traceback
				print(f"🔍 デバッグ: エラー詳細: {traceback.format_exc()}")

		print(f"🔍 デバッグ: 最終結果 - 処理済み: {len(processed_files)}, 出力: {len(output_files)}, エラー: {len(errors)}")
		
		# Step Functions用の出力形式
		response = {
			"success": len(errors) == 0,
			"input_files": input_files,
			"processed_files": processed_files,
			"output_files": output_files,
			"count": len(processed_files),
			"errors": errors,
			"timestamp": datetime.now().strftime("%Y%m%d%H%M%S")
		}
		
		print(f"🔍 デバッグ: レスポンス: {response}")
		return response
		
	except Exception as e:
		error_msg = f"予期しないエラー: {str(e)}"
		print(f"❌ {error_msg}")
		import traceback
		print(f"🔍 デバッグ: エラー詳細: {traceback.format_exc()}")
		
		return {
			"success": False,
			"error": error_msg
		}
