import json
import os
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any

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

def load_group_companies_config_from_s3(bucket: str, key: str) -> Dict[str, Any]:
	"""S3からグループ会社設定JSONファイルを読み込む"""
	try:
		s3 = _get_s3_client()
		if s3 is None:
			return {}
		response = s3.get_object(Bucket=bucket, Key=key)
		config_data = json.loads(response['Body'].read().decode('utf-8'))
		print(f"✅ 設定ファイル読み込み成功: {key}")
		return config_data
	except Exception as e:
		print(f"❌ 設定ファイル読み込み失敗: {e}")
		return {}

def get_company_name_from_group_id(group_id: str, config: Dict[str, Any]) -> str:
	"""グループIDから会社の正式名称を取得"""
	group_companies = config.get("group_companies", {})
	if group_id in group_companies:
		return group_companies[group_id].get("name", group_id)
	return group_id




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


def _add_site_pie_inplace(target_path: Path, chart_title: str = "サイト分類別割合", group_id: str = None, company_name: str = None) -> Dict:
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
			ts = datetime.now().strftime("%Y%m")
			pivot_name = f"ピボット_サイト分類円_{data_ws.title}_{ts}"
			# Excelのシート名制限（31文字）に合わせる
			if len(pivot_name) > 31:
				pivot_name = f"サイト分類円_{data_ws.title}_{ts}"[:31]
			
			# 既存のシートをチェックして上書き
			if pivot_name in wb.sheetnames:
				print(f"    📊 既存シートを上書き: '{pivot_name}'")
				wb.remove(wb[pivot_name])
			
			print(f"    📊 新しいシート作成: '{pivot_name}'")
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


def _get_target_months_from_data(data: List[Dict]) -> List[Tuple[int, int]]:
    """直近2か月の対象月を取得する関数（7月と8月）"""
    try:
        from zoneinfo import ZoneInfo
        JST = ZoneInfo("Asia/Tokyo")
    except:
        from datetime import timezone, timedelta
        JST = timezone(timedelta(hours=9))
    
    # 現在時刻をJSTで取得
    now = datetime.now(JST)
    current_year = now.year
    current_month = now.month
    
    # 直近2か月を計算（7月と8月）
    target_months = []
    
    # 2か月前の月（7月）
    if current_month <= 2:
        two_months_ago_year = current_year - 1
        two_months_ago_month = current_month + 10  # 12 + (month - 2)
    else:
        two_months_ago_year = current_year
        two_months_ago_month = current_month - 2
    
    # 前の月（8月）
    if current_month == 1:
        previous_month_year = current_year - 1
        previous_month = 12
    else:
        previous_month_year = current_year
        previous_month = current_month - 1
    
    # 対象月を追加（7月と8月）
    target_months.append((two_months_ago_year, two_months_ago_month))
    target_months.append((previous_month_year, previous_month))
    
    print(f"🔍 対象月: {target_months}")
    return target_months


def parse_datetime(date_str: str) -> Optional[datetime]:
	"""日時文字列を datetime に変換（ISO8601/TZ対応）"""
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


def load_csv_from_s3(bucket: str, key: str) -> Optional[List[Dict]]:
	"""S3 から CSV を読み込んで辞書リストとして返す"""
	s3_client = _get_s3_client()
	if s3_client is None:
		return None
	try:
		obj = s3_client.get_object(Bucket=bucket, Key=key)
		body = obj["Body"].read().decode("utf-8-sig")
		import csv
		import io
		reader = csv.DictReader(io.StringIO(body))
		return list(reader)
	except Exception:
		try:
			obj = s3_client.get_object(Bucket=bucket, Key=key)
			body = obj["Body"].read().decode("utf-8")
			import csv
			import io
			reader = csv.DictReader(io.StringIO(body))
			return list(reader)
		except Exception:
			return None


def _list_last_two_months_others(s3, bucket: str, prefix: str) -> List[Dict[str, str]]:
	"""直近2か月分の統合データファイルを検索する関数（複数グループ会社対応）"""
	# 直近2か月の対象月を取得（7月と8月）
	target_months = _get_target_months_from_data([])
	print(f"🔍 対象月: {target_months}")
	
	# 統合データファイルを検索（グループID対応）
	pattern = re.compile(r"(?:(\w+)_)?(\d+)月_統合データ\.xlsx$")
	by_month: Dict[Tuple[int, int], List[Dict[str, str]]] = {}
	
	# 現在の月のフォルダ内で検索（2025_09/）
	from datetime import datetime
	current_date = datetime.now()
	current_year_month = f"{current_date.year}_{current_date.month:02d}"
	current_prefix = f"{current_year_month}/"
	print(f"🔍 検索中: {current_prefix}")
	
	for page in s3.get_paginator("list_objects_v2").paginate(Bucket=bucket, Prefix=current_prefix):
		for obj in page.get("Contents", []):
			key = obj["Key"]
			name = key.split("/")[-1]
			m = pattern.search(name)
			if m:
				group_id = m.group(1)  # グループID（オプション）
				month_num = int(m.group(2))
				
				# 対象月に含まれるかチェック（7月と8月）
				for year, month in target_months:
					if month_num == month:
						file_info = {
							"bucket": bucket, 
							"key": key, 
							"name": name,
							"group_id": group_id,
							"s3_path": f"s3://{bucket}/{key}"
						}
						by_month.setdefault((year, month), []).append(file_info)
						print(f"  ✅ 発見: {name} (グループ: {group_id})")
						break
	
	if not by_month:
		print("⚠️ 直近2か月の統合データファイルが見つかりませんでした")
		return []
	
	# 対象月の順で結果を返す
	result: List[Dict[str, str]] = []
	for year, month in target_months:
		if (year, month) in by_month:
			files = by_month[(year, month)]
			result.extend(files)
			print(f"✅ {year}年{month}月の統合データファイル: {len(files)}件")
		else:
			print(f"⚠️ {year}年{month}月の統合データファイルが見つかりませんでした")
	
	return result


def lambda_handler(event, context):
	"""Step Functionsから呼び出されるLambda関数 - サイト分類別円グラフ作成"""
	try:
		print(f"🔍 デバッグ: イベント受信: {event}")
		
		# 環境変数から取得、なければデフォルト値を設定
		s3_bucket = os.environ.get('S3_BUCKET_NAME', 'monthly-report-output-group')
		# 直近2か月の年月を動的に生成（7月と8月）
		from datetime import datetime
		current_date = datetime.now()
		current_year = current_date.year
		current_month = current_date.month
		
		# 2か月前の月を計算（7月と8月）
		if current_month <= 2:
			two_months_ago_year = current_year - 1
			two_months_ago_month = current_month + 10  # 12 + (month - 2)
		else:
			two_months_ago_year = current_year
			two_months_ago_month = current_month - 2
		
		# 前の月を計算
		if current_month == 1:
			previous_month_year = current_year - 1
			previous_month = 12
		else:
			previous_month_year = current_year
			previous_month = current_month - 1
		
		# 直近2か月のフォルダを検索対象とする（7月と8月）
		two_months_ago_year_month = f"{two_months_ago_year}_{two_months_ago_month:02d}"
		previous_year_month = f"{previous_month_year}_{previous_month:02d}"
		
		s3_prefix = os.environ.get('S3_OUTPUT_PREFIX', f'{two_months_ago_year_month}/')
		upload_prefix = os.environ.get('S3_UPLOAD_PREFIX', f'{previous_year_month}/')
		chart_title = (event or {}).get("chart_title", "サイト分類別割合")
		chart_style = (event or {}).get("chart_style", "vibrant")
		
		# グループ会社設定を読み込み
		config_key = "bigquery-data/group_companies_config.json"
		group_config = load_group_companies_config_from_s3(s3_bucket, config_key)
		
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
		
		# グループID別にファイルを分類
		group_files = {}
		for file_info in files_to_process:
			filename = file_info['filename']
			# ファイル名からグループIDを抽出 (例: aoba_11月_統合データ.xlsx)
			name_parts = filename.replace('.xlsx', '').split('_')
			if len(name_parts) >= 2:
				group_id = name_parts[0]
				if group_id not in group_files:
					group_files[group_id] = []
				group_files[group_id].append(file_info)
			else:
				# グループIDがない場合は"default"として処理
				if "default" not in group_files:
					group_files["default"] = []
				group_files["default"].append(file_info)
		
		print(f"📁 グループ別ファイル: {[(group, [f['filename'] for f in files]) for group, files in group_files.items()]}")
		
		# 各グループのファイルを処理
		processed_files = []
		output_files = []
		errors = []
		
		for group_id, files in group_files.items():
			print(f"\n=== グループ {group_id} の処理 ===")
			
			# グループIDから会社の正式名称を取得
			company_name = get_company_name_from_group_id(group_id, group_config)
			print(f"🏢 会社名: {company_name}")
			
			for i, file_info in enumerate(files, 1):
				try:
					print(f"🔍 デバッグ: 処理中 ({i}/{len(files)}): {file_info['s3_path']}")
					
					# ファイルをダウンロード
					local_path = Path("/tmp") / file_info['filename']
					s3.download_file(file_info['bucket'], file_info['key'], str(local_path))
					print(f"🔍 デバッグ: ダウンロード成功: {local_path}")
					
					# グラフを追加
					print(f"🎯 グラフ作成開始: {file_info['filename']}")
					result = _add_site_pie_inplace(local_path, chart_title, group_id, company_name)
					
					if result.get('success'):
						print(f"✅ グラフ作成完了: {result.get('message', '')}")
						processed_files.append(file_info['s3_path'])
						
						# 年月と会社名を含むパスで保存
						from datetime import datetime
						current_date = datetime.now()
						year_month = f"{current_date.year}_{current_date.month:02d}"
						company_folder = company_name or group_id or "unknown"
						new_prefix = f"{year_month}/{company_folder}/"
						
						# ファイル名を新しいパスで保存
						filename = file_info['key'].split('/')[-1]
						new_key = f"{new_prefix}{filename}"
						s3.upload_file(str(local_path), s3_bucket, new_key)
						
						output_path = f"s3://{s3_bucket}/{new_key}"
						output_files.append(output_path)
						print(f"✅ S3保存成功: {output_path}")
					else:
						error_msg = f"グラフ作成失敗: {result.get('error', '')}"
						print(f"❌ {error_msg}")
						errors.append(f"{file_info['s3_path']}: {error_msg}")
					
					# 一時ファイルを削除
					if local_path.exists():
						local_path.unlink()
						
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
			"timestamp": datetime.now().strftime("%Y%m")
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
