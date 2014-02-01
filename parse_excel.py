# coding:utf-8
import os
from os import path
import sys
import re
import xlrd
import json
import codecs

'''
24年産水陸稲の時期別作柄及び収穫量（全国農業地域別・都道府県別）_水陸稲計
http://www.e-stat.go.jp/SG1/estat/Xlsdl.do?sinfid=000021160231

24年産麦類（子実用）の収穫量（全国農業地域別・都道府県別）_4麦計
http://www.e-stat.go.jp/SG1/estat/Xlsdl.do?sinfid=000021160241

24年産豆類（乾燥子実）及びそばの収穫量（全国農業地域別・都道府県別）_大豆
http://www.e-stat.go.jp/SG1/estat/Xlsdl.do?sinfid=000021160247

を解析するスクリプト
'''
def rm_unneed_char(u_str):
	'''全角の数字や空白、カタカナを削除'''
	return re.sub(u'[０１２３４５６７８９（）0123456789()アイウエオ　 ]+', '', u_str)

def crop_report_genre(u_str):
	'''全角スペースや半角スペースで区切られた最後の文字列を取得'''
	if u'　' in u_str:
		return u_str.split(u'　')[-1]
	elif u' ' in u_str:
		return u_str.split(u' ')[-1]
	else:
		return u_str

def excel_to_json(genre, excel_dir_path):
	if path.exists(excel_dir_path) == False: return
	for dirpath, dirnames, filenames in os.walk(excel_dir_path):
		for filename in filenames:
			filename_without_ext, ext = path.splitext(filename)
			if ext != '.xls' and ext != '.xlsx':
				print filename + ': extension has to be xls or xlsx'
				continue
			root_path = path.dirname(dirpath)
			json_dir_path = path.join(root_path, 'json')
			csv_dir_path = path.join(root_path, 'csv')
			json_file_path = path.join(json_dir_path, filename_without_ext + '.json')
			csv_file_path = path.join(csv_dir_path, filename_without_ext + '.csv')
			if path.exists(json_file_path) or path.exists(csv_file_path):
				print filename + ': file is already exported'
				continue
			if path.exists(json_dir_path) == False: os.mkdir(json_dir_path)
			if path.exists(csv_dir_path) == False: os.mkdir(csv_dir_path)
			
			excel_file_path = path.join(dirpath, filename)
			wb = xlrd.open_workbook(excel_file_path, formatting_info=True, on_demand=True)
			#print 'Number of Sheets ', len(wb.sheets())
			# 一つ目のシートだけをパース対象とする
			sheet = wb.sheet_by_index(0)

			# 文書の最大の列、行の数を求める
			max_col_idx = sheet.ncols
			max_row_idx = sheet.nrows
			# 出力されるjsonデータ
			json_data = {'data':[]}

			# 解析を始める行
			start_row = 4

			if genre in ['rice','wheat','soybean', 'soba']:
				json_data['goods'] = rm_unneed_char(sheet.cell(0,0).value if genre == 'rice' else sheet.cell(2,0).value )
				json_data['report_genre'] = crop_report_genre(sheet.cell(1,0).value),
				json_data['extraInfo'] = rm_unneed_char(sheet.cell(2,0).value if genre == 'rice' else sheet.cell(0,0).value )

				while sheet.cell(start_row, 0).value != u'全国':
					start_row += 1
					if start_row == max_row_idx:
						print "can't apply the algorism to this data"
						sys.exit(-1)

				# 重複している地域があるため、それを省く
				area_list = []
				for row_idx in range(start_row, max_row_idx):
					area = sheet.cell(row_idx, 0).value
					if area == '' or sheet.cell(row_idx, 1).value == '' or u'　' in sheet.cell(row_idx, 0).value:
						continue
					else:
						if genre in ['rice', 'wheat']:
							json_data['data'].append({	'area': area,
														'areaUnderCultivation': {'value':sheet.cell(row_idx, 1).value,'unit':'ha'}, 
														'yield': {'value':sheet.cell(row_idx, 2).value, 'unit':'t'}
							})
						elif genre in ['soybean', 'soba']:
							json_data['data'].append({	'area': area,
														'areaUnderCultivation': {'value':sheet.cell(row_idx, 1).value,'unit':'ha'},
														'yield_per_10a': {'value':sheet.cell(row_idx, 2).value, 'unit':'kg'},
														'yield': {'value':sheet.cell(row_idx, 3).value, 'unit':'t'}
							})
			elif genre in ['fruits']:
				json_data['report_genre'] = crop_report_genre(sheet.cell(1,0).value)
				json_data['goods'] = rm_unneed_char(sheet.cell(3,0).value)
				#json_data['extraInfo'] = sub_whitespace_number(sheet.cell(4,0).value)
				
				for row_idx in range(start_row, max_row_idx):
					read_col = 2
					# "全国"だけ異なる列にある
					if row_idx == start_row: read_col = 1
					area = sheet.cell(row_idx, read_col).value

					if area == '' or sheet.cell(row_idx, 4).value == '':
						continue
					else:
						json_data['data'].append({'area': area,
												'fruitingTreeArea': {'value':sheet.cell(row_idx, 4).value,'unit':'ha'},
												'yield_per_10a': {'value':sheet.cell(row_idx, 5).value, 'unit':'kg'},
												'yield': {'value':sheet.cell(row_idx, 6).value, 'unit':'t'},
												'shipment': {'value':sheet.cell(row_idx, 7).value, 'unit': 't'}
						})
			elif genre in ['vegetable']:
				json_data['report_genre'] = crop_report_genre(sheet.cell(2,0).value)
				json_data['goods'] = rm_unneed_char(sheet.cell(4,0).value)
				print json_data['goods']
				#json_data['extraInfo'] = sub_whitespace_number(sheet.cell(4,0).value)
				
				start_row = 14

				for row_idx in range(start_row, max_row_idx):
					area = sheet.cell(row_idx, 0).value

					if area == '' or sheet.cell(row_idx, 2).value == '' or sheet.cell(row_idx, 2).value == u'…':
						continue
					else:
						json_data['data'].append({'area': area,
												'plantedArea': {'value':sheet.cell(row_idx, 2).value,'unit':'ha'},
												'yield_per_10a': {'value':sheet.cell(row_idx, 3).value, 'unit':'kg'},
												'yield': {'value':sheet.cell(row_idx, 4).value, 'unit':'t'},
												'shipment': {'value':sheet.cell(row_idx, 5).value, 'unit': 't'}
						})


			#print json.dumps(json_data, ensure_ascii=False)
			with codecs.open(json_file_path, 'w', 'utf-8') as json_file:
				json.dump(json_data, json_file, indent=2, sort_keys=True, ensure_ascii=False)

			with codecs.open(csv_file_path, 'w', 'utf-8') as csv_file:
				# 品目コード, 品目名, 産地, 生産量(t), 出荷量(t), データの年度, 元データURL, 流通量(t), 市場シェア(%)
				for one_area_dict in json_data['data']:
					if 'shipment' not in one_area_dict:
						csv_file.write(', '.join(['', json_data['goods'], one_area_dict['area'], str(one_area_dict['yield']['value']), '', '2012', '', '', ''])+'\n')
					else:
						csv_file.write(', '.join(['', json_data['goods'], one_area_dict['area'], str(one_area_dict['yield']['value']), str(one_area_dict['shipment']['value']), '2012', '', '', ''])+'\n')


if __name__ == '__main__':
	args_len = len(sys.argv) 
	if args_len != 3:
		print 'usage: python parse_excel.py genre excel_dir_path'
		print 'genre: rice, wheat, soybean, soba, fruits, vegetable'
	else:
		genre = sys.argv[1]
		excel_dir_path = sys.argv[2]
		excel_to_json(genre, excel_dir_path)


