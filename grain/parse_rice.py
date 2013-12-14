# coding:utf-8
import os
from os import path
import sys
import re
import xlrd
import json
import codecs

'''
水陸稲の時期別作柄及び収穫量_全国農業地域別_都道府県別_水陸稲計
http://www.e-stat.go.jp/SG1/estat/Xlsdl.do?sinfid=000021160231
を解析するスクリプト
'''
def sub_whitespace_number(u_str):
	'''全角の数字や空白を削除'''
	result = re.sub(u'[１２３４５６７８９　]+', '', u_str)
	return result

def excel_to_json(excel_file_path):
	wb = xlrd.open_workbook(excel_file_path,formatting_info=True,on_demand=True)
	#print 'Number of Sheets ', len(wb.sheets())

	# 一つ目のシートだけをパース対象とする. (row, col) (行, 列)
	sheet = wb.sheet_by_index(0)
	json_data = {'goods': sub_whitespace_number(sheet.cell(0,0).value), 
				'report_name': u'24年産水陸稲の時期別作柄及び収穫量（全国農業地域別・都道府県別）',
				'extraInfo': u'水陸稲計',
				'data' : [] }
	max_col_idx = 3#sheet.ncols
	max_row_idx = sheet.nrows
	# print 'col ', max_col_idx, ', row ', max_row_idx
	# 重複している地域があるため、それを省く
	area_list = []
	for row_idx in range(11, max_row_idx):
		area = sheet.cell(row_idx, 0).value
		if area == '' or u'　' in area or u'(' in area or area in area_list:
			continue
		else:
			json_data['data'].append({'area': sheet.cell(row_idx, 0).value, 
									'areaUnderCultivation': {'value':sheet.cell(row_idx, 1).value,'unit':'ha'} , 
									'yield': {'value':sheet.cell(row_idx, 2).value, 'unit':'t'}})
	
	json_file_path = path.join(path.dirname(excel_file_path), '水陸稲の時期別作柄及び収穫量_全国農業地域別_都道府県別_水陸稲計.json')
	#print json.dumps(json_data, ensure_ascii=False)
	with codecs.open(json_file_path, 'w', 'utf-8') as json_file:
		json.dump(json_data, json_file, indent=2, sort_keys=True, ensure_ascii=False)

if __name__ == '__main__':
	args_len = len(sys.argv) 
	if args_len != 2:
		print 'usage: python parse_rice.py excel_path'
	else:
		excel_path = sys.argv[1]
		ext = path.splitext(excel_path)[1]
		if ext == '.xls' or ext == '.xlsx':
			excel_to_json(excel_path)
		else:
			print 'extension has to be xls or xlsx'

