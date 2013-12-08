# coding:utf-8
import os
from os import path
import sys
import re
import urlparse
from bs4 import BeautifulSoup
import urllib2
import xlrd
import json
import codecs


'''
htmlからリンクを抽出し、そのリンクからエクセルファイルをダウンロード、
ダウンロードしたエクセルファイルを解析するプログラムのサンプルです。
例外処理などは殆どしていない ＆ 読み取るエクセルデータの形式が決まっていますが、
適宜変更を加えて、自分のプログラムに組み込んで下さい。

このプログラムが対象としたサイトは、農林水産省の輸入データを統計としてまとめた、
http://www.e-stat.go.jp/SG1/estat/GL08020103.do?_toGL08020103_&listID=000001108886&disp=Other&requestSender=dsearch
です。このサイトの
'''

def parse_excel_in_html(html_file_path, html_downloaded_url):
	json_list = []

	def get_excel_urls(html_file_path, html_downloaded_url = None):
		'''
		htmlからexcelに関係しそうなurlを取得する
		'''
		if path.exists(html_file_path) == False: raise Exception('No such path ' + html_file_path)
		with open(html_file_path, 'r') as html_file:
			html = html_file.read()
		soup = BeautifulSoup(html)
		# urlにxlsDownloadと含まれている物のみを取得する
		anchor_tags = soup.find_all('a', href = re.compile('xlsDownload'))
		print 'number of excel files link in this html ',len(anchor_tags)
		if html_downloaded_url == None:
			return [tag.attrs['href'] for tag in anchor_tags]
		else:
			# urlが相対パスの場合、ダウンロード元のurlも必要
			return [urlparse.urljoin(html_downloaded_url, tag.attrs['href']) for tag in anchor_tags]

	def download_and_process_from_urls(urls, tmp_folder_path, file_processing_func):
		'''
		URLのリストからファイルをダウンロードして、処理を行う。ダウンロードしたファイルは処理後に削除される。
		file_processing_func には、tmpファイルパスを引数とした関数を渡す
		
		'''
		if urls == None or len(urls) == 0: raise Exception("failed getting urls")
		if path.exists(tmp_folder_path) == False:
			os.mkdir(tmp_folder_path)
		for idx, url in enumerate(urls):
			tmp_file_path = path.join(tmp_folder_path, str(idx))
			response = urllib2.urlopen(url)
			with open(tmp_file_path, "wb") as tmp_file:
				tmp_file.write(response.read())
			file_processing_func(tmp_file_path)
			os.remove(tmp_file_path)
		os.rmdir(tmp_folder_path)

	def excel_to_json(excel_file_path):
		'''
		エクセルファイルを処理する（ファイル形式に応じて変更を加えること）
		'''
		wb = xlrd.open_workbook(excel_file_path,formatting_info=True,on_demand=True)
		#print 'Number of Sheets ', len(wb.sheets())
		# 一つ目のシートだけをパース対象とする
		sheet = wb.sheet_by_index(0)
		data_dict = {'report_name':sheet.cell(0,0).value, 'time':sheet.cell(1,0).value, 'title':sheet.cell(3,0).value, 'goods' : {} }
		col_idx = 1
		max_col_idx = sheet.ncols
		max_row_idx = sheet.nrows
		# print 'col ', max_col_idx, ', row ', max_row_idx
		while col_idx < max_col_idx:
			goods_name = sheet.cell(6, col_idx).value
			data_dict['goods'][goods_name] = []
			row_idx = 8
			while row_idx < max_row_idx:
				row_data = {}
				row_data[sheet.cell(6, 0).value] = sheet.cell(row_idx, 0).value
				row_data[sheet.cell(7, 1).value] = sheet.cell(row_idx, col_idx).value
				row_data[sheet.cell(7, 2).value] = sheet.cell(row_idx, col_idx+1).value
				row_data[sheet.cell(7, 3).value] = sheet.cell(row_idx, col_idx+2).value
				row_data[sheet.cell(7, 4).value] = sheet.cell(row_idx, col_idx+3).value
				row_data[sheet.cell(7, 5).value] = sheet.cell(row_idx, col_idx+4).value
				# 値が入っている場合だけ抽出
				if row_data[sheet.cell(7, 5).value]: data_dict['goods'][goods_name].append(row_data)
				row_idx += 1
			col_idx += 5
		json_list.append(data_dict)
	
	# urlの後部を削除（不必要なため）
	html_url = html_url.rsplit('/', 1)
	excel_urls = get_excel_urls(html_file_path, html_url)
	# print excel_urls[0:10]
	tmp_folder_path = path.join(path.dirname(html_file_path), 'tmp')
	download_and_process_from_urls(excel_urls, tmp_folder_path, excel_to_json)
	json_file_path = path.join(path.dirname(html_file_path), 'result.json')
	with codecs.open(json_file_path, 'w', 'utf-8') as json_file:
		json.dump(json_list, json_file, indent=2, sort_keys=True, ensure_ascii=False)


if __name__ == '__main__':
	if len(sys.argv) != 3:
		raise Exception('usage: python parse_excel_in_html.py html_file_path html_downloaded_url')
	html_file_path = sys.argv[1]
	html_url = sys.argv[2]
	parse_excel_in_html(html_file_path, html_url)	
