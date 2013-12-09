parse_excel_in_html
===================

This is a Simple Sample for parsing html, get the urls for excel files, download them and parse the excel files.
Please modify this program to use it in your code.

You need to import several packages to run this program.
* BeautifulSoup4
* xlrd

The easiest way to install these packages is using pip install

* pip install beautifulsoup4
* pip install xlrd

You can parse the html file of this link (http://www.e-stat.go.jp/SG1/estat/GL08020103.do?_toGL08020103_&listID=000001108886&disp=Other&requestSender=dsearch) which is unfortunately, Japanese. So it would be better you read the python code.

usage: python parse_excel_in_html.py html_file_path html_downloaded_url
example: python parse_excel_in_html.py  “OO.html”  “http://www.e-stat.go.jp/SG1/estat/GL08020103.do?_toGL08020103_&listID=000001108886&disp=Other&requestSender=dsearch”

You need to double quote the html file path and the url.
