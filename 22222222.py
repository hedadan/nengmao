import os
import xlrd
import xlwt
import requests
from bs4 import BeautifulSoup

def read_excel(dir_path, filename, result):
	file_path = os.path.join(dir_path, filename)
	print(file_path)
	wb = xlrd.open_workbook(file_path)
	sh = wb.sheet_by_index(0)

	nrows = sh.nrows
	ncols = sh.ncols

	for i in range(0, nrows):
		row = sh.row_values(i)
		if row[0]:
			result.append(row)

	return result

def write_xls(filepath, data):
	wb = xlwt.Workbook(encoding='utf-8')    
	sh = wb.add_sheet('Sheet')  
	  
	for i in range(len(data)): 
		for j in range(0, len(data[i])):
			sh.write(i, j, data[i][j])
			print("(%d/%d) data writed." % (i+1, len(data)))
	wb.save(filepath)









def request(url):
	price=[]
	u = requests.get(url)
	html_doc = u.text
	soup = BeautifulSoup(html_doc, 'html.parser')
	for link in soup.find_all('span',class_ = "tm-price"):
		price.depend(link.get_text())
	return price

def main():
	root_path = os.getcwd()
	file_list = os.listdir(root_path)
	result = []
	a = read_excel(root_path,file_list[0],result)
	for i in a:
		price = request(i[0])
		result[i].append(price)
	write_xls('result.xls', result)

if __name__ == '__main__':
	main()

