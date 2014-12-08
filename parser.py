# coding: utf-8
try:
	from urllib.request import urlopen
	from urllib.parse import urljoin
except ImportError:
	from urllib2 import urlopen	
	from urlparse import urljoin
import xlsxwriter


from lxml.html import fromstring

#paragraph for sotovik
URL = 'http://shop.siriust.ru/index.php/cPath/372'
ITEM_PATH = '.category_item a'
ITEM_PATH_MODEL = '.mikro'
ITEM_PATH_KAZAN = '.nalichie_box'
items = ['']

def parse_courses():
	f = urlopen(URL)
	list_html = u'%s' %f.read().decode('cp1251')
	list_doc = fromstring(list_html)
	model_last = ''
	for elem in list_doc.cssselect(ITEM_PATH):
		a = elem.cssselect('a')[0]
		href = (a.get('href').split('?'))[0]
		span = elem.cssselect('.menu_link')[0]
		name = span.text_content()
		name = (name.replace('\n', '')).replace(' ', '', 11).replace('.', '')

		sotovik = {'name': name, 'href': href}
		firmselect_html = urlopen(href).read().decode('cp1251')
		firm_doc = fromstring(firmselect_html)
		for elem in firm_doc.cssselect(ITEM_PATH):
			a = elem.cssselect('a')[0]
			href1 = a.get('href').split('?')[0]
			i=0
			breaker = False
			if href1.count('_')==2:
				span = elem.cssselect('.menu_link')[0]
				firm_name = span.text_content()
				firm_name = (firm_name.replace('\n', '')).replace(' ', '', 11).replace('.', '')
				items.append(firm_name)
				no=1
				if not(breaker):
					while no < 11:
						no+=1
						modelselect_html = urlopen('%s/sort/1a/page/%s' %(href1, no)).read().decode('cp1251')				 
						model_doc = fromstring(modelselect_html)
						for elem in model_doc.cssselect(ITEM_PATH_MODEL):
							a = elem.cssselect('a')[0]
							href2 = a.get('href').split('?')[0]
							model_name = a.text_content()
							if not('%s| %s| %s' %(name, firm_name , model_name) in items):
								model_last = model_name

								kazan_html = urlopen(href2).read().decode('cp1251')
								kazan_doc = fromstring(kazan_html)
								breaker_kaz = False
								for elem in kazan_doc.cssselect(ITEM_PATH_KAZAN):
									a = elem.cssselect('li a')
									for n in a:
										href3 = n.get('href')
										nalichie = n.text_content()
										if href3.count('kazan')>0:
											item = '%s| %s| %s' %(name, firm_name , model_name)
											if not(item in items):
												print '%s, %s, %s' %(name, firm_name , model_name)
												items.append(item)
												breaker_kaz = True
												break
											else:
												i=+1
												print i
												if i == 1:
													breaker = True
													breaker_kaz = True
													print i
													break
									if breaker_kaz:
										break
				else:
					print('bingo')
					break
	return items

def export_excel(filename, parse_courses):
	workbook = xlsxwriter.Workbook(filename)
	worksheet = workbook.add_worksheet()
	i=0
	for item in items:
		worksheet.write(i, 0, item)
		i+=1
	workbook.close()

def main():
	items = parse_courses()
	export_excel('items.xlsx', items)


if __name__ == '__main__':
	main()