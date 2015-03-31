# -*- coding: UTF-8 -*-
from pyxll import xl_func, xl_menu, get_config, xlcAlert, get_active_object
import pandas	
import json
import jsonpickle
import requests
import pickle
from xlwings import Workbook, Sheet, Range

from win32com.client import Dispatch
from win32com.client import constants as c

from pywintypes import com_error

from pprint import pprint
from objbrowser import browse

@xl_func("string name: string")
def hello(name):
	""" return a familiar greeting"""
	return 'Hello, %s' % name

@xl_menu("Another example menu item", order=1)
def on_example_menu_item_2():
    xlcAlert("Hello again from PyXLL")

@xl_menu("Publish", menu="APIquitous", menu_order=1)
def publish_to_cloud():
	"""
	Publish a data structure to APIquitous' 
	"""

	# Initialize the application
	app = Dispatch("Excel.Application")
	wb = app.ActiveWorkbook
	ws = wb.ActiveSheet
	sheet_name = wb.ActiveSheet.Name

	wings_wb = Workbook(xl_workbook=wb)

	# Create lists needed to iterate over the template
	template_rows = wb.Sheets(sheet_name + '.AQ').Range('A1').End(-4121).Row
	publishing_keys_cells = ['B' + str(x) for x in range(1, template_rows + 1)]
	publishing_keys = wb.Sheets(sheet_name + '.AQ').Range('A1:A' + str(template_rows)).Value
	publishing_keys = ['aq_' + x[0].lower().replace(' ', '_') for x in publishing_keys]

	# Collect information from the template
	data = {}
	for i, j in zip(publishing_keys_cells, publishing_keys):
		k = str(wb.Sheets(sheet_name + '.AQ').Range(i).Value)
		if j == 'aq_name' or j == 'aq_organization':
			data[j] = k.replace(' ', '_')
		else:
			data[j] = k
	# TO-DO: This is a problematic part. You should not pickle on the client, because the expectation is
	# that the client when fetching in the later stage should receive the resources in JSON format.
	# Possible solution is to unpickle at the server, JSONfy and then send it to the client.
	
	#data['aq_data'] = Range('Data', 'E3:E23').value
	#data['aq_data'] = Range('Data', 'E3:E23').value
	#print data['aq_data'][0]
	#print type(data['aq_data'][0])
	print app.Selection.Address
	data['aq_data'] = Range(wb.ActiveSheet.Name, app.Selection.Address).value

	# Find out what type of data is is being published
	cols = app.Selection.Columns.Count
	print cols
	if cols == 1:
		data['aq_type'] = 'List'
	elif cols == 2:
		data['aq_type'] = 'Dictionary'
	else:
		data['aq_type'] = 'Table'

	print data

	# Send data to APIquitous
	response = requests.post("http://localhost:8080/{}.publish".format(data['aq_type']),
						 headers={'content_type':'application/json'},
						 data=json.dumps(data))

	print response.content

	# Send a success message to the user.
	alert_data = [data['aq_organization'],
				  data['aq_name'],
				  data['aq_type']]
	xlcAlert('Your data has been published at http://apiquitous.appspot.com/display.\n\nIf you would like to share your data with others, you can use ID {}.{}.{}'.format(*alert_data))

@xl_menu("Fetch all", menu="APIquitous", menu_order=2)
def fetch_all():

	app = Dispatch("Excel.Application")
	wb = app.ActiveWorkbook
	wings_wb = Workbook(xl_workbook=wb)

	try:
		ws = wb.Sheets.Item('Configuration')

	except com_error as e:
		if e[0] == -2147352567:
			xlcAlert('There is no "Configuration" sheet in this document. Without this sheet you can\'t fetch any data from APIquitous.')
	
	# This formatting should be abstracted to a function
	conf_keys = [ 'aq_' + x.lower().replace(' ', '_') for x in Range('Configuration', 'A1').horizontal.value]
	conf_table = []
	row_to_check = 2
	while Range('Configuration', 'A' + str(row_to_check)).value != None:
		conf_values = Range('Configuration', 'A' + str(row_to_check)).horizontal.value
		if conf_values[-1] == 'No':
			row_to_check += 1
			continue
		row_to_check += 1
		conf_dict = {}
		for i, j in zip(conf_keys, conf_values):
			conf_dict[i] = j
		conf_table.append(conf_dict)  

	responses = []
	for i in conf_table:
		print 'Artur ', i
		response = requests.post('http://localhost:8080/' + i['aq_id'].split('.', 2)[2] + '.fetch',
							 headers={'content_type':'application/json'},
							 data=json.dumps(i))
		responses.append(response.json())

	for i in responses:
		print i['aq_data']
		Range(i['aq_sheet_name'], i['aq_destination_cell']).value = [[x] for x in i['aq_data']]
		Range('Data', 'N3').value = [[1],[2],[3]]



	'''
	Range('Data', 'X99').value = ['a','b','c']
	wb.Sheets('Data').Range('X100').Value = ['a','b','c']
	aaa = Range('Data', 'K86:K88').value
	print aaa
	print type(aaa)
	Range('Data', 'L86:L88').value = tuple(aaa)
	'''
	#wb = Workbook('C:\\Users\\chbapie\\Desktop\\Bartosz\\apiquitous\\spreadsheets\\PublishAList.xlsx')

	#conf_keys2 = Range('Configuration', 'C8').value
	

	#try:
	#	wb.Sheets.Item('Conf')
	#except IOError:
   	#	print "Error: can\'t find file or read data"
	#print wb.Sheets.items()
	#browse(wb.Sheets)
	#ws = wb.Sheet(''
	#xlcAlert("There is no APIquitous feeds here.")





	#Range('Sheet1', 'H3').value = selection
	#Range('Sheet1', 'H6').value = type(selection)
	#Range('Sheet1', 'H9').value = dir(selection)
	#Range('Sheet1', 'K16').value = pickle.loads(response.json()['from_db'])
	#xlcAlert("Hello again from PyXLL")
    #xlcAlert("List has been published")

@xl_menu("Load publishing template", menu="APIquitous", menu_order=3)
def load_a_publishing_template():
	"""
	Generate a template to create a new data structure to be published at apiquitous.
	"""

	# Initialize the application and add a new sheet
	app = Dispatch("Excel.Application")
	current_sheet_name = app.ActiveWorkbook.ActiveSheet.Name
	app.ActiveWorkbook.Worksheets.Add(After=app.ActiveWorkbook.Worksheets(1))
	new_sheet = app.ActiveWorkbook.ActiveSheet
	new_sheet.Name = current_sheet_name + ".AQ"
	publishing_keys = ['Name', 'Description', 'Organization', 'Created by']
	publishing_keys_cells = ['A' + str(x) for x in range(1, len(publishing_keys) + 1)]

	# Formatting of the new sheet
	keys_count = str(len(publishing_keys))
	new_sheet.Cells.Interior.Color = 16777215.0
	new_sheet.Range('A1:A' + keys_count).Interior.Color = 0.0
	new_sheet.Range('B1:B' + keys_count).Interior.Color = 55295.0
	new_sheet.Range('A1:A' + keys_count).Font.Color = 55295.0
	new_sheet.Range('B1:B' + keys_count).Font.Color = 0.0
	for i, j in zip(publishing_keys_cells, publishing_keys):
		new_sheet.Range(i).Value = j
	new_sheet.Range('B1').Value = 'Type values in the yellow cells'
	new_sheet.Range('A1:B' + keys_count).Columns.AutoFit()

