from pyxll import xl_func, xl_menu, get_config, xlcAlert, get_active_object
import pandas	
import json
import requests
import pickle
from xlwings import Workbook, Sheet, Range
import win32com.client

@xl_func("string name: string")
def hello(name):
	""" return a familiar greeting"""
	return 'Hello, %s' % name

@xl_menu("Another example menu item", order=1)
def on_example_menu_item_2():
    xlcAlert("Hello again from PyXLL")

@xl_menu("Publish a list to APIquitous", menu="Ubiquitous", menu_order=1)
def publish_a_list_to_cloud():
	xl_window = get_active_object()
	xl_app = win32com.client.Dispatch(xl_window).Application
	
	selection = xl_app.Selection.Value
	aqlist = pickle.dumps(selection)
	wb = Workbook('C:\\Users\\chbapie\\Desktop\\Bartosz\\apiquitous\\spreadsheets\\PublishAList.xlsx')
	#Range('Sheet1', 'D1').value = selection
    #response = requests.post("http://localhost:8080/HelloService.aqlist",
	#					 headers={'content_type':'application/json'},
	#					 data=json.dumps({'aqlist':'aqlist'}))
	response = requests.post("http://localhost:8080/HelloService.aqlist",
						 headers={'content_type':'application/json'},
						 data=json.dumps({'aqlist':aqlist}))
	#xlcAlert("Hello again from PyXLL")
    #xlcAlert("List has been published")

@xl_menu("Get a list from APIquitous", menu="Ubiquitous", menu_order=2)
def get_a_list_from_cloud():
	#xl_window = get_active_object()
	#xl_app = win32com.client.Dispatch(xl_window).Application
	#selection = xl_app.Selection
	response = requests.post("http://localhost:8080/HelloService.from_db",
						 headers={'content_type':'application/json'},
						 data=json.dumps({'from_db':'from_db'}))
	wb = Workbook('C:\\Users\\chbapie\\Desktop\\Bartosz\\apiquitous\\spreadsheets\\PublishAList.xlsx')
	#Range('Sheet1', 'H3').value = selection
	#Range('Sheet1', 'H6').value = type(selection)
	#Range('Sheet1', 'H9').value = dir(selection)
	Range('Sheet1', 'K16').value = pickle.loads(response.json()['from_db'])
	#xlcAlert("Hello again from PyXLL")
    #xlcAlert("List has been published")