# -*- coding: UTF-8 -*-
import pandas
import json
import requests
import pickle
from xlwings import Workbook, Sheet, Range

from win32com.client import Dispatch
from win32com.client import constants as c

from pywintypes import com_error

from pprint import pprint
from objbrowser import browse

import logging


def publish_to_cloud(data):
    """
	Publish a data structure to APIquitous' 
	"""

    # Initialize the application
    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    ws = wb.ActiveSheet
    sheet_name = wb.ActiveSheet.Name
    wings_wb = Workbook(xl_workbook=wb)

    # Fetch data from workbook
    data['aq_data'] = Range(wb.ActiveSheet.Name, app.Selection.Address).value
    data['aq_workbook_path'] = wb.Path
    data['aq_workbook'] = wb.Name
    data['aq_worksheet'] = wb.ActiveSheet.Name
    data['aq_destination_cell'] = app.Selection.Address

    # Find out what type of data is is being published
    cols = app.Selection.Columns.Count
    print cols
    if cols == 1:
        data['aq_type'] = 'List'
    elif cols == 2:
        data['aq_type'] = 'Dictionary'
    else:
        data['aq_type'] = 'Table'


    print 'Is this what i think'
    print data

    # Send data to APIquitous
            response = requests.post("http://localhost:8080/{}.publish".format(data['aq_type']),
                             headers={'content_type': 'application/json'},
                             data=json.dumps(data))

    return data['aq_organization'] + data['aq_name'] + data['aq_type']

    logging.info(response.content)


    # Send a success message to the user.
    # alert_data = [data['aq_organization'],
    #			  data['aq_name'],
    #			  data['aq_type']]
    #xlcAlert('Your data has been published at http://apiquitous.appspot.com/display.\n\nIf you would like to share your data with others, you can use ID {}.{}.{}'.format(*alert_data))


def get_publishing_list():
    """
    Get a list of all items which have been published.
    """
    pass

def fetch_new(conf):
    """
    Fetch Blueberry ID from Excel, request data from GAE, save it to excel and save a new fetch configuration to GAE.
    """

    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    wings_wb = Workbook(xl_workbook=wb)

    # 'GTO.List_of_tanks.List'
    # 'GTO.SNL_Banks.List'

    #conf['workbook_path'] = 'C:\\Users\\chbapie\\Desktop\\Bartosz\\apiquitous\\APIquitousAFO\\APIquitousAFO\\bin\\Debug'
    conf['workbook_path'] = wb.Path
    conf['workbook'] = wb.Name
    print wb.Path
    print wb.Name
    print wb.ActiveSheet.Name
    conf['worksheet'] = wb.ActiveSheet.Name

    print app.Selection.Address
    conf['destination_cell'] = app.Selection.Address
    #conf['destination_cell'] = 'A1'

    print wb
    print dir(wb)



    response = requests.post('http://localhost:8080/' + conf['aq_id'].split('.', 2)[2] + '.fetch',
                         headers={'content_type': 'application/json'},
                         data=json.dumps(conf))

    response = response.json()

    print response
    logging.info(response)


    #wings_wb = Workbook(response['workbook_path'] + response['workbook'])


    Range(response['worksheet'], response['destination_cell']).value = [[x] for x in response['aq_data']]
    ws = wb.ActiveSheet
    #Dim LR As Long, LC As Long
    #LR = ActiveCell.End(xlDown).Row
    #LC = ActiveCell.End(xlToRight).Column
    #Range(ActiveCell, Cells(LR, LC)).Select
    endCell = ws.Range(response['destination_cell']).End(-4121).Address
    paste_range = ws.Range(response['destination_cell'] + ":" + endCell)
    paste_range.Interior.Color = 5907204.0
    paste_range.Font.Color = 16510126.0
    paste_range.Columns.AutoFit()


    #Range(response['worksheet'], response['destination_cell']).color = (174, 236, 251)

    #new_sheet = app.ActiveWorkbook.ActiveSheet
    #new_sheet.Range('B1:B' + keys_count).Interior.Color = 55295.0

    return response

def get_fetched():
    """
    Get a list of all items which have been fetched in the current sheet.
    """

    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    wings_wb = Workbook(xl_workbook=wb)

    conf = {'workbook_path': wb.Path,
            'workbook': wb.Name}

    response = requests.post('http://localhost:8080/List.get_fetched',
                         headers={'content_type': 'application/json'},
                         data=json.dumps(conf))

    return response.json()


def get_published():
    """
    Get a list of all items which have been published in the current sheet.
    """

    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    wings_wb = Workbook(xl_workbook=wb)

    conf = {'workbook_path': wb.Path,
            'workbook': wb.Name}

    response = requests.post('http://localhost:8080/List.get_published',
                         headers={'content_type': 'application/json'},
                         data=json.dumps(conf))

    return response.json()

def fetch_many(confs):

    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    wings_wb = Workbook(xl_workbook=wb)

    data_from_cloud = []
    for i in confs['ids']:
        print i
        data = {'skip_new_conf': True,
                'aq_id': i}
        response = requests.post('http://localhost:8080/' + i.split('.', 2)[2] + '.fetch',
                         headers={'content_type': 'application/json'},
                         data=json.dumps(data))
        response = response.json()['aq_data']
        data_from_cloud.append(response)

    confs['aq_data'] = data_from_cloud
    print len(confs['aq_data'])

    for i in range(len(confs['aq_data'])):
        Range(confs['worksheets'][i], confs['destination_cells'][i]).value = [[x] for x in confs['aq_data'][i]]
        ws = wb.ActiveSheet
        endCell = ws.Range(confs['destination_cells'][i]).End(-4121).Address
        paste_range = ws.Range(confs['destination_cells'][i] + ":" + endCell)
        paste_range.Interior.Color = 5907204.0
        paste_range.Font.Color = 16510126.0
        paste_range.Columns.AutoFit()

def publish_many(confs):

    # Initialize the application
    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    ws = wb.ActiveSheet
    sheet_name = wb.ActiveSheet.Name
    wings_wb = Workbook(xl_workbook=wb)

    sss = []

    for i in range(len(confs['ids'])):
        # Fetch data from workbook
        data = {}

        data['aq_id'] = confs['ids'][i]
        data['aq_type'] = confs['data_types'][i]
        #data['user'] = confs['users'][i]
        #data['description'] = confs['descriptions'][i]
        #data['organization'] = confs['organization'][i]
        #data['aq_workbook_path'] = confs['aq_workbook_path'][i]
        #data['aq_workbook'] = confs['aq_workbook'][i]
        #data['aq_worksheet'] = confs['aq_worksheet'][i]
        #data['aq_destination_cell'] = confs['aq_destination_cell'][i]

        destination_cell = confs['destination_cells'][i].split(':')[0]

        data['aq_data'] = Range(sheet_name, destination_cell).vertical.value

        print 'Is this what i think'
        print data

        # Send data to APIquitous
        response = requests.post("http://localhost:8080/{}.publish".format(data['aq_type']),
                                 headers={'content_type': 'application/json'},
                                 data=json.dumps(data))
        sss.append(i)

    print sss

    #return conf['aq_organization'] +conf['aq_name'] + conf['aq_type']

    ##### delete the following ###

"""

def fetch_many(confs):

    app = Dispatch("Excel.Application")
    wb = app.ActiveWorkbook
    wings_wb = Workbook(xl_workbook=wb)

    try:
        ws = wb.Sheets.Item('Configuration')

    except com_error as e:
        if e[0] == -2147352567:
            pass
            #xlcAlert(
            #    'There is no "Configuration" sheet in this document. Without this sheet you can\'t fetch any data from APIquitous.')

    # This formatting should be abstracted to a function
    conf_keys = ['aq_' + x.lower().replace(' ', '_') for x in Range('Configuration', 'A1').horizontal.value]
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
                                 headers={'content_type': 'application/json'},
                                 data=json.dumps(i))
        responses.append(response.json())

    for i in responses:
        print i['aq_data']
        Range(i['aq_sheet_name'], i['aq_destination_cell']).value = [[x] for x in i['aq_data']]
        Range('Data', 'N3').value = [[1], [2], [3]]

"""


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

