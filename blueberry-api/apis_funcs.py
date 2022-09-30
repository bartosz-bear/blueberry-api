__author__ = 'Bartosz Piechnik'

import json
import pickle

from collections import OrderedDict
from google.appengine.api.users import User
from models import FetchConfigurations, PublishConfigurations

import logging
import pdb

CELL_ERRORS = [-2146826281,
              -2146826246,
              -2146826259,
              -2146826288,
              -2146826252,
              -2146826265,
              -2146826273]

def create_a_dict_for_fetch(request, queried_data):
    """
    Takes a HTTP request and a data from database and prepares a dictionary which will be packed into the HTTP response.
    """
    FetchResponseDictionary = {}
    request_keys = ['bapi_id',
                    'description',
                    'destination_cell',
                    'organization',
                    'data',
                    'user',
                    'workbook',
                    'workbook_path',
                    'worksheet',
                    'headers_list']
    items = [request.bapi_id,
             queried_data[0].description,
             request.destination_cell,
             queried_data[0].organization,
             queried_data[1],
             request.user,
             request.workbook,
             request.workbook_path,
             request.worksheet,
             queried_data[2]]
    for i, j in zip(request_keys, items):
        FetchResponseDictionary[i] = j

    return FetchResponseDictionary


def publish_and_collect(request, class_):
    """
    Check if the ID was already published. If not, then create a Publish Configuration.
    """

    request_data = json.loads(request.data.pop())
    request.data.append(json.dumps(errors_to_nulls(request_data)))

    is_published = PublishConfigurations.query(PublishConfigurations.bapi_id == request.bapi_id).count()
    if is_published == 0:
        # In case this is the first 'Publish' for this id, create a 'PublishConfiguration'.
        name = request.name
        user = request.user
        description = request.description
        organization = request.organization
        data_type = request.data_type
        bapi_id = request.bapi_id
        PublishConfigurations(bapi_id=bapi_id,
                              user=user,
                              name=name,
                              description=description,
                              organization=organization,
                              workbook_path=request.workbook_path,
                              workbook=request.workbook,
                              worksheet=request.worksheet,
                              destination_cell=request.destination_cell,
                              data_type=data_type
                              ).put()
        return_string = 'Data has been uploaded.'
    else:
        # It's not the first 'Publish', therefore fetch info about this ID from the database.
        logging.info(class_)
        try:
            item = class_.query(class_.bapi_id == request.bapi_id).get()
            name = item.name
            user = item.user
            description = item.description
            organization = item.organization
            return_string = 'Data has been uploaded.'
        except AttributeError:
            name = request.name
            user = request.user
            description = request.description
            organization = request.organization
            bapi_id = request.bapi_id
            return_string = 'Data did not exist, but has been uploaded.'

    class_(name=name,
           user=user,
           description=description,
           organization=organization,
           bapi_id=request.bapi_id,
           headers=json.dumps(request.headers_list),
           data=json.dumps(request.data)).put()

    return return_string


def query_and_configure(request, class_, data_type):
    """
    Query one of the BAPI data structures. If fetch is of 'repetetive' type which means
    a user will most likely fetch it again in the future, then create create a Fetch Configuration
    which saves what type(ID) and where(among others name and path of the excel spreadsheet) data should be fetched.
    """

    # Query one BAPI data structures.
    queried_list = class_.query(
        class_.bapi_id == request.bapi_id).order(-class_.last_updated).get()


    # Check if the 'Fetch' request is a one-off or a repetitive request.
    # In case it's a repetitive, save it to the Datastore.
    if not request.skip_new_conf:
        FetchConfigurations(name=queried_list.name,
                            user=request.user,
                            data_type=data_type,
                            bapi_id=request.bapi_id,
                            workbook_path=request.workbook_path,
                            workbook=request.workbook,
                            worksheet=request.worksheet,
                            destination_cell=request.destination_cell,
                            organization=queried_list.organization,
                            description=queried_list.description
                            ).put()

    queried_list_pickled = json.loads(queried_list.data)
    queried_list_headers = json.loads(queried_list.headers)

    return [queried_list, queried_list_pickled, queried_list_headers]


def query_configurations(request, published_args_keys, published_args_values, class_):
    """
    Query PublishConfigurations datastore class to get all information about data which have
    been previously published from a particular workbook.
    """

    workbook_path = request.workbook_path
    workbook = request.workbook

    published_configurations = class_.query(
        class_.workbook_path == workbook_path,
        class_.workbook == workbook
    )

    configurations_dict = OrderedDict((x, list()) for x in published_args_keys)

    for i in published_configurations.iter():
        for j, k in zip(configurations_dict, published_args_values):
            configurations_dict[j].append(eval(k))

    return configurations_dict

def is_ID_valid(request, class_):
    """
    Validates whether the BAPI ID exists.
    :return:
    """
    if class_.query(class_.bapi_id == request.bapi_id).count() > 0:
        return True
    else:
        return False

def errors_to_nulls(request_data):
    """
    Iterates over a list of list which corresponds to data sent from excel, and replaces all error cells with null.
    :param request_data:
    :return:
    """
    cell_coordinates = []

    for r in range(len(request_data)):
        for c in range(len(request_data[0])):
            current_item = request_data[r][c]
            if type(current_item) == int:
                if current_item in CELL_ERRORS:
                    cell_coordinates.append((r, c))

    new_request_data = request_data
    for co in cell_coordinates:
        new_request_data[co[0]][co[1]] = None

    return new_request_data