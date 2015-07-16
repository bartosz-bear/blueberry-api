__author__ = 'CHBAPIE'

from google.appengine.api.users import User

from models import FetchConfigurations, PublishConfigurations

import pickle
import logging
import pdb

from collections import OrderedDict

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
                    'worksheet']
    items = [request.bapi_id,
             queried_data[0].description,
             request.destination_cell,
             queried_data[0].organization,
             queried_data[1],
             request.user,
             request.workbook,
             request.workbook_path,
             request.worksheet]
    for i, j in zip(request_keys, items):
        FetchResponseDictionary[i] = j

    return FetchResponseDictionary


def publish_and_collect(request, class_):
    """
    Check if the ID was already published. If not, then create a Publish Configuration.
    """
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
    else:
        # It's not the first 'Publish', therefore fetch info about this ID from the database.
        logging.info(class_)
        item = class_.query(class_.bapi_id == request.bapi_id).get()
        name = item.name
        user = item.user
        description = item.description
        organization = item.organization

    class_(name=name,
           user=user,
           description=description,
           organization=organization,
           bapi_id=request.bapi_id,
           data=pickle.dumps(request.data)).put()


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



    queried_list_pickled = pickle.loads(queried_list.data)

    logging.info('Checking pickled data: ')
    logging.info(queried_list_pickled)

    return [queried_list, queried_list_pickled]


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
    jan = class_.query(class_.bapi_id == request.bapi_id).count()
    logging.info(jan)
    if class_.query(class_.bapi_id == request.bapi_id).count() > 0:
        return True
    else:
        return False

