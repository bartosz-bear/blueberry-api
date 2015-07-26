from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service

from google.appengine.api.users import User
from google.appengine.ext import ndb

import models
from models import BAPIList, BAPIScalar, BAPIDictionary, BAPITable, FetchConfigurations, PublishConfigurations
import apis_funcs

import logging
import pickle
import sys
import pdb

fetch_args_keys = ['bapi_id',
                      'description',
                      'destination_cell',
                      'organization',
                      'data',
                      'user',
                      'workbook',
                      'workbook_path',
                      'worksheet']

fetch_args_values = ['request.bapi_id',
               'queried_data[0].description',
               'request.destination_cell',
               'queried_data[0].organization',
               'queried_data[1]',
               'request.user',
               'request.workbook',
               'request.workbook_path',
               'request.worksheet']


class PublishRequest(messages.Message):
    """
    Request Object which collects information from JSON passed via HTTP request.
    """
    name = messages.StringField(1)
    user = messages.StringField(2)
    description = messages.StringField(3)
    organization = messages.StringField(4)
    data = messages.StringField(5, repeated=True)
    workbook_path = messages.StringField(6)
    workbook = messages.StringField(7)
    worksheet = messages.StringField(8)
    destination_cell = messages.StringField(9)
    bapi_id = messages.StringField(10)
    data_type = messages.StringField(11)


class PublishResponse(messages.Message):
    """
    Response Object which informs the requester about the successful processing.
    """
    response = messages.StringField(1, required=True)


class FetchRequest(messages.Message):
    """
    Request object which delivers data needed to fetch data from Datastore.
    """
    bapi_id = messages.StringField(1, required=True)
    user = messages.StringField(2)
    organization = messages.StringField(3)
    description = messages.StringField(4)
    workbook_path = messages.StringField(5)
    workbook = messages.StringField(6)
    worksheet = messages.StringField(7)
    destination_cell = messages.StringField(8)
    skip_new_conf = messages.BooleanField(9)


class FetchResponse(messages.Message):
    """
    Response object which returns data from Datastore.
    """
    bapi_id = messages.StringField(1, required=True)
    user = messages.StringField(2)
    organization = messages.StringField(3)
    description = messages.StringField(4)
    workbook_path = messages.StringField(5)
    workbook = messages.StringField(6)
    worksheet = messages.StringField(7)
    destination_cell = messages.StringField(8)
    data = messages.StringField(9, repeated=True)
    info = messages.StringField(10)


class GetPublishedRequest(messages.Message):
    """
    Request object which delivers data needed to fetch data from Datastore.
    """
    workbook_path = messages.StringField(1, required=True)
    workbook = messages.StringField(2, required=True)


class GetPublishedResponse(messages.Message):
    """
    Response object which returns data from Datastore.
    """
    ids = messages.StringField(1, repeated=True)
    users = messages.StringField(2, repeated=True)
    names = messages.StringField(3, repeated=True)
    descriptions = messages.StringField(4, repeated=True)
    organizations = messages.StringField(5, repeated=True)
    workbook_paths = messages.StringField(6, repeated=True)
    workbooks = messages.StringField(7, repeated=True)
    worksheets = messages.StringField(8, repeated=True)
    destination_cells = messages.StringField(9, repeated=True)
    data_types = messages.StringField(10, repeated=True)


class GetFetchedRequest(messages.Message):
    """
    Request object which delivers data needed to fetch data from Datastore.
    """
    workbook_path = messages.StringField(1, required=True)
    workbook = messages.StringField(2, required=True)


class GetFetchedResponse(messages.Message):
    """
    Response object which returns data from Datastore.
    """
    names = messages.StringField(1, repeated=True)
    users = messages.StringField(2, repeated=True)
    organizations = messages.StringField(3, repeated=True)
    descriptions = messages.StringField(4, repeated=True)
    workbook_paths = messages.StringField(5, repeated=True)
    workbooks = messages.StringField(6, repeated=True)
    worksheets = messages.StringField(7, repeated=True)
    destination_cells = messages.StringField(8, repeated=True)
    bapi_ids = messages.StringField(9, repeated=True)


class IsIDUsedRequest(messages.Message):
    """
    Request object which delivers data needed to get all IDs for a particular user.
    """
    bapi_id = messages.StringField(1)
    user = messages.StringField(2)


class IsIDUsedResponse(messages.Message):
    """
    Response object which returns a Boolean value indicating whether a particular ID was used by any user other
    than the one which has sent the request.
    """
    response = messages.BooleanField(1)


class Data(remote.Service):
    """
    A RPC Service which handles all requests related to all data structures.
    """

    published_args_keys = ['ids',
                   'users',
                   'names',
                   'workbook_paths',
                   'workbooks',
                   'worksheets',
                   'descriptions',
                   'organizations',
                   'destination_cells',
                   'data_types']

    published_args_values = ['i.bapi_id',
                             'i.user',
                             'i.name',
                             'i.workbook_path',
                             'i.workbook',
                             'i.worksheet',
                             'i.description',
                             'i.organization',
                             'i.destination_cell',
                             'i.data_type']

    fetched_args_keys = ['bapi_ids',
                     'names',
                     'users',
                     'workbook_paths',
                     'workbooks',
                     'worksheets',
                     'destination_cells',
                     'organizations',
                     'descriptions']

    fetched_args_values = ['i.bapi_id',
                           'i.name',
                           'i.user',
                           'i.workbook_path',
                           'i.workbook',
                           'i.worksheet',
                           'i.destination_cell',
                           'i.organization',
                           'i.description']


    @remote.method(GetPublishedRequest, GetPublishedResponse)
    def get_published(self, request):
        """
        Get a list of Published Configurations. This list will be used to update
        all previously defined BAPI data.
        """
        class_ = getattr(models, 'PublishConfigurations')
        published_items_dict = apis_funcs.query_configurations(request, self.published_args_keys,
                                                               self.published_args_values, class_)

        return GetPublishedResponse(**{x: published_items_dict[y]
                                       for x, y in zip(self.published_args_keys, list(published_items_dict))})

    @remote.method(GetFetchedRequest, GetFetchedResponse)
    def get_fetched(self, request):
        """
        Get a list of Fetched Configurations. This list will be used to fetch
        all previously fetched BAPI data.
        """
        class_ = getattr(models, 'FetchConfigurations')
        fetched_items_dict = apis_funcs.query_configurations(request, self.fetched_args_keys,
                                                             self.fetched_args_values, class_)

        return GetFetchedResponse(**{x: fetched_items_dict[y]
                                       for x, y in zip(self.fetched_args_keys, list(fetched_items_dict))})

    @remote.method(IsIDUsedRequest, IsIDUsedResponse)
    def is_id_used(self, request):
        """
        Check if a requested ID was already used by a different user than the requesting one.
        """
        is_id_used = PublishConfigurations.query(ndb.AND(PublishConfigurations.bapi_id == request.bapi_id,
                                                 PublishConfigurations.user != request.user)).count()

        if is_id_used > 0:
            response = True
        else:
            response = False

        return IsIDUsedResponse(response=response)


class Scalar(remote.Service):
    """
    A RPC Service which handles all requests related to a data structure Scalar.
    """

    @remote.method(PublishRequest, PublishResponse)
    def publish(self, request):
        """
        Publish method receives a request with a Scalar and saves it to a Datastore.
        """

        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)

        return PublishResponse(response=apis_funcs.publish_and_collect(request, class_))

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        """
        Fetch method receives a request from a client and returns a BAPI Scalar.
        """

        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)
        if not apis_funcs.is_ID_valid(request, class_):
            return FetchResponse(bapi_id="BAPI Info Message", data=[u'1', u'2'], info="Incorrect BAPI ID")

        queried_data = apis_funcs.query_and_configure(request, class_, data_type)

        evaluated_items = [eval(x) for x in fetch_args_values]
        logging.info(evaluated_items)
        return FetchResponse(**{x: y for x, y in zip(fetch_args_keys, evaluated_items)})


class List(remote.Service):
    """
    A RPC Service which handles all requests related to a data structure List.
    """

    @remote.method(PublishRequest, PublishResponse)
    def publish(self, request):
        """
        Publish method receives a request with a List and saves it to a Datastore.
        """
        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)

        return PublishResponse(response=apis_funcs.publish_and_collect(request, class_))

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        """
        Fetch method receives a request from a client and returns a BAPI List.
        """

        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)
        if not apis_funcs.is_ID_valid(request, class_):
            return FetchResponse(bapi_id="BAPI Info Message", data=[u'1', u'2'], info="Incorrect BAPI ID")

        queried_data = apis_funcs.query_and_configure(request, class_, data_type)

        evaluated_items = [eval(x) for x in fetch_args_values]
        return FetchResponse(**{x: y for x, y in zip(fetch_args_keys, evaluated_items)})


class Dictionary(remote.Service):
    """
    A RPC Service which handles all requests related to a data structure Dictionary.
    """

    @remote.method(PublishRequest, PublishResponse)
    def publish(self, request):
        """
        Publish method receives a request with a List and saves it to a Datastore.
        """
        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)

        return PublishResponse(response=apis_funcs.publish_and_collect(request, class_))

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        """
        Fetch method receives a request from a client and returns a BAPI List.
        """

        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)
        if not apis_funcs.is_ID_valid(request, class_):
            return FetchResponse(bapi_id="BAPI Info Message", data=[u'1', u'2'], info="Incorrect BAPI ID")

        queried_data = apis_funcs.query_and_configure(request, class_, data_type)

        evaluated_items = [eval(x) for x in fetch_args_values]
        logging.info(evaluated_items)
        return FetchResponse(**{x: y for x, y in zip(fetch_args_keys, evaluated_items)})


class Table(remote.Service):
    """
    A RPC Service which handles all requests related to a data structure Table.
    """

    @remote.method(PublishRequest, PublishResponse)
    def publish(self, request):
        """
        Publish method receives a request with a Table and saves it to a Datastore.
        """
        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)

        return PublishResponse(response=apis_funcs.publish_and_collect(request, class_))

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        """
        Fetch method receives a request from a client and returns a BAPI Table.
        """
        data_type = request.bapi_id.split('.')[2]
        class_ = getattr(models, 'BAPI' + data_type)
        if not apis_funcs.is_ID_valid(request, class_):
            return FetchResponse(bapi_id="BAPI Info Message", data=[u'1', u'2'], info="Incorrect BAPI ID")

        queried_data = apis_funcs.query_and_configure(request, class_, data_type)

        evaluated_items = [eval(x) for x in fetch_args_values]
        logging.info(evaluated_items)
        return FetchResponse(**{x: y for x, y in zip(fetch_args_keys, evaluated_items)})

app = service.service_mappings([('/Data.*', Data),
                                ('/Scalar.*', Scalar),
                                ('/List.*', List),
                                ('/Dictionary.*', Dictionary),
                                ('/Table.*', Table)
                                ])
