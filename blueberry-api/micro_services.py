__author__ = 'Bartosz Piechnik'

import pickle
import sys
import json
import numpy as np
import models
import apis_funcs

from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service
from google.appengine.api.users import User
from models import BAPIList, BAPIScalar, BAPIDictionary, FetchConfigurations, PublishConfigurations, Pipeline

import logging
import pdb


class SelectColumnsRequest(messages.Message):
    """
    Request object with information necessary to execute 'SelectColumns' service.
    """
    data = messages.StringField(1, repeated=True)
    headers_list_all = messages.StringField(2, repeated=True)
    headers_list_selected = messages.StringField(3, repeated=True)


class SelectColumnsResponse(messages.Message):
    """
    Response object received upon executing 'SelectColumns' service.
    """
    data = messages.StringField(1, repeated=True)
    headers_list = messages.StringField(2, repeated=True)


class SelectColumns(remote.Service):
    """
    Receive a list of lists, a list of existing headers and a sublist of existing headers. Returns data with headers
    from a sublist and a sublist itself.
    """
    @remote.method(SelectColumnsRequest, SelectColumnsResponse)
    def select_columns(self, request):
        #user = json.loads(request.user)
        #column = request.column
        #filter_values = request.filter_values
        x = np.array([1, 2, 3])
        return SelectColumnsResponse(data=json.dumps(x), headers_list=["jan", "maria"])


class AddFilterRequest(messages.Message):
    """
    Request object to send information necessary to set up a new pipeline.filter configuration.
    """
    user = messages.StringField(1)
    column = messages.StringField(2)
    filter_values = messages.StringField(3, repeated=True)
    pipeline = messages.StringField(4, required=True)

class AddFilterResponse(messages.Message):
    """
    Response object to AddFilterRequest.
    """
    response = messages.StringField(1)


class Filter(remote.Service):
    """
    RPC Service which handles all requeste related to pipeline.Filter.
    """
    @remote.method(AddFilterRequest, AddFilterResponse)
    def add_filter(self, request):
        """
        Add a filter to Pipeline and Filter datastore.
        """


        pdb.set_trace()




class SortingRequestSchema(messages.Message):
    """
    Request Object which collects information from JSON passed via HTTP request responsible for sorting.
    """
    bapi_id = messages.StringField(1)
    text = messages.StringField(2)


class SortingResponseSchema(messages.Message):
    """
    Request Object which collects information from JSON passed via HTTP request responsible for sorting.
    """
    response = messages.StringField(1)


class Sorting(remote.Service):
    """
    A RPC Service which handles all requests related to all data structures.
    """

    @remote.method(SortingRequestSchema, SortingResponseSchema)
    def get_published(self, request):
        """
        Get a list of Published Configurations. This list will be used to update
        all previously defined BAPI data.
        """
        bapi_id = request.bapi_id
        bapi_id_after_processing = bapi_id + "end_tag" + "finish"
        bapi_id_after_processing = str(bapi_id_after_processing)

        return bapi_id_after_processing

app = service.service_mappings([('/Sorting.*', Sorting),
                                ('/MicroService/SelectColumns.*', SelectColumns)])
