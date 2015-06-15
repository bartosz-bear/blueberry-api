
__author__ = 'CHBAPIE'

from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service

from google.appengine.api.users import User

import models
from models import BAPIList, BAPIScalar, BAPIDictionary, FetchConfigurations, PublishConfigurations
import apis_funcs

import logging
import pickle
import sys
import pdb


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
        bapi_id_after_processing = bapi_id + "end_tag" + "kurwa"
        bapi_id_after_processing = str(bapi_id_after_processing)

        return bapi_id_after_processing

app = service.service_mappings([('/Sorting.*', Sorting)])
