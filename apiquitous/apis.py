from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service

from google.appengine.api.users import User

from models import AqList

import logging
import pickle
import codecs

# Create the request string containing user's name
class ListRequest(messages.Message):
    """
    Request Object which collects information from JSON passed via HTTP request.
    """
    aq_name = messages.StringField(1, required=True)
    aq_created_by = messages.StringField(2, required=True)
    aq_description = messages.StringField(3, required=True)
    aq_organization = messages.StringField(4, required=True)
    aq_data = messages.StringField(5, repeated=True)


class ListResponse(messages.Message):
    """
    Response Object which informs the requester about the successful processing.
    """
    response = messages.StringField(1, required=True)


class FetchRequest(messages.Message):
    """
    Request object which delivers data needed to fetch data from Datastore.
    """
    aq_id = messages.StringField(1, required=True)
    aq_sheet_name = messages.StringField(2)
    aq_destination_cell = messages.StringField(3)


class FetchResponse(messages.Message):
    """
    Response object which returns data from Datastore.
    """
    aq_data = messages.StringField(1, repeated=True)
    aq_sheet_name = messages.StringField(2)
    aq_destination_cell = messages.StringField(3)


class List(remote.Service):
    """
    A RPC Service which handles all requests related to a data structure List.
    """
    @remote.method(ListRequest, ListResponse)
    def publish(self, request):
        """
        Publish method receives a request with a List and saves it to a Datastore.
        """

        # Validate that data in request.aq_data is one of acceptable data formats by

        AqList(name=request.aq_name,
               user=User(request.aq_created_by),
               description=request.aq_description,
               organization=request.aq_organization,
               data=pickle.dumps(request.aq_data)).put()

        return ListResponse(response='The list has been uploaded')

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        aq_organization, aq_name, aq_type = request.aq_id.split('.')
        del(aq_type)
        from_db = AqList.query(
            AqList.organization == aq_organization,
            AqList.name == aq_name).order(AqList.last_updated)
        from_db = from_db.get().data
        from_db = pickle.loads(from_db)
        return FetchResponse(aq_data=from_db,
                             aq_sheet_name=request.aq_sheet_name,
                             aq_destination_cell=request.aq_destination_cell)

app = service.service_mappings([('/List.*', List)])
