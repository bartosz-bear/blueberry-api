from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service

from google.appengine.api.users import User

from models import BAPIList, FetchConfigurations, PublishConfigurations

import logging
import pickle

# Create the request string containing user's name
class ListRequest(messages.Message):
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


class ListResponse(messages.Message):
    """
    Response Object which informs the requester about the successful processing.
    """
    response = messages.StringField(1, required=True)


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


class List(remote.Service):
    """
    A RPC Service which handles all requests related to a data structure List.
    """
    @remote.method(ListRequest, ListResponse)
    def publish(self, request):
        """
        Publish method receives a request with a List and saves it to a Datastore.
        """
        is_configured = PublishConfigurations.query(PublishConfigurations.bapi_id == request.bapi_id).count()
        if is_configured == 0:
            # In case this is the first 'Publish' for this id, create a 'PublishConfiguration'.
            name = request.name
            user = User(request.user)
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
            item = BAPIList.query(BAPIList.bapi_id == request.bapi_id).get()
            name = item.name
            user = item.user
            description = item.description
            organization = item.organization

        BAPIList(name=name,
               user=user,
               description=description,
               organization=organization,
               bapi_id=request.bapi_id,
               data=pickle.dumps(request.data)).put()

        return ListResponse(response='The list has been uploaded.')

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        """
        Fetch method receives a request from a client and returns a BAPI List.
        """
        organization, name, data_type = request.bapi_id.split('.')
        del data_type

        # Get the BAPI list.
        queried_list = BAPIList.query(
            BAPIList.organization == organization,
            BAPIList.name == name.replace('_', ' ')).order(-BAPIList.last_updated).get()

        # Check if the 'Fetch' request is a one-off or a repetitive request.
        # In case it's a repetitive, save it to the Datastore.
        if not request.skip_new_conf:
            FetchConfigurations(name=request.bapi_id,
                                user=User(request.user),
                                workbook_path=request.workbook_path,
                                workbook=request.workbook,
                                worksheet=request.worksheet,
                                destination_cell=request.destination_cell,
                                organization=queried_list.organization,
                                description=queried_list.description
                                ).put()

        queried_list_pickled = pickle.loads(queried_list.data)
        return FetchResponse(data=queried_list_pickled,
                             bapi_id=request.bapi_id,
                             user=request.user,
                             workbook_path=request.workbook_path,
                             workbook=request.workbook,
                             worksheet=request.worksheet,
                             organization=queried_list.organization,
                             description=queried_list.description,
                             destination_cell=request.destination_cell)

    @remote.method(GetPublishedRequest, GetPublishedResponse)
    def get_published(self, request):
        """
        Get a list of Published Configurations. This list will be used to update
        all previously defined BAPI data.
        """

        workbook_path = request.workbook_path
        workbook = request.workbook

        published_configurations = PublishConfigurations.query(
            PublishConfigurations.workbook_path == workbook_path,
            PublishConfigurations.workbook == workbook
        )

        ids = []
        users = []
        names = []
        workbook_paths = []
        workbooks = []
        worksheets = []
        descriptions = []
        organizations = []
        destination_cells = []
        data_types = []

        for i in published_configurations.iter():
            ids.append(i.bapi_id)
            users.append(i.user.email())
            names.append(i.name)
            workbook_paths.append(i.workbook_path)
            workbooks.append(i.workbook)
            worksheets.append(i.worksheet)
            descriptions.append(i.description)
            organizations.append(i.organization)
            destination_cells.append(i.destination_cell)
            data_types.append(i.data_type)

        return GetPublishedResponse(ids=ids,
                                    users=users,
                                    names=names,
                                    workbook_paths=workbook_paths,
                                    workbooks=workbooks,
                                    worksheets=worksheets,
                                    descriptions=descriptions,
                                    organizations=organizations,
                                    destination_cells=destination_cells,
                                    data_types=data_types)



    @remote.method(GetFetchedRequest, GetFetchedResponse)
    def get_fetched(self, request):
        """
        Get a list of Fetched Configurations. This list will be used to fetch
        all previously fetched BAPI data.
        """

        workbook_path = request.workbook_path
        workbook = request.workbook

        fetched_configurations = FetchConfigurations.query(
            FetchConfigurations.workbook_path == workbook_path,
            FetchConfigurations.workbook == workbook
        )

        names = []
        users = []
        workbook_paths = []
        workbooks = []
        worksheets = []
        destination_cells = []
        organizations = []
        descriptions = []

        for i in fetched_configurations.iter():
            names.append(i.name)
            users.append(i.user.email())
            workbook_paths.append(i.workbook_path)
            workbooks.append(i.workbook)
            worksheets.append(i.worksheet)
            organizations.append(i.organization)
            descriptions.append(i.description)
            destination_cells.append(i.destination_cell)

        return GetFetchedResponse(names=names,
                                  users=users,
                                  workbook_paths=workbook_paths,
                                  workbooks=workbooks,
                                  worksheets=worksheets,
                                  organizations=organizations,
                                  descriptions=descriptions,
                                  destination_cells=destination_cells)

app = service.service_mappings([('/List.*', List)])
