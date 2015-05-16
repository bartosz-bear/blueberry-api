from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service

from google.appengine.api.users import User

from models import AqList, FetchConfigurations, PublishConfigurations

import logging
import pickle
import codecs

# Create the request string containing user's name
class ListRequest(messages.Message):
    """
    Request Object which collects information from JSON passed via HTTP request.
    """
    aq_name = messages.StringField(1)
    aq_created_by = messages.StringField(2)
    aq_description = messages.StringField(3)
    aq_organization = messages.StringField(4)
    aq_data = messages.StringField(5, repeated=True)
    aq_workbook_path = messages.StringField(6)
    aq_workbook = messages.StringField(7)
    aq_worksheet = messages.StringField(8)
    aq_destination_cell = messages.StringField(9)
    aq_id = messages.StringField(10)
    aq_type = messages.StringField(11)


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
    aq_id = messages.StringField(1, required=True)
    user = messages.StringField(2)
    workbook_path = messages.StringField(3)
    workbook = messages.StringField(4)
    worksheet = messages.StringField(5)
    destination_cell = messages.StringField(6)
    organization = messages.StringField(7)
    description = messages.StringField(8)
    skip_new_conf = messages.BooleanField(9)


class FetchResponse(messages.Message):
    """
    Response object which returns data from Datastore.
    """
    aq_id = messages.StringField(1, required=True)
    user = messages.StringField(2)
    workbook_path = messages.StringField(3)
    workbook = messages.StringField(4)
    worksheet = messages.StringField(5)
    destination_cell = messages.StringField(6)
    organization = messages.StringField(7)
    description = messages.StringField(8)
    aq_data = messages.StringField(9, repeated=True)


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
    workbook_paths = messages.StringField(3, repeated=True)
    workbooks = messages.StringField(4, repeated=True)
    worksheets = messages.StringField(5, repeated=True)
    organizations = messages.StringField(6, repeated=True)
    descriptions = messages.StringField(7, repeated=True)
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

        # Validate that data in request.aq_data is one of acceptable data formats by

        logging.info(request)
        logging.info(request.aq_name)
        logging.info(request.aq_data)

        is_configured = PublishConfigurations.query(PublishConfigurations.bapi_id == request.aq_id).count()

        logging.info(is_configured)

        if is_configured == 0:
            # In case this is the first 'Publish' for this id, create a 'PublishConfiguration'
            name = request.aq_name
            user = User(request.aq_created_by)
            description = request.aq_description
            organization = request.aq_organization
            data_type = request.aq_type
            bapi_id = request.aq_id
            #bapi_id = request.aq_organization + '.' + request.aq_name.replace(' ', '_') + '.List'
            logging.info('Bartosz 0')
            PublishConfigurations(name=name,
                                  user=user,
                                  description=description,
                                  organization=organization,
                                  bapi_id=bapi_id,
                                  workbook_path=request.aq_workbook_path,
                                  workbook=request.aq_workbook,
                                  worksheet=request.aq_worksheet,
                                  destination_cell=request.aq_destination_cell,
                                  data_type=data_type
                                  ).put()
        else:
            # It's not the first 'Publish', therefore fetch info about this ID from the database.
            item = AqList.query(AqList.bapi_id == request.aq_id).get()
            logging.info('Henryk')
            name = item.name
            user = item.user
            description = item.description
            organization = item.organization

        AqList(name=name,
               user=user,
               description=description,
               organization=organization,
               bapi_id=request.aq_id,
               data=pickle.dumps(request.aq_data)).put()

        #logging.info(check_conf.count())

        #check_conf = PublishConfigurations.query(PublishConfigurations.name == request.aq_name)

        return ListResponse(response='The list has been uploaded')

    @remote.method(GetPublishedRequest, GetPublishedResponse)
    def get_published(self, request):

        workbook_path = request.workbook_path
        workbook = request.workbook

        from_db = PublishConfigurations.query(
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

        for i in from_db.iter():
            logging.info(type(i.user))
            logging.info(type(i.user.email))
            logging.info(dir(i.user))
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

    @remote.method(FetchRequest, FetchResponse)
    def fetch(self, request):
        aq_organization, aq_name, aq_type = request.aq_id.split('.')
        del(aq_type)
        logging.info(aq_name)
        logging.info(request.skip_new_conf)
        from_db = AqList.query(
            AqList.organization == aq_organization,
            AqList.name == aq_name.replace('_', ' ')).order(-AqList.last_updated).get()
        if not request.skip_new_conf:
            logging.info('Skipping, then why?')
            FetchConfigurations(name=request.aq_id,
                                user=User(request.user),
                                workbook_path=request.workbook_path,
                                workbook=request.workbook,
                                worksheet=request.worksheet,
                                destination_cell=request.destination_cell,
                                organization=from_db.organization,
                                description=from_db.description
                                ).put()


        print 'Jan Tomaszewski', from_db
        from_db_pickled = pickle.loads(from_db.data)
        return FetchResponse(aq_data=from_db_pickled,
                             aq_id=request.aq_id,
                             user=request.user,
                             workbook_path=request.workbook_path,
                             workbook=request.workbook,
                             worksheet=request.worksheet,
                             organization=from_db.organization,
                             description=from_db.description,
                             destination_cell=request.destination_cell)

    @remote.method(GetFetchedRequest, GetFetchedResponse)
    def get_fetched(self, request):

        workbook_path = request.workbook_path
        workbook = request.workbook

        from_db = FetchConfigurations.query(
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

        for i in from_db.iter():
            logging.info(type(i.user))
            logging.info(type(i.user.email))
            logging.info(dir(i.user))
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
