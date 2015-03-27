"""Hello World API implemented using Google Cloud Endpoints.

Defined here are the ProtoRPC messages needed to define Schemas for methods
as well as those methods defined in an API.
"""


import endpoints
from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service

from google.appengine.ext import ndb
from google.appengine.api import users
from google.appengine.api.users import User

from endpoints_proto_datastore.ndb import EndpointsModel

import logging
import pickle

bar = "Bartoszelomew"

logging.info("value of my var is %s", str(messages.__file__))

import pickle

# TODO: Replace the following lines with client IDs obtained from the APIs
# Console or Cloud Console.
WEB_CLIENT_ID = 'replace this with your web client application ID'
ANDROID_CLIENT_ID = 'replace this with your Android client ID'
IOS_CLIENT_ID = 'replace this with your iOS client ID'
ANDROID_AUDIENCE = WEB_CLIENT_ID

package = 'Hello'
DEFAULT_USER = user = users.User('bartosz.piechnik@ch.abb.com')

class AqList(ndb.Model):
    """
    It's a Python list stored in ndb.
    """
    list_name = ndb.StringProperty()
    user_id = ndb.UserProperty(default=DEFAULT_USER)
    last_updated = ndb.DateTimeProperty(auto_now_add=True)
    list_description = ndb.StringProperty()
    aq_list = ndb.StringProperty()

my_list = ['Bartosz', 'Artur', 'Lazarus']
my_list = pickle.dumps(my_list)

# Create the request string containing user's name
class HelloRequest(messages.Message):
    aqlist = messages.StringField(1, required=True)

# Create the response string
class HelloResponse(messages.Message):
    aqlist = messages.StringField(1, required=True)

# Create the request string containing user's name
class FromDBRequest(messages.Message):
    from_db = messages.StringField(1, required=True)

# Create the response string
class FromDBResponse(messages.Message):
    from_db = messages.StringField(1, required=True)

# Create the RPC service to exchange messages
class HelloService(remote.Service):

    @remote.method(HelloRequest, HelloResponse)
    def aqlist(self, request):
        logging.info(type(request.aqlist))
        logging.info(request.aqlist)
        list_description = 'This list contains all banks ABB has a relationship with. This includes deposits'\
                           'guarantees or letters of credits.'
        AqList(aq_list=str(request.aqlist), list_name='List of ABB banks', list_description=list_description).put()
        return HelloResponse(aqlist='The list has been uploaded')

    @remote.method(FromDBRequest, FromDBResponse)
    def from_db(self, request):
        from_db = AqList.query()
        for result in from_db.iter():
            from_db = result.aq_list
        logging.info(from_db)
        logging.info(request.from_db)
        return FromDBResponse(from_db=from_db)


























'''



class Task(EndpointsModel):
    """
    My first EndpointsModel
    """
    name = ndb.StringProperty(required=True)
    owner = ndb.StringProperty()

@endpoints.api(name='tasks', version='vGDL', description='API for Task Management')
class TaskApi(remote.Service):

    @Task.method(name='task.insert',
                 path='task')
    def insert_task(self, task):
        task.put()
        task
'''
class Greeting(messages.Message):
    """Greeting that stores a message."""
    message = messages.StringField(1)


class GreetingCollection(messages.Message):
    """Collection of Greetings."""
    items = messages.MessageField(Greeting, 1, repeated=True)


STORED_GREETINGS = GreetingCollection(items=[
    Greeting(message=pickle.dumps([0,1,2])),
    Greeting(message='goodbye world!'),
    Greeting(message='Bartosz was here'),
])


@endpoints.api(name='helloworld', version='v1',
               allowed_client_ids=[WEB_CLIENT_ID, ANDROID_CLIENT_ID,
                                   IOS_CLIENT_ID],
               audiences=[ANDROID_AUDIENCE])
class HelloWorldApi(remote.Service):
    """Helloworld API v1."""

    ####################
    ### GREETING GET ###
    ####################
    ID_RESOURCE = endpoints.ResourceContainer(
            message_types.VoidMessage,
            id=messages.IntegerField(1, variant=messages.Variant.INT32))

    @endpoints.method(ID_RESOURCE, Greeting,
                      path='hellogreeting/{id}', http_method='GET',
                      name='greetings.getGreeting')
    def greeting_get(self, request):
        try:
            return STORED_GREETINGS.items[request.id]
        except (IndexError, TypeError):
            raise endpoints.NotFoundException('Greeting %s not found.' %
                                              (request.id,))
    ##########################
    ### GREETINGS MULTIPLY ###
    ##########################
    MULTIPLY_METHOD_RESOURCE = endpoints.ResourceContainer(
            Greeting,
            times=messages.IntegerField(2, variant=messages.Variant.INT32,
                                        required=True))

    @endpoints.method(MULTIPLY_METHOD_RESOURCE, Greeting,
                      path='hellogreeting/{times}', http_method='POST',
                      name='greetings.multiply')
    def greetings_multiply(self, request):
        return Greeting(message=request.message * request.times)

    ######################
    ### GREETINGS LIST ###
    ######################
    @endpoints.method(message_types.VoidMessage, GreetingCollection,
                      path='hellogreeting', http_method='GET',
                      name='greetings.listGreeting')
    def greetings_list(self, unused_request):
        return STORED_GREETINGS

    ##############################
    ### GREETING AUTHORIZATION ###
    ##############################
    @endpoints.method(message_types.VoidMessage, Greeting,
                      path='hellogreeting/authed', http_method='POST',
                      name='greetings.authed')
    def greeting_authed(self, request):
        current_user = endpoints.get_current_user()
        email = (current_user.email() if current_user is not None
                 else 'Anonymous')
        return Greeting(message='hello %s' % (email,))


APPLICATION = endpoints.api_server([HelloWorldApi])

# Map the RPC service and path (/hello)
app = service.service_mappings([('/HelloService.*', HelloService)])
