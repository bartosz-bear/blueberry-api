from google.appengine.ext import ndb
from google.appengine.api import users
from google.appengine.api.users import User

DEFAULT_USER = User('bartosz.piechnik@ch.abb.com')


class AqList(ndb.Model):
    """
    Model designed to store Python lists in ndb along with descriptive information: name of the list,
    created by, last updated at and a short description.
    """
    name = ndb.StringProperty()
    user = ndb.UserProperty(default=DEFAULT_USER)
    last_updated = ndb.DateTimeProperty(auto_now_add=True)
    description = ndb.StringProperty()
    organization = ndb.StringProperty()
    data = ndb.BlobProperty()
    bapi_id = ndb.StringProperty()


class FetchConfigurations(ndb.Model):
    """
    Stores configuration for for Data Fetch.
    """
    name = ndb.StringProperty()
    user = ndb.UserProperty(default=DEFAULT_USER)
    workbook_path = ndb.StringProperty(required=True)
    workbook = ndb.StringProperty(required=True)
    worksheet = ndb.StringProperty(required=True)
    destination_cell = ndb.StringProperty(required=True)
    description = ndb.StringProperty(required=True)
    organization = ndb.StringProperty(required=True)


class PublishConfigurations(ndb.Model):
    """
    Stores configuration for for Data Fetch.
    """
    name = ndb.StringProperty()
    user = ndb.UserProperty(default=DEFAULT_USER)
    organization = ndb.StringProperty(required=True)
    description = ndb.StringProperty(required=True)
    bapi_id = ndb.StringProperty(required=True)
    workbook_path = ndb.StringProperty(required=True)
    workbook = ndb.StringProperty(required=True)
    worksheet = ndb.StringProperty(required=True)
    destination_cell = ndb.StringProperty(required=True)
    data_type = ndb.StringProperty(required=True)