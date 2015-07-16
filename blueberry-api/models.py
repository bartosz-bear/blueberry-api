from google.appengine.ext import ndb
from google.appengine.api import users
from google.appengine.api.users import User

import webapp2_extras.appengine.auth.models as auth_models

DEFAULT_USER = User('bartosz.piechnik@ch.abb.com')


class BAPIScalar(ndb.Model):
    """
    Model designed to store scalar values of any type in ndb along with descriptive information: name of the scalar,
    created by, last updated at and a short description.
    """
    bapi_id = ndb.StringProperty()
    name = ndb.StringProperty()
    user = ndb.StringProperty(required=True)
    last_updated = ndb.DateTimeProperty(auto_now_add=True)
    description = ndb.StringProperty()
    organization = ndb.StringProperty()
    data = ndb.BlobProperty()


class BAPIList(ndb.Model):
    """
    Model designed to store lists in ndb along with descriptive information: name of the list,
    created by, last updated at and a short description.
    """
    bapi_id = ndb.StringProperty()
    name = ndb.StringProperty()
    user = ndb.StringProperty(required=True)
    last_updated = ndb.DateTimeProperty(auto_now_add=True)
    description = ndb.StringProperty()
    organization = ndb.StringProperty()
    data = ndb.BlobProperty()


class BAPIDictionary(ndb.Model):
    """
    Model designed to store dictionaries in ndb along with descriptive information: name of the dictionary,
    created by, last updated at and a short description.
    """
    bapi_id = ndb.StringProperty()
    name = ndb.StringProperty()
    user = ndb.StringProperty(required=True)
    last_updated = ndb.DateTimeProperty(auto_now_add=True)
    description = ndb.StringProperty()
    organization = ndb.StringProperty()
    data = ndb.BlobProperty()


class BAPITable(ndb.Model):
    """
    Model designed to store tables in ndb along with descriptive information: name of the table,
    created by, last updated at and a short description.
    """
    bapi_id = ndb.StringProperty()
    name = ndb.StringProperty()
    user = ndb.StringProperty(required=True)
    last_updated = ndb.DateTimeProperty(auto_now_add=True)
    description = ndb.StringProperty()
    organization = ndb.StringProperty()
    data = ndb.BlobProperty()


class FetchConfigurations(ndb.Model):
    """
    Stores configuration for for Data Fetch.
    """
    name = ndb.StringProperty()
    user = ndb.StringProperty(required=True)
    organization = ndb.StringProperty(required=True)
    description = ndb.StringProperty(required=True)
    bapi_id = ndb.StringProperty(required=True)
    workbook_path = ndb.StringProperty(required=True)
    workbook = ndb.StringProperty(required=True)
    worksheet = ndb.StringProperty(required=True)
    destination_cell = ndb.StringProperty(required=True)
    data_type = ndb.StringProperty(required=True)


class PublishConfigurations(ndb.Model):
    """
    Stores configuration for for Data Fetch.
    """
    name = ndb.StringProperty()
    user = ndb.StringProperty(required=True)
    organization = ndb.StringProperty(required=True)
    description = ndb.StringProperty(required=True)
    bapi_id = ndb.StringProperty(required=True)
    workbook_path = ndb.StringProperty(required=True)
    workbook = ndb.StringProperty(required=True)
    worksheet = ndb.StringProperty(required=True)
    destination_cell = ndb.StringProperty(required=True)
    data_type = ndb.StringProperty(required=True)


# Is this class really necessary?
class User(ndb.Model):
    """
    Class for storing User information.
    """
    username = ndb.StringProperty()
    password = ndb.StringProperty()


class BAPIUser(auth_models.User):
    """
    User class inherited from webapp2 extras.
    """
    email = ndb.StringProperty()

"""

Possibly this can be used later to optimize query on /display

class BAPICounts(ndb.Model):
    ""
    Model responsible for counting number of items in each of the fundamental BAPI data types.
    ""
    bapi_scalar_count = ndb.IntegerProperty()
    bapi_list_count = ndb.IntegerProperty()
    bapi_dictionary_count = ndb.IntegerProperty()
    bapi_table_count = ndb.IntegerProperty()
"""