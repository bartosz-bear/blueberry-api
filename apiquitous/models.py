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