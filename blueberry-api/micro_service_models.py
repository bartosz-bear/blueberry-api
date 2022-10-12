import webapp2_extras.appengine.auth.models as auth_models

from google.appengine.ext import ndb
from google.appengine.api import users
from google.appengine.api.users import User

DEFAULT_USER = User('bartosz.piechnik@email.com')


class Filter(ndb.Model):
    """
    Store configurations for a filter pipeline.
    """
    user = ndb.StringProperty(required=True)
    column = ndb.StringProperty()
    filter_value = ndb.StringProperty()
