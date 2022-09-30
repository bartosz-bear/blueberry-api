__author__ = 'CHBAPIE'

import os
import json
import urllib
import webapp2
import jinja2
import models

from webapp2 import redirect
from google.appengine.ext import ndb
from google.appengine.api import urlfetch
from models import BAPIScalar, BAPIList, BAPIDictionary, BAPITable, PublishConfigurations, FetchConfigurations, BAPIUser, FavoriteIDs
from users import BaseHandler, login_required
from constants import BAPI_DATA_TYPES
from apis_funcs import is_ID_valid

import pdb
import logging

JINJA_ENVIRONMENT = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.dirname(__file__)),
    extensions=['jinja2.ext.autoescape'],
    autoescape=True)

config = {}
config['webapp2_extras.sessions'] = {
    'secret_key': 'MG1VKMXtBpKG'
}
config['webapp2_extras.auth'] = {
    'user_model': BAPIUser
}


class SelectHeadersProxy(BaseHandler):
    """
    Add a new favorite Blueberry ID. The list of favorite Blueberry IDs is used in Blueberry Add-in.
    """
    @login_required
    def post(self):

        request = json.loads(self.request.body)
        data = request["data"]
        headers_list = request["headers_list_all"]
        headers_list_selected = request["headers_list_selected"]

        form_fields = {
              "data": data,
              "headers_list": headers_list,
              "headers_list_selected": headers_list_selected
            }

        url = "http://ec2-54-186-191-26.us-west-2.compute.amazonaws.com/"

        result = urlfetch.fetch(url)
        result = str(result)

        """
        form_data = urllib.urlencode(form_fields)

        url = "http://riskcontrol.pythonanywhere.com/"
        result = urlfetch.fetch(url=url,
                                payload=form_data,
                                method=urlfetch.POST,
                                headers={'Content-Type': 'application/x-www-form-urlencoded'})

        pdb.set_trace()
        """

        self.response.body = result

application = webapp2.WSGIApplication([
    ('/proxy/select_headers', SelectHeadersProxy)
], debug=True, config=config)

