__author__ = 'Bartosz Piechnik'

import os
import jinja2
import webapp2
import models

from google.appengine.ext import ndb
from webapp2 import redirect
from constants import BAPI_DATA_TYPES
from models import BAPIScalar, BAPIList, BAPIDictionary, BAPITable, PublishConfigurations, FetchConfigurations, BAPIUser, FavoriteIDs
from users import BaseHandler, login_required
from apis_funcs import is_ID_valid

import pdb
import logging

JINJA_ENVIRONMENT = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.dirname(__file__)),
    extensions=['jinja2.ext.autoescape'],
    autoescape=True)

config = {}
config['webapp2_extras.sessions'] = {
    'secret_key': 'secret_key_string'
}
config['webapp2_extras.auth'] = {
    'user_model': BAPIUser
}


class AddFavorite(BaseHandler):
    """
    Add a new favorite Blueberry ID. The list of favorite Blueberry IDs is used in Blueberry Add-in.
    """
    @login_required
    def post(self):
        bapi_id = self.request.POST['bapi_id']
        user_email = BAPIUser.get_by_id(self.user['user_id']).email
        MyDictClass = type('object', (), {})
        request_as_class = MyDictClass()
        request_as_class.bapi_id = bapi_id

        try:
            data_type = self.request.POST['bapi_id'].split('.')[2]
        except IndexError:
            self.response.body = "ID is incorrect."
            return redirect('/add-in_configurations')
        class_ = getattr(models, 'BAPI' + data_type)
        if is_ID_valid(request_as_class, class_):
            if FavoriteIDs.query(FavoriteIDs.user == user_email).count() < 5:
                FavoriteIDs(user=user_email, bapi_id=bapi_id, name=bapi_id.split('.')[1]).put()
                self.response.body = "ID has been added to your favorites."
            else:
                self.response.body = "You can only have 5 favorite IDs."
            return redirect('/add-in_configurations')
        self.response.body = "ID doesn't exist."
        return redirect('/add-in_configurations')


class RemoveFavorite(BaseHandler):
    """
    Remove a Blueberry ID from a favorites list.
    """
    @login_required
    def post(self):
        bapi_id = self.request.POST['bapi_id']
        user_email = BAPIUser.get_by_id(self.user['user_id']).email
        fav_id = FavoriteIDs.query(FavoriteIDs.bapi_id == bapi_id).get()
        if fav_id.user == user_email:
            fav_id.key.delete()
            self.response.body = "ID has been deleted."
            return redirect('/add-in_configurations')
        else:
            self.response.body = "You don't have permissions to delete this ID."
            return redirect('/add-in_configurations')

application = webapp2.WSGIApplication([
    ('/add-in/add_favorite', AddFavorite),
    ('/add-in/remove_favorite', RemoveFavorite)
], debug=True, config=config)

