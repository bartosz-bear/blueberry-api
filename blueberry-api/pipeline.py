__author__ = 'Bartosz Piechnik'

import sys
import pickle
import json
import webapp2
import models
import apis_funcs

from protorpc import messages
from protorpc import message_types
from protorpc import remote
from protorpc.wsgi import service
from google.appengine.api.users import User
from google.appengine.ext import deferred
from webapp2 import redirect
from models import BAPIList, BAPIScalar, BAPIDictionary, FetchConfigurations, PublishConfigurations, Pipeline, BAPIUser
from users import BaseHandler, login_required

import logging
import pdb

config = {}
config['webapp2_extras.sessions'] = {
    'secret_key': 'secret_key_string'
}
config['webapp2_extras.auth'] = {
    'user_model': BAPIUser
}


class AddPipeline(BaseHandler):
    """
    Add a new pipeline to a datastore.
    """

    @login_required
    def post(self):
        name = self.request.POST['name']
        user_email = BAPIUser.get_by_id(self.user['user_id']).email
        if Pipeline.query(Pipeline.user == user_email and Pipeline.name == name).count() == 1:
            self.response.body = "Pipeline name already exists."
        else:
            Pipeline(user=user_email, name=name).put()
            self.response.body = "Pipeline has been added."
        return redirect('/pipelines')

class DeletePipeline(BaseHandler):
    """
    Remove a Blueberry ID from a favorites list.
    """
    @login_required
    def post(self):
        name = self.request.POST['name']
        user_email = BAPIUser.get_by_id(self.user['user_id']).email
        fav_id = Pipeline.query(Pipeline.name == name).get()
        #pdb.set_trace()
        if fav_id.user == user_email:
            fav_id.key.delete()
            self.response.body = "Pipeline has been deleted."
            return redirect('/add-in_configurations')
        else:
            self.response.body = "You don't have permissions to delete this pipeline."
            return redirect('/add-in_configurations')

application = webapp2.WSGIApplication([
    ('/pipelines/add_pipeline', AddPipeline),
    ('/pipelines/delete_pipeline', DeletePipeline)
], debug=True, config=config)