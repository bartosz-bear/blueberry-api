__author__ = 'CHBAPIE'

import os
import pickle
from datetime import timedelta

import webapp2
import jinja2

import models
from google.appengine.ext import ndb
from google.appengine.ext.db import Query
from models import BAPIScalar, BAPIList, BAPIDictionary, BAPITable, PublishConfigurations, FetchConfigurations, BAPIUser
from google.appengine.api.users import User, get_current_user
from users import BaseHandler, login_required
from constants import BAPI_DATA_TYPES

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


def remove_duplicates(query):
    """
    Takes a query of a Datastore table as a parameter
    and return a list of results without duplicates.
    """

    fetched_query = query.fetch()
    ids = {i.bapi_id for i in fetched_query}
    return [query.filter(BAPIList.bapi_id == i).order(-BAPIList.last_updated).get() for i in ids]

def time_minus_2(value):
    """
    It's a temporary JINJA2 custom filter to change the display date in the /browse URL.
    Later it should be replaced by a more permanent solution.
    :param value:
    :return:
    """
    value = (value + timedelta(hours=2))
    value = value.strftime('%Y/%m/%d %H:%M:%S')
    return value

JINJA_ENVIRONMENT.filters['time_minus_2'] = time_minus_2


class IndexPage(webapp2.RequestHandler):
    """
    Displays the home page.
    """
    def get(self):
        template = JINJA_ENVIRONMENT.get_template('templates/index.html')
        self.response.write(template.render())

class MainPage(webapp2.RequestHandler):
    """
    Display the most recent list.
    """

    def get(self):

        """
        data_summary = {}
        for i in BAPI_DATA_TYPES:
            data_summary[i] =
        """
        from_db = BAPIList.query().order(BAPIList.last_updated)
        if from_db.count() == 0:
            from_db = ''
            template_values = {}
        else:
            template_values = {}
            for result in from_db.iter():
                template_values['name'] = result.name
                template_values['id'] = result.bapi_id
                template_values['organization'] = result.organization
                template_values['user'] = result.user
                template_values['last_updated'] = result.last_updated
                template_values['description'] = result.description
            template_values['from_db'] = pickle.loads(result.data)


        template = JINJA_ENVIRONMENT.get_template('templates/display.html')

        self.response.write(template.render(template_values))


class BrowsePage(BaseHandler):
    """
    Display the available lists
    """
    @login_required
    def get(self):

        classes = [getattr(models, i) for i in 'BAPIScalar', 'BAPIList', 'BAPIDictionary', 'BAPITable']
        data0 = [remove_duplicates(j.query()) for j in classes]
        data = []
        for i in data0:
            data += i

        template_values = {'data': data}
        template = JINJA_ENVIRONMENT.get_template('templates/browse.html')

        self.response.write(template.render(template_values))

class PublishConfigurationsPage(BaseHandler):
    """
    Display the available lists
    """
    @login_required
    def get(self):
        user_email = BAPIUser.get_by_id(self.user['user_id']).email
        from_db = PublishConfigurations.query(PublishConfigurations.user == user_email).fetch()
        if len(from_db) == 0:
            from_db = ''
            template_values = {}
        else:
            template_values = {}
            for result in from_db:
                template_values['data'] = from_db

        template = JINJA_ENVIRONMENT.get_template('templates/publish_configurations.html')

        self.response.write(template.render(template_values))


class FetchConfigurationsPage(BaseHandler):
    """
    Display lists which have been fetched in the past and are expected to be fetched in the future.
    """
    @login_required
    def get(self):
        user_email = BAPIUser.get_by_id(self.user['user_id']).email
        from_db = FetchConfigurations.query(FetchConfigurations.user == user_email).fetch()
        if len(from_db) == 0:
            from_db = ''
            template_values = {}
        else:
            template_values = {}
            for result in from_db:
                template_values['data'] = from_db

        template = JINJA_ENVIRONMENT.get_template('templates/fetch_configurations.html')

        self.response.write(template.render(template_values))


application = webapp2.WSGIApplication([
    ('/', IndexPage),
    ('/display', MainPage),
    ('/browse', BrowsePage),
    ('/publish_configurations', PublishConfigurationsPage),
    ('/fetch_configurations', FetchConfigurationsPage)
], debug=True, config=config)

