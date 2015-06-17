__author__ = 'CHBAPIE'

from google.appengine.ext import ndb
from google.appengine.ext.db import Query
from models import BAPIScalar, BAPIList, BAPIDictionary, BAPITable, PublishConfigurations, FetchConfigurations


import pickle
import os
import webapp2
import jinja2
import logging
import pdb
import models
from datetime import timedelta

JINJA_ENVIRONMENT = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.dirname(__file__)),
    extensions=['jinja2.ext.autoescape'],
    autoescape=True)



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

class MainPage(webapp2.RequestHandler):
    """
    Display the most recent list.
    """

    def get(self):
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


class BrowsePage(webapp2.RequestHandler):
    """
    Display the available lists
    """

    def get(self):

        classes = [getattr(models, i) for i in 'BAPIScalar', 'BAPIList', 'BAPIDictionary', 'BAPITable']
        data0 = [remove_duplicates(j.query()) for j in classes]
        data = []
        for i in data0:
            data += i

        template_values = {'data': data}
        template = JINJA_ENVIRONMENT.get_template('templates/browse.html')

        self.response.write(template.render(template_values))

class PublishConfigurationsPage(webapp2.RequestHandler):
    """
    Display the available lists
    """

    def get(self):

        from_db = PublishConfigurations.query().fetch()
        logging.info("Bartosz")
        logging.info(len(from_db))
        if len(from_db) == 0:
            from_db = ''
            template_values = {}
        else:
            template_values = {}
            for result in from_db:
                template_values['data'] = from_db
                """
                template_values['id'] = result.bapi_id
                template_values['user'] = result.user
                template_values['name'] = result.name
                template_values['description'] = result.description
                template_values['workbook_path'] = result.workbook_path
                template_values['workbook'] = result.workbook
                template_values['worksheet'] = result.worksheet
                template_values['destination_cell'] = result.destination_cell
                template_values['data_type'] = result.data_type
                """


        template = JINJA_ENVIRONMENT.get_template('templates/publish_configurations.html')

        self.response.write(template.render(template_values))


class FetchConfigurationsPage(webapp2.RequestHandler):
    """
    Display lists which have been fetched in the past and are expected to be fetched in the future.
    """

    def get(self):

        from_db = FetchConfigurations.query().fetch()
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
    ('/display', MainPage),
    ('/browse', BrowsePage),
    ('/publish_configurations', PublishConfigurationsPage),
    ('/fetch_configurations', FetchConfigurationsPage)
], debug=True)

