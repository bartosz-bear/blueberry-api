__author__ = 'CHBAPIE'

from google.appengine.ext import ndb
from google.appengine.ext.db import Query
from apis import AqList

import pickle
import os
import webapp2
import jinja2
import logging

JINJA_ENVIRONMENT = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.dirname(__file__)),
    extensions=['jinja2.ext.autoescape'],
    autoescape=True)

class MainPage(webapp2.RequestHandler):
    """
    Display the most recent list.
    """

    def get(self):
        from_db = AqList.query().order(AqList.last_updated)
        #logging.info(from_db.count())
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

        template = JINJA_ENVIRONMENT.get_template('display.html')

        self.response.write(template.render(template_values))


class BrowsePage(webapp2.RequestHandler):
    """
    Display the available lists
    """

    def get(self):

        from_db = AqList.query().fetch()

        template_values = {'data': from_db}

        #for result in from_db.iter():



        """
        template_values['names'] = result.name
        template_values['descriptions'] = result.description
        template_values['ids'] = result.bapi_id
        template_values['organizations'] = result.organization
        template_values['users'] = result.user
        template_values['last_updated'] = result.last_updated
        """

        template = JINJA_ENVIRONMENT.get_template('browse.html')

        self.response.write(template.render(template_values))

application = webapp2.WSGIApplication([
    ('/display', MainPage),
    ('/browse', BrowsePage)
], debug=True)
