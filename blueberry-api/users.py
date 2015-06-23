__author__ = 'CHBAPIE'

import webapp2
import logging
import json
import cgi
from models import User
import pdb
#import jinja2
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'lib'))

from wtforms import Form, TextField, PasswordField, validators

from models import BAPIUser

from webapp2_extras import sessions, auth
from webapp2_extras import jinja2 as jinja2

config = {}
config['webapp2_extras.sessions'] = {
    'secret_key': 'MG1VKMXtBpKG'
}
config['webapp2_extras.auth'] = {
    'user_model': BAPIUser,
}

"""
config['webapp2_extras.jinja2'] = {
    'template_path'
}
"""

logger = logging.getLogger(__name__)


def login_required(handler):
    """
    Requires that a user be logged in to access the resource.
    """
    def check_login(self, *args, **kwargs):
        if not self.user:
            return self.redirect('/login')
        else:
            return handler(self, *args, **kwargs)
    return check_login

def jinja2_factory(app):
    """
    True ninja method for attaching globals/filters to jinja.
    """
    j = jinja2.Jinja2(app)
    j.environment.globals.update({
        'uri_for': webapp2.uri_for
    })
    return j

class BaseHandler(webapp2.RequestHandler):

    def dispatch(self):

        self.session_store = sessions.get_store(request=self.request)

        try:
            # Dispatch the request.
            webapp2.RequestHandler.dispatch(self)
        finally:
            # Save all sessions.
            self.session_store.save_sessions(self.response)

    @webapp2.cached_property
    def session(self):
        # Returns a session using the default cookie key.
        return self.session_store.get_session()


class UserAwareHandler(webapp2.RequestHandler):
    @webapp2.cached_property
    def session_store(self):
        return sessions.get_store(request=self.request)

    @webapp2.cached_property
    def session(self):
        return self.session_store.get_session(backend="datastore")

    def dispatch(self):
        try:
            super(UserAwareHandler, self).dispatch()
        finally:
            self.session_store.save_sessions(self.response)

    @webapp2.cached_property
    def auth(self):
        return auth.get_auth(request=self.request)

    @webapp2.cached_property
    def user(self):
        user = self.auth.get_user_by_session()
        return user

    @webapp2.cached_property
    def user_model(self):
        user_model, timestamp = self.auth.store.user_model.get_by_auth_token(
                self.user['user_id'],
                self.user['token']) if self.user else (None, None)
        return user_model

    @webapp2.cached_property
    def jinja2(self):
        return jinja2.get_jinja2(factory=jinja2_factory, app=self.app)

    def render_response(self, _template, **context):
        ctx = {'user': self.user_model}
        ctx.update(context)
        rv = self.jinja2.render_template(_template, **ctx)
        self.response.write(rv)


class SignupForm(Form):
    email = TextField('Email', [validators.Required(), validators.Email()])
    password = PasswordField('Password', [validators.Required(), validators.EqualTo('confirm_password',
                                                                                    message="Passwords must match.")])
    confirm_password = PasswordField('Confirm Password', [validators.Required()])


class SignupHandler(UserAwareHandler):
    """
    Serves up a signup form, creates new users
    """
    def get(self):
        self.render_response("register.html", form=SignupForm())


    def post(self):
        form = SignupForm(self.request.POST)
        error = None
        if form.validate():
            success, info = self.auth.store.user_model.create_user(
                "auth:" + form.email.data,
                unique_properties=['email'],
                email=form.email.data,
                password_raw=form.password.data)
            if success:
                #self.auth.get_user_by_password("auth:"+form.email.data,
                #                               form.password.data)
                return self.redirect('/register')
            else:
                error = "That email is already in use."
        self.render_response("register.html", form=form, error=error)


class LoginForm(Form):
    email = TextField('Email', [validators.Required(), validators.Email()])
    password = PasswordField('Password', [validators.Required()])


class LoginHandler(UserAwareHandler):
    def get(self):
        self.render_response("log_in.html", form=LoginForm())

    def post(self):
        form = LoginForm(self.request.POST)
        error = None
        if form.validate():
            try:
                self.auth.get_user_by_password(
                    "auth:"+form.email.data,
                    form.password.data)
                logging.info("bbbbbbbbbbbbbbb")
                return self.redirect('/login')
            except (auth.InvalidAuthIdError, auth.InvalidPasswordError):
                error = "Invalid Error/Password"
        self.render_response("log_in.html",
                             form=form,
                             error=error)


class LogoutHandler(UserAwareHandler):
    """Destroy the user session and return them to the login screen."""
    @login_required
    def get(self):
        self.auth.unset_session()
        self.redirect('/login')


class AfterLogoutTest(UserAwareHandler):
    @login_required
    def get(self):
        self.redirect('/display')



class SessionTest(BaseHandler):
    """
    Testing
    """
    def post(selfs):
        pdb.set_trace()



class LogIn(BaseHandler):
    """
    Handling user log in and session management process.
    """

    def post(self):

        #logging.info(self.session)
        #logging.info(dir(self.session))
        #self.session['foo'] = 'bar'

        jdata = json.loads(cgi.escape(self.request.body))
        username = jdata['username']
        password = jdata['password']

        if type(User.query(User.username == username, User.password == password).get()).__class__ is type(User):
            if self.session.get('counter'):
                self.response.out.write('Session is in place')
                counter = self.session.get('counter')
                self.session['counter'] = counter + 1
                self.response.out.write('Counter = ' + str(self.session.get('counter')))
            else:
                self.response.out.write('Fresh Session')
                self.session['counter'] = 1
                self.response.out.write('Counter = ' + str(self.session.get('counter')))
                data = {'message': 'Correct password.'}
        else:
            data = {'message': 'Wrong username or password'}

        #self.response.set_cookie('session_id', self.session.get('counter'), max_age=360, domain='localhost:8080')

        self.response.headers['Content-Type'] = 'application/json'
        self.response.out.write(json.dumps(data))

    @staticmethod
    def testing_method(self):
        logging.info("Testing Class method")


class Register(webapp2.RequestHandler):
    """
    Temporary User registration handler
    """

    def post(self):

        jdata = json.loads(cgi.escape(self.request.body))
        username = jdata['username']
        password = jdata['password']

        User(username=username, password=password).put()


app = webapp2.WSGIApplication([
    ('/logging', LogIn),
    ('/register', SignupHandler),
    ('/logging_testing', SessionTest),
    ('/login', LoginHandler),
    ('/logout', LogoutHandler),
    ('/test', AfterLogoutTest)
], debug=True, config=config)