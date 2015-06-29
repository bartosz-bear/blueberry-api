__author__ = 'CHBAPIE'

import os
import sys
import json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'lib'))

from webob.multidict import MultiDict
import webapp2
from webapp2_extras import sessions, auth
from webapp2_extras import jinja2 as jinja2
from wtforms import Form, TextField, PasswordField, validators

from models import BAPIUser

import pdb
import logging

config = {}
config['webapp2_extras.sessions'] = {
    'secret_key': 'MG1VKMXtBpKG'
}
config['webapp2_extras.auth'] = {
    'user_model': BAPIUser
}


def login_required(handler):
    """
    Decorator which performs a user authentication before a particular handler is executed.
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
    """
    Basehandler provides several methods and properties for all RequestHandlers which require sessions
    and user management functionalities.
    """
    @webapp2.cached_property
    def session_store(self):
        return sessions.get_store(request=self.request)

    @webapp2.cached_property
    def session(self):
        return self.session_store.get_session(backend="datastore")

    def dispatch(self):
        try:
            super(BaseHandler, self).dispatch()
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


class RegisterForm(Form):
    """
    WTForm which provides validations for user registration process. It makes sure that email, password and password confirmation
    has been entered by the user. It also checks if password and password confirmation is the same. Finally it checks
    if the email is a valid email address.
    """
    email = TextField('Email', [validators.Required(), validators.Email()])
    password = PasswordField('Password', [validators.Required(), validators.EqualTo('confirm_password',
                                                                                    message="Passwords must match.")])
    confirm_password = PasswordField('Confirm Password', [validators.Required()])


class RegisterHandler(BaseHandler):
    """
    RegisterHandler is responsible for creating a new user in a datastore.
    """
    def get(self):
        self.render_response("register.html", form=RegisterForm())

    def post(self):
        form = RegisterForm(self.request.POST)
        error = None
        if form.validate():
            success, info = self.auth.store.user_model.create_user(
                "auth:" + form.email.data,
                unique_properties=['email'],
                email=form.email.data,
                password_raw=form.password.data)
            if success:
                return self.redirect('/register')
            else:
                error = "That email is already in use."
        self.render_response("register.html", form=form, error=error)


class LoginForm(Form):
    """
    WTForm which provides validation for user login process. It makes sure that both email and password has been
    entered by the user. It also checks that the email was entered in the valid email address form.
    """
    email = TextField('Email', [validators.Required(), validators.Email()])
    password = PasswordField('Password', [validators.Required()])


class LoginHandler(BaseHandler):
    """
    LoginHandler checks the user and the password against the users datastore. If email or password is incorrect/not
    matching the handler returns "Invalid Error/Password" error.
    """
    def get(self):
        self.render_response("login.html", form=LoginForm())

    def post(self):
        # Check if the request comes from a browser or from Excel.
        # Potential risk exists here that this handler can be exploited. Some security measure should be applied here.
        form = LoginForm(self.request.POST)
        error = None
        if form.validate():
            try:
                self.auth.get_user_by_password(
                    "auth:"+form.email.data,
                    form.password.data)
                # If self.request.user_agent is None then it's a request coming from Excel Add-in, otherwise
                # the request is coming from a browser.
                if not self.request.user_agent:
                    self.response.headers['Content-Type'] = 'application/json'
                    return self.response.write(json.dumps(self.user))
                self.redirect('/display')
            except (auth.InvalidAuthIdError, auth.InvalidPasswordError):
                error = "Invalid username or password"
                # If self.request.user_agent is None then it's a request coming from Excel Add-in, otherwise
                # the request is coming from a browser.
                if not self.request.user_agent:
                    return self.response
                self.render_response("login.html",
                                     form=form,
                                     error=error)


class LogoutHandler(BaseHandler):
    """
    Terminate the user session and return them to the login screen.
    """
    @login_required
    def get(self):
        self.auth.unset_session()
        if not self.request.user_agent:
            return self.response.write('OK')
        self.redirect('/login')


app = webapp2.WSGIApplication([
    ('/register', RegisterHandler),
    ('/login', LoginHandler),
    ('/logout', LogoutHandler)
], debug=True, config=config)