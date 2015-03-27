import sys
sys.path.append('C:\Program Files (x86)\Google\google_appengine')

from google.appengine.ext.remote_api import remote_api_stub
from guestbook import Greeting
import getpass

def auth_func():
  return (raw_input('Username:'), getpass.getpass('Password:'))

remote_api_stub.ConfigureRemoteApi(None, '/remote_api', auth_func,
                               'apiquitous.appspot.com')


greeting = Greeting(author='Bartosz', content='Hello there, again and again')
greeting.put()

# Fetch the most recent 10 guestbook entries
#entries = helloworld.Greeting.all().order("-date").fetch(10)
# Create our own guestbook entry
#helloworld.Greeting(content="A greeting").put()