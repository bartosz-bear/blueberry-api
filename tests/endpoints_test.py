import pprint
import pickle

from apiclient.discovery import build

def main():
  # Build a service object for interacting with the API.
  #api_root = 'https://apiquitous.appspot.com/_ah/api'
  api_root = 'http://localhost:8080/_ah/api'
  api = 'helloworld'
  version = 'v1'
  discovery_url = '%s/discovery/v1/apis/%s/%s/rest' % (api_root, api, version)
  service = build(api, version, discoveryServiceUrl=discovery_url)

  # Fetch all greetings and print them out.
  #response = service.greetings().list().execute()
  #pprint.pprint(response)

  # Fetch a single greeting and print it out.
  #response = service.greetings().get(id='0').execute()
  #pprint.pprint(service.greetings().listGreeting().execute().viewitems())
  #pprint.pprint(pickle.loads((service.greetings().getGreeting(id='0').execute().get('message'))))
  message = service.greetings().getGreeting(id='0').execute().get('message').decode("utf-8")
  print type(pickle.loads(message))
  #print pickle.loads(str(message))

if __name__ == '__main__':
  main()