import json
import requests
import pickle

aqlist = pickle.dumps(['Bartosz', 'Artur', 'Lazarus'])
#aqlist = ['Bartosz', 'Artur', 'Lazarus']
#aqlist = 'Bart'

#response = requests.post("http://localhost:8080/HelloService.aqlist",
#						 headers={'content_type':'application/json'},
#						 data=json.dumps({'aqlist':aqlist}))

#response = requests.post("http://apiquitous.appspot.com/HelloService.aqlist",
#						 headers={'content_type':'application/json'},
#						 data=json.dumps({'aqlist':aqlist}))

response = requests.post("http://localhost:8080/HelloService.from_db",
						 headers={'content_type':'application/json'},
						 data=json.dumps({'from_db':'from_db'}))

#print response.
#print response.json()['aqlist']

print pickle.loads(response.json()['from_db'])
print type(pickle.loads(response.json()['from_db']))