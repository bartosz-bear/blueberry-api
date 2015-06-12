from bottle import route, run, template, hook, response, request
from test import lets_see
from apiquitous import publish_to_cloud, fetch_new, load_a_publishing_template, get_fetched, fetch_many, get_published, publish_many

import logging

import json
from objbrowser import browse


@hook('before_request')
def enable_cors():
    response.headers['Access-Control-Allow-Origin'] = '*'
    response['Access-Control-Allow-Methods'] = 'POST, GET, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Origin, Accept, Content-Type, X-Requested-With, X-CSRF-Token'
    response['Access-Control-Allow-Headers'] = '*'

@route('/hello/<name>')
def index(name):
    return template('<b>Hello {{name}}</b>!', name=name)


@route('/bartosz')
def my_test():
    # lets_see()
    print 'Janusz and Przemus'

@route('/fetch_new', method=['OPTIONS', 'POST'])
def fetch_new_route():

    #logging.info(request)
    #print request
    #print '\n\n'

    #browse(request)

    #PythonDict = {}
    #for item in request.forms:
    #    PythonDict[item]=request.forms.get(item)

    #print request.POST
    #print type(request.json)

    
    r = fetch_new(request.json)
    response.content_type = 'application/json'

    print response
    return json.dumps(r)

@route('/fetch_many', method=['OPTIONS', 'POST'])
def fetch_many_route():
    fetch_many(request.json)

@route('/publish_many', method=['OPTIONS', 'POST'])
def publish_many_route():
    publish_many(request.json)


@route('/get_fetched', 'POST')
def get_fetched_route():

    print 'Bartosz Artur'

    print 'Jan Jan Jan', get_fetched()

    r = get_fetched()
    response.content_type = 'application/json'

    return json.dumps(r)

@route('/get_published', 'POST')
def get_published_route():

    print 'Published Bartosz sa'

    print 'Jan Jan Jan', get_published()

    r = get_published()
    response.content_type = 'application/json'

    return json.dumps(r)

"""
This is fetch_new which works with JSON-P

@route('/fetch_new')
def fetch_new_route():

    logging.info(request)
    browse(request)

    #for i in request.query.keys():
    #    if i != '_' and i != 'callback':
    #        params = i


    fetch_new(json.loads(params))
"""

# return template('<b>Hello {{ai}}</b>!', ai='Waj')

@route('/publish', method='POST')
def publish():
    """
	Receive JSON from Excel, call publish_to_cloud() to save in the datastore.
	"""
    #print request.query.keys()
    #browse(request)
    #publish_to_cloud(json.loads(request.query.keys()[2]))
    publish_to_cloud(request.json)


#print'Hayek and Mises'
#print json.dumps(request.query.callback)
#req = request.query.keys()[1]
#print req
#j_req = json.loads(req)
#print 'Is it? ', j_req['aq_name']

#browse(request)
#print request.query.keys()[1]
#print 'Most likely callback ', request.query.keys()[0]
#print json.loads(request.query.keys()[1])
#jan = publish_to_cloud(json.loads(request.query.keys()[1]))
#print jan

#print dir(response)
#response.content_type = 'application/json; charset=utf-8'
#response.body = json.dumps({'Albert': 'Japa'})
#print 'response ', response
#print 'response type ', response.content_type
#print 'response body ', response.body
#d = json.dumps(dict(a='Olbrycht'))
#return 'showData(' + d + ');'


#browse(request)
#callback = request.args.get('callback')
#print 'callback={0}&({1})'.format(request.params['callback'], json.dumps({'a':1, 'b':2}))
#return 'callback={0}&{1}'.format(request.params['callback'], json.dumps({'a':1, 'b':2}))
#return 'jQuery({})'.format(json.dumps({'a':1, 'b':2}))
#getthedata({ "hello" : "Hi, I'm JSON. Who are you?"})
#return response.body

@route('/fetch')
def publish():
    fetch_all()


@route('/load_template')
def publish():
    load_a_publishing_template()


run(host='localhost', port=8001)