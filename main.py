#!/usr/bin/python3

import json
import urllib.request
import urllib.parse
import requests
import mailparser
import yaml

## Load configuration
with open("config.yml", 'r') as stream:
    try:
        config = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(exc)

appId = config['appId']
appSecret = config['appSecret']
tenantId = config['tenantId']
user = config['user']
inbox_id = config['inbox_id']
deleteditems_id = config['deleteditems_id']

## Azure Active Directory token endpoint.
url = "https://login.microsoftonline.com/%s/oauth2/v2.0/token" % (tenantId)
body = {
    'client_id': appId,
    'client_secret': appSecret,
    'grant_type': 'client_credentials',
    'scope': 'https://graph.microsoft.com/.default'
}

## Authenticate and obtain AAD Token for future calls
data = urllib.parse.urlencode(body).encode("utf-8")  # encodes the data into a 'x-www-form-urlencoded' type
req = urllib.request.Request(url, data)
response = urllib.request.urlopen(req)
jsonResponse = json.loads(response.read().decode())

## Grab the token from the response then store it in the headers dict.
aadToken = jsonResponse["access_token"]
headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Authorization': "Bearer " + aadToken
}

api_root = "https://graph.microsoft.com/v1.0/"

if len(aadToken) > 0:
    print("Access token acquired.")

## HTTP functions
def make_request(url):
    """
    Makes a GET request.

    :param url: Url of the request.
    :returns: json response.
    :raises HTTPError: raises an exception
    """
    url_sanitized = urllib.parse.quote(url, safe="%/:=&?~#+!$,;'@()*[]")  # Url encode spaces
    req = urllib.request.Request(url_sanitized, headers=headers)
    print()
    print("########################################################################################")
    print("Calling the Microsoft Graph  API...")
    print()
    print('GET "%s"' % url_sanitized)
    print()
    print("Headers :")
    print(json.dumps(headers, indent=4))
    print("########################################################################################")

    try:
        response = urllib.request.urlopen(req)
    except urllib.error.HTTPError as e:
        raise e

    return response.read().decode()

def make_request_post(url, data):
    """
    Makes a GET request.

    :param url: Url of the request.
    :returns: json response.
    :raises HTTPError: raises an exception
    """
    url_sanitized = urllib.parse.quote(url, safe="%/:=&?~#+!$,;'@()*[]")  # Url encode spaces
    req = urllib.request.Request(url, data=data.encode(), headers=headers)
    print()
    print("########################################################################################")
    print("Calling the Microsoft Graph  API...")
    print()
    print('POST "%s"' % url_sanitized)
    print()
    print("Headers :")
    print(json.dumps(headers, indent=4))
    print("########################################################################################")

    try:
        response = urllib.request.urlopen(req)
    except urllib.error.HTTPError as e:
        raise e

    return response

## Open the output file
output = open('mail.json', 'a')

## Retrieve messages
messages = "https://graph.microsoft.com/v1.0/users('{}')/mailFolders/{}/messages?$select=id".format(user, inbox_id)
response = make_request(messages)
jsonResponse = json.loads(response)

## Loop through messages and process
for mail in jsonResponse['value']:
    message = "https://graph.microsoft.com/v1.0/users('{}')/messages/{}/$value".format(user, mail['id'])
    response = make_request(message)
    mime = response
    m = mailparser.parse_from_string(mime)
    json.dump(json.loads(m.mail_json), output)
    output.write("\r\n")

    delete = "https://graph.microsoft.com/v1.0/users('{}')/messages/{}/move".format(user, mail['id'])
    response = make_request_post(delete, '{"destinationId": "' + deleteditems_id + '"}')
    print(response.read())