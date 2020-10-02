# Python Microsoft 365 mail to JSON-file through

Reads a Microsoft 365 mailbox through Microsoft Graph and writes the e-mails in JSON-format to a file (line by line).
The output file can be read by filebeat, to forward the e-mails to logstash / elasticsearch.
Adding the script to a cronjob allows to repeat this every X.

Note: The script automatically moves all e-mail messages processed.

## Requirements

``pip3 install mail-parser``

## Usage

Rename example.yml to config.yml and set the values.

The values explained:
- appId: the Azure AD application ID used to connect to Graph
- appSecret: the secret connected for the Azure AD application
- tenantId: the tenantID that contains the application
- user: the ID of the user to read the e-mail from
- inbox_id: the ID of the Inbox folder
- deleteditems_id: the ID of the Deleted items folder

After configuration, run the script with:

``python3 main.py``

## Some useful calls to obtain the required config values:

### Retrieve user ID's

```
messages = "https://graph.microsoft.com/v1.0/users"
response = make_request(messages)
print(response)
```

### Retrieve mailFolder ID's

```
messages = "https://graph.microsoft.com/v1.0/users('{}')/mailFolders".format(user)
response = make_request(messages)
print(response)
```