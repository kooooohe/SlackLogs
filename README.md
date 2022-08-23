# SlackLogs

## Set up Clasp

https://github.com/google/clasp

make .clasp.json 

ex.
```
{
  "scriptId": "",
  "rootDir": "build/",
  "fileExtension": "ts",
}
```

## Slack Token and set permissons
make SlackApp & get token
OAuth & Permissions > Bot Token Scopes
```
channels:history
channels:join
channels:read
files:read
users:read
```

## Set Goole App Script Propaties
FOLDER_ID=XXXXX // Your own Folder
SLACK_TOKEN=XXXXX

## Deploy
clasp push

