import json
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.deployments',
    'https://www.googleapis.com/auth/script.external_request',
    'https://www.googleapis.com/auth/script.scriptapp',
]

flow = InstalledAppFlow.from_client_secrets_file(
    '/Users/hiraokawashin/.config/gcp-oauth.keys.json',
    scopes=SCOPES
)
creds = flow.run_local_server(port=8080)

data = {
    'access_token':  creds.token,
    'refresh_token': creds.refresh_token,
    'scope':         ' '.join(SCOPES),
    'token_type':    'Bearer',
    'expiry_date':   9999999999999
}
json.dump(data, open('/Users/hiraokawashin/.config/gdrive-server-credentials.json', 'w'))
print('認証完了！')
