import glob
import os
import sys

from google_auth_oauthlib.flow import InstalledAppFlow
from google.cloud import bigquery

SCOPES = ['https://www.googleapis.com/auth/bigquery']


def _find_secret_file():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    matches = glob.glob(os.path.join(base_dir, 'client_secret*.json'))
    if not matches:
        print(f"Could not find client_secret*.json in {base_dir}")
        sys.exit(1)
    return matches[0]

def get_credentials():
    secret_file = _find_secret_file()
    flow = InstalledAppFlow.from_client_secrets_file(
        client_secrets_file=secret_file,
        scopes=SCOPES
    )
    return flow.run_local_server(port=0)

def get_bq_client(project_id, location):
    creds = get_credentials()
    return bigquery.Client(credentials=creds, project=project_id, location=location)