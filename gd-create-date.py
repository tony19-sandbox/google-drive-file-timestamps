from __future__ import print_function
import sys, httplib2, os, datetime, io
from time import gmtime, strftime
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from datetime import date

#########################################################################
# Fixing OSX el capitan bug ->AttributeError: 'Module_six_moves_urllib_parse' object has no attribute 'urlencode'
os.environ["PYTHONPATH"] = "/Library/Python/2.7/site-packages"
#########################################################################

CLIENT_SECRET_FILE = 'client_secrets.json'
TOKEN_FILE="drive_api_token.json"
SCOPES = 'https://www.googleapis.com/auth/drive'
APPLICATION_NAME = 'Drive File API - Python'
OUTPUT_DIR=str(date.today())+"_drive_backup"

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, TOKEN_FILE)
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def prepDest():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        return True
    return False

def setFileCreationTime(fname, newtime):
    """http://stackoverflow.com/a/4996407/6277151"""
    if os.name != 'nt':
        # file creation time can only be changed in Windows
        return

    import pywintypes, win32file, win32con

    wintime = pywintypes.Time(newtime)
    winfile = win32file.CreateFile(
        fname, win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None, win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL, None)

    win32file.SetFileTime(winfile, wintime, None, None)

    winfile.close()

def setFileModificationTime(fname, newtime):
    # Set access time to same values as modified time,
    # since Google doesn't provide access time
    os.utime(fname, (newtime, newtime))


def dateToSeconds(dateTime):
    return int(datetime.datetime.strptime(dateTime, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%s"))

def setFileTimestamps(fname, createdTime, modifiedTime):
    ctime = dateToSeconds(createdTime)
    mtime = dateToSeconds(modifiedTime)
    setFileCreationTime(fname, ctime)
    setFileModificationTime(fname, mtime)


def downloadFile(file_name, file_id, file_createdTime, modifiedTime, mimeType, service):
    request = service.files().get_media(fileId=file_id)
    if "application/vnd.google-apps" in mimeType:
        if "document" in mimeType:
            request = service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            file_name = file_name + ".docx"
        else:
            request = service.files().export_media(fileId=file_id, mimeType='application/pdf')
            file_name = file_name + ".pdf"
    print("Downloading -- " + file_name)
    response = request.execute()
    prepDest()
    fname = os.path.join(OUTPUT_DIR, file_name)
    with open(fname, "wb") as wer:
        wer.write(response)

    setFileTimestamps(fname, file_createdTime, modifiedTime)


def listFiles(service):
    def getPage(pageTok):
        return service.files().list(q="mimeType != 'application/vnd.google-apps.folder'",
                                    pageSize=1000, pageToken=pageTok, fields="nextPageToken,files(id,name, createdTime, modifiedTime, mimeType)").execute()
    pT = ''; files=[]
    while pT is not None:
        results = getPage(pT)
        pT = results.get('nextPageToken')
        files = files + results.get('files', [])
    return files

def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    for item in listFiles(service):
        downloadFile(item.get('name'), item.get('id'), item.get('createdTime'), item.get('modifiedTime'), item.get('mimeType'), service)

if __name__ == '__main__':
    main()