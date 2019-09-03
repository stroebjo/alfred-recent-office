# encoding=utf8
import plistlib
from os.path import expanduser, dirname, basename, isfile, splitext
import urllib
import urllib.parse
import sys
import re
import json

office = {
    'excel': {
        'bookmarks': "/Library/Containers/com.microsoft.Excel/Data/Library/Preferences/com.microsoft.Excel.securebookmarks.plist",
        'resources': "/Applications/Microsoft Excel.app/Contents/Resources/"
    },
    'word': {
        'bookmarks': "/Library/Containers/com.microsoft.Word/Data/Library/Preferences/com.microsoft.Word.securebookmarks.plist",
        'resources': "/Applications/Microsoft Word.app/Contents/Resources/"
    },
        'powerpoint': {
        'bookmarks': "/Library/Containers/com.microsoft.Powerpoint/Data/Library/Preferences/com.microsoft.Powerpoint.securebookmarks.plist",
        'resources': "/Applications/Microsoft PowerPoint.app/Contents/Resources/"
    }
}

# which office app are we using?
app = sys.argv[1]

# File name search
arg = sys.argv[2]


try:
    with open(expanduser("~") + office[app]['bookmarks'], 'rb') as f:
        plist = plistlib.load(f)
except Exception as e:
    data = {
        'items': [{
            'title': "Uups! Coudn't fetch recent files.",
            'subtitle': str(e)
        }]
    }

    print(json.dumps(data))
    exit(0)

data = {
    'items': []
}

for key,value in plist.items():
    file_path  = urllib.parse.unquote(key.replace("file://", ""))
    file_name  = basename(file_path)
    clean_path = dirname(file_path).replace(expanduser("~"), "~")
    extension  = splitext(file_path)[1][1:]

    # check if file exists
    if not isfile(file_path):
        continue

    # do we have a search query and does it match?
    if arg and not re.search(arg, file_path, re.IGNORECASE):
        continue

    item  = {
        'arg': file_path,
        'type': 'file',
        'title': file_name,
        'subtitle': clean_path,
        'icon': file_path
    }

    data['items'].append(item)

print(json.dumps(data))
