# encoding=utf8
import plistlib
from os.path import expanduser, dirname, basename, isfile, splitext
import urllib
import sys
import re
from xml.etree.ElementTree import Element, SubElement, Comment, tostring

reload(sys)  
sys.setdefaultencoding('utf8')

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
    plist = plistlib.readPlist(expanduser("~") + office[app]['bookmarks'] )
except Exception, e:
    root = Element('items')
    
    item  = SubElement(root, 'item');
    
    title         = SubElement(item, 'title')
    title.text    = "Uups! Coudn't fetch recent files."
    subtitle      = SubElement(item, 'subtitle')
    subtitle.text =  str(e) 
    print tostring(root)
    exit(0)

root = Element('items')

for key,value in plist.iteritems():

    file_path  = urllib.unquote(key.replace("file://", ""))
    file_name  = basename(file_path)
    clean_path = dirname(file_path).replace(expanduser("~"), "~")
    extension  = splitext(file_path)[1][1:]

    # check if file exists
    if not isfile(file_path):
        continue

    # do we have a search query and does it match?
    if arg and not re.search(arg, file_path, re.IGNORECASE):
        continue


    item  = SubElement(root, 'item', { 'arg': file_path });
    
    title         = SubElement(item, 'title')
    title.text    = file_name
    subtitle      = SubElement(item, 'subtitle')
    subtitle.text = clean_path
    icon          = SubElement(item, 'icon', {'type': 'fileicon'})
    icon.text     = file_path

print tostring(root)
