import plistlib
from os.path import expanduser, dirname, basename, isfile, splitext
import urllib
import sys
import re

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
    print "<?xml version=\"1.0\"?>\n<items>\n"
    print '<item>'
    print "\t<title>" + "Uups! Coudn't fetch recent files." + "</title>"
    print "\t<subtitle>" + str(e) + "</subtitle>"
    print"</item>\n"
    print "</items>"
    exit(0)

print "<?xml version=\"1.0\"?>\n<items>\n"

for key,value in plist.iteritems():

    file_path  = urllib.unquote(key.replace("file://", ""))
    file_name  = basename(file_path)
    clean_path = dirname(file_path).replace(expanduser("~"), "~")
    extension  = splitext(file_path)[1][1:]

    # Office stores filetype icons in their App resources
    # in the format "<EXTENSION>.icns".
    icon_path  = office[app]['resources'] +  extension.upper() + ".icns"

    # Word hover doesn't save them in the name of the extion of the file.
    # So we allways use the same icon.
    if "word" == app:
        icon_path = '/Applications/Microsoft Word.app/Contents/Resources/WXBN.icns'


    # check if file exists
    if not isfile(file_path):
        continue

    # do we have a search query and does it match?
    if arg and not re.search(arg, file_path, re.IGNORECASE):
        continue

    # check if icns for filetype icon is present in default MS Office location
    if not isfile(icon_path):
        icon_path = "icon/unkown.png"


    print '<item arg="' + file_path + '">'
    print "\t<title>" + file_name + "</title>"
    print "\t<subtitle>" + clean_path +  "</subtitle>"
    print "\t<icon>" + icon_path + "</icon>";
    print"</item>\n"

print "</items>"


