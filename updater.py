import requests
from packaging import version
import os
import win32api

def check_and_run_updater(fname):
    # current_version = getFileProperties(fname=fname)["FileVersion"]
    latest_version = get_latest_version_from_github()

    # print(current_version)
    print()
    print(latest_version)

    # if version.parse(latest_version) > version.parse(current_version):
    #     print("New Version")


def getFileProperties(fname):
    """
    Read all properties of the given file return them as a dictionary.
    """
    propNames = ('Comments', 'InternalName', 'ProductName',
        'CompanyName', 'LegalCopyright', 'ProductVersion',
        'FileDescription', 'LegalTrademarks', 'PrivateBuild',
        'FileVersion', 'OriginalFilename', 'SpecialBuild')

    props = {'FixedFileInfo': None, 'StringFileInfo': None, 'FileVersion': None}

    try:

        # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
        fixedInfo = win32api.GetFileVersionInfo(fname, '\\')
        props['FixedFileInfo'] = fixedInfo
        props['FileVersion'] = "%d.%d.%d.%d" % (fixedInfo['FileVersionMS'] / 65536,
                fixedInfo['FileVersionMS'] % 65536, fixedInfo['FileVersionLS'] / 65536,
                fixedInfo['FileVersionLS'] % 65536)

        # \VarFileInfo\Translation returns list of available (language, codepage)
        # pairs that can be used to retreive string info. We are using only the first pair.
        lang, codepage = win32api.GetFileVersionInfo(fname, '\\VarFileInfo\\Translation')[0]

        # any other must be of the form \StringfileInfo\%04X%04X\parm_name, middle
        # two are language/codepage pair returned from above

        strInfo = {}
        for propName in propNames:
            strInfoPath = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, propName)
            ## print str_info
            strInfo[propName] = win32api.GetFileVersionInfo(fname, strInfoPath)

        props['StringFileInfo'] = strInfo

        return props

    except:
        pass

def get_latest_version_from_github():

    url = "https://api.github.com/repos/mpetty9991/mgr_creator/releases/latest"   
    response = requests.get(url)
    response = response.json()
    latest_version = response
    return latest_version
    
def get_version_from_file():
    version_file = os.path.join('version.txt')

    if os.path.isfile(version_file):
        with open(version_file, 'r') as f:
            current_version = f.read().strip(' \n\r')
    else:
        current_version = None

    return current_version

check_and_run_updater(r"gui.exe")