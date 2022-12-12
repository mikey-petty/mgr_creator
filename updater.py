import requests
from packaging import version
import win32api
import shutil
import subprocess

def check_and_run_updater(fname):
    
    # get the version details of latest release
    latest_version_info = get_latest_version_info_from_github()

    # read the current release
    current_version = get_file_properties(fname=fname)["FileVersion"]
    latest_version = latest_version_info["tag_name"]

    if version.parse(latest_version) > version.parse(current_version):
        get_latest_installer(latest_version_info)
        run_latest_installer()


def run_latest_installer():
    try: 
        subprocess.run(args=["installer.exe", "/SILENT"])
    except: 
        print("Error running installer")
        return

def get_latest_installer(version_info: dict):
    """ Downloads and saves the latest installer

    Args:
        version_info (dict): version info of the latest release
    """
    download_url = version_info["assets"][0]["browser_download_url"]

    header = {"accept": "application/octet-stream"}

    response = requests.get(url=download_url, headers=header)

    with requests.get(download_url, stream=True) as r:
        with open("installer.exe", 'wb') as f:
            shutil.copyfileobj(r.raw, f)

    del response

    return

def get_file_properties(fname: str) -> dict:
    """ Read all properties of the given file return them as a dictionary.

    Args:
        fname (str): filename of executable to read properties of

    Returns:
        dict: properties of executable file
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

def get_latest_version_info_from_github() -> dict:
    """ Send API request to get the latest installer for application

    Returns:
        str: Returns latest version of the application as a string
    """
    
    # get request for latest release
    headers = {"accept": "application/vnd.github+jso"}
    url = "https://api.github.com/repos/mpetty9991/mgr_creator/releases/latest"
    response = requests.get(url=url, headers=headers)
    
    # return None if response is invalid
    if response.status_code != 200: return None

    return response.json()


# Developer testing
if __name__ == "__main__":
    check_and_run_updater(r"C:\dev\mgr_creator\dist\gui\gui.exe")

else: 
    check_and_run_updater(r"gui.exe")