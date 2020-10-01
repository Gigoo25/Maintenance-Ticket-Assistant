import os
import os.path
import sys
import time
import wget

# Define function to go up directories
def dir_up(path,n): # here 'path' is your path, 'n' is number of dirs up you want to go
    for _ in range(n):
        path = dir_up(path.rpartition("\\")[0], 0) # second argument equal '0' ensures that 
                                                        # the function iterates proper number of times
    return(path)

# Set repo variables
REPO_URL = "https://raw.githubusercontent.com/Gigoo25/Maintenance-Ticket-Assistant/master/"

# Set tool verison check variables
CURRENT_VERSION = "unidentified"
ONLINE_VERSION = "unidentified"

# Text files local variables
Readme_Local = "unidentified"
Version_Local = "unidentified"
Batch_Script_Local = "unidentified"
Python_Script_Local = "unidentified"
Requirements_Local = "unidentified"
Update_Script_Local = "unidentified"

# Text files online variables
Readme_Online = "unidentified"
Version_Online = "unidentified"
Batch_Script_Online = "unidentified"
Python_Script_Online = "unidentified"
Requirements_Online = "unidentified"
Update_Script_Online = "unidentified"

# Set folder variables
curDir, _ = os.path.split(os.path.abspath(__file__))
Tools = dir_up(curDir,1)
Root = dir_up(curDir,2)

# Check for version file
if os.path.isfile(Tools+'\Version.txt'):
    # Open version file
    CURRENT_VERSION_FILE = open(Tools+'\Version.txt', "r")
    # Set current version
    lines=CURRENT_VERSION_FILE.readlines()
    CURRENT_VERSION=lines[0]
    # Close version file
    CURRENT_VERSION_FILE.close
else:
    print("Version file was not found.")
    time.sleep(10)
    sys.exit()

# Delete version check file if found
if os.path.isfile(Tools+'\Version_Check.txt') and os.access(Tools+'\Version_Check.txt', os.R_OK):
	os.remove(Tools+'\Version_Check.txt')

# Download version to compare from online
wget.download(REPO_URL+"Tools/Version.txt", out=Tools+'\Version_Check.txt')

# Check for version check file
if os.path.isfile(Tools+'\Version_Check.txt'):
    # Open version check file
    ONLINE_VERSION_FILE = open(Tools+'\Version_Check.txt', "r")
    # Set online version
    lines=ONLINE_VERSION_FILE.readlines()
    ONLINE_VERSION=lines[0]
    # Close online file
    ONLINE_VERSION.close
else:
    print("Version check file was not found.")
    time.sleep(10)
    sys.exit()

# Compare versions and set variables 
if ONLINE_VERSION > CURRENT_VERSION:
	print("Update was found!")
	print("Updating...")
    # Open local version file
    CURRENT_VERSION_FILE = open(Tools+'\Version.txt', "r")
    # Set local version variables
    lines=CURRENT_VERSION_FILE.readlines()
    Readme_Local=lines[12]
    Version_Local=lines[13]
    Batch_Script_Local=lines[14]
    Python_Script_Local=lines[15]
    Requirements_Local=lines[16]
    Update_Script_Local=lines[17]
    # Close local version file
    CURRENT_VERSION_FILE.close

    # Open online version check file
    ONLINE_VERSION_FILE = open(Tools+'\Version_Check.txt', "r")
    # Set online version variables
    lines=ONLINE_VERSION_FILE.readlines()
    Readme_Online=lines[12]
    Version_Online=lines[13]
    Batch_Script_Online=lines[14]
    Python_Script_Online=lines[15]
    Requirements_Online=lines[16]
    Update_Script_Local=lines[17]
    # Close online version file
    ONLINE_VERSION.close
else:
	print("No update was found.")
    time.sleep(10)
    sys.exit()

# Update Readme file
if Readme_Online > Readme_Local:
    wget.download(REPO_URL+"README.md", out=Root+'\README.md')

# Update Version file
if Readme_Online > Readme_Local:
    wget.download(REPO_URL+"Tools/Version.txt", out=Root+'\Tools\Version.txt')

# Update Batch Script
if Readme_Online > Readme_Local:
    wget.download(REPO_URL+"RunMe.bat", out=Root+'\RunMe.bat')

# Update Python Script
if Readme_Online > Readme_Local:
    wget.download(REPO_URL+"script.py", out=Root+'\script.py')

# Update Requirements
if Readme_Online > Readme_Local:
    wget.download(REPO_URL+"Tools/Requirements.txt", out=Root+'\Tools\Requirements.txt')

# Update Requirements
if Readme_Online > Readme_Local:
    wget.download(REPO_URL+"Tools/Functions/Update_Script.py", out=Root+'\Tools\Functions\Update_Script.py')