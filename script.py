from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pynput.keyboard import Key, Controller
from pathlib import Path
from PIL import Image
import openpyxl
from openpyxl import Workbook

import csv
import time
import os
import sys

###########################################################

#get script directory
dir_path = os.path.dirname(os.path.abspath(__file__))
#set controller variable
keyboard = Controller()

###########################################################

def quit_chrome():
	time.sleep(3)
	browser.quit()
	time.sleep(1)
	browserExe = "Chrome"

def quit_chrome_exception():
	#### This exception occurs if the element are not found in the webpage.
	print("###########################################################")
	print("Error: An element was not found.")
	print("###########################################################")
	print("Exiting chrome in 10 seconds.")
	print("###########################################################")
	print('\n')
	#### Quit Browser
	time.sleep(10)
	quit_chrome()

###########################################################

# Delete old files if they exist
ModemReportCSV = dir_path+"\Maintenance Ticket Files\ModemReport.csv"
ModemReportCSV1 = dir_path+"\Maintenance Ticket Files\ModemReport (1).csv"
ModemReport_Red = dir_path+"\Maintenance Ticket Files\ModemReport_Red.xlsx"
ModemReport_RedYellow = dir_path+"\Maintenance Ticket Files\ModemReport_RedYellow.xlsx"
Full_Map_Element = dir_path+"\Maintenance Ticket Files\Full_Map_Element.png"
PEA_critical_modem_map = dir_path+"\Maintenance Ticket Files\PEA critical modem map.png"
PEA_critical_modems_XLSX = dir_path+"\Maintenance Ticket Files\PEA critical modems.xlsx"
PEA_vTDR_XLSX = dir_path+"\Maintenance Ticket Files\PEA vTDR.xlsx"

if os.path.isfile(ModemReportCSV) and os.access(ModemReportCSV, os.R_OK):
	os.remove(ModemReportCSV)

if os.path.isfile(ModemReportCSV1) and os.access(ModemReportCSV1, os.R_OK):
	os.remove(ModemReportCSV1)

if os.path.isfile(ModemReport_Red) and os.access(ModemReport_Red, os.R_OK):
	os.remove(ModemReport_Red)

if os.path.isfile(ModemReport_RedYellow) and os.access(ModemReport_RedYellow, os.R_OK):
	os.remove(ModemReport_RedYellow)

if os.path.isfile(Full_Map_Element) and os.access(Full_Map_Element, os.R_OK):
	os.remove(Full_Map_Element)

if os.path.isfile(PEA_critical_modem_map) and os.access(PEA_critical_modem_map, os.R_OK):
	os.remove(PEA_critical_modem_map)

if os.path.isfile(PEA_critical_modems_XLSX) and os.access(PEA_critical_modems_XLSX, os.R_OK):
	os.remove(PEA_critical_modems_XLSX)

if os.path.isfile(PEA_vTDR_XLSX) and os.access(PEA_vTDR_XLSX, os.R_OK):
	os.remove(PEA_vTDR_XLSX)

Full_Spectrum_MAX_Drop = dir_path+"\Maintenance Ticket Files\Full_Spectrum_Max_Element.png"
Full_Waterfall_MAX_Drop = dir_path+"\Maintenance Ticket Files\Full_Waterfall_Max_Element.png"

Full_Spectrum_AVG_Drop = dir_path+"\Maintenance Ticket Files\Full_Spectrum_Average_Element.png"
Full_Waterfall_AVG_Drop = dir_path+"\Maintenance Ticket Files\Full_Waterfall_Avg_Element.png"

Spectrum_MAX_Drop = dir_path+"\Maintenance Ticket Files\Spectrum_Max.png"
Waterfall_MAX_Drop = dir_path+"\Maintenance Ticket Files\Waterfall_Max.png"

Spectrum_AVG_Drop = dir_path+"\Maintenance Ticket Files\Spectrum_Average.png"
Waterfall_AVG_Drop = dir_path+"\Maintenance Ticket Files\Waterfall_Average.png"

if os.path.isfile(Full_Spectrum_MAX_Drop) and os.access(Full_Spectrum_MAX_Drop, os.R_OK):
	os.remove(Full_Spectrum_MAX_Drop)

if os.path.isfile(Full_Waterfall_MAX_Drop) and os.access(Full_Waterfall_MAX_Drop, os.R_OK):
	os.remove(Full_Waterfall_MAX_Drop)

if os.path.isfile(Full_Spectrum_AVG_Drop) and os.access(Full_Spectrum_AVG_Drop, os.R_OK):
	os.remove(Full_Spectrum_AVG_Drop)

if os.path.isfile(Full_Waterfall_AVG_Drop) and os.access(Full_Waterfall_AVG_Drop, os.R_OK):
	os.remove(Full_Waterfall_AVG_Drop)

if os.path.isfile(Spectrum_MAX_Drop) and os.access(Spectrum_MAX_Drop, os.R_OK):
	os.remove(Spectrum_MAX_Drop)

if os.path.isfile(Waterfall_MAX_Drop) and os.access(Waterfall_MAX_Drop, os.R_OK):
	os.remove(Waterfall_MAX_Drop)

if os.path.isfile(Spectrum_AVG_Drop) and os.access(Spectrum_AVG_Drop, os.R_OK):
	os.remove(Spectrum_AVG_Drop)

if os.path.isfile(Waterfall_AVG_Drop) and os.access(Waterfall_AVG_Drop, os.R_OK):
	os.remove(Waterfall_AVG_Drop)

###########################################################

# Check for login file & create a template if not present
print("###########################################################")
print("Checking for the credentials file")
print("###########################################################")
PATH = './login.txt'
if os.path.isfile(PATH) and os.access(PATH, os.R_OK):
    print("The credentials file exists and is readable.")
    print('\n')
else:
    print("The credentials file is missing or not readable.")
    print('\n')
    time.sleep(3)
    print("Generating new credentials file.")
    print('\n')
    credentials= open("login.txt","w+")
    credentials.write("#enter your PEA_username in the line below." + '\n')
    credentials.write('\n')
    credentials.write("#enter your PEA_password in the line below." + '\n')
    credentials.write('\n')
    credentials.write("#enter your Viewpoint_username in the line below." + '\n')
    credentials.write('\n')
    credentials.write("#enter your Viewpoint_password in the line below." + '\n')
    credentials.write('\n')
    credentials.close
    print("Please enter your credentials in the file named 'login.txt'.")
    print('\n')
    input("Press Enter when done...")

###########################################################

# Open passwords file
passwords=open(dir_path+"/login.txt", "r")
# Set password vars
lines=passwords.readlines()
PEA_username=lines[1]
PEA_password=lines[3]
Viewpoint_username=lines[5]
Viewpoint_password=lines[7]
# Close passwords file
passwords.close

###########################################################

print('\n')
print("###########################################################")
print("Typing small tutorial.")
print("###########################################################")
print("At anytime press 'CTRL + C' to exit the script.")
print("Do not click out of the open Chrome window.")
print("Be patient.")
print('\n')
print("###########################################################")
print("Starting script.")
print("###########################################################")
time.sleep(1)

###########################################################

# Create input to ask for HUB/Node
print("###########################################################")
Short_HUB_Input = input("Type the HUB & Node you would like to create a maintenance ticket for\nThis should be entered as a 2/3 letter HUB -Space- Node.\nFor example \"MAU 1\"\n###########################################################\n") 

###########################################################

# Convert input to full HUB name for Viewpoint
Input_Starts_With_ALX = Short_HUB_Input.startswith("ALX")
Input_Starts_With_ARL = Short_HUB_Input.startswith("ARL")
Input_Starts_With_BED = Short_HUB_Input.startswith("BED")
Input_Starts_With_ET = Short_HUB_Input.startswith("ET")
Input_Starts_With_ERI = Short_HUB_Input.startswith("ERI")
Input_Starts_With_ANG = Short_HUB_Input.startswith("ANG")
Input_Starts_With_MAU = Short_HUB_Input.startswith("MAU")
Input_Starts_With_NE = Short_HUB_Input.startswith("NE")
Input_Starts_With_ORG = Short_HUB_Input.startswith("ORG")
Input_Starts_With_OWN = Short_HUB_Input.startswith("OWN")
Input_Starts_With_PER = Short_HUB_Input.startswith("PER")
Input_Starts_With_SPR = Short_HUB_Input.startswith("SPR")
Input_Starts_With_SYL = Short_HUB_Input.startswith("SYL")
Input_Starts_With_UT = Short_HUB_Input.startswith("UT")
Input_Starts_With_WAT = Short_HUB_Input.startswith("WAT")

# Set variable with full name
if Input_Starts_With_ALX == True:
	Viewpoint_Full_HUB = "Alexis"
elif Input_Starts_With_ARL == True:
	Viewpoint_Full_HUB = "Arlington"
elif Input_Starts_With_BED == True:
	Viewpoint_Full_HUB = "Bedford"
elif Input_Starts_With_ET == True:
	Viewpoint_Full_HUB = "East Toledo"
elif Input_Starts_With_ERI == True:
	Viewpoint_Full_HUB = "Erie"
elif Input_Starts_With_ANG == True:
	Viewpoint_Full_HUB = "HE (Angola)"
elif Input_Starts_With_MAU == True:
	Viewpoint_Full_HUB = "Maumee"
elif Input_Starts_With_NE == True:
	Viewpoint_Full_HUB = "NorthEast"
elif Input_Starts_With_ORG == True:
	Viewpoint_Full_HUB = "Oregon"
elif Input_Starts_With_OWN == True:
	Viewpoint_Full_HUB = "Owens"
elif Input_Starts_With_PER == True:
	Viewpoint_Full_HUB = "Perrysburg"
elif Input_Starts_With_SPR == True:
	Viewpoint_Full_HUB = "Springfield"
elif Input_Starts_With_SYL == True:
	Viewpoint_Full_HUB = "Sylvania"
elif Input_Starts_With_UT == True:
	Viewpoint_Full_HUB = "UT"
elif Input_Starts_With_WAT == True:
	Viewpoint_Full_HUB = "Waterville"

# Split Short HUB & Node
HUB_Node_Split = Short_HUB_Input.split()
# Convert to individual lists
Short_HUB_Split_List = [HUB_Node_Split[0]]
Short_Node_Split_List = [HUB_Node_Split[1]]
# Convert to String
out_str = ""
Short_HUB_Split_String = out_str.join(Short_HUB_Split_List)
Short_Node_Split_String = out_str.join(Short_Node_Split_List)

# Add missing 0's to the Node for Viewpoint
if len(Short_Node_Split_String) == 3:
    Corrected_Short_Node_Split = "0"+Short_Node_Split_String
elif len(Short_Node_Split_String) == 1:
    Corrected_Short_Node_Split = "0"+Short_Node_Split_String
else:
	Corrected_Short_Node_Split = Short_Node_Split_String

# Merge corrected short HUB & Node for Viewpoint
Viewpoint_Short_HUB_Node = Short_HUB_Split_String+" "+Corrected_Short_Node_Split

# Add "Node " infront of shortened HUB for PEA.
PEA_Short_HUB_Node = "Node "+Short_HUB_Input

print('Printing Vars for testing')
print(Viewpoint_Full_HUB)
print(Viewpoint_Short_HUB_Node)

###########################################################

# Convert HUB/Node to MD
if Viewpoint_Short_HUB_Node == "ALX 01P1":
	Mac_Domain = "7:0/0"
elif Viewpoint_Short_HUB_Node == " ALX 01P4":
	Mac_Domain = "7:0/1"

print("Testing MacD conversion")
print(Mac_Domain)
time.sleep(2500)

###########################################################

# Open chrome
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : dir_path+"\Maintenance Ticket Files\\"}
chromeOptions.add_experimental_option("prefs",prefs)
#chromeOptions.add_argument("--window-size=1600,600")
chromedriver = dir_path+"/chromedriver.exe"
browser = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeOptions)
# Open PEA in Chrome.
browser.get(("https://buckeye.pea.zcorum.com/pnm4/index.php/detailView"))

# Grab screenshot of the PEA map for the node & export the bad visible modems.
try:
	# Log into PEA
	print('\n')
	print("###########################################################")
	print("Logging into PEA.")
	print("###########################################################")
	
	# Type username
	PEA_username_element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div[5]/form/div[1]/div/input")))
	print("Typing username.")
	PEA_username_element.send_keys(PEA_username)
	
	# Type password
	PEA_username_element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div[5]/form/div[2]/div/input")))
	print("Typing password.")
	PEA_username_element.send_keys(PEA_password)
	
	# Enter HUB in the "HUB" field
	# Add extra time for page to load
	time.sleep(3)
	HUB_Field = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[1]/div/div/div[2]/div[2]/div/input[2]")))
	print("Entering HUB.")
	HUB_Field.send_keys(PEA_Short_HUB_Node)
	
	# Click the suggested Node
	# Add extra time for page to load
	time.sleep(1)
	suggestButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[1]/div/div/div[2]/div[2]/div/select/option[1]")))
	print("Clicking Suggest button.")
	suggestButton.click()
	
	# Click the search button
	# Add extra time for page to load
	time.sleep(1)
	searchButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[1]/div/div/div[2]/div[2]/button")))
	print("Clicking Search button.")
	searchButton.click()
	
	# Click the upstream search button
	upstreamsearchButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[7]/div[2]/div/div[3]/button[2]")))
	print("Clicking Upstream Search button.")
	upstreamsearchButton.click()
	
	# Click the red button
	redButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[1]/div/div/div[2]/div[4]/span/input[1]")))
	print("Clicking Red button.")
	redButton.click()

	# Take screenshot of the map
	print("Take screenshot.")
	# Add extra time for canvas to load
	time.sleep(3)
	element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/div/div[1]/div[2]/div/div/div/div[1]/div[1]/canvas")))
	location = element.location
	size = element.size
	browser.save_screenshot(dir_path+"/Maintenance Ticket Files/Full_Map_Element.png")
	# Crop screenshot of the map
	print("Crop screenshot.")
	x = location['x']
	y = location['y']
	width = location['x']+size['width']
	height = location['y']+size['height']
	im = Image.open(dir_path+"/Maintenance Ticket Files/Full_Map_Element.png")
	im = im.crop((int(x), int(y), int(width), int(height)))
	im.save(dir_path+"/Maintenance Ticket Files/PEA critical modem map.png")
	
	# Click on data tab
	dataTab = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/ul/li[2]/a")))
	print("Click on data tab.")
	dataTab.click()
	
	# Click on sorting
	VTDR_Element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "//th[contains(text(),'Outside Plant')]")))
	print("Sorting by VTDR.")
	VTDR_Element.click()
	time.sleep(2)
	VTDR_Element.click()
	
	# Wait for page to load
	time.sleep(4)
	
	# Click on Export
	Export_Element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/div/div[2]/div/div/div[1]/div[6]/div/button")))
	print("Click Export.")
	Export_Element.click()

	# Click on Only Visible Data
	Only_Visible_Data_Element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/div/div[2]/div/div/div[1]/div[6]/div/ul/li[2]/a")))
	print("Click Only Visible Data.")
	Only_Visible_Data_Element.click()
	
	# Click the yellow button
	yellowButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[1]/div/div/div[2]/div[4]/span/input[2]")))
	print("Clicking Yellow button.")
	yellowButton.click()
	
	# Click on data tab
	dataTab = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/ul/li[2]/a")))
	print("Click on data tab.")
	dataTab.click()
	
	# Click on sorting
	VTDR_Element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "//th[contains(text(),'Outside Plant')]")))
	print("Sorting by VTDR.")
	VTDR_Element.click()
	time.sleep(2)
	VTDR_Element.click()
	
	# Wait for page to load
	time.sleep(4)
	
	# Click on Export
	Export_Element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/div/div[2]/div/div/div[1]/div[6]/div/button")))
	print("Click Export.")
	Export_Element.click()

	# Click on Only Visible Data
	Only_Visible_Data_Element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[6]/div[3]/div/div[3]/div/div[2]/div/div/div[1]/div[6]/div/ul/li[2]/a")))
	print("Click Only Visible Data.")
	Only_Visible_Data_Element.click()
	
	#### Quit Browser
	time.sleep(2)
	print("Exiting PEA.")
	quit_chrome()
except Exception:
	quit_chrome_exception()
	sys.exit()

# Convert CSV to XLSX
print("Converting file to XLSX.")
wb = Workbook()
ws = wb.active
with open(dir_path+"\Maintenance Ticket Files\ModemReport.csv", 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
wb.save(dir_path+"\Maintenance Ticket Files\ModemReport_Red.xlsx")

with open(dir_path+"\Maintenance Ticket Files\ModemReport (1).csv", 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
wb.save(dir_path+"\Maintenance Ticket Files\ModemReport_RedYellow.xlsx")

# Remove unecessary columns for PEA vTDR
print("Deleting unecessary columns for PEA vTDR.")
wb = openpyxl.load_workbook(dir_path+"\Maintenance Ticket Files\ModemReport_RedYellow.xlsx")
ws = wb.active
# Delete B,C,D
ws.delete_cols(2,3)
# Delete O-U
ws.delete_cols(12,7)
# Delete W-AG
ws.delete_cols(13,11)
# Save XLSX file
wb.save(dir_path+"\Maintenance Ticket Files\PEA vTDR.xlsx")
print("Saving file.")

# Remove unecessary columns for critical modems
print("Deleting unecessary columns for critical modems.")
wb = openpyxl.load_workbook(dir_path+"\Maintenance Ticket Files\ModemReport_Red.xlsx")
ws = wb.active
# Delete B-U
ws.delete_cols(2,20)
# Delete C-M
ws.delete_cols(3,11)
# Save XLSX file
wb.save(dir_path+"\Maintenance Ticket Files\PEA critical modems.xlsx")
print("Saving file.")

# Delete extra files
print("Deleting extra files.")
os.remove(dir_path+"\Maintenance Ticket Files\ModemReport.csv")
os.remove(dir_path+"\Maintenance Ticket Files\ModemReport (1).csv")
os.remove(dir_path+"\Maintenance Ticket Files\ModemReport_Red.xlsx")
os.remove(dir_path+"\Maintenance Ticket Files\ModemReport_RedYellow.xlsx")
os.remove(dir_path+"\Maintenance Ticket Files\Full_Map_Element.png")

###########################################################

# Open chrome
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : dir_path+"\Maintenance Ticket Files\\"}
chromeOptions.add_experimental_option("prefs",prefs)
#chromeOptions.add_argument("--window-size=1600,600")
chromedriver = dir_path+"/chromedriver.exe"
browser = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeOptions)
# Open Viewpoint in Chrome.
browser.get(("http://10.6.10.12/ViewPoint/site/Site/Login"))

# 
try:
	# Log into Viewpoint
	print('\n')
	print("###########################################################")
	print("Logging into Viewpoint.")
	print("###########################################################")

	# Type username
	Viewpoint_username_element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[3]/input[1]")))
	print("Typing username.")
	Viewpoint_username_element.send_keys(Viewpoint_username)

	# Type password
	Viewpoint_password_element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[3]/input[2]")))
	print("Typing password.")
	Viewpoint_password_element.send_keys(Viewpoint_password)

	# Click "login"
	loginButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[3]/div[1]/a")))
	print("Clicking login.")
	loginButton.click()

	# Click "RPM"
	RPMButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div/div/div/ul/ul/li/a[1]")))
	print("Clicking RPM.")
	RPMButton.click()

	# Click "Buckeye Cable"
	BuckeyeCableButton = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div/div/div/ul/ul/ul/li/a[1]")))
	print("Clicking Buckeye Cable.")
	BuckeyeCableButton.click()

	# Click "Alexis"
	Full_HUB = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Full_HUB)))
	print("Clicking on "+Viewpoint_Full_HUB+".")
	Full_HUB.click()

	# Click "Node"
	Shortened_Node = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Short_HUB_Node)))
	print("Clicking on "+Viewpoint_Short_HUB_Node+".")
	Shortened_Node.click()

	# Click "Return Spectrum"
	Return_Spectrum_Button = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[4]/div[1]/div/div/div/div/ul/li")))
	print("Clicking on Return Spectrum.")
	Return_Spectrum_Button.click()
	
	# Click "Mode"
	Mode_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[2]/li[2]/select")))
	print("Clicking on Mode.")
	Mode_Dropdown.click()

	# Click "Historical"
	Historical_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[2]/li[2]/select/option[2]")))
	print("Clicking on Historical.")
	Historical_Dropdown.click()

	# Click "Display"
	Display_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[1]/li[2]/select")))
	print("Clicking on Display.")
	Display_Dropdown.click()

	# Click "Spectrum"
	Spectrum_Dropdown_Option = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[1]/li[2]/select/option[1]")))
	print("Clicking on Spectrum.")
	Spectrum_Dropdown_Option.click()

	# Click "Back button for 15 minute increments"
	Back_15_Minute_Button = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[4]/div[1]/div/div/div/div[2]/ul/li[2]")))
	print("Clicking on back button.")
	Back_15_Minute_Button.click()

	# Take screenshot of the average spectrum
	print("Taking screenshots.")
	# Add extra time for canvas to load
	time.sleep(3)
	element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[4]/div[1]/div/div/div/div[4]/div/div/div/div[1]")))
	location = element.location
	size = element.size
	browser.save_screenshot(Full_Spectrum_AVG_Drop)
	# Crop screenshot of the average spectrum
	print("Crop screenshot.")
	x = location['x']
	y = location['y']
	width = location['x']+size['width']
	height = location['y']+size['height']
	im = Image.open(Full_Spectrum_AVG_Drop)
	im = im.crop((int(x), int(y), int(width), int(height)))
	im.save(Spectrum_AVG_Drop)

	# Click "Active Trace"
	Active_Trace_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[2]/ul/li[2]/select")))
	print("Clicking on Active Trace.")
	Active_Trace_Dropdown.click()

	# Click "MAX"
	Max_Dropdown_Option = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[2]/ul/li[2]/select/option[1]")))
	print("Clicking on MAX.")
	Max_Dropdown_Option.click()

	# Take screenshot of the max spectrum
	print("Taking screenshots.")
	# Add extra time for canvas to load
	time.sleep(3)
	element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[4]/div[1]/div/div/div/div[4]/div/div/div/div[1]")))
	location = element.location
	size = element.size
	browser.save_screenshot(Full_Spectrum_MAX_Drop)
	# Crop screenshot of the max spectrum
	print("Crop screenshot.")
	x = location['x']
	y = location['y']
	width = location['x']+size['width']
	height = location['y']+size['height']
	im = Image.open(Full_Spectrum_MAX_Drop)
	im = im.crop((int(x), int(y), int(width), int(height)))
	im.save(Spectrum_MAX_Drop)

	# Click "Display"
	Display_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[1]/li[2]/select")))
	print("Clicking on Display.")
	Display_Dropdown.click()

	# Click "Waterfall"
	Waterfall_Dropdown_Option = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[1]/li[2]/select/option[2]")))
	print("Clicking on Waterfall.")
	Waterfall_Dropdown_Option.click()

	# Click "Time Span"
	Time_Span_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[1]/li[4]/select")))
	print("Clicking on Time Span.")
	Time_Span_Dropdown.click()

	# Click "24 Hr"
	Twenty_four_hours_dropdown_option = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[1]/li[4]/select/option[5]")))
	print("Clicking on 24 HR.")
	Twenty_four_hours_dropdown_option.click()

	# Take screenshot of the 24 hr MAX waterfall
	print("Taking screenshots.")
	# Add extra time for canvas to load
	time.sleep(3)
	element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[4]/div[1]/div/div/div/div[4]/canvas[1]")))
	location = element.location
	size = element.size
	browser.save_screenshot(Full_Waterfall_MAX_Drop)
	# Crop screenshot of the 24 hr MAX waterfall
	print("Crop screenshot.")
	x = location['x']
	y = location['y']
	width = location['x']+size['width']
	height = location['y']+size['height']
	im = Image.open(Full_Waterfall_MAX_Drop)
	im = im.crop((int(x), int(y), int(width), int(height)))
	im.save(Waterfall_MAX_Drop)

	# Click "Active Trace"
	Active_Trace_Dropdown = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[2]/ul/li[2]/select")))
	print("Clicking on Active Trace.")
	Active_Trace_Dropdown.click()

	# Click "AVG"
	Average_Dropdown_Option = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/ul[3]/ul[2]/ul/li[2]/select/option[2]")))
	print("Clicking on Average.")
	Average_Dropdown_Option.click()
	
	# Take screenshot of the 24 hr AVG waterfall
	print("Taking screenshots.")
	# Add extra time for canvas to load
	time.sleep(3)
	element = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[4]/div[1]/div/div/div/div[4]/canvas[1]")))
	location = element.location
	size = element.size
	browser.save_screenshot(Full_Waterfall_AVG_Drop)
	# Crop screenshot of the 24 hr AVG waterfall
	print("Crop screenshot.")
	x = location['x']
	y = location['y']
	width = location['x']+size['width']
	height = location['y']+size['height']
	im = Image.open(Full_Waterfall_AVG_Drop)
	im = im.crop((int(x), int(y), int(width), int(height)))
	im.save(Waterfall_AVG_Drop)

	# Delete extra files
	print("Deleting extra files.")
	os.remove(Full_Spectrum_MAX_Drop)
	os.remove(Full_Spectrum_AVG_Drop)
	os.remove(Full_Waterfall_MAX_Drop)
	os.remove(Full_Waterfall_AVG_Drop)

	#### Quit Browser
	print("Exiting Viewpoint.")
	quit_chrome()
except Exception:
	quit_chrome_exception()
	sys.exit()