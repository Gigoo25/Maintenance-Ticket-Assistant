import PySimpleGUI as sg
import os
import time
import sys
import tkinter as tk
import wget
from shutil import copyfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import xlwings as xw

## Set Variables
# Set theme for the windows.
sg.theme('DarkGrey4')
# Set global directories.
Root = os.path.dirname(os.path.abspath(__file__))
Backend = Root+"/Backend"

## General functions
# Define function to quit Chrome
def quit_chrome():
	browser.quit()

# Define error window
def error_window():
    # Themes & Colors
    sg.theme('DarkGrey4')
    DARK_COLOR = '#1e1e1e'
    LIGHT_COLOR = '#333333'

    # Images
    error_img = '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'

    # Header block
    top_banner = [[sg.Image(data=error_img, background_color=DARK_COLOR), sg.Text('Error - Unknown', font=('Calibri', 20), background_color=DARK_COLOR)]]

    # Main block
    body = [
                [sg.Text('Something went wrong during execution.', font=('Calibri', 13), background_color=LIGHT_COLOR)],
                [sg.Text('The following exception occured:', font=('Calibri', 13), background_color=LIGHT_COLOR)],
                [sg.Multiline(size=(73, 5), key='textbox', background_color=LIGHT_COLOR, text_color='white', autoscroll=False, border_width=0)],
                [sg.Text('Please press \"Exit\" to close out of MaintenaceBoi.', font=('Calibri', 13), background_color=LIGHT_COLOR)]
              ]

    # Define Layout
    layout = [
                [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                [sg.Column([[sg.Column(body, size=(550,150), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                [sg.Button('Exit', size=(32,2), pad=(150,15), button_color=('white',LIGHT_COLOR), font=('Calibri', 12))]
             ]

    # Create window
    window = sg.Window('MaintenanceBoi - Error', layout, background_color=DARK_COLOR, keep_on_top=True, finalize=True)
    window['textbox'].print(str_exception, text_color='white')
    # Event Loop
    while True:             
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
    window.close()
    sys.exit()

## Prep functions
def del_old_files():
    # Delete old files if they exist
    ModemReportCSV = Root+"\\Maintenance Ticket Files\\ModemReport.csv"
    ModemReportCSV1 = Root+"\\Maintenance Ticket Files\\ModemReport (1).csv"
    ModemReport_Red = Root+"\\Maintenance Ticket Files\\ModemReport_Red.xlsx"
    ModemReport_RedYellow = Root+"\\Maintenance Ticket Files\\ModemReport_RedYellow.xlsx"
    Full_Map_Element = Root+"\\Maintenance Ticket Files\\Full_Map_Element.png"
    PEA_critical_modem_map = Root+"\\Maintenance Ticket Files\\PEA critical modem map.png"
    PEA_critical_modems_XLSX = Root+"\\Maintenance Ticket Files\\PEA critical modems.xlsx"
    PEA_vTDR_XLSX = Root+"\\Maintenance Ticket Files\\PEA vTDR.xlsx"
    Full_Spectrum_MAX_Drop = Root+"\\Maintenance Ticket Files\\Full_Spectrum_Max_Element.png"
    Full_Waterfall_MAX_Drop = Root+"\\Maintenance Ticket Files\\Full_Waterfall_Max_Element.png"
    Full_Spectrum_AVG_Drop = Root+"\\Maintenance Ticket Files\\Full_Spectrum_Average_Element.png"
    Full_Waterfall_AVG_Drop = Root+"\\Maintenance Ticket Files\\Full_Waterfall_Avg_Element.png"
    Spectrum_MAX_Drop = Root+"\\Maintenance Ticket Files\\Spectrum_Max.png"
    Waterfall_MAX_Drop = Root+"\\Maintenance Ticket Files\\Waterfall_Max.png"
    Spectrum_AVG_Drop = Root+"\\Maintenance Ticket Files\\Spectrum_Average.png"
    Waterfall_AVG_Drop = Root+"\\Maintenance Ticket Files\\Waterfall_Average.png"

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

def create_missing_dirs():
    # Make output directory if not found
    if not os.path.exists(Root+"\\Maintenance Ticket Files"):
        os.makedirs(Root+"\\Maintenance Ticket Files")
    # Make chrome driver directory if not found
    if not os.path.exists(Root+"\\Backend\\Chrome_Driver"):
        os.makedirs(Root+"\\Backend\\Chrome_Driver")
    # Make functions directory if not found
    if not os.path.exists(Root+"\\Backend\\Functions"):
        os.makedirs(Root+"\\Backend\\Functions")

def check_credentials():
    # Define the path for the credentials file
    PATH = Root+"\\Backend\\Credentials.txt"

    # Define lines for the credentials file
    line1 = "# Enter your PEA_username in the line below."
    line3 = "# Enter your PEA_password in the line below."
    line5 = "# Enter your Viewpoint_username in the line below."
    line7= "# Enter your Viewpoint_password in the line below"
    line9 = "# Enter your Grafana_username in the line below."
    line11 = "# Enter your Grafana_password in the line below."

    if os.path.isfile(PATH) and os.access(PATH, os.R_OK):
        print("The credentials file exists and is readable.")
    else:
        # Themes & Colors
        sg.theme('DarkGrey4')
        DARK_COLOR = '#1e1e1e'
        LIGHT_COLOR = '#333333'

        # Images
        error_img = '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'

        # Set enter key variables to allow enter key for windows.
        QT_ENTER_KEY1 = 'special 16777220'
        QT_ENTER_KEY2 = 'special 16777221'
        
        # Header block
        top_banner = [[sg.Image(data=error_img, background_color=DARK_COLOR), sg.Text('Error - Credentials Missing', font=('Calibri', 20), background_color=DARK_COLOR)]]

        # Main block
        body = [
                    [sg.Text('The credentials file is missing or not readable.', font=('Calibri', 13), background_color=LIGHT_COLOR, pad=((0,0),(5,0)))],
                    [sg.Text('Please enter your credentials below to generate a new file.', font=('Calibri', 13), background_color=LIGHT_COLOR, pad=((0,0),(0,15)))],
                    [sg.Text("Pre-Equalization Analyzer", font=("Calibri", 20), background_color=LIGHT_COLOR, pad=((0,0),(5,2)))],
                    [sg.Text("Username:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((4,0),(0,0))), sg.Text("Password:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((54,0),(0,0)))],
                    [sg.InputText('', key='pea_usr', enable_events=True, font=("Calibri", 12), size=(13, 1)), sg.InputText('', key='pea_pass', enable_events=True, font=("Calibri", 12), size=(13, 1))],
                    [sg.Text("Viewpoint", font=("Calibri", 20), background_color=LIGHT_COLOR, pad=((0,0),(5,2)))],
                    [sg.Text("Username:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((4,0),(0,0))), sg.Text("Password:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((54,0),(0,0)))],
                    [sg.InputText('', key='viw_usr', enable_events=True, font=("Calibri", 12), size=(13, 1)), sg.InputText('', key='viw_pass', enable_events=True, font=("Calibri", 12), size=(13, 1))],
                    [sg.Text("Grafana", font=("Calibri", 20), background_color=LIGHT_COLOR, pad=((0,0),(5,2)))],
                    [sg.Text("Username:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((4,0),(0,0))), sg.Text("Password:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((54,0),(0,0)))],
                    [sg.InputText('', key='grf_usr', enable_events=True, font=("Calibri", 12), size=(13, 1)), sg.InputText('', key='grf_pass', enable_events=True, font=("Calibri", 12), size=(13, 1))]
                ]

        # Define Layout
        layout = [
                    [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                    [sg.Column([[sg.Column(body, size=(550,380), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                    [sg.Button('Ok', size=(32,2), pad=((10,22),(10,10)), button_color=('white',LIGHT_COLOR), font=('Calibri', 12)), sg.Button('Cancel', size=(32,2), pad=((0,0),(10,10)), button_color=('white',LIGHT_COLOR), font=('Calibri', 12))]
                ]

        # Create window
        window = sg.Window('MaintenanceBoi - Error - Credentials Missing', layout, return_keyboard_events=True, background_color=DARK_COLOR, keep_on_top=True)

        # Event Loop to process "events" and get the "values" of the inputs
        while True:
            event, values = window.read()
            # When "Cancel" button is pressed
            if event == sg.WIN_CLOSED or event == 'Cancel':
                window.close()
                sys.exit()
            # Check for ENTER key & input
            if event in ('\r', QT_ENTER_KEY1, QT_ENTER_KEY2):
                # Find the ok button & click it
                elem = window.find_element("Ok")
                elem.Click()
            # When "Ok" button is pressed
            if event == 'Ok':
                # Gather variables from user input
                line2 = values['pea_usr']
                line4 = values['pea_pass']
                line6 = values['viw_usr']
                line8 = values['viw_pass']
                line10 = values['grf_usr']
                line12 = values['grf_pass']
                # Write to credentials file
                credentials = open(PATH, "w+")
                credentials.write(line1 + "\n" + line2 + "\n" + line3 + "\n" + line4 + "\n" + line5 + "\n" + line6 + "\n" + line7 + "\n" + line8 + "\n" + line9 + "\n" + line10 + "\n" + line11 + "\n" + line12)
                credentials.close()
                # Continue
                break

        window.close()

def obtain_passwords():
    # Open passwords file
    passwords=open(Backend+"/Credentials.txt", "r")

    # Set password vars
    lines=passwords.readlines()

    global PEA_username
    global PEA_password
    global Viewpoint_username
    global Viewpoint_password
    global Grafana_username
    global Grafana_password

    PEA_username=lines[1]
    PEA_password=lines[3]
    Viewpoint_username=lines[5]
    Viewpoint_password=lines[7]
    Grafana_username=lines[9]
    Grafana_password=lines[11]

    # Close passwords file
    passwords.close

## Update Functions

## Analyze Functions
def main_window():
    try:
        # Themes & Colors
        sg.theme('DarkGrey4')
        DARK_COLOR = '#1e1e1e'
        LIGHT_COLOR = '#333333'
        DARK_COLOR = '#1e1e1e'
        LIGHT_COLOR = '#333333'

        # Images
        error_img = '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'
        
        # Set enter key variables to allow enter key for windows.
        QT_ENTER_KEY1 = 'special 16777220'
        QT_ENTER_KEY2 = 'special 16777221'

        # Header block
        top_banner = [[sg.Image(data=error_img, background_color=DARK_COLOR), sg.Text('Welcome to MaintenanceBoi', font=('Calibri', 20), background_color=DARK_COLOR)]]

        # Main block
        body = [
                    [sg.Text('Enter a few basic details for the drop that you would like to investigate.', font=('Calibri', 13), background_color=LIGHT_COLOR, pad=((5,0),(0,15)))],
                    [sg.Text("Mac Domain:", font=("Calibri", 12), background_color=LIGHT_COLOR, pad=((5,0),(0,0))), sg.InputText('', enable_events=True,  key='-INPUT-', tooltip='For example \"7:0/0\"', font=("Calibri", 12), size=(50,1))],
                    [sg.Text("Minutes since drop:", font=("Calibri", 12), background_color=LIGHT_COLOR, pad=((5,0),(17,0))), sg.Slider(range=(15, 120),  key='-SLIDE-', font=("Calibri", 12), orientation='h', size=(40, 20), default_value=15, background_color=LIGHT_COLOR)]
                ]

        # Define Layout
        layout = [
                    [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                    [sg.Column([[sg.Column(body, size=(550,150), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                    [sg.Button('Ok', size=(32,2), pad=((10,22),(10,10)), button_color=('white',LIGHT_COLOR), font=('Calibri', 12)), sg.Button('Cancel', size=(32,2), pad=((0,0),(10,10)), button_color=('white',LIGHT_COLOR), font=('Calibri', 12))]
                ]

        # Create window
        window = sg.Window('MaintenanceBoi', layout, return_keyboard_events=True, background_color=DARK_COLOR, keep_on_top=True)

        # Event Loop to process "events" and get the "values" of the inputs
        while True:
            event, values = window.read()
            # When "Cancel" button is pressed
            if event == sg.WIN_CLOSED or event == 'Cancel':
                window.close()
                sys.exit()
            # Check for ENTER key & input
            if event in ('\r', QT_ENTER_KEY1, QT_ENTER_KEY2) and len(values['-INPUT-']) and values['-INPUT-'][-1] not in (''):
                # Find the ok button & click it
                elem = window.find_element("Ok")
                elem.Click()
            # When "Ok" button is pressed
            if event == 'Ok' and len(values['-INPUT-']) and values['-INPUT-'][-1] not in (''):
                # Set variables as global
                global Mac_Domain
                global Time_Of_Drop_Int
                # Gather "Mac Domain" user input.
                Mac_Domain = values['-INPUT-']
                # Gather the time of drop from the user.
                Time_Of_Drop = values['-SLIDE-']
                # Convert the time of the drop to an int.
                Time_Of_Drop_Int = int(Time_Of_Drop)
                # Continue
                break
            # Check if input is "0123456789/:"
            if len(values['-INPUT-']) and values['-INPUT-'][-1] not in ('0123456789/:'):
                # Delete the last character 
                window['-INPUT-'].update(values['-INPUT-'][:-1])

        window.close()
    except Exception as e:
        # Export exception to variable
        global str_exception
        str_exception = str(e)
        # Run functions
        window.Close()
        error_window()

def user_input_conversion():
    try:
        # Set variable to global
        global Viewpoint_TimeFrame_Count
        global Viewpoint_Full_HUB
        global Viewpoint_Short_HUB_Node
        global Short_HUB_Input
        global Cluster
        global Short_HUB_Split_String
        global Mac_Domain_Split_First_Number_String
        global Mac_Domain_Split_Second_Number_String
        global Mac_Domain_Split_Third_Number_String

        # Set the amount of clicks for Viewpoint
        if Time_Of_Drop_Int >= 15 and Time_Of_Drop_Int <= 29:
            Viewpoint_TimeFrame_Count = 0
        elif Time_Of_Drop_Int >= 30 and Time_Of_Drop_Int <= 44:
            Viewpoint_TimeFrame_Count = 1
        elif Time_Of_Drop_Int >= 45 and Time_Of_Drop_Int <= 59:
            Viewpoint_TimeFrame_Count = 2
        elif Time_Of_Drop_Int >= 60 and Time_Of_Drop_Int <= 74:
            Viewpoint_TimeFrame_Count = 3
        elif Time_Of_Drop_Int >= 75 and Time_Of_Drop_Int <= 89:
            Viewpoint_TimeFrame_Count = 4
        elif Time_Of_Drop_Int >= 90 and Time_Of_Drop_Int <= 104:
            Viewpoint_TimeFrame_Count = 5
        elif Time_Of_Drop_Int >= 105 and Time_Of_Drop_Int <= 119:
            Viewpoint_TimeFrame_Count = 6
        elif Time_Of_Drop_Int == 120:
            Viewpoint_TimeFrame_Count = 7

        testingggg = "ALX 1P1"
        
        nodes = {
            # Alexis
                # Cluster 1
                    "7:0/0": {'shrt_hub': testingggg, 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/1": {'shrt_hub': 'ALX 1P4', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/2": {'shrt_hub': 'ALX 2', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/3": {'shrt_hub': 'ALX 3P1', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/4": {'shrt_hub': 'ALX 3P4', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/5": {'shrt_hub': 'ALX 4', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/6": {'shrt_hub': 'ALX 5P1', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/7": {'shrt_hub': 'ALX 6', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/8": {'shrt_hub': 'ALX 7', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/9": {'shrt_hub': 'ALX 8P1', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/10": {'shrt_hub': 'ALX 8P4', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                    "7:0/11": {'shrt_hub': 'ALX 9P1', 'clstr': 'COS-1', 'viwp_long': 'Alexis'},
                }
        
        # Set variables according to dictionary entry & user input
        Short_HUB_Input = nodes[Mac_Domain]['shrt_hub']
        Cluster = nodes[Mac_Domain]['clstr']
        Viewpoint_Full_HUB = nodes[Mac_Domain]['viwp_long']

        # Set an an empty var to set other variables as strings later
        out_str = ""

        # Split "Short_HUB_Input" into a HUB & Node ('ALX', '1P1')
        HUB_Node_Split = Short_HUB_Input.split()
        
        # Convert HUB variable to a list & convert to string ('ALX')
        Short_HUB_Split_List = [HUB_Node_Split[0]]
        Short_HUB_Split_String = out_str.join(Short_HUB_Split_List)
        # Convert Node variable to a list & convert to string ('1P1')
        Short_Node_Split_List = [HUB_Node_Split[1]]
        Short_Node_Split_String = out_str.join(Short_Node_Split_List)

        # Add missing 0's to the Node for Viewpoint base on length of the node number (01P1)
        if len(Short_Node_Split_String) == 1 or len(Short_Node_Split_String) == 3:
            Corrected_Short_Node_Split = "0"+Short_Node_Split_String
        else:
            Corrected_Short_Node_Split = Short_Node_Split_String

        # Merge corrected short HUB & Node for Viewpoint (ALX 01P1)
        Viewpoint_Short_HUB_Node = Short_HUB_Split_String+" "+Corrected_Short_Node_Split

        # Add "Node" infront of shortened HUB for PEA (Node ALX 1P1)
        PEA_Short_HUB_Node = "Node "+Short_HUB_Input

        # Convert the "/" in the input to a ":" for easier splitting later (7:0:0)
        Mac_Domain_Corrected = Mac_Domain.replace("/", ":")
        # Split the Mac Domain into individual numbers ('7', '0', '0')
        Mac_Domain_Split = Mac_Domain_Corrected.split(":")

        # Place first split number into variables for manipulation & convert to string ('7')
        Mac_Domain_Split_First_Number = [Mac_Domain_Split[0]]
        Mac_Domain_Split_First_Number_String = out_str.join(Mac_Domain_Split_First_Number)
        # Place second split number into variables for manipulation & convert to string ('0')
        Mac_Domain_Split_Second_Number = [Mac_Domain_Split[1]]
        Mac_Domain_Split_Second_Number_String = out_str.join(Mac_Domain_Split_Second_Number)
        # Place third split number into variables for manipulation & convert to string ('0')
        Mac_Domain_Split_Third_Number = [Mac_Domain_Split[2]]
        Mac_Domain_Split_Third_Number_String = out_str.join(Mac_Domain_Split_Third_Number)
    except Exception as e:
        # Export exception to variable
        global str_exception
        str_exception = str(e)
        # Run functions
        error_window()

def grab_data():
    try:
        # Themes & Colors
        sg.theme('DarkGrey4')
        DARK_COLOR = '#1e1e1e'
        LIGHT_COLOR = '#333333'

        # Images
        exclamation_img = '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'
        
        # Header block
        top_banner = [[sg.Image(data=exclamation_img, background_color=DARK_COLOR), sg.Text('Loading...', font=('Calibri', 20), background_color=DARK_COLOR)]]

        # Main block
        body = [
                    [sg.Text('Please be patient while I grab the requested data.', font=('Calibri', 13), background_color=LIGHT_COLOR)]
                ]

        # Define Layout
        layout = [
                    [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                    [sg.Column([[sg.Column(body, size=(550,150), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                    [sg.ProgressBar(1, orientation='h', size=(49, 10), pad=(15,15), key='progress')]
                ]

        # Create window
        window = sg.Window('MaintenanceBoi - Loading', layout, background_color=DARK_COLOR, keep_on_top=True, disable_close=True).Finalize()
        progress_bar = window.FindElement('progress')
    
        # Update progress bar
        progress_bar.UpdateBar(0, 5)
        
        # Open chrome
        global browser
        # Set chromedriver variables
        chromeOptions = webdriver.ChromeOptions()
        chromeOptions.add_argument("--window-size=1,1")
        chromedriver = Root+"/Backend/Chrome_Driver/chromedriver.exe"
        browser = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)

        # Set window outside of view
        browser.set_window_position(-2000,-2000)

        # Set window size to normal
        browser.set_window_size(1544, 1368)

        # First Tab
        browser.get(("https://harmonicinc.okta.com/login/login.htm?fromURI=/oauth2/v1/authorize/redirect?okta_key=3ZRZV6ACBLrdKnZEm6VZM1vShPqJ5nwPTs7cnMeLObc"))
        
        ## Log into Grafana
        # Type username
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-username"]'))).send_keys(Grafana_username)

        # Type password
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-password"]'))).send_keys(Grafana_password)
        time.sleep(1)

        # Click login
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-submit"]'))).click()

        # Wait for error page to show
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="content"]/div[2]')))

        # Update progress bar
        progress_bar.UpdateBar(1, 5)

        # Set link vars
        mac_domain_counter = "https://buckeye.cableos-operations.com/d/core-mac-domain-counters/core-mac-domain-counters?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Short_HUB_Split_String+"-"+Cluster+"&var-mdName=Md"+Mac_Domain_Split_First_Number_String+":"+Mac_Domain_Split_Second_Number_String+"%2F"+Mac_Domain_Split_Third_Number_String+".0"
        cm_states = "https://buckeye.cableos-operations.com/d/core-cm-states-per-mac-domain/core-cm-states-per-mac-domain?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Short_HUB_Split_String+"-"+Cluster+"&var-mdName=Md"+Mac_Domain_Split_First_Number_String+":"+Mac_Domain_Split_Second_Number_String+"%2F"+Mac_Domain_Split_Third_Number_String+".0"
        core_upstreams = "https://buckeye.cableos-operations.com/d/core-upstream-metrics-mh/core-upstream-metrics?orgId=1&refresh=1m&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Short_HUB_Split_String+"-"+Cluster+"&var-us_rf_port=Us"+Mac_Domain_Split_First_Number_String+":"+Mac_Domain_Split_Second_Number_String+"%2F"+Mac_Domain_Split_Third_Number_String
        viewpoint_home = "http://10.6.10.12/ViewPoint/site/Site/Login"

        # Open first link
        browser.execute_script("window.open('" + mac_domain_counter +"');")
        # Switch to new tab
        browser.switch_to.window(browser.window_handles[1])
        # Wait for page to load
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '/html/body/grafana-app/sidemenu/a/img')))

        # Open second link
        browser.execute_script("window.open('" + cm_states +"');")
        # Switch to new tab
        browser.switch_to.window(browser.window_handles[2])
        # Wait for page to load
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '/html/body/grafana-app/sidemenu/a/img')))
        
        # Open third link
        browser.execute_script("window.open('" + core_upstreams +"');")
        # Switch to new tab
        browser.switch_to.window(browser.window_handles[3])
        # Wait for page to load
        WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, '/html/body/grafana-app/sidemenu/a/img')))

        # Open fourth link
        browser.execute_script("window.open('" + viewpoint_home +"');")
        # Switch to new tab
        browser.switch_to.window(browser.window_handles[4])

        # Update progress bar
        progress_bar.UpdateBar(2, 5)

        ## Log into Viewpoint
        # Type username
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys(Viewpoint_username)

        # Type password
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Viewpoint_password)

        # Click "login"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginButton"]'))).click()

        # Click "RPM"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteOrgTreeContent"]/div/ul/ul/li[2]/a[1]'))).click()

        # Click "Buckeye Cable"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteOrgTreeContent"]/div/ul/ul/ul/li[1]/a[1]'))).click()

        #Update progress bar
        progress_bar.UpdateBar(3, 5)

        # Click "Alexis"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Full_HUB))).click()

        # Click "Node"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Short_HUB_Node))).click()

        # Click "Return Spectrum"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewContent"]/div/ul/li'))).click()

        # Click "Mode"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[2]/li[2]'))).click()
        
        # Click "Historical"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[2]/li[2]/select/option[2]'))).click()

        # Click "Display"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[2]/select'))).click()

        #Update progress bar
        progress_bar.UpdateBar(4, 5)

        # Click "Spectrum"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[2]/select/option[1]'))).click()

        # Click "Time Span"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[4]'))).click()
        
        # Click "15 min"
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[4]/select/option[1]'))).click()

        # Click "Back button for 15 minute increments"
        i = 0
        while i < Viewpoint_TimeFrame_Count:
            WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewContent"]/div[2]/ul/li[2]'))).click()
            time.sleep(.5)
            i += 1

        time.sleep(Viewpoint_TimeFrame_Count)

        # Close First tab
        browser.switch_to.window(browser.window_handles[0])
        browser.close()

        # Switch to First tab again
        browser.switch_to.window(browser.window_handles[0])
    
        # Update progress bar
        progress_bar.UpdateBar(5, 5)

        # Set window back into view
        browser.set_window_position(593,60)
        
        # Open Excel file
        xw.App(visible=True, add_book=False).books.open("\\\\taz\\cabout$\\Network Surveillance\\Reports On Demand\\RF Impairments\\RF Impairment Watchlist.xlsm")
    
        # Close The Window
        window.Close()
    except Exception as e:
        # Export exception to variable
        global str_exception
        str_exception = str(e)
        # Run functions
        quit_chrome()
        window.Close()
        error_window()
        
def completion_window():
    # Themes & Colors
    sg.theme('DarkGrey4')
    DARK_COLOR = '#1e1e1e'
    LIGHT_COLOR = '#333333'

    # Images
    checkmark_img = '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'

    # Header block
    top_banner = [[sg.Image(data=checkmark_img, background_color=DARK_COLOR), sg.Text('Completed!', font=('Calibri', 20), background_color=DARK_COLOR)]]

    # Main block
    body = [
                [sg.Text('Please review the gathered data.', font=('Calibri', 13), background_color=LIGHT_COLOR)],
                [sg.Text('Press \"Exit\" to get rid of this window.', font=('Calibri', 13), background_color=LIGHT_COLOR)]
              ]

    # Define Layout
    layout = [
                [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                [sg.Column([[sg.Column(body, size=(550,150), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                [sg.Button('Exit', size=(32,2), pad=(150,15), button_color=('white',LIGHT_COLOR), font=('Calibri', 12))]
             ]

    # Create window
    window = sg.Window('MaintenanceBoi - Completed', layout, background_color=DARK_COLOR, keep_on_top=True)

    # Event Loop
    while True:             
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
    window.close()

def version_check():
    # Set repo variables
    REPO_URL = "https://raw.githubusercontent.com/Gigoo25/Maintenance-Ticket-Assistant/test/"

    # Set tool verison check variables
    CURRENT_VERSION = "1.0"
    ONLINE_VERSION = "unidentified"

    # Delete version check file if found
    if os.path.isfile(Root+"\\version.txt") and os.access(Root+"\\version.txt", os.R_OK):
        os.remove(Root+"\\version.txt")

    # Download version to compare from online
    wget.download(REPO_URL+"version.txt", out=Root+"\\version.txt")

    # Check for version check file
    if os.path.isfile(Root+"\\version.txt"):
        # Open version check file
        ONLINE_VERSION_FILE = open(Root+"\\version.txt", "r")
        # Set online version
        lines=ONLINE_VERSION_FILE.readlines()
        ONLINE_VERSION=lines[0]
        # Close online file
        ONLINE_VERSION_FILE.close
    else:
        print("\n")
        print("Version check file was not found.")
        print("Skipping...")
        time.sleep(10)
        sys.exit()

    # Convert version number to floats for correct comparison
    CURRENT_VERSION_FLT = float(CURRENT_VERSION)
    ONLINE_VERSION_FLT = float(ONLINE_VERSION)

    # Compare versions and set variables 
    if ONLINE_VERSION_FLT > CURRENT_VERSION_FLT:
        update_window()
    else:
        print("\n")
        print("No update was found.")
        
def update_window():
    # Themes & Colors
    sg.theme('DarkGrey4')
    DARK_COLOR = '#1e1e1e'
    LIGHT_COLOR = '#333333'

    # Images
    download_img = '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'
    
    # Header block
    top_banner = [[sg.Image(data=download_img, background_color=DARK_COLOR), sg.Text('New version found!', font=('Calibri', 20), background_color=DARK_COLOR)]]

    # Main block
    body = [
                [sg.Text('There is a new version avalible!', font=('Calibri', 13), background_color=LIGHT_COLOR)],
                [sg.Text('Press \"Download\" to get the latest version.', font=('Calibri', 13), background_color=LIGHT_COLOR)]
              ]

    # Define Layout
    layout = [
                [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                [sg.Column([[sg.Column(body, size=(550,150), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                [sg.Button('Download', size=(32,2), pad=(150,15), button_color=('white',LIGHT_COLOR), font=('Calibri', 12))]
             ]

    # Create window
    window = sg.Window('MaintenanceBoi - Update', layout, background_color=DARK_COLOR, keep_on_top=True)

    # Event Loop
    while True:             
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            window.close()
            sys.exit()
        if event == 'Download':
            window.close()
            # Browser Variables
            chromedriver = Root+"/Backend/Chrome_Driver/chromedriver.exe"
            browser = webdriver.Chrome(executable_path=chromedriver)
            browser.get(("https://github.com/Gigoo25/Maintenance-Ticket-Assistant/releases"))

def progress():
    # Themes & Colors
    sg.theme('DarkGrey4')
    DARK_COLOR = '#1e1e1e'
    LIGHT_COLOR = '#333333'

    # Images
    exclamation_img = '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'
    exclamation_ico = 'AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAgBAAAAAAAAAAAAAAAAAAAAAAAAAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AMjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AMjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AMjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AMjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AMjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AMjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/wAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAyMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/ADIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8yMtz/MjLc/zIy3P8AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA='
    
    exclamation_ico = Root+"\\Backend\\Icons\\exclamation_mark-32.ico"
    
    # Header block
    top_banner = [[sg.Image(data=exclamation_img, background_color=DARK_COLOR), sg.Text('Loading...', font=('Calibri', 20), background_color=DARK_COLOR)]]

    # Main block
    body = [
                [sg.Text('Please be patient while I grab the requested data.', font=('Calibri', 13), background_color=LIGHT_COLOR)],
                [sg.Text('Press \"Exit\" to stop & close out of MaintenaceBoi.', font=('Calibri', 13), background_color=LIGHT_COLOR)]
              ]

    # Define Layout
    layout = [
                [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                [sg.Column([[sg.Column(body, size=(550,150), background_color=LIGHT_COLOR)]], background_color=DARK_COLOR)],
                [sg.ProgressBar(1, orientation='h', size=(49, 10), pad=(15,15), key='progress')]
             ]

    # Create window
    window = sg.Window('MaintenanceBoi - Loading', layout, background_color=DARK_COLOR, keep_on_top=True, disable_close=True, icon=r'H:\NOC\Development\Maintenance-Ticket-AssistantV2\Backend\Icons\red_plus.ico').Finalize()
    progress_bar = window.FindElement('progress')
 
    #This Updates the Window
    #progress_bar.UpdateBar(Current Value to show, Maximum Value to show)
    progress_bar.UpdateBar(0, 5)
    #adding time.sleep(length in Seconds) has been used to Simulate adding your script in between Bar Updates
    time.sleep(.5)
 
    progress_bar.UpdateBar(1, 5)
    time.sleep(.5)
 
    progress_bar.UpdateBar(2, 5)
    time.sleep(.5)
 
    progress_bar.UpdateBar(3, 5)
    time.sleep(.5)
 
    progress_bar.UpdateBar(4, 5)
    time.sleep(.5)
 
    progress_bar.UpdateBar(5, 5)
    time.sleep(.5)
    
    #I paused for 3 seconds at the end to give you time to see it has completed before closing the window
    time.sleep(3)
 
    #This will Close The Window
    window.Close()

### Execute functions defined above
## Prep functions
del_old_files()
create_missing_dirs()
check_credentials()
obtain_passwords()
## Update functions

## Analyze functions
main_window()
user_input_conversion()
grab_data()
completion_window()