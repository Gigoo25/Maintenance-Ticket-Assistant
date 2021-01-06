import PySimpleGUI as sg
import os
import time
import sys
import tkinter as tk
import wget
import sys
import xlwings as xw
import _thread
import threading
import re
from webdriver_manager.chrome import ChromeDriverManager
from threading import Thread
from shutil import copyfile
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from subprocess import CREATE_NO_WINDOW
from multiprocessing import Process

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
	os.system('taskkill /f /im chromedriver.exe')

## General window functions
def error_window():
    # Themes & Colors
    sg.theme('DarkGrey4')
    DARK_COLOR = '#1e1e1e'
    LIGHT_COLOR = '#333333'

    # Images
    error_img = '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'

    # Header block
    top_banner = [[sg.Image(data=error_img, background_color=DARK_COLOR), sg.Text('Error', font=('Calibri', 20), background_color=DARK_COLOR)]]

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
    os.system('taskkill /f /im chromedriver.exe')

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
    # Make backend directory if not found
    if not os.path.exists(Root+"\\Backend"):
        os.makedirs(Root+"\\Backend")

def download_chromedriver():
    # Set repo variables
    REPO_URL = "https://raw.githubusercontent.com/Gigoo25/Maintenance-Ticket-Assistant/test/Backend/"

    # Delete old Chrome driver if found
    if os.path.isfile(Root+"\\Backend\\chromedriver.exe") and os.access(Root+"\\Backend\\chromedriver.exe", os.R_OK):
        os.remove(Root+"\\Backend\\chromedriver.exe")

    # Download Chrome driver from repo
    print("Downloading Chrome driver.")
    wget.download(REPO_URL+"chromedriver.exe", out=Root+"\\Backend\\chromedriver.exe")

    # Check the downloaded Chrome driver file
    if os.path.isfile(Root+"\\Backend\\chromedriver.exe"):
        print("\n")
        print("Chrome driver found.")
    else:
        print("\n")
        print("Version check file was not found.")
        error_window()

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
                    [sg.Text("Username:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((4,0),(0,0))), sg.Text("Password:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((208,0),(0,0)))],
                    [sg.InputText('', key='pea_usr', enable_events=True, font=("Calibri", 12), size=(32, 1)), sg.InputText('', password_char="*", key='pea_pass', enable_events=True, font=("Calibri", 12), size=(32, 1))],
                    [sg.Text("Viewpoint", font=("Calibri", 20), background_color=LIGHT_COLOR, pad=((0,0),(5,2)))],
                    [sg.Text("Username:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((4,0),(0,0))), sg.Text("Password:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((208,0),(0,0)))],
                    [sg.InputText('', key='viw_usr', enable_events=True, font=("Calibri", 12), size=(32, 1)), sg.InputText('', password_char="*", key='viw_pass', enable_events=True, font=("Calibri", 12), size=(32, 1))],
                    [sg.Text("Grafana", font=("Calibri", 20), background_color=LIGHT_COLOR, pad=((0,0),(5,2)))],
                    [sg.Text("Username:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((4,0),(0,0))), sg.Text("Password:", font=("Calibri", 10), background_color=LIGHT_COLOR, pad=((208,0),(0,0)))],
                    [sg.InputText('', key='grf_usr', enable_events=True, font=("Calibri", 12), size=(32, 1)), sg.InputText('', password_char="*", key='grf_pass', enable_events=True, font=("Calibri", 12), size=(32, 1))],
                    [sg.Text(background_color=LIGHT_COLOR, pad=((4,0),(4,0)), text_color=('red'), font=("Calibri", 12), size=(50,1), key='-OUTPUT-')]
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
                # Check to make sure something is entered in all input fields
                if len(values['pea_usr']) == 0:
                    # Output error
                    window['-OUTPUT-'].update("Missing PEA Username.")
                elif len(values['pea_pass']) == 0:
                    # Output error
                    window['-OUTPUT-'].update("Missing PEA Password.")
                elif len(values['viw_usr']) == 0:
                    # Output error
                    window['-OUTPUT-'].update("Missing Viewpoint Username.")
                elif len(values['viw_pass']) == 0:
                    # Output error
                    window['-OUTPUT-'].update("Missing Viewpoint Password.")
                elif len(values['grf_usr']) == 0:
                    # Output error
                    window['-OUTPUT-'].update("Missing Grafana Username.")
                elif len(values['grf_pass']) == 0:
                    # Output error
                    window['-OUTPUT-'].update("Missing Grafana Password.")
                else:
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
def version_check():
    # Set repo variables
    REPO_URL = "https://raw.githubusercontent.com/Gigoo25/Maintenance-Ticket-Assistant/test/Backend/"

    # Set tool verison check variables
    CURRENT_VERSION = "1.0"
    ONLINE_VERSION = "unidentified"

    # Delete version check file if found
    if os.path.isfile(Root+"\\Backend\\Version.txt") and os.access(Root+"\\Backend\\Version.txt", os.R_OK):
        os.remove(Root+"\\Backend\\Version.txt")

    # Download version to compare from online
    print("Downloading version file.")
    wget.download(REPO_URL+"Version.txt", out=Root+"\\Backend\\Version.txt")

    # Check for version check file
    if os.path.isfile(Root+"\\Backend\\Version.txt"):
        # Open version check file
        ONLINE_VERSION_FILE = open(Root+"\\Backend\\Version.txt", "r")
        # Set online version
        lines=ONLINE_VERSION_FILE.readlines()
        ONLINE_VERSION=lines[0]
        # Close online file
        ONLINE_VERSION_FILE.close
    else:
        print("\n")
        print("Version check file was not found.")
        error_window()

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
    download_img = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAABmJLR0QA/wD/AP+gvaeTAAAAB3RJTUUH5AwYCTcyRl2rfgAAChFJREFUWMOtmXtwVPUVxz/nd3ez2eyGQEI0JAFDgAQIGEFgbCuPTimCKLWltlP7cvqwdmprfczYam2nr5nWPrTT2traWrV1qkVEUQRrFBXwAYXwqGDABEpJyDvZPDbZ7N7f6R+7eW02Yan9zdzkd+e3957vnvM9z5XVq1cZ4ksATbVX0KAd4JlFm+wvXl8qd05dMz0mZpnCCpDFwEwgF8hMPNMPtAMnQQ8IvOaofevrnW/U/6LiAXvN8btMj8lIJU8S94N7KwmAgyDHLAG+2bHHvT1/nbfWm3eJK3KtwlpgFuBP/b2S5WgfUCuww1HdNCfaevDvDc9Gv3nBKkeZcKmsXr3KGfHG4RMgz+3TlX0n9db8KytiYm5S+ChIfpLwc4Abs28ReMqj9v5ftmx7+1V/qbQ5fpFxADqlpSWSDFCBVX0n7eZghX9zcMH1rshvgNUggZGAVDV+2cT/wRMhlRYHTwPAEityxY5AeaTNyTq6seft6L+9k1OCHANQgds7drs35191Ycjx/1iFu4Cpo4WAtUrQn8nCkmLeP28282cUEvD56OrrJxKNMSxtXGpPAdaETUb+3szifd9t39lb7SsaA3KUiQfBbSy8rjginvsU+dhYDcTXBxfO5daPrGFZWSk5WX4AQuE+3qyp5d6tL7Lz8DuotSCCMTIhHQTd4tPYzY80Pln/QM4yI6lMPGjWm/OvurDfeH89HjhV+MTlS/nD1z7PotklZGVmMBBzsWrJCfgpK5rGhyvnc8HkbMqKpiEiNHV2Y61FRMbj6jxXTMm2QPnOT3cf7j01wtxDAHPdPt0crPCHHP+PFflcKiewVrmkdAYP3vR5pufn0RLq4nfPv8JPNz/PE7v3crYtxJzCCymYksMH5pexfslCrrlsETn+TPadODXC9Cn5OS8mTuBYRv7OsmhrrF+8MgjQCMhVvTW6KXvh9QnOOcngVBVUuf1ja7lyaSVd4T5u/dMT/HzLC5yob+TdhmaqDh+job2TlQvKCfdHGIi55GUHed/cWfT2R3ijpg7XdUEkSZtDoCv7TEb9bR279//LVyCAOjNLS+QbnW/oFwo2zk9469SxmrNMCWbxieXLuHHtKqZkB6g6eJTv/+0ZYq7FGIMkmPNOfSMHak/z+x2v8syb1cydPo0ZF+RxSekMFpXOIOjP5HRLO+FIJAFylPM4CvO3B8pf+n3z0027/CXGWVhSJJuyF3iaPcHvAR9OBa68uID7b/wMt2xYQ/7kbEB4+s0DbN9/BGPMkKlEBKuWd+ubONsR4tTZJozjcPWySrJ8PiouKmL9pRdTObOY6rrTtIS6ETHJWpyiIrIvs/jFqW5YnZprD2uofddiK/LDVHGuYEoOD950PVcuuRgjwvH6Jl49UsMzbx3kREPTCE4NPycmDlZFiLouGR4PJ+obEYT8nGzKi6cxt2gaLx56m56+SNI7BOCiDuN/+djMO844z/bfK3+bVPkNkDXJYUBVuXHdB7lh7Upc1+Whqt189Xd/4aGq3Ryvb0yYddw4hwg0h7p5dt9htrxZzda9B8nKzKBy5nRmFeTT0tXDnqMnkkwNQEBFOu8+de9L5uOF1xUrrBudfRVVyAn42bCsEmMMe4+f5Dt/3UJdYwuutYl3jQ8uOZRYtZxuaeO7jz3NWzV1iDFcvbSSnEAWqmM9W+GKe3JXFJmYmKXxxD86rKgqU7ODzMjPA+Dlw8do7gglOHc+4IaFGyO0dHax88g7AEzPzyUvOxCPEGNyNrOjYpaaeMlEVgqXHxEOIGYtyVpOR3Op3htz7dCdkXGLiyyFFQZkUSohIkJ7dy8N7Z0AXD5/DjnBLKy1/zM4ay05wQArKsoAONsRoq27d0RMTH5OFhugNJUQEejsDbPjwBFQZfn8Mu7YeCV5k4JYq4nLptjbpP3wed6kIHdsXMfl82eDKtv3H6GzNzyisBiTWmd6EpVwSu6oKg+/tIe1ixdy2dxZ3HbNGlZWlLG/9t/0R6NpGXtwZXq9LJldwqWzLyLD6+Wtmloeefl1VHVELBxDh1wxG75o43epC0xrLUvmlHDflz7F++fOQozhvSy1ljdqarnlj4+z9/jJUYE+hSU1AVBlokrYWktR3hQ+tWIZaxZVUJg7Gec8gbrW0tAe4oXqf/HErr2cae1IlGETt0NiNnyxN+7FTEBywSaKhUCmj6A/E8dIcok4GKDiHE5qbKxVuvv66e2PgICRiTQ3ZOKwB2iLA5zIAyGeNIRwZIBwZCAJmOIxDjdcsZKLZxZzsO40f/zHa8SsHQVVYEQc1XTCVZsB6s4du0anL0nEx8G9QfA6DtdctoivrF3F1UsrcYyDGfpc/LOInA84QOsM6IF0wU30RRTFtW6Cb5rkw+krICmSHDACu0DD7wXc2LDy/wCnYYHXjKN2L3DivaYvRrnG2Dx8/grgXa+6/zR/adzUILD9f9UciX5YE4F92G3iDVbiz3lbR2Db0/WP1ZtPXrZbHbVPAk3nC86IIZDpi18+H44T91DHGAI+39CZGVM1n1NGk6P2qbUrT6hs+NAHzH88OZ7DvoJfKnwtbbOqkpnh5dvXrmfVgnJUlXnTC8nLDtDa3cM7/zmLiFB16Cg/3bydgWgs4cXnHpUIev+CSNMtxbGQNWHxyn3Nz0c9an8LvJse5xQEwpEBntt3iKK8ySxfUE5udgCrSl52kOULyrlwyiSe23eI/oFo2uBAT3jU3n9n+ytuRDzxvnhXVonzk9YdzU9mLwwrXAHiSTcmnmntoKG9k9WV8/BneNFEjdfW3cNNDzzGK0dqRuTbc4KLGLjzoaanqv6cc6kZ1bgfz5gqIZN5rMf4ckGWpUtsSbSajjEsr5iDYwzRmMsPHt/Koy+/PqLfODd1BB4oiHXf22My3H7xDDfugPSLR9b3Ho9W+wr3xcSUgFSkRew4HamuOz3U/z5ctZsf/f05ojE3xRBpXHCbsmz0W18O7euq9eYOlhHDsxkBzngmyW0de8IvBmbvccVMByrSCgkiRKJRDp06Q9R1+dmWF2gNdafoX1KZNQ4uU2O3/LZ5a/PzgXIj40234o8Kn+2qtl+/4Oq8sPHerXAD4EtnQKmqeB2HqOsmTKvn4ly/wINZNvqj3zRvbX100mIjo0skHTLxSEMc8hXIxp63w7XevFd6je+0wjwgb+LpqSICVjUNs8a91cBd02Ld930p9M+u7YEyI+lOWOPmzpFZ0fbYr5qfO/B4duVLKmKBi4BgOmlvoiAs8LBX7W2PNm6q6nIy3bphzqWcsJpUM2oBIuKRPf4Sc0/r9uaajPyqVidQpSLtQHb8kow0M0QvcFTgIY/auxcOND16T+uOlj/nLDH94hlvPj3EQZloyj8oyacxti140N22+2LZWHjdtKg4S4HlGv8ZojROgVE/Q7SB1gEHBF7zqru/ru6RhsJ1jbq++lonIg7nKHkA9L+Rf/PHJVywiwAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyMC0xMi0yNFQwOTo1NTo0MSswMDowMC+U76UAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjAtMTItMjRUMDk6NTU6NDErMDA6MDBeyVcZAAAAAElFTkSuQmCC'
    
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
            # Set chromedriver location
            service = Service(Root+"/Backend/chromedriver.exe")
            # Set chromediver to not make an additional window
            service.creationflags = CREATE_NO_WINDOW
            # Open chromedriver
            browser = webdriver.Chrome(service=service)
            # Open release webpage
            browser.get(("https://github.com/Gigoo25/Maintenance-Ticket-Assistant/releases"))

def open_browser_background():
    global browser
    # Set chromedriver variables
    chromeOptions = webdriver.ChromeOptions()
    # Set window size to normal
    chromeOptions.add_argument("--window-size=1544,1368")
    # Set window outside of view
    chromeOptions.add_argument('--window-position=-2000,-2000')
    # Set chromedriver location
    service = Service(Root+"/Backend/chromedriver.exe")
    #service = Service(ChromeDriverManager().install())
    # Set chromediver to not make an additional window
    service.creationflags = CREATE_NO_WINDOW
    # Open chromedriver
    browser = webdriver.Chrome(service=service, options=chromeOptions)

## Analyze Functions
def main_window():
    try:
        def collapse(layout, key, visible):
            return sg.pin(sg.Column(layout, key=key, visible=visible, pad=(0,0), size=(549,70), background_color=LIGHT_COLOR))
            
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
        top_banner = [[sg.Image(data=error_img, background_color=DARK_COLOR), sg.Text('Welcome to MaintenanceBoi', font=('Calibri', 20), background_color=DARK_COLOR)]]

        # Main block
        body = [
                    [sg.Text('Enter a few basic details for the drop that you would like to investigate.', font=('Calibri', 13), background_color=LIGHT_COLOR, pad=((5,0),(0,15)))],
                    [sg.Text("Mac Domain:", font=("Calibri", 12), background_color=LIGHT_COLOR, pad=((5,0),(0,0))), sg.InputText('', enable_events=True,  key='-INPUT-', tooltip='For example \"7:0/0\"', font=("Calibri", 12), size=(50,1))],         
                    [sg.Text(background_color=LIGHT_COLOR, pad=((97,0),(0,0)), text_color=('red'), font=("Calibri", 12), size=(50,1), key='-OUTPUT-')],
                    [sg.Text("Event Type:", font=("Calibri", 12), background_color=LIGHT_COLOR, pad=((5,0),(0,0))), sg.Radio('Maintenance Ticket', "RADIO1", key="-RADIO1-", enable_events=True, background_color=LIGHT_COLOR, font=("Calibri", 12), default=True), sg.Radio('Node Outage', "RADIO1", key="-RADIO2-", enable_events=True, background_color=LIGHT_COLOR, font=("Calibri", 12)), sg.Radio('HUB Outage', "RADIO1", key="-RADIO3-", enable_events=True, background_color=LIGHT_COLOR, font=("Calibri", 12))],
                ]

        section1 = [
                    [sg.Text("Minutes since drop:", font=("Calibri", 12), background_color=LIGHT_COLOR, pad=((5,0),(17,0))), sg.Slider(range=(15, 120),  key='-SLIDE-', font=("Calibri", 12), orientation='h', size=(40, 20), default_value=15, background_color=LIGHT_COLOR)]
                ]

        # Define Layout
        layout = [
                    [sg.Column(top_banner, size=(550, 60), background_color=DARK_COLOR)],
                    [sg.Column([[sg.Column(body, size=(550,130), pad=((5,0),(0,0)), background_color=LIGHT_COLOR)]], pad=((0,0),(0,0)), background_color=DARK_COLOR)],
                    [sg.Column([[collapse(section1, 'sec_1', True)]], pad=((5,0),(0,0)), background_color=DARK_COLOR)],
                    [sg.Button('Ok', size=(32,2), pad=((5,22),(10,5)), button_color=('white',LIGHT_COLOR), font=('Calibri', 12)), sg.Button('Cancel', size=(32,2), pad=((0,0),(10,5)), button_color=('white',LIGHT_COLOR), font=('Calibri', 12))]
                ]

        # Create window
        window = sg.Window('MaintenanceBoi', layout, return_keyboard_events=True, background_color=DARK_COLOR, keep_on_top=True)

        # Pre-Load browser
        open_browser_background()

        # Event Loop to process "events" and get the "values" of the inputs
        while True:
            event, values = window.read()
            # When "Maintenance ticket" is selected unhide the "Minutes since drop" section
            if event == '-RADIO1-':
                window['sec_1'].update(visible=True)
            # When any other radio option is selected hide the "Minutes since drop" section
            elif event == '-RADIO2-' or event == '-RADIO3-':
                window['sec_1'].update(visible=False)
            # When "Cancel" button is pressed close chrome, window & exit script
            elif event == sg.WIN_CLOSED or event == 'Cancel':
                window.close()
                quit_chrome()
                sys.exit()
            # Check for ENTER key & input
            elif event in ('\r', QT_ENTER_KEY1, QT_ENTER_KEY2) and len(values['-INPUT-']) and values['-INPUT-'][-1] not in (''):
                # Find the ok button & click it
                elem = window.find_element("Ok")
                elem.Click()
            # When "Ok" button is pressed validate input against a regex
            elif event == 'Ok' and len(values['-INPUT-']) and values['-INPUT-'][-1] not in (''):
                if re.match(r"(^[0-9]{2}|^[0-9]{1}):([0-9]{2}|[0-9]{1})[\/]([0-9]{2}|[0-9]{1})", values['-INPUT-']):
                    # Set variables as global
                    global mac_domain
                    global time_of_drop_int
                    global event_type
                    # Gather "Mac Domain" user input.
                    mac_domain = values['-INPUT-']
                    # Gather the time of drop from the user.
                    Time_Of_Drop = values['-SLIDE-']
                    # Convert the time of the drop to an int.
                    time_of_drop_int = int(Time_Of_Drop)
                    if values["-RADIO1-"] == True:
                        # Set event type
                        event_type = "maintenance"
                        # Continue
                        break
                    elif values["-RADIO2-"] == True:
                        # Set event type
                        event_type = "node_outage"
                        # Continue
                        break
                    elif values["-RADIO3-"] == True:
                        # Set event type
                        event_type = "hub_outage"
                        # Continue
                        break
                else:
                    window['-INPUT-'].update(values['-INPUT-'][:-8])
                    window['-OUTPUT-'].update("'" + values['-INPUT-'] + "' is not a valid input!")
            # Only allow the following characters "0123456789/:"
            elif len(values['-INPUT-']) and values['-INPUT-'][-1] not in ('0123456789/:'):
                # Delete the last character 
                window['-INPUT-'].update(values['-INPUT-'][:-1])
            # Limit input length
            elif len(values['-INPUT-']) > 8:
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
        global mac_domain_Split_First_Number_String
        global mac_domain_Split_Second_Number_String
        global mac_domain_Split_Third_Number_String

        # Set the amount of clicks for Viewpoint
        if time_of_drop_int >= 15 and time_of_drop_int <= 29:
            Viewpoint_TimeFrame_Count = 0
        elif time_of_drop_int >= 30 and time_of_drop_int <= 44:
            Viewpoint_TimeFrame_Count = 1
        elif time_of_drop_int >= 45 and time_of_drop_int <= 59:
            Viewpoint_TimeFrame_Count = 2
        elif time_of_drop_int >= 60 and time_of_drop_int <= 74:
            Viewpoint_TimeFrame_Count = 3
        elif time_of_drop_int >= 75 and time_of_drop_int <= 89:
            Viewpoint_TimeFrame_Count = 4
        elif time_of_drop_int >= 90 and time_of_drop_int <= 104:
            Viewpoint_TimeFrame_Count = 5
        elif time_of_drop_int >= 105 and time_of_drop_int <= 119:
            Viewpoint_TimeFrame_Count = 6
        elif time_of_drop_int == 120:
            Viewpoint_TimeFrame_Count = 7
        
        nodes = {
            # Alexis
                # Cluster 1
                "7:0/0": {'shrt_hub': 'ALX 1P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/1": {'shrt_hub': 'ALX 1P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/2": {'shrt_hub': 'ALX 2', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/3": {'shrt_hub': 'ALX 3P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/4": {'shrt_hub': 'ALX 3P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/5": {'shrt_hub': 'ALX 4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/6": {'shrt_hub': 'ALX 5P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/7": {'shrt_hub': 'ALX 6', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/8": {'shrt_hub': 'ALX 7', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/9": {'shrt_hub': 'ALX 8P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/10": {'shrt_hub': 'ALX 8P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:0/11": {'shrt_hub': 'ALX 9P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/0": {'shrt_hub': 'ALX 9P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/1": {'shrt_hub': 'ALX 10P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/2": {'shrt_hub': 'ALX 10P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/3": {'shrt_hub': 'ALX 11P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/4": {'shrt_hub': 'ALX 11P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/5": {'shrt_hub': 'ALX 12 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/6": {'shrt_hub': 'ALX 13 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/7": {'shrt_hub': 'ALX 14 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/8": {'shrt_hub': 'ALX 15P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/9": {'shrt_hub': 'ALX 15P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/10": {'shrt_hub': 'ALX 16P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:1/11": {'shrt_hub': 'ALX 17P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/0": {'shrt_hub': 'ALX 17P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/1": {'shrt_hub': 'ALX 18P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/2": {'shrt_hub': 'ALX 18P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/3": {'shrt_hub': 'ALX 19 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/4": {'shrt_hub': 'ALX 20 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/5": {'shrt_hub': 'ALX 21 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/6": {'shrt_hub': 'ALX 22 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/7": {'shrt_hub': 'ALX 23P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/8": {'shrt_hub': 'ALX 23P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/9": {'shrt_hub': 'ALX 24 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/10": {'shrt_hub': 'ALX 25P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:2/11": {'shrt_hub': 'ALX 25P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/0": {'shrt_hub': 'ALX 26P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/1": {'shrt_hub': 'ALX 26P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/2": {'shrt_hub': 'ALX 27P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/3": {'shrt_hub': 'ALX 27P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/4": {'shrt_hub': 'ALX 28P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/5": {'shrt_hub': 'ALX 28P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/6": {'shrt_hub': 'ALX 29 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/7": {'shrt_hub': 'ALX 30P1', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/8": {'shrt_hub': 'ALX 30P4', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/9": {'shrt_hub': 'ALX 31 ', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/10": {'shrt_hub': 'ALX 32', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:3/11": {'shrt_hub': 'ALX 33', 'clstr': 'ALX-COS-1', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                # Cluster 2
                "7:4/0": {'shrt_hub': ' ALX 34', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/1": {'shrt_hub': ' ALX 35P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/2": {'shrt_hub': ' ALX 35P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/3": {'shrt_hub': ' ALX 36', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/4": {'shrt_hub': ' ALX 37', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/5": {'shrt_hub': ' ALX 38', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/6": {'shrt_hub': ' ALX 39', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/7": {'shrt_hub': ' ALX 40', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/8": {'shrt_hub': ' ALX 41', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/9": {'shrt_hub': ' ALX 42', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/10": {'shrt_hub': 'ALX 43', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:4/11": {'shrt_hub': 'ALX 44', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/0": {'shrt_hub': ' ALX 45', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/1": {'shrt_hub': ' ALX 46', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/2": {'shrt_hub': ' ALX 47', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/3": {'shrt_hub': ' ALX 48', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/4": {'shrt_hub': ' ALX 49', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/5": {'shrt_hub': ' ALX 50P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/6": {'shrt_hub': ' ALX 50P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/7": {'shrt_hub': ' ALX 51P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/8": {'shrt_hub': ' ALX 51P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/9": {'shrt_hub': ' ALX 52', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/10": {'shrt_hub': 'ALX 53', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:5/11": {'shrt_hub': 'ALX 54', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/0": {'shrt_hub': ' ALX 55', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/1": {'shrt_hub': ' ALX 56P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/2": {'shrt_hub': ' ALX 56P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/3": {'shrt_hub': ' ALX 57', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/4": {'shrt_hub': ' ALX 58P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/5": {'shrt_hub': ' ALX 58P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/6": {'shrt_hub': ' ALX 59', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/7": {'shrt_hub': ' ALX 60', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/8": {'shrt_hub': ' ALX 61P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/9": {'shrt_hub': ' ALX 62P1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/10": {'shrt_hub': 'ALX 62P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:9/11": {'shrt_hub': 'ALX 63', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/0": {'shrt_hub': 'ALX FUT1', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/1": {'shrt_hub': 'ALX 5P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/2": {'shrt_hub': 'ALX 61P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/3": {'shrt_hub': 'ALX 18P3', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/4": {'shrt_hub': 'ALX 18P6', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/5": {'shrt_hub': 'ALX 16P4', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/6": {'shrt_hub': 'ALX FUT7', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/7": {'shrt_hub': 'ALX FUT8', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/8": {'shrt_hub': 'ALX FUT9', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/9": {'shrt_hub': 'ALX FUT10', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/10": {'shrt_hub': 'ALX FUT11', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
                "7:10/11": {'shrt_hub': 'ALX FUT12', 'clstr': 'ALX-COS-2', 'viwp_long': 'Alexis', 'viwp_short': 'ALX'},
            # Angola
                # Cluster 1
                "14:0/0": {'shrt_hub': 'ANG 1 ', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/1": {'shrt_hub': 'ANG 2 ', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/2": {'shrt_hub': 'ANG 3 ', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/3": {'shrt_hub': 'ANG 5P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/4": {'shrt_hub': 'ANG 5P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/5": {'shrt_hub': 'ANG 6 ', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/6": {'shrt_hub': 'ANG 7P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/7": {'shrt_hub': 'ANG 7P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/8": {'shrt_hub': 'ANG 8P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/9": {'shrt_hub': 'ANG 8P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/10": {'shrt_hub': 'ANG 9', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:0/11": {'shrt_hub': 'ANG 10', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/0": {'shrt_hub': 'ANG 11P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/1": {'shrt_hub': 'ANG 11P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/2": {'shrt_hub': 'ANG 12P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/3": {'shrt_hub': 'ANG 13P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/4": {'shrt_hub': 'ANG 13P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/5": {'shrt_hub': 'ANG 14', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/6": {'shrt_hub': 'ANG 15', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/7": {'shrt_hub': 'ANG 16P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/8": {'shrt_hub': 'ANG 16P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/9": {'shrt_hub': 'ANG 17P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/10": {'shrt_hub': 'ANG 18P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:1/11": {'shrt_hub': 'ANG 18P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/0": {'shrt_hub': 'ANG 19', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/1": {'shrt_hub': 'ANG 20P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/2": {'shrt_hub': 'ANG 20P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/3": {'shrt_hub': 'ANG 21P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/4": {'shrt_hub': 'ANG 21P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/5": {'shrt_hub': 'ANG 22', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/6": {'shrt_hub': 'ANG 23', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/7": {'shrt_hub': 'ANG 24', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/8": {'shrt_hub': 'ANG 25', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/9": {'shrt_hub': 'ANG 26', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/10": {'shrt_hub': 'ANG 27', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:2/11": {'shrt_hub': 'ANG 28', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/0": {'shrt_hub': 'ANG 29', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/1": {'shrt_hub': 'ANG 30P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/2": {'shrt_hub': 'ANG 30P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/3": {'shrt_hub': 'ANG 31', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/4": {'shrt_hub': 'ANG 32', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/5": {'shrt_hub': 'ANG 33', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/6": {'shrt_hub': 'ANG 34', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/7": {'shrt_hub': 'ANG 35', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/8": {'shrt_hub': 'ANG 36P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/9": {'shrt_hub': 'ANG 36P4', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/10": {'shrt_hub': 'ANG 37', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:3/11": {'shrt_hub': 'ANG 38P1', 'clstr': 'ANG-COS-1', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                # Cluster 2
                "14:4/0": {'shrt_hub': 'ANG 38P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/1": {'shrt_hub': 'ANG 39', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/2": {'shrt_hub': 'ANG 40', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/3": {'shrt_hub': 'ANG 41P1', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/4": {'shrt_hub': 'ANG 41P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/5": {'shrt_hub': 'ANG 42P1', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/6": {'shrt_hub': 'ANG 42P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/7": {'shrt_hub': 'ANG 43', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/8": {'shrt_hub': 'ANG 44', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/9": {'shrt_hub': 'ANG 45', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/10": {'shrt_hub': 'ANG 46P1', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:4/11": {'shrt_hub': 'ANG 46P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/0": {'shrt_hub': 'ANG 47P1', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/1": {'shrt_hub': 'ANG 47P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/2": {'shrt_hub': 'ANG 48', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/3": {'shrt_hub': 'ANG 49P1', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/4": {'shrt_hub': 'ANG 49P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/5": {'shrt_hub': 'ANG 50', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/6": {'shrt_hub': 'ANG 51', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/7": {'shrt_hub': 'ANG 52', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/8": {'shrt_hub': 'SWYCK + NOC UP', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/9": {'shrt_hub': 'NOC BLDG LOWER', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/10": {'shrt_hub': 'ANG FUT1', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:5/11": {'shrt_hub': 'ANG 17P4', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/0": {'shrt_hub': 'ANG 47P3', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/1": {'shrt_hub': 'ANG 47P6', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/2": {'shrt_hub': 'ANG FUT5', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/3": {'shrt_hub': 'ANG FUT6', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/4": {'shrt_hub': 'ANG FUT7', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/5": {'shrt_hub': 'ANG FUT8', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/6": {'shrt_hub': 'ANG FUT9', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/7": {'shrt_hub': 'ANG FUT10', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/8": {'shrt_hub': 'ANG FUT11', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/9": {'shrt_hub': 'ANG FUT12', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/10": {'shrt_hub': 'ANG FUT13', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
                "14:9/11": {'shrt_hub': 'ANG FUT14', 'clstr': 'ANG-COS-2', 'viwp_long': 'HE (Angola)', 'viwp_short': 'HE'},
            # Arlington
                # Cluster 1
                "12:0/0": {'shrt_hub': 'ARL 1 ', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/1": {'shrt_hub': 'ARL 2P1', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/2": {'shrt_hub': 'ARL 2P4', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/3": {'shrt_hub': 'ARL 3 ', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/4": {'shrt_hub': 'ARL 4P1', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/5": {'shrt_hub': 'ARL 4P4', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/6": {'shrt_hub': 'ARL 5 ', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/7": {'shrt_hub': 'ARL 6 ', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/8": {'shrt_hub': 'ARL 7 ', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/9": {'shrt_hub': 'ARL 8 ', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/10": {'shrt_hub': 'ARL 9', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:0/11": {'shrt_hub': 'ARL 10', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/0": {'shrt_hub': 'ARL 11P1', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/1": {'shrt_hub': 'ARL 11P4', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/2": {'shrt_hub': 'ARL 12', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/3": {'shrt_hub': 'ARL 13P1', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/4": {'shrt_hub': 'ARL 13P4', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/5": {'shrt_hub': 'ARL 14', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/6": {'shrt_hub': 'ARL 15', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/7": {'shrt_hub': 'ARL 16', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/8": {'shrt_hub': 'ARL 17', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/9": {'shrt_hub': 'ARL 18', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/10": {'shrt_hub': 'ARL 19', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:1/11": {'shrt_hub': 'ARL 20', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/0": {'shrt_hub': 'ARL 21', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/1": {'shrt_hub': 'ARL 22', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/2": {'shrt_hub': 'ARL 23', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/3": {'shrt_hub': 'ARL 24', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/4": {'shrt_hub': 'ARL 25', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/5": {'shrt_hub': 'ARL 26', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/6": {'shrt_hub': 'ARL 27', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/7": {'shrt_hub': 'ARL 28', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/8": {'shrt_hub': 'ARL 29', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/9": {'shrt_hub': 'ARL 30', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/10": {'shrt_hub': 'ARL 31', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
                "12:2/11": {'shrt_hub': 'ARL FUT1', 'clstr': 'ARL-COS-1', 'viwp_long': 'Arlington', 'viwp_short': 'ARL'},
            # Bedford
                # Cluster 1
                "13:0/0": {'shrt_hub': 'BED 1P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/1": {'shrt_hub': 'BED 2P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/2": {'shrt_hub': 'BED 2P3', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/3": {'shrt_hub': 'BED 2P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/4": {'shrt_hub': 'BED 2P6', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/5": {'shrt_hub': 'BED 3P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/6": {'shrt_hub': 'BED 4P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/7": {'shrt_hub': 'BED 5P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/8": {'shrt_hub': 'BED 5P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/9": {'shrt_hub': 'BED 6P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/10": {'shrt_hub': 'BED 6P3', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:0/11": {'shrt_hub': 'BED 6P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/0": {'shrt_hub': 'BED 6P6', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/1": {'shrt_hub': 'BED 7 ', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/2": {'shrt_hub': 'BED 8 ', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/3": {'shrt_hub': 'BED 9 ', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/4": {'shrt_hub': 'BED 10', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/5": {'shrt_hub': 'BED 11', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/6": {'shrt_hub': 'BED 12', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/7": {'shrt_hub': 'BED 13P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/8": {'shrt_hub': 'BED 13P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/9": {'shrt_hub': 'BED 14', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/10": {'shrt_hub': 'BED 15P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:1/11": {'shrt_hub': 'BED 15P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/0": {'shrt_hub': 'BED 16P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/1": {'shrt_hub': 'BED 16P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/2": {'shrt_hub': 'BED 17P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/3": {'shrt_hub': 'BED 18P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/4": {'shrt_hub': 'BED 18P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/5": {'shrt_hub': 'BED 19', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/6": {'shrt_hub': 'BED 20', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/7": {'shrt_hub': 'BED 23', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/8": {'shrt_hub': 'BED 27P1', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/9": {'shrt_hub': 'BED 29', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/10": {'shrt_hub': 'BED 30', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:2/11": {'shrt_hub': 'BED 31', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/0": {'shrt_hub': 'BED 33', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/1": {'shrt_hub': 'BED 34', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/2": {'shrt_hub': 'BED 35', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/3": {'shrt_hub': 'BED 36', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/4": {'shrt_hub': 'BED 37', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/5": {'shrt_hub': 'BED 38', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/6": {'shrt_hub': 'BED 39', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/7": {'shrt_hub': 'BED 40', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/8": {'shrt_hub': 'BED 41', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/9": {'shrt_hub': 'BED 42', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/10": {'shrt_hub': 'BED 43P4', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:3/11": {'shrt_hub': 'BED 44', 'clstr': 'BDF-COS-1', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                # Cluster 2
                "13:4/0": {'shrt_hub': 'BED 45P4', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/1": {'shrt_hub': 'BED 46', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/2": {'shrt_hub': 'BED 47', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/3": {'shrt_hub': 'BED 48', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/4": {'shrt_hub': 'BED 49P1', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/5": {'shrt_hub': 'BED 49P4', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/6": {'shrt_hub': 'BED 50', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/7": {'shrt_hub': 'BED 51', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/8": {'shrt_hub': 'BED 201A RFoG ', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/9": {'shrt_hub': 'BED 201B RFoG ', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/10": {'shrt_hub': 'BED 202A RFoG', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:4/11": {'shrt_hub': 'BED 202B RFoG', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/0": {'shrt_hub': 'BED 4P4', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/1": {'shrt_hub': 'BED 17P4', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/2": {'shrt_hub': 'BED 27P4', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/3": {'shrt_hub': 'BED 15P3', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/4": {'shrt_hub': 'BED 15P6', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/5": {'shrt_hub': 'BED 3P4', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/6": {'shrt_hub': 'BED 1P1', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/7": {'shrt_hub': 'BED 3P3', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/8": {'shrt_hub': 'BED 3P6', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/9": {'shrt_hub': 'BED 45P1', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/10": {'shrt_hub': 'BED 43P1', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
                "13:5/11": {'shrt_hub': 'BED FUT12', 'clstr': 'BDF-COS-2', 'viwp_long': 'Bedford', 'viwp_short': 'BED'},
            # East Toledo
                # Cluster 1
                "11:0/0": {'shrt_hub': 'ET 1', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/1": {'shrt_hub': 'ET 2', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/2": {'shrt_hub': 'ET 3P1', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/3": {'shrt_hub': 'ET 3P4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/4": {'shrt_hub': 'ET 4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/5": {'shrt_hub': 'ET 5', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/6": {'shrt_hub': 'ET 6P1', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/7": {'shrt_hub': 'ET 6P4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/8": {'shrt_hub': 'ET 7', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/9": {'shrt_hub': 'ET 8', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/10": {'shrt_hub': 'ET 9 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:0/11": {'shrt_hub': 'ET 10', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/0": {'shrt_hub': 'ET 11 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/1": {'shrt_hub': 'ET 12 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/2": {'shrt_hub': 'ET 13 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/3": {'shrt_hub': 'ET 14 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/4": {'shrt_hub': 'ET 15P4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/5": {'shrt_hub': 'ET 16 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/6": {'shrt_hub': 'ET 17 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/7": {'shrt_hub': 'ET 18 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/8": {'shrt_hub': 'ET 19 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/9": {'shrt_hub': 'ET 20 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/10": {'shrt_hub': 'ET 21', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:1/11": {'shrt_hub': 'ET 22P1', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/0": {'shrt_hub': 'ET 22P4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/1": {'shrt_hub': 'ET 23 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/2": {'shrt_hub': 'ET 24 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/3": {'shrt_hub': 'ET 25 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/4": {'shrt_hub': 'ET 26 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/5": {'shrt_hub': 'ET 27 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/6": {'shrt_hub': 'ET 28P1', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/7": {'shrt_hub': 'ET 28P4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/8": {'shrt_hub': 'ET 29 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/9": {'shrt_hub': 'ET 30 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/10": {'shrt_hub': 'ET 31', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:2/11": {'shrt_hub': 'ET 32', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/0": {'shrt_hub': 'ET 33 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/1": {'shrt_hub': 'ET 34 ', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/2": {'shrt_hub': 'ET 15P1', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/3": {'shrt_hub': 'ET FUT2', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/4": {'shrt_hub': 'ET FUT3', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/5": {'shrt_hub': 'ET FUT4', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/6": {'shrt_hub': 'ET FUT5', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/7": {'shrt_hub': 'ET FUT6', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/8": {'shrt_hub': 'ET FUT7', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/9": {'shrt_hub': 'ET FUT8', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/10": {'shrt_hub': 'ET FUT9', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
                "11:3/11": {'shrt_hub': 'ET FUT10', 'clstr': 'ET-COS-1', 'viwp_long': 'East Toledo', 'viwp_short': 'ET'},
            # Erie
                # Cluster 1
                "16:0/0": {'shrt_hub': 'ERI 1P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/1": {'shrt_hub': 'ERI 1P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/2": {'shrt_hub': 'ERI 2 ', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/3": {'shrt_hub': 'ERI 3 ', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/4": {'shrt_hub': 'ERI 4 ', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/5": {'shrt_hub': 'ERI 5 ', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/6": {'shrt_hub': 'ERI 6P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/7": {'shrt_hub': 'ERI 7P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/8": {'shrt_hub': 'ERI 7P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/9": {'shrt_hub': 'ERI 8P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/10": {'shrt_hub': 'ERI 8P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:0/11": {'shrt_hub': 'ERI 9P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/0": {'shrt_hub': 'ERI 9P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/1": {'shrt_hub': 'ERI 10', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/2": {'shrt_hub': 'ERI 11', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/3": {'shrt_hub': 'ERI 12', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/4": {'shrt_hub': 'ERI 13', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/5": {'shrt_hub': 'ERI 14', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/6": {'shrt_hub': 'ERI 15', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/7": {'shrt_hub': 'ERI 16P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/8": {'shrt_hub': 'ERI 16P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/9": {'shrt_hub': 'ERI 17', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/10": {'shrt_hub': 'ERI 18', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:1/11": {'shrt_hub': 'ERI 19P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/0": {'shrt_hub': 'ERI 19P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/1": {'shrt_hub': 'ERI 20P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/2": {'shrt_hub': 'ERI 21P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/3": {'shrt_hub': 'ERI 21P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/4": {'shrt_hub': 'ERI 22P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/5": {'shrt_hub': 'ERI 22P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/6": {'shrt_hub': 'ERI 23P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/7": {'shrt_hub': 'ERI 23P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/8": {'shrt_hub': 'ERI 24', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/9": {'shrt_hub': 'ERI 25P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/10": {'shrt_hub': 'ERI 25P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:2/11": {'shrt_hub': 'ERI 26P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/0": {'shrt_hub': 'ERI 26P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/1": {'shrt_hub': 'ERI 27P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/2": {'shrt_hub': 'ERI 28', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/3": {'shrt_hub': 'ERI 29P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/4": {'shrt_hub': 'ERI 30P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/5": {'shrt_hub': 'ERI 30P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/6": {'shrt_hub': 'ERI 31', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/7": {'shrt_hub': 'ERI 32P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/8": {'shrt_hub': 'ERI 32P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/9": {'shrt_hub': 'ERI 33P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/10": {'shrt_hub': 'ERI 33P4', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:3/11": {'shrt_hub': 'ERI 34P1', 'clstr': 'ERIE-COS-1', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                # Cluster 2
                "16:4/0": {'shrt_hub': 'ERI 34P4', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/1": {'shrt_hub': 'ERI 35P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/2": {'shrt_hub': 'ERI 35P4', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/3": {'shrt_hub': 'ERI 36P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/4": {'shrt_hub': 'ERI 36P4', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/5": {'shrt_hub': 'ERI 37P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/6": {'shrt_hub': 'ERI 37P4', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/7": {'shrt_hub': 'ERI 38', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/8": {'shrt_hub': 'ERI 39', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/9": {'shrt_hub': 'ERI 40', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/10": {'shrt_hub': 'ERI 41', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:4/11": {'shrt_hub': 'ERI 42', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/0": {'shrt_hub': 'ERI 43', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/1": {'shrt_hub': 'ERI 44', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/2": {'shrt_hub': 'ERI 45', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/3": {'shrt_hub': 'ERI 46', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/4": {'shrt_hub': 'ERI 47', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/5": {'shrt_hub': 'ERI 48', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/6": {'shrt_hub': 'ERI 49', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/7": {'shrt_hub': 'ERI 50', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/8": {'shrt_hub': 'ERI 51', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/9": {'shrt_hub': 'ERI 52', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/10": {'shrt_hub': 'ERI 53', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:5/11": {'shrt_hub': 'ERI 54', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/0": {'shrt_hub': 'ERI 55', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/1": {'shrt_hub': 'ERI 57', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/2": {'shrt_hub': 'ERI 6P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/3": {'shrt_hub': 'ERI 20P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/4": {'shrt_hub': 'ERI 27P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/5": {'shrt_hub': 'ERI 19P3', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/6": {'shrt_hub': 'ERI 19P6', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/7": {'shrt_hub': 'ERI 29P1', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/8": {'shrt_hub': 'ERI 1P3', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/9": {'shrt_hub': 'ERI 1P6', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/10": {'shrt_hub': 'ERI FUT9', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
                "16:9/11": {'shrt_hub': 'ERI FUT10', 'clstr': 'ERIE-COS-2', 'viwp_long': 'Erie', 'viwp_short': 'ECC'},
            # Maumee
                # Cluster 1
                "4:0/0": {'shrt_hub': 'MAU 1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/1": {'shrt_hub': 'MAU 2P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/2": {'shrt_hub': 'MAU 2P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/3": {'shrt_hub': 'MAU 3P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/4": {'shrt_hub': 'MAU 3P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/5": {'shrt_hub': 'MAU 4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/6": {'shrt_hub': 'MAU 5P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/7": {'shrt_hub': 'MAU 6', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/8": {'shrt_hub': 'MAU 7', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/9": {'shrt_hub': 'MAU 8', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/10": {'shrt_hub': 'MAU 9P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:0/11": {'shrt_hub': 'MAU 9P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/0": {'shrt_hub': 'MAU 10 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/1": {'shrt_hub': 'MAU 11P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/2": {'shrt_hub': 'MAU 11P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/3": {'shrt_hub': 'MAU 12 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/4": {'shrt_hub': 'MAU 13 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/5": {'shrt_hub': 'MAU 14 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/6": {'shrt_hub': 'MAU 15P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/7": {'shrt_hub': 'MAU 15P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/8": {'shrt_hub': 'MAU 16 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/9": {'shrt_hub': 'MAU 17 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/10": {'shrt_hub': 'MAU 18', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:1/11": {'shrt_hub': 'MAU 19P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/0": {'shrt_hub': 'MAU 20 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/1": {'shrt_hub': 'MAU 5P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/2": {'shrt_hub': 'MAU 23 ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/3": {'shrt_hub': 'MAU 24 Dana HQ ', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/4": {'shrt_hub': 'MAU 201RFoG', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/5": {'shrt_hub': 'MAU 202A RFoG', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/6": {'shrt_hub': 'MAU 202B RFoG', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/7": {'shrt_hub': 'MAU 21P1', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/8": {'shrt_hub': 'MAU 21P3', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/9": {'shrt_hub': 'MAU 21P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/10": {'shrt_hub': 'MAU 21P6', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
                "4:2/11": {'shrt_hub': 'MAU 19P4', 'clstr': 'MAU-COS-1', 'viwp_long': 'Maumee', 'viwp_short': 'MAU'},
            # Northeast
                # Cluster 1
                "15:0/0": {'shrt_hub': 'NE 1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/1": {'shrt_hub': 'NE 2', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/2": {'shrt_hub': 'NE 3', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/3": {'shrt_hub': 'NE 4', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/4": {'shrt_hub': 'NE 5', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/5": {'shrt_hub': 'NE 6', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/6": {'shrt_hub': 'NE 7', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/7": {'shrt_hub': 'NE 8', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/8": {'shrt_hub': 'NE 9', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/9": {'shrt_hub': 'NE 10 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/10": {'shrt_hub': 'NE 11', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:0/11": {'shrt_hub': 'NE 12', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/0": {'shrt_hub': 'NE 13 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/1": {'shrt_hub': 'NE 14 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/2": {'shrt_hub': 'NE 17 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/3": {'shrt_hub': 'NE 18 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/4": {'shrt_hub': 'NE 19 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/5": {'shrt_hub': 'NE 20 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/6": {'shrt_hub': 'NE 21 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/7": {'shrt_hub': 'NE 22P1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/8": {'shrt_hub': 'NE 23P1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/9": {'shrt_hub': 'NE 23P4', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/10": {'shrt_hub': 'NE FUT 1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:1/11": {'shrt_hub': 'NE 24P4', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/0": {'shrt_hub': 'NE 25P1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/1": {'shrt_hub': 'NE 25P4', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/2": {'shrt_hub': 'NE 26 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/3": {'shrt_hub': 'NE 27P1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/4": {'shrt_hub': 'NE 27P4', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/5": {'shrt_hub': 'NE 28 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/6": {'shrt_hub': 'NE 29P1', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/7": {'shrt_hub': 'NE 29P4', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/8": {'shrt_hub': 'NE 30 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/9": {'shrt_hub': 'NE 31 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/10": {'shrt_hub': 'NE 32', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:2/11": {'shrt_hub': 'NE 33', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/0": {'shrt_hub': 'NE 34 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/1": {'shrt_hub': 'NE 35 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/2": {'shrt_hub': 'NE 36 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/3": {'shrt_hub': 'NE 37 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/4": {'shrt_hub': 'NE 38 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/5": {'shrt_hub': 'NE 39 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/6": {'shrt_hub': 'NE 41 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/7": {'shrt_hub': 'NE 42 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/8": {'shrt_hub': 'NE 43 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/9": {'shrt_hub': 'NE 44 ', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/10": {'shrt_hub': 'NE 45', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:3/11": {'shrt_hub': 'NE 46', 'clstr': 'NE-COS-1', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                # Cluster 2
                "15:4/0": {'shrt_hub': ' NE 48', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/1": {'shrt_hub': ' NE 49', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/2": {'shrt_hub': ' NE 50', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/3": {'shrt_hub': ' NE 51P1', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/4": {'shrt_hub': ' NE 51P4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/5": {'shrt_hub': ' NE 52', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/6": {'shrt_hub': ' NE 53P1', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/7": {'shrt_hub': ' NE 53P4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/8": {'shrt_hub': ' NE 54', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/9": {'shrt_hub': ' NE 55', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/10": {'shrt_hub': 'NE 56', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:4/11": {'shrt_hub': 'NE 57', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/0": {'shrt_hub': ' NE 58', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/1": {'shrt_hub': ' NE 59', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/2": {'shrt_hub': ' NE 60', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/3": {'shrt_hub': ' NE 61', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/4": {'shrt_hub': ' NE 62', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/5": {'shrt_hub': ' NE 63', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/6": {'shrt_hub': ' NE 64', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/7": {'shrt_hub': ' NE 65', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/8": {'shrt_hub': ' NE 66', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/9": {'shrt_hub': ' NE 67', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/10": {'shrt_hub': 'NE 68', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:5/11": {'shrt_hub': 'NE 69', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/0": {'shrt_hub': ' NE 70', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/1": {'shrt_hub': ' NE 71', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/2": {'shrt_hub': ' NE 72P1', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/3": {'shrt_hub': ' NE 72P4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/4": {'shrt_hub': ' NE 73', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/5": {'shrt_hub': ' NE 74', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/6": {'shrt_hub': ' NE 75P1', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/7": {'shrt_hub': ' NE 75P4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/8": {'shrt_hub': ' NE 76P1', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/9": {'shrt_hub': ' NE 76P4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/10": {'shrt_hub': 'NE 77', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:9/11": {'shrt_hub': 'NE 78', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/0": {'shrt_hub': 'NE 79', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/1": {'shrt_hub': 'NE 80', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/2": {'shrt_hub': 'NE 81', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/3": {'shrt_hub': 'NE 82', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/4": {'shrt_hub': 'NE 83', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/5": {'shrt_hub': 'NE 84', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/6": {'shrt_hub': 'NE 22P4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/7": {'shrt_hub': 'NE FUT2', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/8": {'shrt_hub': 'NE 24P1', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/9": {'shrt_hub': 'NE FUT4', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/10": {'shrt_hub': 'NE FUT5', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
                "15:10/11": {'shrt_hub': 'NE FUT6', 'clstr': 'NE-COS-2', 'viwp_long': 'NorthEast', 'viwp_short': 'NE'},
            # UT
                # Cluster 1
                "9:0/0": {'shrt_hub': ' UT 1', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/1": {'shrt_hub': ' UT 2', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/2": {'shrt_hub': ' UT 3', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/3": {'shrt_hub': ' UT 4P1', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/4": {'shrt_hub': ' UT 4P4', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/5": {'shrt_hub': ' UT 5', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/6": {'shrt_hub': ' UT 6', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/7": {'shrt_hub': ' UT 7', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/8": {'shrt_hub': ' UT 8', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/9": {'shrt_hub': ' UT 9', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/10": {'shrt_hub': 'UT 10 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:0/11": {'shrt_hub': 'UT 11 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/0": {'shrt_hub': ' UT 12 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/1": {'shrt_hub': ' UT 13 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/2": {'shrt_hub': ' UT 14 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/3": {'shrt_hub': ' UT 15 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/4": {'shrt_hub': ' UT 16 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/5": {'shrt_hub': ' UT 17 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/6": {'shrt_hub': ' UT 18 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/7": {'shrt_hub': ' UT 19 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/8": {'shrt_hub': ' UT 20 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/9": {'shrt_hub': ' UT 21 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/10": {'shrt_hub': 'UT 22 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:1/11": {'shrt_hub': 'UT 23 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:2/0": {'shrt_hub': ' UT 24 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:2/1": {'shrt_hub': ' UT 25 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:2/2": {'shrt_hub': ' UT 26 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:2/3": {'shrt_hub': ' UT 30 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:2/4": {'shrt_hub': ' UT FUT1', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "9:2/5": {'shrt_hub': ' UT FUT2', 'clstr': 'UTOW-COS-1', 'viwp_long': 'UT', 'viwp_short': 'UT'},
                "10:0/0": {'shrt_hub': 'OWN 1P1', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/1": {'shrt_hub': 'OWN 1P3', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/2": {'shrt_hub': 'OWN 1P4', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/3": {'shrt_hub': 'OWN 1P6', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/4": {'shrt_hub': 'OWN 2P1', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/5": {'shrt_hub': 'OWN 2P4', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/6": {'shrt_hub': 'OWN 3 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/7": {'shrt_hub': 'OWN 4 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/8": {'shrt_hub': 'OWN 5 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/9": {'shrt_hub': 'OWN 6 ', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/10": {'shrt_hub': 'OWN 7', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:0/11": {'shrt_hub': 'OWN 8', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:1/0": {'shrt_hub': 'OWN 10', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:1/1": {'shrt_hub': 'OWN 11', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:1/2": {'shrt_hub': 'OWN HamptonCou', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:1/3": {'shrt_hub': 'OWN FUT1', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:1/4": {'shrt_hub': 'OWN FUT2', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
                "10:1/5": {'shrt_hub': 'OWN FUT3', 'clstr': 'UTOW-COS-1', 'viwp_long': 'Owens', 'viwp_short': 'OWN'},
            # Perrysburg
                # Cluster 1
                "5:0/0": {'shrt_hub': 'PER 1P1', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/1": {'shrt_hub': 'PER 1P4', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/2": {'shrt_hub': 'PER 2', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/3": {'shrt_hub': 'PER 3', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/4": {'shrt_hub': 'PER 4', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/5": {'shrt_hub': 'PER 5P1', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/6": {'shrt_hub': 'PER 5P4', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/7": {'shrt_hub': 'PER 6', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/8": {'shrt_hub': 'PER 7', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/9": {'shrt_hub': 'PER 8', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/10": {'shrt_hub': 'PER 9 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:0/11": {'shrt_hub': 'PER 10', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/0": {'shrt_hub': 'PER 11P1', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/1": {'shrt_hub': 'PER 11P4', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/2": {'shrt_hub': 'PER 12 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/3": {'shrt_hub': 'PER 13 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/4": {'shrt_hub': 'PER 14 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/5": {'shrt_hub': 'PER 15P1', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/6": {'shrt_hub': 'PER 15P4', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/7": {'shrt_hub': 'PER 17P1', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/8": {'shrt_hub': 'PER 17P4', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/9": {'shrt_hub': 'PER 18 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/10": {'shrt_hub': 'PER 19', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:1/11": {'shrt_hub': 'PER 20', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/0": {'shrt_hub': 'PER 21 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/1": {'shrt_hub': 'PER 22 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/2": {'shrt_hub': 'PER 23 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/3": {'shrt_hub': 'PER 24 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/4": {'shrt_hub': 'PER 25 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/5": {'shrt_hub': 'PER 26 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/6": {'shrt_hub': 'PER 27 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/7": {'shrt_hub': 'PER 29 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/8": {'shrt_hub': 'PER 33 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/9": {'shrt_hub': 'PER 34 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/10": {'shrt_hub': 'PER 36', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:2/11": {'shrt_hub': 'PER 37', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/0": {'shrt_hub': 'PER 40 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/1": {'shrt_hub': 'PER 41 ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/2": {'shrt_hub': 'PER OI HQ', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/3": {'shrt_hub': 'PER Hotel 860', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/4": {'shrt_hub': 'PER 201A RFoG', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/5": {'shrt_hub': 'PER 201B RFoG', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/6": {'shrt_hub': 'PER 1P3', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/7": {'shrt_hub': 'PER 1P6', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/8": {'shrt_hub': 'PER 201', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/9": {'shrt_hub': 'PER SG-46', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/10": {'shrt_hub': 'PER SG-47', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
                "5:3/11": {'shrt_hub': 'PER SG-48', 'clstr': 'PER-COS-1', 'viwp_long': 'Perrysburg', 'viwp_short': 'PER'},
            # Springfield
                # Cluster 1
                "8:0/0": {'shrt_hub': 'SPF 1', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/1": {'shrt_hub': 'SPF 2', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/2": {'shrt_hub': 'SPF 3', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/3": {'shrt_hub': 'SPF 4P1', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/4": {'shrt_hub': 'SPF 4P4', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/5": {'shrt_hub': 'SPF 5', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/6": {'shrt_hub': 'SPF 6', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/7": {'shrt_hub': 'SPF 8', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/8": {'shrt_hub': 'SPF 10P1', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/9": {'shrt_hub': 'SPF 10P3', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/10": {'shrt_hub': 'SPF 10P4', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:0/11": {'shrt_hub': 'SPF 10p6', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/0": {'shrt_hub': 'SPF 11 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/1": {'shrt_hub': 'SPF 12 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/2": {'shrt_hub': 'SPF 13 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/3": {'shrt_hub': 'SPF 14 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/4": {'shrt_hub': 'SPF 15 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/5": {'shrt_hub': 'SPF 16 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/6": {'shrt_hub': 'SPF 17 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/7": {'shrt_hub': 'SPF 18 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/8": {'shrt_hub': 'SPF 19 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/9": {'shrt_hub': 'SPF 20 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/10": {'shrt_hub': 'SPF 21P1', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:1/11": {'shrt_hub': 'SPF 21P4', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/0": {'shrt_hub': 'SPF 22 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/1": {'shrt_hub': 'SPF 23 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/2": {'shrt_hub': 'SPF 24 ', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/3": {'shrt_hub': 'SPF 25P1', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/4": {'shrt_hub': 'SPF 25P4', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/6": {'shrt_hub': 'SPF FUT1', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/7": {'shrt_hub': 'SPF FUT2', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/8": {'shrt_hub': 'SPF FUT3', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/9": {'shrt_hub': 'SPF FUT4', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/10": {'shrt_hub': 'SPF FUT5', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
                "8:2/11": {'shrt_hub': 'SPF FUT6', 'clstr': 'SPR-COS-1', 'viwp_long': 'Springfield', 'viwp_short': 'SPG'},
            # Sylvania
                # Cluster 1
                "6:0/0": {'shrt_hub': 'SYL 1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/1": {'shrt_hub': 'SYL 2', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/2": {'shrt_hub': 'SYL 3', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/3": {'shrt_hub': 'SYL 4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/4": {'shrt_hub': 'SYL 5', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/5": {'shrt_hub': 'SYL 6P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/6": {'shrt_hub': 'SYL 6P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/7": {'shrt_hub': 'SYL 7P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/8": {'shrt_hub': 'SYL 7P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/9": {'shrt_hub': 'SYL 8', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/10": {'shrt_hub': 'SYL 9 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:0/11": {'shrt_hub': 'SYL 10P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/0": {'shrt_hub': 'SYL 10P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/1": {'shrt_hub': 'SYL 11 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/2": {'shrt_hub': 'SYL 12P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/3": {'shrt_hub': 'SYL 12P3', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/4": {'shrt_hub': 'SYL 12P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/5": {'shrt_hub': 'SYL 12P6', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/6": {'shrt_hub': 'SYL 13P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/7": {'shrt_hub': 'SYL 13P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/8": {'shrt_hub': 'SYL 14 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/9": {'shrt_hub': 'SYL 15P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/10": {'shrt_hub': 'SYL 16P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:1/11": {'shrt_hub': 'SYL 17P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/0": {'shrt_hub': 'SYL 17P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/1": {'shrt_hub': 'SYL 18 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/2": {'shrt_hub': 'SYL 19P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/3": {'shrt_hub': 'SYL 19P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/4": {'shrt_hub': 'SYL 20 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/5": {'shrt_hub': 'SYL 21 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/6": {'shrt_hub': 'SYL 22P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/7": {'shrt_hub': 'SYL 22P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/8": {'shrt_hub': 'SYL 23P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/9": {'shrt_hub': 'SYL 23P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/10": {'shrt_hub': 'SYL 24', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:2/11": {'shrt_hub': 'SYL 25', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/0": {'shrt_hub': 'SYL 26 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/1": {'shrt_hub': 'SYL 27P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/2": {'shrt_hub': 'SYL 27P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/3": {'shrt_hub': 'SYL 28P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/4": {'shrt_hub': 'SYL 28P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/5": {'shrt_hub': 'SYL 29P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/6": {'shrt_hub': 'SYL 29P3', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/7": {'shrt_hub': 'SYL 30 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/8": {'shrt_hub': 'SYL 31 ', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/9": {'shrt_hub': 'SYL 32P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/10": {'shrt_hub': 'SYL 32P4', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:3/11": {'shrt_hub': 'SYL 33P1', 'clstr': 'SYL-COS-1', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                # Cluster 2
                "6:4/0": {'shrt_hub': 'SYL 34p1', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/1": {'shrt_hub': 'SYL 34p3', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/2": {'shrt_hub': 'SYL 34p4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/3": {'shrt_hub': 'SYL 34p6', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/4": {'shrt_hub': 'SYL 35P1', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/5": {'shrt_hub': 'SYL 36 ', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/6": {'shrt_hub': 'SYL 37 ', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/7": {'shrt_hub': 'SYL 38P1', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/8": {'shrt_hub': 'SYL 35P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/9": {'shrt_hub': 'SYL 40 ', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/10": {'shrt_hub': 'SYL 42', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:4/11": {'shrt_hub': 'SYL 33P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/0": {'shrt_hub': 'SYL 201A RFoG', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/1": {'shrt_hub': 'SYL 201B RFoG', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/2": {'shrt_hub': 'SYL 202A RFoG', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/3": {'shrt_hub': 'SYL 202B RFoG', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/4": {'shrt_hub': 'SYL 27P3', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/5": {'shrt_hub': 'SYL 39P1', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/6": {'shrt_hub': 'SYL 39P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/7": {'shrt_hub': 'SYL 39P3', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/8": {'shrt_hub': 'SYL 39P6', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/9": {'shrt_hub': 'SYL 27P6', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/10": {'shrt_hub': 'SYL 6P3', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:5/11": {'shrt_hub': 'SYL 6P6', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/0": {'shrt_hub': 'SYL 28P3', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/1": {'shrt_hub': 'SYL 28P6', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/2": {'shrt_hub': 'SYL 29P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/3": {'shrt_hub': 'SYL 29P6', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/4": {'shrt_hub': 'SYL 16P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/5": {'shrt_hub': 'SYL 38P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/6": {'shrt_hub': 'SYL 15P4', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/7": {'shrt_hub': 'SYL SG-80', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/8": {'shrt_hub': 'SYL SG-81', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/9": {'shrt_hub': 'SYL SG-82', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/10": {'shrt_hub': 'SYL SG-83', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
                "6:9/11": {'shrt_hub': 'SYL SG-84', 'clstr': 'SYL-COS-2', 'viwp_long': 'Sylvania', 'viwp_short': 'SYL'},
            # Watergon
                # Cluster 1
                "1:0/0": {'shrt_hub': 'WAT 1', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/1": {'shrt_hub': 'WAT 2', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/2": {'shrt_hub': 'WAT 3', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/3": {'shrt_hub': 'WAT 4', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/4": {'shrt_hub': 'WAT 5', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/5": {'shrt_hub': 'WAT 6', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/6": {'shrt_hub': 'WAT 7', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/7": {'shrt_hub': 'WAT 8', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/8": {'shrt_hub': 'WAT 9', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/9": {'shrt_hub': 'WAT 10 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/10": {'shrt_hub': 'WAT 11', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:0/11": {'shrt_hub': 'WAT 12', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:1/0": {'shrt_hub': 'WAT 13 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:1/1": {'shrt_hub': 'WAT 14 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:1/2": {'shrt_hub': 'WAT FUT1', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "1:1/3": {'shrt_hub': 'WAT FUT2', 'clstr': 'WAT-COS-1', 'viwp_long': 'Waterville', 'viwp_short': 'WAT'},
                "3:0/0": {'shrt_hub': 'ORG 1', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/1": {'shrt_hub': 'ORG 2', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/2": {'shrt_hub': 'ORG 3', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/3": {'shrt_hub': 'ORG 4', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/4": {'shrt_hub': 'ORG 5', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/5": {'shrt_hub': 'ORG 6', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/6": {'shrt_hub': 'ORG 7', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/7": {'shrt_hub': 'ORG 8', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/8": {'shrt_hub': 'ORG 9', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/9": {'shrt_hub': 'ORG 10 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/10": {'shrt_hub': 'ORG 11', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:0/11": {'shrt_hub': 'ORG 12', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/0": {'shrt_hub': 'ORG 13 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/1": {'shrt_hub': 'ORG 14 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/2": {'shrt_hub': 'ORG 15 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/3": {'shrt_hub': 'ORG 16 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/4": {'shrt_hub': 'ORG 17 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/5": {'shrt_hub': 'ORG 18 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/6": {'shrt_hub': 'ORG 19 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/7": {'shrt_hub': 'ORG 20 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/8": {'shrt_hub': 'ORG 21 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/9": {'shrt_hub': 'ORG 22 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/10": {'shrt_hub': 'ORG 25', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:1/11": {'shrt_hub': 'ORG 26', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:2/0": {'shrt_hub': 'ORG 27 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:2/1": {'shrt_hub': 'ORG 28 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:2/2": {'shrt_hub': 'ORG 30 ', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'},
                "3:2/3": {'shrt_hub': 'ORG FUT 1', 'clstr': 'WAT-COS-1', 'viwp_long': 'Oregon', 'viwp_short': 'ORG'}
                }
        
        # Set variables according to dictionary entry & user input
        Short_HUB_Input = nodes[mac_domain]['shrt_hub']
        Cluster = nodes[mac_domain]['clstr']
        Viewpoint_Full_HUB = nodes[mac_domain]['viwp_long']
        Viewpoint_Short_HUB = nodes[mac_domain]['viwp_short']

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
        Viewpoint_Short_HUB_Node = Viewpoint_Short_HUB+" "+Corrected_Short_Node_Split

        # Add "Node" infront of shortened HUB for PEA (Node ALX 1P1)
        PEA_Short_HUB_Node = "Node "+Short_HUB_Input

        # Convert the "/" in the input to a ":" for easier splitting later (7:0:0)
        mac_domain_Corrected = mac_domain.replace("/", ":")
        # Split the Mac Domain into individual numbers ('7', '0', '0')
        mac_domain_Split = mac_domain_Corrected.split(":")

        # Place first split number into variables for manipulation & convert to string ('7')
        mac_domain_Split_First_Number = [mac_domain_Split[0]]
        mac_domain_Split_First_Number_String = out_str.join(mac_domain_Split_First_Number)
        # Place second split number into variables for manipulation & convert to string ('0')
        mac_domain_Split_Second_Number = [mac_domain_Split[1]]
        mac_domain_Split_Second_Number_String = out_str.join(mac_domain_Split_Second_Number)
        # Place third split number into variables for manipulation & convert to string ('0')
        mac_domain_Split_Third_Number = [mac_domain_Split[2]]
        mac_domain_Split_Third_Number_String = out_str.join(mac_domain_Split_Third_Number)
    except Exception as e:
        # Export exception to variable
        global str_exception
        str_exception = str(e)
        # Run functions
        error_window()

def determine_maintenance():
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
        window = sg.Window('MaintenanceBoi - Loading', layout, background_color=DARK_COLOR, keep_on_top=True, no_titlebar=True, grab_anywhere=True).Finalize()
        progress_bar = window.FindElement('progress')
    
        # Update progress bar
        progress_bar.UpdateBar(0, 5)
        
        # Set gloval var
        global browser

        # Set link vars
        mac_domain_counter = "https://buckeye.cableos-operations.com/d/core-mac-domain-counters/core-mac-domain-counters?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-mdName=Md"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String+".0"
        cm_states = "https://buckeye.cableos-operations.com/d/core-cm-states-per-mac-domain/core-cm-states-per-mac-domain?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-mdName=Md"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String+".0"
        core_upstreams = "https://buckeye.cableos-operations.com/d/core-upstream-metrics-mh/core-upstream-metrics?orgId=1&refresh=1m&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-us_rf_port=Us"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String
        viewpoint_home = "http://10.6.10.12/ViewPoint/site/Site/Login"

        # Open first link
        browser.get((mac_domain_counter))

        ## Log into Grafana
        # Type username
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-username"]'))).send_keys(Grafana_username)

        # Type password
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-password"]'))).send_keys(Grafana_password)
        time.sleep(.5)

        # Click login
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-submit"]'))).click()

        # Update progress bar
        progress_bar.UpdateBar(1, 5)

        # Wait for page to load
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/grafana-app/sidemenu/a/img')))

        # Open second link
        browser.execute_script("window.open('" + cm_states +"');")
        
        # Open third link
        browser.execute_script("window.open('" + core_upstreams +"');")

        # Open fourth link
        browser.execute_script("window.open('" + viewpoint_home +"');")

        # Update progress bar
        progress_bar.UpdateBar(2, 5)

        ## Log into Viewpoint
        # Switch driver to the correct tab
        browser.switch_to.window(browser.window_handles[1])

        # Wait for page to load
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="siteLoadingLogo"]')))

        # Type username
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys(Viewpoint_username)

        # Type password
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Viewpoint_password)

        # Click "login"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginButton"]'))).click()

        # Click "RPM"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteOrgTreeContent"]/div/ul/ul/li[2]/a[1]'))).click()

        # Click "Buckeye Cable"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteOrgTreeContent"]/div/ul/ul/ul/li[1]/a[1]'))).click()

        #Update progress bar
        progress_bar.UpdateBar(3, 5)

        # Click the correct HUB
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Full_HUB))).click()
        
        # Click the correct Node
        try:
            WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Short_HUB_Node)))
        except Exception:
            ActionChains(browser).key_down(Keys.PAGE_DOWN).key_up(Keys.PAGE_DOWN).perform()
        finally:
            WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="%s"]' % Viewpoint_Short_HUB_Node))).click()
        
        # Click "Return Spectrum"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewContent"]/div/ul/li'))).click()

        # Click "Mode"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[2]/li[2]'))).click()
        
        # Click "Historical"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[2]/li[2]/select/option[2]'))).click()

        # Click "Display"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[2]/select'))).click()

        #Update progress bar
        progress_bar.UpdateBar(4, 5)

        # Click "Spectrum"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[2]/select/option[1]'))).click()

        # Click "Time Span"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[4]'))).click()
        
        # Click "15 min"
        WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewPanelContent"]/ul[3]/ul[1]/li[4]/select/option[1]'))).click()

        # Click "Back button for 15 minute increments"
        i = 0
        while i < Viewpoint_TimeFrame_Count:
            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="siteViewContent"]/div[2]/ul/li[2]'))).click()
            time.sleep(.5)
            i += 1

        time.sleep(Viewpoint_TimeFrame_Count)

        # Switch to First tab again
        browser.switch_to.window(browser.window_handles[0])
    
        # Update progress bar
        progress_bar.UpdateBar(5, 5)
        
        # Set window back into view
        browser.set_window_position(593,60)

        # Open Excel file
        xw.App(visible=True, add_book=False).books.open("\\\\taz\\cabout$\\Network Surveillance\\Reports On Demand\\RF Impairments\\RF Impairment Watchlist.xlsm").sheets[Viewpoint_Full_HUB].activate()

        # Close The Window
        window.Close()
    except Exception as e:
        # Set exception as global variable
        global str_exception
        # Set credentials file location
        Credentials_File = Root+"\\Backend\\Credentials.txt"
        # Check if errors are present on the page, otherwise disaplay generic error
        if browser.find_elements(By.XPATH, '//*[text()="Unable to sign in"]'):
            # Export exception to variable
            str_exception = "Error with your username or password! \n\n I've deleted the credentials file. Please restart MaintenanceBoi."
            # Remove credentials file
            if os.path.isfile(Credentials_File) and os.access(Credentials_File, os.R_OK):
                os.remove(Credentials_File)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()
        elif browser.find_elements(By.XPATH, '//div[@class="error active"]'):
            # Export exception to variable
            str_exception = "Error with your username or password! \n\n I've deleted the credentials file. Please restart MaintenanceBoi."
            # Remove credentials file
            if os.path.isfile(Credentials_File) and os.access(Credentials_File, os.R_OK):
                os.remove(Credentials_File)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()
        else:   
            # Export exception to variable
            str_exception = str(e)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()

def node_outage_open():
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
        window = sg.Window('MaintenanceBoi - Loading', layout, background_color=DARK_COLOR, keep_on_top=True, no_titlebar=True, grab_anywhere=True).Finalize()
        progress_bar = window.FindElement('progress')
    
        # Update progress bar
        progress_bar.UpdateBar(0, 5)
        
        # Set gloval var
        global browser

        # Set link vars
        mac_domain_counter = "https://buckeye.cableos-operations.com/d/core-mac-domain-counters/core-mac-domain-counters?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-mdName=Md"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String+".0"
        cm_states = "https://buckeye.cableos-operations.com/d/core-cm-states-per-mac-domain/core-cm-states-per-mac-domain?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-mdName=Md"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String+".0"
        power_map = "http://outages.firstenergycorp.com/oh.html"
        
        # Open first link
        browser.get((mac_domain_counter))

        ## Log into Grafana
        # Type username
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-username"]'))).send_keys(Grafana_username)

        # Type password
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-password"]'))).send_keys(Grafana_password)
        time.sleep(.5)

        # Click login
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-submit"]'))).click()

        # Update progress bar
        progress_bar.UpdateBar(1, 5)

        # Wait for page to load
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/grafana-app/sidemenu/a/img')))

        # Open second link
        browser.execute_script("window.open('" + cm_states +"');")

        # Update progress bar
        progress_bar.UpdateBar(2, 5)

        # Open third link
        browser.execute_script("window.open('" + power_map +"');")

        # Switch to First tab again
        browser.switch_to.window(browser.window_handles[0])
    
        # Update progress bar
        progress_bar.UpdateBar(5, 5)
        
        # Set window back into view
        browser.set_window_position(593,60)

        # Close The Window
        window.Close()
    except Exception as e:
        # Set exception as global variable
        global str_exception
        # Set credentials file location
        Credentials_File = Root+"\\Backend\\Credentials.txt"
        # Check if errors are present on the page, otherwise disaplay generic error
        if browser.find_elements(By.XPATH, '//*[text()="Unable to sign in"]'):
            # Export exception to variable
            str_exception = "Error with your username or password! \n\n I've deleted the credentials file. Please restart MaintenanceBoi."
            # Remove credentials file
            if os.path.isfile(Credentials_File) and os.access(Credentials_File, os.R_OK):
                os.remove(Credentials_File)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()
        else:   
            # Export exception to variable
            str_exception = str(e)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()

def hub_outage_open():
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
        window = sg.Window('MaintenanceBoi - Loading', layout, background_color=DARK_COLOR, keep_on_top=True, no_titlebar=True, grab_anywhere=True).Finalize()
        progress_bar = window.FindElement('progress')
    
        # Update progress bar
        progress_bar.UpdateBar(0, 5)
        
        # Set gloval var
        global browser

        # Set link vars
        mac_domain_counter = "https://buckeye.cableos-operations.com/d/core-mac-domain-counters/core-mac-domain-counters?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-mdName=Md"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String+".0"
        cm_states = "https://buckeye.cableos-operations.com/d/core-cm-states-per-mac-domain/core-cm-states-per-mac-domain?orgId=1&from=now-24h&to=now&var-ds_zabbix=default&var-ds_postgres=ZabbixDirect&var-group=buckeye&var-setup="+Cluster+"&var-mdName=Md"+mac_domain_Split_First_Number_String+":"+mac_domain_Split_Second_Number_String+"%2F"+mac_domain_Split_Third_Number_String+".0"
        # Open first link
        browser.get((mac_domain_counter))

        ## Log into Grafana
        # Type username
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-username"]'))).send_keys(Grafana_username)

        # Type password
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-password"]'))).send_keys(Grafana_password)
        time.sleep(.5)

        # Click login
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="okta-signin-submit"]'))).click()

        # Update progress bar
        progress_bar.UpdateBar(1, 5)

        # Wait for page to load
        WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/grafana-app/sidemenu/a/img')))

        # Open second link
        browser.execute_script("window.open('" + cm_states +"');")

        # Update progress bar
        progress_bar.UpdateBar(2, 5)

        # Switch to First tab again
        browser.switch_to.window(browser.window_handles[0])
    
        # Update progress bar
        progress_bar.UpdateBar(5, 5)
        
        # Set window back into view
        browser.set_window_position(593,60)

        # Close The Window
        window.Close()
    except Exception as e:
        # Set exception as global variable
        global str_exception
        # Set credentials file location
        Credentials_File = Root+"\\Backend\\Credentials.txt"
        # Check if errors are present on the page, otherwise disaplay generic error
        if browser.find_elements(By.XPATH, '//*[text()="Unable to sign in"]'):
            # Export exception to variable
            str_exception = "Error with your username or password! \n\n I've deleted the credentials file. Please restart MaintenanceBoi."
            # Remove credentials file
            if os.path.isfile(Credentials_File) and os.access(Credentials_File, os.R_OK):
                os.remove(Credentials_File)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()
        else:   
            # Export exception to variable
            str_exception = str(e)
            # Run functions
            quit_chrome()
            window.Close()
            error_window()

### Execute functions defined above
## Prep functions
del_old_files()
create_missing_dirs()
download_chromedriver()
check_credentials()
obtain_passwords()
## Update functions
version_check()
## Analyze functions
main_window()
user_input_conversion()                
if event_type == "maintenance":
    print("Working Maintenance Ticket")
    determine_maintenance()
if event_type == "node_outage":
    print("Working Node Outage")
    node_outage_open()
if event_type == "hub_outage":
    print("Working HUB Outage")
    hub_outage_open()
completion_window()