"""
New Residential Fiber Form, NeRFF

Purpose: 

This GUI script captures necessary information for account creation, allocation of resources, and installation dispatch for new
Fiber Subscribers by pre-populating a Residential EULA/ToS .docx, printing plain text in organized manner to Notepad for quick copy
and paste into Platypus and Service Fusion by office personnel.


"""


import datetime
from pathlib import Path

import PySimpleGUI as sg
from docxtpl import DocxTemplate

document_path = Path(__file__).parent / "residential_fiber_paperwork.docx"
doc = DocxTemplate(document_path)

today = datetime.datetime.today()
installationFee = 99.00
paperBillingFee = 4
fiber100 = 49.95
fiber500 = 69.95
fiber1000 = 79.95
voipBundle = 19.95
smartWifi = 5.00
wifiExtender = 6.95
salesTax = 0.07

layout1 = [[sg.Text('Customer Information')],
           [sg.Text("First Name:"), sg.Input(
               key='customerFirstName', do_not_clear=True)],
           [sg.Text("Middle Initial:"), sg.Input(
               key="customerMiddleInitial", do_not_clear=True)],
           [sg.Text("Last Name:"), sg.Input(
               key='customerLastName', do_not_clear=True)],
           [sg.Text("Phone Number:"), sg.Input(
               key="customerPhoneNumber", do_not_clear=True)],
           [sg.Text("Email:"), sg.Input(
               key="customerEmailAddress", do_not_clear=True)],
           [sg.Text("Installation Street Address:"), sg.Input(
               key="installationStreetAddress", do_not_clear=True)],
           [sg.Text("Installation City:"), sg.Input(
               key="installationCity", do_not_clear=True)],
           [sg.Text("Installation State"), sg.Input(
               key="installationState", do_not_clear=True)],
           [sg.Text("Installation Zip Code:"), sg.Input(
               key="installationZipCode", do_not_clear=True)],
           [sg.Text("Billing Street Address:"), sg.Input(
               key="billingStreetAddress", do_not_clear=True)],
           [sg.Text("Billing City:"), sg.Input(
               key="billingCity", do_not_clear=True)],
           [sg.Text("Billing State:"), sg.Input(
               key="billingState", do_not_clear=True)],
           [sg.Text("Billing Zip Code:"), sg.Input(
               key="billingZipCode", do_not_clear=True)],
           [sg.Text("Payment Method")],
           [sg.CB('Credit Card', key='creditCard')],
           [sg.CB('Bank Draft', key='bankDraft')],
           [sg.CB('e-Statements', key='eStatement')],
           [sg.CB('Paper, $4/mo fee', key='paperBilling')],
           [sg.Text("Master Account Username:"), sg.Input(
               key="masterAccountUsername", do_not_clear=True)],
           [sg.Text("Master Account Password:"), sg.Input(
               key="masterAccountPassword", do_not_clear=True)],
           ]

layout2 = [[sg.Text('Service Information')],
           [sg.Text('Installation Date')],
           [sg.Text("Month XX:"),
            sg.Input(key="installMonth", do_not_clear=True)],
           [sg.Text("Day XX:"),
            sg.Input(key="installDay", do_not_clear=True)],
           [sg.Text("Year XXXX:"),
            sg.Input(key="installYear", do_not_clear=True)], 
           [sg.Text('Service Level:')],
           [sg.CB('Fiber 100 $49.95/mo', key='fiber100')],
           [sg.CB('Fiber 500 $69.95/mo', key='fiber500')],
           [sg.CB('Fiber 1GB $79.95/mo', key='fiber1000')],
           [sg.CB('Digital Voice Bundle $19.99/mo', key='voipService')],
           [sg.CB('Plume Wi-Fi $5/mo:', key='smartWifi')],
           [sg.CB('Wi-Fi Extender $6.95/mo:', key='wifiExtender')],
           [sg.Text("Wi-Fi/SSID Name"),
            sg.Input(key="ssid", do_not_clear=True)],
           [sg.Text("Wi-Fi Password"),
            sg.Input(key="wifiPSK", do_not_clear=True)],
           [sg.Text("Additional Item 1:"), sg.Input(
               key="additionalItem1", do_not_clear=True)],
           [sg.Text("Add Item 1 Price:"), sg.Input(
               key="addItemPrice1", do_not_clear=True)],
           [sg.Text("Additional Item 2:"), sg.Input(
               key="additionalItem2", do_not_clear=True)],
           [sg.Text("Add Item 2 Price:"), sg.Input(
               key="addItemPrice2", do_not_clear=True)],
           [sg.Text("Additional Item 3:"), sg.Input(
               key="additionalItem3", do_not_clear=True)],
           [sg.Text("Add Item 3 Price:"), sg.Input(
               key="addItemPrice3", do_not_clear=True)],
           ]


layout3 = [[sg.Text('Property Information')]]

# ----------- Create actual layout using Columns and a row of Buttons
layout = [[sg.Column(layout1, key='-COLCustomer Information-'), sg.Column(layout2, visible=False, key='-COLService Information-'), sg.Column(layout3, visible=False, key='-COLProperty Information-')],
          [sg.Button('Customer Information'), sg.Button('Service Information'), sg.Button(
              'Property Information'), sg.Button("Create Form"), sg.Exit()],
          ]

window = sg.Window("New Residential Fiber Form", layout,
                   element_justification="right")

#Variable declaration to denote page for Buttons
page = "Customer Information"

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    #Layout selector
    if event in 'Customer InformationService InformationProperty Information':
        window[f'-COL{page}-'].update(visible=False)
        page = str(event)
        window[f'-COL{page}-'].update(visible=True)
    elif event == "Create Form":
    
    #Prints today's date as MM/DD/YYYY format    
        values["intakeDate"] = today.strftime("%m-%d-%Y")
    
    #Determines what services get printed to the agreement
        if values['smartWifi'] == True:
            values['smartWifi'] = str('X')
        else:
            values['smartWifi'] = str(' ')
        if values['wifiExtender'] == True:
            values['wifiExtender'] = str('X')
        else:
            values['wifiExtender'] = str(' ')
        if values['fiber100'] == True:
            values['fiber100'] = str('Fiber 100 $49.95/mo')
        else:
            values['fiber100'] = str(' ')
        if values['fiber500'] == True:
            values['fiber500'] = str('Fiber 500 $69.95/mo')
        else:
            values['fiber500'] = str(' ')
        if values['fiber1000'] == True:
            values['fiber1000'] = str('Fiber 1GB $79.95/mo')
        else:
            values['fiber1000'] = str(' ')
        if values['voipService'] == True:
            values['voipService'] = str('Digital Voice Bundle $19.99/mo')
        else:
            values['voipService'] = str(' ')

    #Prints method of Billing

        if values['creditCard'] == True:
            values['paymentMethod'] = str('Credit Card')
        if values['bankDraft'] == True:
            values['paymentMethod'] = str(
                'Bank Draft (Complete ACH Authorization form)')
        if values['eStatement'] == True:
            values['paymentMethod'] = str('E-Statements')
        if values['paperBilling'] == True:
            values['paymentMethod'] = str('Paper, $4/mo fee')

    #Defines variables used in Tabulation
        if values['fiber100'] == str('Fiber 100 $49.95/mo'):
            fiberRate = fiber100
        if values['fiber500'] == ('Fiber 500 $69.95/mo'):
            fiberRate = fiber500
        if values['fiber1000'] == ('Fiber 1000 $79.95/mo'):
            fiberRate = fiber1000
        if values['smartWifi'] == True:
            smartWifi = 5.00
        if values['smartWifi'] == False:
            smartWifi = 0
        if values['wifiExtender'] == True:
            wifiExtender = 6.95
        if values['wifiExtender'] == False:
            wifiExtender = 0
        if values['voipService'] == False:
            voipBundle = 0
        if values["addItemPrice1"] == str(''):
            addItem1 = 0
        if values["addItemPrice2"] == str(''):
            addItem2 = 0
        if values["addItemPrice3"] == str(''):
            addItem3 = 0
        if values["addItemPrice1"] != str(''):
            addItem1 = values["addItemPrice1"]
            addItem1 = float(addItem1)
        if values["addItemPrice2"] != str(''):
            addItem2 = values["addItemPrice2"]
            addItem2 = float(addItem2)
        if values["addItemPrice3"] != str(''):
            addItem3 = values["addItemPrice3"]
            addItem3 = float(addItem3)

        #BIG CALENDAR MATH O.O
        
        #input data type conversion for math
        installMonth = int(values['installMonth'])
        installDay = int(values['installDay'])
        installYear = int(values['installYear'])

        #calculation for last day of installMonth and data type conversion
        installMonthLastDay = datetime.date(installYear + installMonth // 12, 
              installMonth % 12 + 1, 1) - datetime.timedelta(1)
        installMonthLastDay = int(installMonthLastDay.strftime("%d"))
        proRateDays = installMonthLastDay - installDay + 1

        #Tabulation
        dailyRate = (fiberRate + smartWifi + wifiExtender + voipBundle) / installMonthLastDay
        additionalItemsTotal = (addItem1 + addItem2 + addItem3) *(1 + salesTax)
        values['firstMonthService'] = round(dailyRate * proRateDays, 2)
        values['totalSalesTax'] = (addItem1 + addItem2 + addItem3) * salesTax
        values['installTotal'] = round(values['firstMonthService'] + installationFee + additionalItemsTotal, 2)


        #Render the template, save new word document & inform user
        doc.render(values)

        Live version path
        output_path = Path('Z:/Contracts/Fiber/NeRFF Agreements')/ f"{values['customerFirstName'] + values['customerLastName']}-FiberAgreement.docx"
        
        #Work from home path, desktop
        # output_path = Path('C:/Users/matth/OneDrive/Desktop')/ f"{values['customerFirstName'] + values['customerLastName']}-FiberAgreement.docx"
        # "C:\Users\matth\OneDrive\Desktop"

        #Laptop path, "C:/Users/matth/Desktop"
        #output_path = Path('C:/Users/matth/Desktop')/ f"{values['customerFirstName'] + values['customerLastName']}-FiberAgreement.docx"
        
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

window.close()
