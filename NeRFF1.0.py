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
installationFee = float(99.00)
paperBillingFee = float(4)
fiber100 = float(49.95)
fiber500 = float(69.95)
fiber1000 = float(79.95)
voipBundle = float(19.95)
smartWifi = float(5.00)
wifiExtender = float(6.95)


layout1 = [[sg.Text('Customer Information')],
            [sg.Text("First Name:"), sg.Input(key='customerFirstName', do_not_clear=True)],
            [sg.Text("Middle Initial:"), sg.Input(key="customerMiddleInitial", do_not_clear=True)],
            [sg.Text("Last Name:"), sg.Input(key='customerLastName', do_not_clear=True)],
            [sg.Text("Phone Number:"), sg.Input(key="customerPhoneNumber", do_not_clear=True)],
            [sg.Text("Email:"), sg.Input(key="customerEmailAddress", do_not_clear=True)],
            [sg.Text("Installation Street Address:"), sg.Input(key="installationStreetAddress", do_not_clear=True)],
            [sg.Text("Installation City:"), sg.Input(key="installationCity", do_not_clear=True)],
            [sg.Text("Installation State"), sg.Input(key="installationState", do_not_clear=True)],
            [sg.Text("Installation Zip Code:"), sg.Input(key="installationZipCode", do_not_clear=True)],
            [sg.Text("Billing Street Address:"), sg.Input(key="billingStreetAddress", do_not_clear=True)],
            [sg.Text("Billing City:"), sg.Input(key="billingCity", do_not_clear=True)],
            [sg.Text("Billing State:"), sg.Input(key="billingState", do_not_clear=True)],
            [sg.Text("Billing Zip Code:"), sg.Input(key="billingZipCode", do_not_clear=True)],
            [sg.Text("Payment Method")],
            [sg.CB('Credit Card', key = 'creditCard')],
            [sg.CB('Bank Draft', key = 'bankDraft')],
            [sg.CB('e-Statements', key = 'eStatement')],
            [sg.CB('Paper, $4/mo fee', key = 'paperBilling')],
            [sg.Text("Master Account Username:"), sg.Input(key="masterAccountUsername", do_not_clear=True)],
            [sg.Text("Master Account Password:"), sg.Input(key="masterAccountPassword", do_not_clear=True)],
            ]

layout2 = [[sg.Text('Service Information')],
            [sg.Text('Service Level:')],
            [sg.CB('Fiber 100 $49.95/mo', key = 'fiber100')],
            [sg.CB('Fiber 500 $69.95/mo', key = 'fiber500')],
            [sg.CB('Fiber 1GB $79.95/mo', key = 'fiber1000')],
            [sg.CB('Digital Voice Bundle $19.99/mo', key ='voipService')],
            [sg.CB('Plume Wi-Fi $5/mo:', key='smartWifi')],
            [sg.CB('Wi-Fi Extender $6.95/mo:', key = 'wifiExtender')],
            [sg.Text("Wi-Fi/SSID Name"), sg.Input(key="ssid", do_not_clear=True)],
            [sg.Text("Wi-Fi Password"), sg.Input(key="wifiPSK", do_not_clear=True)],
            [sg.Text("Additional Item 1:"), sg.Input(key="additionalItem1", do_not_clear=True)],
            [sg.Text("Add Item 1 Price:"), sg.Input(key="addItemPrice1", do_not_clear=True)],
            [sg.Text("Additional Item 2:"), sg.Input(key="additionalItem2", do_not_clear=True)],
            [sg.Text("Add Item 2 Price:"), sg.Input(key="addItemPrice2", do_not_clear=True)],
            [sg.Text("Additional Item 3:"), sg.Input(key="additionalItem3", do_not_clear=True)],
            [sg.Text("Add Item 3 Price:"), sg.Input(key="addItemPrice3", do_not_clear=True)],
            ]

layout3 = [[sg.Text('Property Information')]]

# ----------- Create actual layout using Columns and a row of Buttons
layout = [[sg.Column(layout1, key='-COLCustomer Information-'), sg.Column(layout2, visible=False, key='-COLService Information-'), sg.Column(layout3, visible=False, key='-COLProperty Information-')],
          [sg.Button('Customer Information'), sg.Button('Service Information'), sg.Button('Property Information'), sg.Button("Create Form"), sg.Exit()],
        ]

window = sg.Window("New Residential Fiber Form", layout, element_justification="right")

"""while True:
    event, values = window.read()
    print(event, values)
    if event in (None, 'Exit'):
        break
    if event in 'Customer Information Service Information Property Information':
        window[f'-COL{layout}-'].update(visible=False)
        layout = str(event)
        window[f'-COL{layout}-'].update(visible=True)
   """     
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event in 'Customer InformationService InformationProperty Information':
        window[f'-COL{layout}-'].update(visible=False)
        layout = str(event)
        window[f'-COL{layout}-'].update(visible=True)
    elif event == "Create Form": 
        values["intakeDate"] = today.strftime("%m-%d-%Y")
    #service Plan section 
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
            values['fiber500'] =str(' ')
        if values['fiber1000'] == True:
            values['fiber1000'] = str('Fiber 1000 $79.95/mo')
        else:
            values['fiber1000'] = str(' ')
        if values['voipService'] == True:
            values['voipService'] = str('Digital Voice Bundle $19.99/mo')
        else:
            values['voipService'] = str(' ')
        
    #Method of Billing

        if values['creditCard'] == True:
            values['paymentMethod'] = str('Credit Card')
        if values['bankDraft'] == True:
            values['paymentMethod'] = str('Bank Draft (Complete ACH Authorization form)')
        if values['eStatement'] == True:
            values['paymentMethod'] = str('E-Statements')
        if values['paperBilling'] == True:
            values['paymentMethod'] = str('Paper, $4/mo fee')

        # Render the template, save new word document & inform user
        doc.render(values)

        output_path = Path(__file__).parent / f"{values['customerFirstName'] + values['customerLastName']}-FiberAgreement.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")


        
    
    
    """output_path = Path(__file__).parent / f"{values['customerFirstName'] + values['customerLastName']}-FiberAgreement.docx"
    doc.save(output_path)
    sg.popup("File saved", f"File has been saved here: {output_path}")"""

window.close()

""" 
   if event in 'Customer Information':
        window[f'-COLService Information-'].update(visible=False)
        window[f'-COLProperty Information-'].update(visible=False)
        layout = layout1
        window[f'-COLCustomer Information-'].update(visible=True)
    if event in 'Service Information':
        window[f'-COLCustomer Information-'].update(visible=False)
        window[f'-COLProperty Information-'].update(visible=False)
        layout = layout2
        window[f'-COLService Information-'].update(visible=True)
    if event in 'Property Information':
        window[f'-COLCustomer Information-'].update(visible=False)
        window[f'-COLService Information-'].update(visible=False)
        layout = layout3
        window[f'-COLProperty Information-'].update(visible=True)   
"""   