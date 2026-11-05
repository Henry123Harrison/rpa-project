# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE

while True:
    directory = single_input("Please input the directory where the `ProductData.xlsx`, `MasterInvoice.pdf`, and `Order` (e.g., `Order1.xlsx`, `Order2.xlsx`, etc.) files are located.")
    if directory[-1] != '\\':
        directory = directory+'\\' 
        try: 
            filenames = os.listdir(directory)
            if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
                break
            else:
                show_message("Error", "The directory does not have the right files. Try again.")
        except:
            show_message("Error", "The directory does not work. Try again.")
    else:   
        try: 
            filenames = os.listdir(directory)
            if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
                break
            else:
                show_message("Error", "The directory does not have the right files. Try again.")
        except:
            show_message("Error", "The directory does not work. Try again.")


# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

df1 = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions') 

for index,row in df1.iterrows():
    item = row.Description
    number = row['Unit number']
    product_descriptions[number] = item
                         
print(product_descriptions)

df2 = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices') 

for index,row in df2.iterrows():
    item = row['Price per unit']
    number = row['Unit number']
    product_prices[number] = item
                         
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet = template['OrderForm']
    countb = 24
    countc = 24
    countx = 0
    dictionary = {}
    Total=0
    while True:
        if selected_sheet[f'B{countb}'].value != None:
            key = selected_sheet[f'B{countb}'].value
            value = selected_sheet[f'C{countb}'].value
            dictionary[f'Description{countx}']=product_descriptions[key]
            dictionary[f'Quantity{countx}']=str(value)
            pricestr = f'${product_prices[key]:,.2f}'
            dictionary[f'Price{countx}']=pricestr
            num = value*product_prices[key]
            numstr = f'${num:,.2f}'
            dictionary[f'Amount{countx}']=numstr
            countb = countb+1
            countc = countc+1
            countx = countx+1
            Total = Total + num
            Totalstr = f'${Total:,.2f}'
            dictionary['Total'] = Totalstr
        else:
            break
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zip = selected_sheet['C19'].value
    compname = selected_sheet['C15'].value

    dictionary['Date'] = date
    dictionary['CustomerAddress1'] = f'{address} {city} {state} {zip}'
    dictionary['CustomerName']=name
    dictionary['CustomerCompany']= compname
    dictionary['InvoiceNum']= invoice_num     

    return (dictionary, name, email, date)

# dictionary, name, email, date  = get_order_data(8515,'Order1.xlsx', directory)
# print(dictionary, name, email, date)


# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice{invoice_num}.pdf', order_data)

# create_invoice(8515,dictionary, directory)


# INSERT ANSWER FOR QUESTION 5 HERE


gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email (invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "henryfharrison14@gmail.com"

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f"Your Order from PythonTech Co - Invoice {invoice_num}"

    # Body of the email

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = f'{directory}Invoice{invoice_num}.pdf'

    message = add_attachment(path_to_file1, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

# send_email(8515, name, 'henryfh@byu.edu', date, directory)

# INSERT ANSWER FOR QUESTION 6 HERE
count = 0
filenames = os.listdir(directory)
for i, file in enumerate(filenames):
    if 'Order' in file and 'xlsx' in file:
        invoice_num = 8515+count
        order_data, name, email, date = get_order_data(invoice_num,file, directory)
        count = count+1
        create_invoice(invoice_num, order_data, directory)
        send_email (invoice_num, name, email, date, directory)

