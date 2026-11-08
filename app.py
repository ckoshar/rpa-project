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
    directory = single_input('Please enter the folder that contains ProductData.xlsx, MasterInvoice.pdf, and Order1-4.xlsx.', width=400)

    if directory.endswith('/') == False:
        directory = directory + '/'

    if os.path.isdir(directory):
        filenames = os.listdir(directory)
        if ('ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and
            'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and
            'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames):
            break

    show_message('Error', 'That folder does not contain the required files. Please try again.', width=400)

print(directory)


# INSERT ANSWER FOR QUESTION 2 HERE

descriptions = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Descriptions')

product_descriptions = {}
for index,row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Prices')

product_prices = {}
for index,row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

print(product_descriptions)
print(product_prices)


df = get_pdf(directory+'MasterInvoice.pdf')
print(df)

def get_order_data(invoice_num, filename, directory):
    order_form = openpyxl.load_workbook(directory+filename)

    selected_sheet = order_form['OrderForm']

    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    email = selected_sheet['C13'].value
    name = selected_sheet['C14'].value

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}"

    total = 0

    if pd.isna(selected_sheet['B24'].value) == False:
        unit_number = selected_sheet['B24'].value
        quantity = selected_sheet['C24'].value
        price = product_prices[unit_number]
        order_data['Description0'] = product_descriptions[unit_number]
        order_data['Price0'] = f'${price:,.2f}'
        order_data['Quantity0'] = str(quantity)
        order_data['Amount0'] = f'${price * quantity:,.2f}'
        total = total + price * quantity

    if pd.isna(selected_sheet['B25'].value) == False:
        unit_number = selected_sheet['B25'].value
        quantity = selected_sheet['C25'].value
        price = product_prices[unit_number]
        order_data['Description1'] = product_descriptions[unit_number]
        order_data['Price1'] = f'${price:,.2f}'
        order_data['Quantity1'] = str(quantity)
        order_data['Amount1'] = f'${price * quantity:,.2f}'
        total = total + price * quantity

    if pd.isna(selected_sheet['B26'].value) == False:
        unit_number = selected_sheet['B26'].value
        quantity = selected_sheet['C26'].value
        price = product_prices[unit_number]
        order_data['Description2'] = product_descriptions[unit_number]
        order_data['Price2'] = f'${price:,.2f}'
        order_data['Quantity2'] = str(quantity)
        order_data['Amount2'] = f'${price * quantity:,.2f}'
        total = total + price * quantity

    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)


# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory):
    create_pdf(f'{directory}MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)


# INSERT ANSWER FOR QUESTION 5 HERE

gmail_user = "carterkoshar@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):

    recipients = [email]
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = f'{directory}Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'carterkoshar@gmail.com', '11/13/2023', directory)


# INSERT ANSWER FOR QUESTION 6 HERE

filenames = os.listdir(directory)
filenames.sort() 
print(filenames)


invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
