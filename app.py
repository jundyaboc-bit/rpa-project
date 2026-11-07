import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format
pd.set_option('display.max_columns', None) 
pd.set_option('display.max_rows', None) 


# Question 1
required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input("Please enter the directory")

    try:
        if directory[-1] != '\\':
            directory = directory + '\\'

        filenames = os.listdir(directory)

        missing_files = []
        for file in required_files:
            if file not in filenames:
                missing_files.append(file)

        if len(missing_files) == 0:
            break
        else:
            show_message("Error", f"The following files were not found in that directory: {', '.join(missing_files)}. Please try again.")
    except:
        show_message("Error", "That directory does not exist. Please try again.")


# Question 2
descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')

product_descriptions = {}
for index,row in descriptions_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')

product_prices = {}
for index,row in prices_df.iterrows():
     product_prices[row['Unit number']] = row['Price per unit']

# Question 3
def get_order_data(invoice_num, filename, directory):
    workbook = openpyxl.load_workbook(directory + filename)
    sheet = workbook['OrderForm']

    date = sheet['C9'].value.strftime('%m/%d/%Y')
    email = sheet['C13'].value
    name = sheet['C14'].value

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = sheet['C15'].value
    order_data['CustomerAddress1'] = str(sheet['C16'].value).strip()
    order_data['CustomerAddress2'] = f"{sheet['C17'].value}, {sheet['C18'].value} {sheet['C19'].value}"


    total = 0
    i = 0
    row_num = 24

    while sheet['B' + str(row_num)].value is not None:
        unit_number = sheet['B' + str(row_num)].value
        quantity = sheet['C' + str(row_num)].value
        price = product_prices[unit_number]
        amount = price * quantity
        total = total + amount

        order_data['Description' + str(i)] = product_descriptions[unit_number]
        order_data['Price' + str(i)] = f'${price:,.2f}'
        order_data['Quantity' + str(i)] = str(quantity)
        order_data['Amount' + str(i)] = f'${amount:,.2f}'

        i = i + 1
        row_num = row_num + 1

    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date

# Question 4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)


# Question 5
gmail_user = "jundy.aboc@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice{invoice_num}'


    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co
10 Python Drive
Pleasant Grove, UT 84602
Phone: (801) 555-2372
Email: sales@pythontechco.com"""


    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)
        

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

# Question 6
filenames = os.listdir(directory)

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
