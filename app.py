while True:
    directory = single_input("Please input the directory where the ProductData.xlsx, MasterInvoice.pdf, and Order files are located.", mask=False, width=400, height=200)
    show_message('Thank you', f'You entered "{directory}". Thank you.')

    directory = directory.strip('"').strip("'")

    if directory[-1] != '\\':
        directory = directory + '\\'
    else:
        directory = directory

    files = os.listdir(directory)

    if 'ProductData.xlsx' in files and 'MasterInvoice.pdf' in files and 'Order1.xlsx' in files and 'Order2.xlsx' in files and 'Order3.xlsx' in files and 'Order4.xlsx' in files:
        break
    else:
        show_message('Error', 'The directory you entered does not contain the correct files. Please try again.')    

print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
descriptions = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')

product_descriptions = {}

for index,row in descriptions.iterrows():
    description = row['Description']
    UnitNum = row['Unit number']
    product_descriptions[UnitNum] = description


prices = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')

product_prices = {}

for i,row in prices.iterrows():
    price = row['Price per unit']
    UnitNum = row['Unit number']
    product_prices[UnitNum] = price


product_descriptions
product_prices

# df = get_pdf('MasterInvoice.pdf')
# df

def get_order_data(invoice_num,filename,directory):
    order = openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet = order['OrderForm']

    CustName = selected_sheet['C14'].value
    CustEmail = selected_sheet['C13'].value
    date = str(selected_sheet['C9'].value)
    date = date.split()[0].split('-')
    OrderDate = f'{date[1]}/{date[2]}/{date[0]}'

    update_dict = {}
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = OrderDate
    update_dict['CustomerAddress1'] = selected_sheet['C16'].value
    update_dict['CustomerName'] = CustName
    update_dict['CustomerCompany'] = selected_sheet['C15'].value
    update_dict['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'

    # Add Item1
    amount0_raw = (selected_sheet['C24'].value * product_prices[selected_sheet['B24'].value])
    update_dict['Price0'] = str(product_prices[selected_sheet['B24'].value])
    update_dict['Quantity0'] = str(selected_sheet['C24'].value)
    update_dict['Amount0'] = f'${amount0_raw:.2f}'
    total_raw = amount0_raw
    update_dict['Total'] = f'${total_raw:.2f}'
    update_dict['Description0'] = product_descriptions[selected_sheet['B24'].value]

    # Add Item2
    if selected_sheet['B25'].value != None:
        amount1_raw = (selected_sheet['C25'].value * product_prices[selected_sheet['B25'].value])
        update_dict['Price1'] = str(product_prices[selected_sheet['B25'].value])
        update_dict['Quantity1'] = str(selected_sheet['C25'].value)
        update_dict['Amount1'] = f'${amount1_raw:.2f}'
        total_raw = amount0_raw + amount1_raw
        update_dict['Total'] = f'${total_raw:.2f}'
        update_dict['Description1'] = product_descriptions[selected_sheet['B25'].value]

    # Add Item3
    if selected_sheet['B26'].value != None:
        amount2_raw = (selected_sheet['C26'].value * product_prices[selected_sheet['B26'].value])
        update_dict['Price2'] = str(product_prices[selected_sheet['B26'].value])
        update_dict['Quantity2'] = str(selected_sheet['C26'].value)
        update_dict['Amount2'] = f'${amount2_raw:.2f}'
        total_raw = amount0_raw + amount1_raw + amount2_raw
        update_dict['Total'] = f'${total_raw:.2f}'
        total_raw = amount0_raw + amount1_raw + amount2_raw
        update_dict['Total'] = f'${total_raw:.2f}'
        update_dict['Description2'] = product_descriptions[selected_sheet['B26'].value]
   
    return update_dict, CustName, CustEmail, OrderDate


def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', 'Invoice ' + str(invoice_num) + '.pdf', order_data)




# Enter your gmail user name and password
gmail_user = "amy1lynn2021@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):

    # List of recipients

    recipient = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

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
    message['To'] = recipient
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+'\\Invoice ' +str(invoice_num)+'.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:

        #get data
        order_data, name, email, date = get_order_data(invoice_num, file, directory)

        #load into pdf template
        create_invoice(invoice_num, order_data, directory)

        #send email
        send_email(invoice_num, name, email, date, directory)

        #adjust invoice number
        invoice_num += 1