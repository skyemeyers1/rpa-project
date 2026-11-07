# INSERT ANSWER FOR QUESTION 1 HERE

files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']
while True:
    try:
        directory = single_input("Please input the directory where the file is located.")
        files_in_dir = os.listdir(directory)
        if directory[-1] != "/":
            directory = directory + "/"
        all_files_exist = 1
        for file in files:
            if file not in files_in_dir:
                all_files_exist = 0
        if all_files_exist == 1:
            break
        else:
            show_message("Error", "Please re-enter the directory and try again.")
    except:
        show_message("Error", "Please re-enter the directory and try again.")
print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices ={}

df = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
for index,row in df.iterrows():
    unit_number = row['Unit number']
    price_per_unit = row['Price per unit']
    product_prices[unit_number]= price_per_unit

df1 = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
for index,row in df1.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    workbook = openpyxl.load_workbook(filename)
    selected_sheet = workbook['OrderForm']

    order_data = {}
    order_data['InvoiceNum'] = invoice_num
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['Date'] = date
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
    name = selected_sheet['C14'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    email = selected_sheet['C13'].value
    order_data['Email'] = email

    unit_number0 = selected_sheet['B24'].value
    unit_number1 = selected_sheet['B25'].value
    unit_number2 = selected_sheet['B26'].value
    Total = 0

    Price0 = product_prices.get(unit_number0)
    order_data['Price0'] = f'${Price0:,.2f}'
    Quantity0 = selected_sheet['C24'].value
    order_data['Quantity0'] = Quantity0
    Amount0 = Price0*Quantity0
    order_data['Amount0'] = f'${Amount0:,.2f}'
    Total = Total + Amount0
    order_data['Description0'] = product_descriptions.get(unit_number0)

    if unit_number1 is not None:
        Price1 = product_prices.get(unit_number1)
        order_data['Price1'] = f'${Price1:,.2f}'
        Quantity1 = selected_sheet['C25'].value
        order_data['Quantity1'] = Quantity1
        Amount1 = Price1*Quantity1
        order_data['Amount1'] = f'${Amount1:,.2f}'
        Total = Total + Amount1
        order_data['Description1'] = product_descriptions.get(unit_number1)

    if unit_number2 is not None:
        Price2 = product_prices.get(unit_number2)
        order_data['Price2'] = f'${Price2:,.2f}'
        Quantity2 = selected_sheet['C26'].value
        order_data['Quantity2'] = Quantity2
        Amount2 = Price2*Quantity2
        order_data['Amount2'] = f'${Amount2:,.2f}'
        Total = Total + Amount2
        order_data['Description2'] = product_descriptions.get(unit_number2)
    
    order_data['Total'] = f'${Total:,.2f}'
    return order_data, name, email, date
    
get_order_data(8515, 'Order1.xlsx', directory)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data = order_data,  directory = directory):
    create_pdf('MasterInvoice.pdf', f'Invoice{invoice_num}.pdf', order_data)
    
create_invoice(8515)

# INSERT ANSWER FOR QUESTION 5 HERE
import smtplib
from email.message import EmailMessage

def send_email(invoice_num, name, email, date, directory):
    gmail_user = "skyem426@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    recipients = [email]
    subject = (f'Your Order from PythonTech Co. - Invoice {invoice_num}')
    name = f'{name}'

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = f'{directory}/Invoice{invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        [order_data, name, email, date] = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1


    
