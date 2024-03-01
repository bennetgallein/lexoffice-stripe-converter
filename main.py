import csv
import stripe
from stripe.error import InvalidRequestError
from datetime import datetime, timedelta, timezone
from dateutil.relativedelta import relativedelta
import time
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
import requests
from dotenv import load_dotenv
from os import getenv
import pytz

load_dotenv()

# Define your stripe key here
STRIPE_KEY = getenv('STRIPE_KEY')
STRIPE_NAME = 'Stripe Technology Europe, Limited'


def send_mail(send_from, send_to, subject, message, files=[],
              server="localhost", port=587, username='', password='',
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


def get_unix_timestamp(dt):
    return int(dt.timestamp())


def get_first_and_last_day_unix_timestamps():
    # Get the current date in UTC
    today_utc = datetime.now(pytz.utc)

    # Get the CEST timezone
    cest_timezone = pytz.timezone('Europe/Paris')

    # Convert current date to CEST
    today_cest = today_utc.astimezone(cest_timezone)

    # Calculate the first day of the previous month
    first_day_previous_month = today_cest.replace(day=1) - timedelta(days=1)
    first_day_previous_month = first_day_previous_month.replace(day=1)

    # Calculate the last day of the previous month
    last_day_previous_month = today_cest.replace(day=1) - timedelta(days=1)

    # Convert the dates to UTC
    first_day_previous_month_utc = first_day_previous_month.astimezone(
        pytz.utc)
    last_day_previous_month_utc = last_day_previous_month.astimezone(pytz.utc)

    # Get the Unix timestamps
    first_day_timestamp = get_unix_timestamp(first_day_previous_month_utc)
    last_day_timestamp = get_unix_timestamp(
        last_day_previous_month_utc.replace(hour=23, minute=59, second=59))

    return first_day_timestamp, last_day_timestamp


def get_client():
    """
    This method sets the api key & return the client
    :return: stripe
    """
    stripe.api_key = STRIPE_KEY
    return stripe


def csv_header():
    """
    This method only returns the csv header for our export
    :return: array
    """
    return [
        # 'id',
        # 'type',
        # 'source',
        'Betrag',
        'Auftraggeber/Empfänger',
        'Buchungsdatum',
        'Wertstellungsdatum',
        'Verwendungszweck',
    ]


def getCustomerByPayment(payment_id: str):
    """
    This method tries to fetch the Customer name from the payment
    if it does not succeed, it returns the stripe name (happens e.g. when there is a chargeback from a customer)
    :param payment_id: => source id in import.csv
    :return:
    """
    try:
        charge = get_client().Charge.retrieve(payment_id)
        if charge['billing_details']['name']:
            return charge['billing_details']['name']

        # load name from customer
        customer = get_client().Customer.retrieve(charge.customer)
        return customer.name
    except InvalidRequestError:
        return STRIPE_NAME


def getDescription(payment_id: str):
    """
    This method fetches the description of the payment_intend based on the charge
    """
    try:
        charge = get_client().Charge.retrieve(payment_id)
        sessions = get_client().checkout.Session.list(
            payment_intent=charge.payment_intent)
        if len(sessions.data) > 0:
            line_items = get_client().checkout.Session.list_line_items(
                sessions.data[0].id)
            return ' + '.join(map((lambda e: e.description), line_items.data))
        return ""
    except InvalidRequestError:
        return ""


def read_csv():
    """
    This method returns all csv lines in import.csv & drops the header.
    :return: CSV Lines in array from import.csv
    """
    csvlines = []
    with open('import.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)

        # Appending rows to csvlines
        for row in reader:
            csvlines.append(row)

    # We remove the header
    csvlines.pop(0)

    return csvlines


def toMoney(am: int):
    """
    Convert the amount part to float and to an easy format, so lexoffice does not get in trouble
    :param am:
    :return:
    """
    return str(am / 100).replace(".", ',')


# Run the script
if __name__ == '__main__':
    everhypeCSV = []
    attachments = []

    first_day_timestamp, last_day_timestamp = get_first_and_last_day_unix_timestamps()

    transactions = get_client().BalanceTransaction.list(
        created={"gte": first_day_timestamp, "lte": last_day_timestamp}, limit=100)

    print(first_day_timestamp, last_day_timestamp)
    print(f'Found {len(transactions.data)} transactions...')

    for line in transactions.data:

        customer = getCustomerByPayment(line.source)

        # --> description
        description = line.description
        if description == "" or description == None:
            description = getDescription(line.source)
        if description == "STRIPE PAYOUT":
            description = "Auszahlung auf Bankkonto"

        if description.startswith('Invoice'):
            # load invoice pdf URL to payment
            charge = get_client().Charge.retrieve(line.source)
            if charge.invoice:
                invoice = get_client().Invoice.retrieve(charge.invoice)
                pdf = requests.get(invoice.invoice_pdf)
                with open(f'invoices/{invoice.id}.pdf', 'wb') as invoice_file:
                    invoice_file.write(pdf.content)
                attachments.append(f'invoices/{invoice.id}.pdf')
                # print(invoice)

        created = datetime.fromtimestamp(
            line.created).strftime('%d.%m.%Y %H:%M:%S')
        available_on = datetime.fromtimestamp(
            line.available_on).strftime('%d.%m.%Y %H:%M:%S')

        everhypeCSV.append([
            # line.id,
            # line.type,
            # line.source,
            toMoney(line.amount or line.fee),
            customer,
            created,
            available_on,
            description
        ])

        # If there are fees, we are generating a new line
        if line.fee != 0:
            everhypeCSV.append([
                # line.id + '_fee',
                # 'Kontoführungsgebühr',
                # line.source + '_fee',
                toMoney(line.fee * -1),
                STRIPE_NAME,
                created,
                available_on,
                f'Gebühren für Zahlung {line.id} -- {description}'
            ])

    # writing to export.csv
    today = datetime.now().date()
    file_name = f'csvs/export-{today.strftime("%Y")}-{today.strftime("%m")}.csv'
    with open(file_name, 'w', newline='', encoding='utf-8') as exportFile:
        writer = csv.writer(exportFile, delimiter=';')

        writer.writerow(csv_header())
        writer.writerows(everhypeCSV)

    attachments.append(file_name)

    send_mail(
        send_from=getenv('MAIL_FROM'),
        send_to=getenv("MAIL_TO"),
        subject=f'Stripe export {file_name} is ready',
        message='Hello,\n\nplease check attached files.',
        files=attachments,
        server=getenv('MAIL_SERVER'),
        username=getenv("MAIL_USER"),
        password=getenv("MAIL_PASSWORD")
    )

    print(
        f'Send email to {getenv("MAIL_TO")} containing all invoices and csv ready for import')
