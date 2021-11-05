# Email via MS exchangelib
# Tested using MS Exchange 2019 server
# Provided by Paul E Davey
# Learn more at https://automationpro.co.uk
#
# Required Dependencies
# exchangelib==4.6.0
# python-dateutil
#
# Default inputs
# Type Name Value
# Action Constant smtpServer The FQDN to the exchange server
# Action Constant smtpSender Email address of the sender (who we are sending an email from)
# Action Constant smtpLoginUsername The username to use to connect to the SMTP / Exchange server with
# Action Constant smtpLoginPassword Password for the smtpLoginUsername account

def handler(context, inputs):
from exchangelib import DELEGATE, Configuration, Message, Mailbox, Credentials, Account, EWSTimeZone, UTC_NOW
from dateutil.tz import gettz
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

# Tell exchangelib to use this adapter class instead of the default
# In my lab environment I do not have a valid HTTPS certificate installed on my MS Exchange 2019 server
# The following line works around this issue by ignoring the certification validation
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

# Credentials to connect to the SMTP server with
credentials = Credentials(
username = inputs["smtpLoginUsername"],
password = context.getSecret(inputs["smtpLoginPassword"])
)

# Configuration for the SMTP connection
config = Configuration(
server=inputs["smtpServer"],
credentials=credentials
)

# Account settings for the SMTP connection
acc = Account(
primary_smtp_address = inputs["smtpSender"],
config = config,
autodiscover = False,
access_type = DELEGATE,
default_timezone = EWSTimeZone.timezone('/'.join(gettz('Europe/London')._filename.split('/')[-2:]))
)

# Construct the email message (sibject and body) that you wish to send
msg = Message(
account=acc,
folder=acc.sent,
subject='E-mail from vRA ABX Action',
body='This email was produced by an ABX action and sent to you via the specified MS Exchange Server',
to_recipients=[Mailbox(email_address='administrator@automationpro.lan')]
)

# Send the message
msg.send_and_save()