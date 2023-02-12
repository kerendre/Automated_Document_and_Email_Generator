import os

# The receiving address, can be of anY mail provider
send_to_address= "keren.drev@gmail.com"

# sending mail with smtp, in order to use gmail via (mail_brain), you may need to create and use app passwords and
# use 2-Step-Verification. follow the instruction: https://support.google.com/mail/answer/185833?hl=en-GB


my_gmail = os.environ.get("MY_GMAIL_APP_ACCOUNT")
my_gmail_app_password = os.environ.get("MY_GMAIL_APP_PASSWORD")