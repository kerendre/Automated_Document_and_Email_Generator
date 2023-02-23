# libraries to be imported
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import personal_setting


# function to send email
def send_mail(to_address, name, company, attached_file, mail_text):
    """
    This function adds a picture to a Word document along with a description and spacing.

    Parameters:
    doc (docx.Document): The Word document to which the picture is added.
    pic_name (str): The file name of the picture.
    pic_description (str): The description to be added under the picture.
    mail_text (str): The body of the email.

    Returns:
    None
    """

    # get the app password from personal setting
    my_gmail_app_password = personal_setting.my_gmail_app_password

    # get the email address to send from
    fromaddr = personal_setting.my_gmail

    # get the email address to send to
    toaddr = to_address

    # instance of MIMEMultipart to store all the details of the email
    msg = MIMEMultipart()

    # storing the senders email address
    msg['From'] = fromaddr

    # storing the receivers email address
    msg['To'] = toaddr

    # storing the subject
    msg['Subject'] = f"Status of Fire Ant Infestation in Greenhouses in Israel for {name} on Behalf of {company}"

    # string to store the body of the mail
    body = f"{name} Shalom \n {mail_text}" \


    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
    filename = attached_file

    # open the attachment file with the given filename
    attachment = open(f"{os.getcwd()}\\{filename}", "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    p.set_payload((attachment).read())

    # encode into base64
    encoders.encode_base64(p)

    # add header to the attachment
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(fromaddr, personal_setting.my_gmail_app_password)

    # Converts the Multipart msg into a string
    text = msg.as_string()

    # sending the mail
    s.sendmail(fromaddr, toaddr, text)

    # terminating the session
    s.quit()

# send_mail()
