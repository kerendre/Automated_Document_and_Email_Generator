Automated Document and Email Generator
Overview

This project contains a set of python scripts that automates the process of creating a Microsoft Word document and sending emails. 
The scripts use the docx and pandas libraries to generate a Word document and the smtplib library to send emails.
Requirements

The following packages need to be installed in order to run these scripts:

    docx
    pandas
    smtplib

Input files

The project uses the following input files:

    A Microsoft Word document named "input_word_file.docx"
    An Excel file named "tourism_israel.xlsx" which contains data to be added to the Word document
    An Excel file named "דף_קשר.xlsx" which contains information about the recipients of the emails

Scripts

    adding_objects_file.py: This script contains functions to add images and tables to a Microsoft Word document.

    send_mail_tab.py: This script contains a function to send emails with attachments.

    main.py: This is the main script that calls the functions from the other scripts to create a Word document and send emails.

How to run

    Clone the repository to your local machine.
    Ensure that you have all the required packages installed.
    Place the input files in the same directory as the scripts.
    Run the main.py script.

Note

The email sending function in the send_mail_tab.py script uses Gmail to send emails. In order to use it, 
you need to provide your Gmail username and Gmail app password in a file named personal_setting.py.

Conclusion

This project provides a streamlined way to create a Word document and send emails. 
It can be easily customized to meet specific needs by modifying the functions in the scripts.
