

import docx
import adding_objects_file
import send_mail_tab
import pandas as pd

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

pic1="pic1.JPG"
pic1_description ="pic1 is description of pic1"

pic2="pic2.PNG"
pic2_description="pic2_description"

table1_origin_location_n_name="DATABASES/input_xlsx_file.xlsx" # tourism_israel.xlsx
table1_sheet_name="Sheet1"
table1_number_of_rows_2_skip=[]
subtitle_to_the_table="This is subtitle to  the table "


#table2_origin_location_n_name="DATABASES/דף_קשר.xlsx"
#table2_sheet_name="אנשי קשר"
table2_number_of_rows_2_skip=[0]
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Open the Word document
#doc = docx.Document("input_word_file.docx")

# Create an instance of a word document
# after i have a word to change this part
doc = docx.Document()

# Adding a picture to the Word document
adding_objects_file.Add_pic(doc, pic_name=pic1, pic_description=pic1_description)

# Adding another picture to the Word document
adding_objects_file.Add_pic(doc, pic_name=pic2, pic_description=pic2_description)

# Reading the data from the excel file into a pandas dataframe, skipping rows 0,1,2 in the input excel file
df = pd.read_excel(table1_origin_location_n_name, sheet_name=table1_sheet_name, skiprows=table1_number_of_rows_2_skip)




# Reading the contact information from another excel file into a pandas dataframe
df_contacts = pd.read_excel("DATABASES/דף_קשר.xlsx", sheet_name="אנשי קשר", skiprows=table2_number_of_rows_2_skip)

# Adding the dataframe as a table to the Word document
doc =adding_objects_file.Add_table(doc, df, subtitle_to_the_table=subtitle_to_the_table)

# add a hyperlink to the header
adding_objects_file.add_hyperlink_to_header(doc, 'https://www.linkedin.com/in/keren-drevin/', 'Linkdein/Keren Drevin')
adding_objects_file.add_hyperlink_to_header(doc, 'https://github.com/kerendre', '                                                                              GitHub/Keren Drevin')
adding_objects_file.add_hyperlink_to_header(doc, 'https://www.linkedin.com/in/keren-drevin/', "\n ________________________________________________________________________________________________________")


# Try to save the file as "doc1.docx"
doc_to_attach = doc.save("doc1.docx")


# Convert the df_contacts dataframe to a list of dictionaries and send emails to each person
dict_list = df_contacts.to_dict('records')
for num in range(0, len(dict_list)):
    dict = (dict_list[num])
    name = dict_list[num][' שם']
    mail = dict_list[num]['מייל']
    company_name = dict_list[num]['חברה']

    # Call the send_mail function to send an email to each person in the list
    #send_mail_tab.send_mail(mail, name, company_name, "doc1.docx")
