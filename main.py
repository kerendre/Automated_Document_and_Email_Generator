

import docx
import adding_objects_file
import send_mail_tab
import pandas as pd

# Open the Word document
#doc = docx.Document("input_word_file.docx")
# Create an instance of a word document
# after i have a word to change this part
doc = docx.Document()

# Adding a picture to the Word document
adding_objects_file.Add_pic(doc, pic_name="pic1.JPG", pic_description="pic is description of pic1")

# Adding another picture to the Word document
adding_objects_file.Add_pic(doc, pic_name="pic2.PNG", pic_description="pic is description of pic2")

# Reading the data from the excel file into a pandas dataframe, skipping rows 0,1,2 in the input excel file
table1 = f"DATABASES/tables_style.xlsx"
df = pd.read_excel(table1, sheet_name="Table 8", skiprows=[0, 1])

# Reading the contact information from another excel file into a pandas dataframe
table2 = f"DATABASES/דף_קשר.xlsx"
df_contacts = pd.read_excel(table2, sheet_name="אנשי קשר", skiprows=[0])

# Adding the dataframe as a table to the Word document
doc =adding_objects_file.Add_table(doc, df, subtitle_to_the_table="this is subtitle")

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
