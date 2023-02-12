
import docx
import adding_objects_file
import send_mail_tab
import pandas as pd

# Open the Word document
doc = docx.Document("input_word_file.docx")

# Adding a picture to the Word document
adding_objects_file.Add_pic(doc,pic_name="pic1.JPG",pic_description="pic is description of pic1")

# Adding another picture to the Word document
adding_objects_file.Add_pic(doc,pic_name="pic2.PNG",pic_description="pic is description of pic2")

# Reading the data from the excel file into a pandas dataframe, skipping rows 0,1,3 in the input excel file
table1=f"DATABASES/tourism_israel.xlsx"
df = pd.read_excel(table1,sheet_name="2.2.28", skiprows=[0,1,2])

# Reading the contact information from another excel file into a pandas dataframe
table2=f"DATABASES/דף_קשר.xlsx"
df_contacts = pd.read_excel(table2,sheet_name="אנשי קשר", skiprows=[0])

# Adding the dataframe as a table to the Word document
adding_objects_file.Add_table(doc,df,subtitle_to_the_table="this is subtitle")

# Saving the changes to the Word document
try:
    # Try to save the file as "doc1.docx"
    doc_to_attach=doc.save("doc1.docx")
except PermissionError:
    # If there is a permission error, save the file as "doc2.docx"
    doc_to_attach=doc.save("doc2.docx")

# Convert the df_contacts dataframe to a list of dictionaries and send emails to each person
dict_list = df_contacts.to_dict('records')
for num in range(0, len(dict_list)):
    dict = (dict_list[num])
    name = dict_list[num][' שם']
    mail = dict_list[num]['מייל']
    company_name = dict_list[num]['חברה']

    # Call the send_mail function to send an email to each person in the list
    send_mail_tab.send_mail(mail,name,company_name,"doc1.docx")



