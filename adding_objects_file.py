
import docx
import os
import pandas as pd

def Add_pic(doc,pic_name,pic_description):
    ''' Add pic to document'''



    # Get the page width
    page_width = doc.sections[0].page_width

    # Add the picture to the Word document
    pic1 = doc.add_picture(f"PICS/{pic_name}")

    # Fit the width of the first picture to the page width
    pic1.width = int(page_width * 0.65)

    # Add a description under the picture pic1
    description_pic= doc.add_paragraph(pic_description)

    # Center the description
    description_pic.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # Add two lines of space
    doc.add_paragraph("\n\n")


def Add_table(doc, df, subtitle_to_the_table):
    # Replace NaN values with an empty string
    df = df.fillna('')

    # Check if the dataframe is empty
    if df.empty:
        return

    paragraph = doc.add_paragraph(subtitle_to_the_table)
    paragraph.style = doc.styles['Heading 1']
    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # Create the table
    table = doc.add_table(rows=1, cols=len(df.columns))

    # Define the header row
    header_row = table.rows[0]
    header_row.style = 'Table Header'
    for i, column_title in enumerate(df.columns):
        header_row.cells[i].text = column_title

    # Add the data to the table
    for i, row in df.iterrows():
        if row.isnull().all():
            continue

        new_row = table.add_row()
        for j, cell_value in enumerate(row):
            new_row.cells[j].text = str(cell_value)




from docx import Document

#
# def Add_table(doc, df, subtitle_to_the_table):
#     # Replace NaN values with an empty string
#     df = df.fillna('')
#
#     paragraph = doc.add_paragraph(subtitle_to_the_table)
#     paragraph.style = doc.styles['Heading 1']
#     paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER



    #
    # # Add the header row
    # #hdr_cells = table.rows[0].cells
    # for header_index, header in enumerate(df.columns):
    #     hdr_cells[header_index].text = header

    # Add the data rows
    #for row_index, row in df.iterrows():
        # data_cells = table.rows[row_index + 1].cells
        # if row.empty:
        # #     continue
        # for data_index, data in enumerate(row):
        #     data_cells[data_index].text = str(data)


        #doc.add_table(table)

    #
    # table = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
    # hdr_cells = table.rows[0].cells
    # for i, column in enumerate(df.columns):
    #     hdr_cells[i].text = column
    # for i, row in df.iterrows():
    #     cells = table.add_row().cells
    #
    #
    #     for j, column in enumerate(df.columns):
    #         cells[j].text = str(row[column])
    #         if all(pd.isna(row)):
    #             continue  # skip this iteration if the row is empty
    #
    # doc.add_paragraph(subtitle_to_the_table)
    # #doc.add_table(table)
    #




# def Add_table(doc,df, subtitle_to_the_table):
#     print("df,shape", (df.shape))
#     # Add the title to the table
#
#     paragraph = doc.add_paragraph(subtitle_to_the_table)
#     paragraph.style = doc.styles['Heading 1']
#     paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
#
#
#
#     # Replace NaN values with an empty string
#     df = df.fillna('')
#
#
#     # Add the dataframe to the document
#     doc.add_table(df.shape[0] + 1, df.shape[1])
#
#     # Add the header row
#     for j in range(df.shape[-1]):
#         header_cell = df.columns[j] if not pd.isna(df.columns[j]) else " "
#         if "Unnamed:" in header_cell:
#             header_cell = " "
#
#         doc.tables[0].cell(0, j).text = header_cell
#         doc.tables[0].cell(0, j).paragraphs[0].runs[0].bold = True
#
#
#     ##
#     print("(df.shape[0])",df.shape[0])
#     print("df.shape[-1]",df.shape[1])
#
#     # Add the data
#     for i in range(df.shape[0]):
#         for j in range(df.shape[-1]):
#
#             print("i,j, str(df.values[i, j])", i,j, str(df.values[i, j]))
#             doc.tables[0].cell(i + 1, j).text = str(df.values[i, j])
#
#







# # #!~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# #
# # import docx
# # import os
# # import pandas as pd
# #
# #
# # def Add_pic(doc, pic_name, pic_description):
# #     '''
# #     Add a picture to the given docx document.
# #
# #     Parameters:
# #         doc (docx.document.Document): The docx document object.
# #         pic_name (str): The name of the picture file.
# #         pic_description (str): The description to be added under the picture.
# #
# #     Returns:
# #         None
# #     '''
# #     # Get the page width of the document
# #     page_width = doc.sections[0].page_width
# #
# #     # Add the picture to the Word document
# #     pic1 = doc.add_picture(f"PICS/{pic_name}")
# #
# #     # Fit the width of the first picture to 65% of the page width
# #     pic1.width = int(page_width * 0.65)
# #
# #     # Add a description under the picture
# #     description_pic = doc.add_paragraph(pic_description)
# #
# #     # Center the description
# #     description_pic.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
# #
# #     # Add two lines of space
# #     doc.add_paragraph("\n")
# #
# #
# # def Add_table(doc, df, subtitle_to_the_table):
# #     """
# #     Add a table to the given docx document.
# #
# #     Parameters:
# #         doc (docx.document.Document): The docx document object.
# #         df (pandas.DataFrame): The pandas dataframe to be added as a table in the document.
# #         subtitle_to_the_table (str): The title or subtitle for the table.
# #
# #     Returns:
# #         None
# #     """
# #
# #     # Add the title/subtitle to the table
# #     paragraph = doc.add_paragraph(subtitle_to_the_table)
# #     paragraph.style = doc.styles['Heading 1']
# #     paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
# #
# #     # Replace NaN values in the dataframe with an empty string
# #     df = df.fillna('')
# #
# #     # Add the dataframe to the document
# #     doc.add_table(df.shape[0] + 1, df.shape[1])
# #
# #     # Add the header row
# #     for j in range(df.shape[-1]):
# #         header_cell = df.columns[j] if not pd.isna(df.columns[j]) else " "
# #         # Replace header cell if it contains "Unnamed:" with an empty string
# #         if "Unnamed:" in header_cell:
# #             header_cell = " "
# #
# #         # Add the header cell to the table
# #         doc.tables[0].cell(0, j).text = header_cell
# #         # Make the header cell bold
# #         doc.tables[0].cell(0, j).paragraphs[0].runs[0].bold = True
# #
# #     # Add the data to the table
# #     for i in range(df.shape[0]):
# #         for j in range(df.shape[-1]):
# #
# #   
# #
# #             print(i,j)
# #
# #             #!~~~~~~~~~~~~~~~~~
# #             # Add the cell data to the table
# #             doc.tables[0].cell(i + 1, j).text = str(df.values[i, j])
# #             print(doc.tables[0].cell(i + 1, j).text)
# #
# #
