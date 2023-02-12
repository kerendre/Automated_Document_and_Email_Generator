
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


