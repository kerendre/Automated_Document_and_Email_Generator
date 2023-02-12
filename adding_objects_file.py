import docx
import os
import pandas as pd


def Add_pic(doc, pic_name, pic_description):
    '''
    Add a picture to the given docx document.

    Parameters:
        doc (docx.document.Document): The docx document object.
        pic_name (str): The name of the picture file.
        pic_description (str): The description to be added under the picture.

    Returns:
        None
    '''
    # Get the page width of the document
    page_width = doc.sections[0].page_width

    # Add the picture to the Word document
    pic1 = doc.add_picture(f"PICS/{pic_name}")

    # Fit the width of the first picture to 65% of the page width
    pic1.width = int(page_width * 0.65)

    # Add a description under the picture
    description_pic = doc.add_paragraph(pic_description)

    # Center the description
    description_pic.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # Add two lines of space
    doc.add_paragraph("\n")


def Add_table(doc, df, subtitle_to_the_table):
    """
    Add a table to the given docx document.

    Parameters:
        doc (docx.document.Document): The docx document object.
        df (pandas.DataFrame): The pandas dataframe to be added as a table in the document.
        subtitle_to_the_table (str): The title or subtitle for the table.

    Returns:
        None
    """

    # Add the title/subtitle to the table
    paragraph = doc.add_paragraph(subtitle_to_the_table)
    paragraph.style = doc.styles['Heading 1']
    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # Replace NaN values in the dataframe with an empty string
    df = df.fillna('')

    # Add the dataframe to the document
    doc.add_table(df.shape[0] + 1, df.shape[1])

    # Add the header row
    for j in range(df.shape[-1]):
        header_cell = df.columns[j] if not pd.isna(df.columns[j]) else " "
        # Replace header cell if it contains "Unnamed:" with an empty string
        if "Unnamed:" in header_cell:
            header_cell = " "

        # Add the header cell to the table
        doc.tables[0].cell(0, j).text = header_cell
        # Make the header cell bold
        doc.tables[0].cell(0, j).paragraphs[0].runs[0].bold = True

    # Add the data to the table
    for i in range(df.shape[0]):
        for j in range(df.shape[-1]):
            # Add the cell data to the table
            doc.tables[0].cell(i + 1, j).text = str(df.values[i, j])

