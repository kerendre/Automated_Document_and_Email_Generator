# Import docx NOT python-docx
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

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

    # Adding a page break
    doc.add_page_break()




def Add_table(doc, df, subtitle_to_the_table):

    # Replace NaN values with an empty string
    df = df.fillna('')

    # # Create an instance of a word document
    # # after i have a word to change this part
    # doc = docx.Document()

    # Add a Title to the document
    for num in range(0,9):
        heading =doc.add_heading(subtitle_to_the_table, num)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Creating a table object
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])

    row = table.rows[0].cells
    for i in range(df.shape[1]):
        row[i].text = df.columns[i]

    for i in range(df.shape[1]):
        row[i].text = df.columns[i]
    for row in range(df.shape[0]):
        if not df.iloc[row].isnull().all():
            row_cells = table.rows[row + 1].cells
            for col in range(df.shape[1]):
                if isinstance(df.iloc[row, col], float):
                    row_cells[col].text = f"{df.iloc[row, col]:.2f}"
                else:
                    row_cells[col].text = str(df.iloc[row, col])

    # Adding style to a table
    table.style = 'Colorful List'

    return doc

    # # Now save the document to a location
    # doc.save('gfg.docx')
