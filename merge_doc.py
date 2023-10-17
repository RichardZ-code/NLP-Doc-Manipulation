# Import the libraries
import os.path
from docxcompose.composer import Composer
from docx import Document as Document_compose


def merge_all_docx(doc_list):
    """""
    This function merges all Word documents in [doc_list] into
    'merged_doc.docx'
    """""


    # Count the number of Word documents in doc_list and call the composer
    number_of_docx = len(doc_list)
    master = Document_compose('merged_doc.docx')
    composer = Composer(master)


    """"
    Loop from the first Word document to the last Word document in doc_list, 
    and append each one of them to 'merged_doc.docx'
    """
    for i in range(0, number_of_docx):
        doc_temp = Document_compose(doc_list[i])
        composer.append(doc_temp)


    # Save 'merged_doc.docx'
    composer.save('merged_doc.docx')


# Create and save 'merged_doc.docx'
merged_doc = Document_compose()
merged_doc.save('merged_doc.docx')


# Example
merge_all_docx(["11D.docx", "12D.docx", "13D.docx", "14D.docx", "15D.docx"])