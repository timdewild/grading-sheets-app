# Written by Tim de Wild, April 2023, University of Groningen, the Netherlands

import streamlit as st
import pandas as pd
import numpy as np
import io
import docx
from docx.shared import Pt
import zipfile
import xlsxwriter # internal dependency openpyxl


#################################################### GLOBAL VARIABLES GO HERE ####################################################

alphabet = {1:'a', 2:'b', 3:'c', 4:'d', 5:'e', 6:'f', 7:'g', 8:'h', 9:'i', 10:'j', 11:'k', 12:'l', 13:'m', 14:'n', 15:'o', 16:'p', 17:'q', 18:'r', 19:'s', 20:'t', 21:'u', 22:'v', 23:'w', 24:'x', 25:'y', 26:'z'} 

#################################################### FUNCTIONS GO HERE ####################################################

def sample_snumbers():

    # generate dataframe

    df = pd.DataFrame(
        {'S-number': [1234567, 2345678, 3456789]}
    )

    buf = io.BytesIO()
    writer = pd.ExcelWriter(buf, engine='xlsxwriter')
    df.to_excel(writer, sheet_name = 'Sheet1', index = False)

    writer.save()

    return buf

def Qindex(subq):
    """
    Returns list with labels of questions 'qlist' based on list with subquestions 'subq'. 
    Example: if subq = [2,2] then qlist = ['1a', '1b', '2a', '2b']. 
    """

    qlist = []

    for n, nsq in enumerate(subq):  #loop over all questions
        if nsq == 0:
            qlist.append(str(n+1))
        else:
            for nn in range(nsq):
                qlist.append(str(n+1)+alphabet[nn+1]) #add corresponding index to qlist

    return qlist

def split_dataframe(df, nTA, qlist):
    """
    Split dataframe df with S-numbers into nTA dataframes and add questions (qlist) as column headers. 
    Returns df_split, a numpy array with splitted dataframes. 
    """

    # split dataframe based on nTA
    df_split = np.array_split(df, nTA)

    # add columns with subquestions
    for df in df_split:
        for n in range(len(qlist)):
                df[qlist[n]]=''

    return df_split

def grading_labels_generator(df_split):
    """
    Generates Word document, stored in buffer (instance of io.BytesIO), containing S-number ranges. 
    The object df_split = np.split_array(df, nTA), the array with different dataframes for the TAs. 
    """

    # define buffer
    buffer_gl = io.BytesIO()

    # define doc with labels
    labels_doc = docx.Document()
    style = labels_doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(42)

    # add grading labels
    for i, df in enumerate(df_split):
        
        # label values
        lab_1 = round(df['S-number'][df.index[0]])
        lab_2 = round(df['S-number'][df.index[-1]])

        paragraph = labels_doc.add_paragraph(str(i+1)+'. S'+str(lab_1)+' - S'+str(lab_2))
        paragraph.style = labels_doc.styles['Normal']
        labels_doc.add_page_break()

    # save doc to buffer
    labels_doc.save(buffer_gl)

    return buffer_gl

def grading_sheets_generator(df_split):
    """
    Generates zip-file, containing grading sheets for all the TAs. 
    The object df_split = np.split_array(df, nTA), the array with different dataframes for the TAs. nTA is the number of TAs.
    """

    # generate grading sheets
    files_gs = []

    for i, df in enumerate(df_split):
        buffer_gs = io.BytesIO()
        writer = pd.ExcelWriter(buffer_gs, engine='xlsxwriter')

        df.to_excel(writer, sheet_name='Sheet1', index = False)

        # add borders to all cells
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df), len(df.columns)-1), {'type': 'no_errors', 'format': border_fmt})

        writer.save()

        files_gs.append(buffer_gs)
    
    # generate zip file
    buffer_zip = io.BytesIO()

    with zipfile.ZipFile(file=buffer_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as z:
        for i, gs in enumerate(files_gs):
            filename = f'grading_sheet_{i+1}.xlsx'
            z.writestr( zinfo_or_arcname = filename, data = gs.getvalue() )

    return buffer_zip

#################################################### APP STARTS HERE ####################################################

st.write("""
# Grading Sheets Generator

This app allows you to generate grading sheets for your TAs and hand-in labels. All you have to provide is an Excel sheet with the student numbers of the students attending your course. Upload your sheet with student numbers below:
""")

S_file = st.file_uploader("", type=['xls','xlsx'])

with st.expander("Sheet Format Details"):
    st.write("""
        ##### Sheet Format
        The Excel sheet with student numbers can be easily made given the official course list. It should have the following format:
    """)
    col1, col2 = st.columns(2, gap="medium")

    with col1:
        format = pd.DataFrame(
        ['S-number', 1234567, 2345678, 3456789],
        columns = ['A']
        )
        format.index += 1

        st.table(format)

    st.write("""
        The first entry contains the column name (i.e. _S-number_), and will be discarded by default. You can download a sample sheet with the correct format below.
        ##### Download Sample Sheet
    """)

    st.download_button(
        label="S-number Sample Sheet",
        data=sample_snumbers(),
        file_name="snumber_sample_sheet.xlsx",
    )

st.write("""
    ## Exam Details
""")

# ask for course name
course_name = st.text_input('What is the course name (or acronym)? This will be part of the filenames.')

# ask for exam name
exam_name = st.text_input('What is the exam name? This will be part of the filenames.' )

# ask for the number of exam questions
nq = st.number_input('How many exam questions are there?', min_value = 0, max_value = 100, step = 1)

if nq is not None:
    subq = np.zeros(nq, dtype=int)
    for q in range(nq):
        nsq = st.number_input(f'How many subquestions does question {q+1} have?', min_value = 0, max_value = 100, step = 1)
        subq[q] = nsq

    qlist = Qindex(subq)

filename_prefix = course_name + "_" + exam_name + "_"

st.write("""
    ## Grading Team Details
""")

#ask for number of TAs
nTA = st.number_input('How many TAs will be grading?', min_value = 1, max_value = 50, step = 1)

st.write("""
    ## Downloads
    The download buttons will appear automatically after you provide all the necessary input above. 
""")

if S_file is not None and course_name is not "" and exam_name is not "":

    # read excel file with s-numbers into dataframe
    df = pd.read_excel(S_file, names = ['S-number'])

    # generate numpy array with splitted dataframe
    df_split = split_dataframe(df, nTA, qlist)

    # generate grading sheets
    buffer_zip = grading_sheets_generator(df_split)

    # download button zip file
    st.download_button(
        label="Download Grading Sheets",
        data=buffer_zip.getvalue(),
        file_name=filename_prefix+"grading_sheets.zip",
    )

    # buffer for grading labels
    buffer_gl = grading_labels_generator(df_split)

    # download document
    st.download_button(
        label="Download Grading Labels",
        data=buffer_gl,
        file_name=filename_prefix+"grading_labels.docx",
    )

    







    


