import  streamlit as st
from docx import Document
import pandas as pd

def data_conv(url,sheet_url,start_q,end_q):
    sheet_adr = url
    sheet_name = sheet_url
    quation_start = int(start_q)
    quation_end = int(end_q)

    url = f"https://docs.google.com/spreadsheets/d/{sheet_adr}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    dataset = pd.read_csv(url, index_col=False)

    data_part_1 = '''
        INSTRUCTIONS:

        PLEASE NOTE THESE ARE SAMPLE QUESTIONS. REPLACE THEM WITH YOUR QUESTION CONTENT. 

        TO ADD QUESTIONS USE THE ENTIRE BLOCK FROM {START  QUESTIONS} TO {END QUESTIONS} AND ADD YOUR QUESTION CONTENT

        TO ADD DIRECTIONAL QUESTION USE THE ENTIRE BLOCK FROM {START DIRECTIONAL  QUESTIONS} TO {END DIRECTIONAL QUESTIONS} AND ADD YOUR QUESTION CONTENT

        TO ADD NORMAL QUESTION USE THE ENTIRE BLOCK FROM {START   NORMAL  QUESTIONS} TO {END NORMAL QUESTIONS} AND ADD YOUR QUESTION CONTENT

        TO ADD MORE QUESTIONS COPY PASTE THE ENTIRE BLOCK FROM {QUESTION BEGINS} TO {QUESTION ENDS} AND ADD YOUR QUESTION CONTENT

        FOLLOWING ARE THE SUPPORTED QUESTION TYPE – NORMAL, DIRECTIONAL

        {START QUESTIONS}


        {START NORMAL QUESTIONS} '''

    data_part_2 = ""

    def replace_all(text, dic):
        for i, j in dic.items():
            text = text.replace(i, j)
        return text

    p = 1
    for i in range(quation_start, quation_end + 1):
        q_no = str(p)
        quation = dataset.iloc[:, 1][i]
        opt_1 = dataset.iloc[:, 2][i]
        opt_2 = dataset.iloc[:, 3][i]
        opt_3 = dataset.iloc[:, 4][i]
        opt_4 = dataset.iloc[:, 5][i]
        ans = str(dataset.iloc[:, 6][i])
        dic = {"q_no": q_no, "quation": quation, "opt_1": opt_1, "opt_2": opt_2, "opt_3": opt_3, "opt_4": opt_4,
               "ans": ans}
        data_part = """
        {QUESTION BEGINS}
        {QUESTION NUMBER} q_no
        {QUESTION TYPE}  NORMAL
        {QUESTION ENGLISH TEXT} quation
        {OPTION ENGLISH 1} opt_1
        {OPTION ENGLISH 2} opt_2
        {OPTION ENGLISH 3} opt_3
        {OPTION ENGLISH 4} opt_4
        {OPTION ENGLISH 5}
        {RIGHT ANSWER} ans
        {EXPLANATION ENGLISH} No explanation available
        {EXPLANATION LINK ENGLISH} No explanation available
        {QUESTION ENDS}
        """
        data_1 = replace_all(data_part, dic)
        data_part_2 = data_part_2 + data_1
        p = p + 1

    data_part_3 = """

        {END QUESTIONS} """

    final_data = data_part_1 + data_part_2 + data_part_3
    try:
        document = Document()
        document.add_paragraph(final_data)
        document.save('demo.docx')
        with open('demo.docx', 'rb') as file:
            doc_data = file.read()
        st.download_button(label='Download Word File', data=doc_data, file_name='my_document.docx')
        st.success("Word File is Done")
    except:
        st.error("Some Mistak in Google Sheet")


st.title("Wscube Tech Question Upload")
url = st.text_input("Enter Your Url")
sheet_name = st.text_input("Enter Sheet Name")
start_q = st.text_input("Start Q. No.")
end_q = st.text_input("End Q. No.")
button = st.button("Done")
if button :
    data_conv(url,sheet_name,start_q,end_q)
