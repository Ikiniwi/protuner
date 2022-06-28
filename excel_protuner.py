import streamlit as st
import pandas as pd
import numpy as np
import numpy as np
from pandas import DataFrame, read_csv
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import pandas as pd

st.title('Excel to protuner')


uploaded_file = st.file_uploader("Choose a file")


def modify_table_for_protuner_input(name_excel):
    # lire data depuis la feuille Con data extrait avec l'addin
    xls = pd.ExcelFile(name_excel)
    df = pd.read_excel(xls, 'Sheet1', index_col=0, header=0, skiprows=0)
    df = df[~df.loop_number.isnull()]  # we remove none data !
    # df.head()

    df = df.drop_duplicates(subset='loop_number', keep="first")  # remove duplicate
    list_loop_number = df.loop_number.unique().tolist()  # list of loop_number
    df = df.sort_values(['loop_number'], ascending=True)  # sort the dataframe by loop_number

    tagname, OPCname, Raw_Min, Raw_Max, EU_MIN, EU_MAX = [], [], [], [], [], []  # init list

    for i in list_loop_number:  #
        for j in ["PV", "SP", "OUT"]:
            tagname.append(i + "_" + df.loop_name[df['loop_number'] == i].to_list()[0] + "_" + j)
            OPCname.append(i + "/PID1/" + j + ".CV")
            if j == "PV" or j == "SP":
                Raw_Min.append(df['Input Eng. Range min'][df['loop_number'] == i].to_list()[0])
                Raw_Max.append(df['Input Eng. Range max'][df['loop_number'] == i].to_list()[0])
            elif j == "OUT":
                Raw_Min.append(df['Output Eng. Range min'][df['loop_number'] == i].to_list()[0])
                Raw_Max.append(df['Output Eng. Range max'][df['loop_number'] == i].to_list()[0])

    DataSet = list(zip(tagname, OPCname, Raw_Min, Raw_Max, Raw_Min, Raw_Max))  # create dataframe
    df_final = pd.DataFrame(data=DataSet, columns=['Tag name', 'OPC Name', 'Raw Min', 'Raw Max', 'EU_MIN', 'EU_MAX'])

    # send in excel file
    #sheet_name = 'protuner'
    #insert_df_in_existing_excel(df_final, sheet_name, name_excel + '.xlsx')

    return df_final

if uploaded_file is None:
    st.write('No excel file')
else :
    st.write('The current excel file is : ', uploaded_file)
    st.write('Here is the new table for protuner')
    df_protuner = modify_table_for_protuner_input(uploaded_file)
    df_protuner

    @st.cache
    def convert_df(df):
        return df.to_csv().encode('utf-8')

    csv = convert_df(df_protuner)

    st.download_button(
        "Press to Download",
        csv,
        "file.csv",
        "text/csv",
        key='download-csv'
    )



