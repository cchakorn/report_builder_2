import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob
import numpy as np
from openpyxl import load_workbook
from datetime import datetime
import openpyxl
import xlsxwriter
from io import BytesIO
import io

###headers dictionary
headersDict = {'Others':
                    # ['BOOKED DATE', 'OTA RN', 'OTA REV', 'OTA ADR', 'B2B RN', 'B2B REV', 'B2B ADR', 
                    # 'TA RN', 'TA REV', 'TA ADR', 'GRTA RN', 'GRTA REV', 'GRTA ADR', 'CORP RN', 'CORP REV', 
                    # 'CORP ADR', 'GRCOR RN', 'GRCOR REV', 'GRCOR ADR', 'CORP TEAM-A RN', 'CORP TEAM-A REV', 
                    # 'CORP TEAM-A ADR', 'CORP TEAM-B RN', 'CORP TEAM-B REV', 'CORP TEAM-B ADR', 'CORP TEAM-C RN', 
                    # 'CORP TEAM-C REV', 'CORP TEAM-C ADR', 'WIN RN', 'WIN REV', 'WIN ADR', 'HWS RN', 'HWS REV', 
                    # 'HWS ADR', 'FIT RN', 'FIT REV', 'FIT ADR', 'EXT RN', 'EXT REV', 'EXT ADR', 'DIS RN', 'DIS REV', 
                    # 'DIS ADR', 'REDM RN', 'REDM REV', 'REDM ADR', 'PRB RN', 'PRB REV', 'PRB ADR', 'BT RN', 'BT REV', 
                    # 'BT ADR', 'GOV RN', 'GOV REV', 'GOV ADR', 'COM RN', 'COM REV', 'COM ADR', 'HOU RN', 'HOU REV', 
                    # 'HOU ADR', 'STAFF RN', 'STAFF REV', 'STAFF ADR', 'FNF RN', 'FNF REV', 'FNF ADR', 'OTH RN', 'OTH REV', 
                    # 'OTH ADR', 'NON code RN', 'NON code REV', 'NON code ADR', 'TOTAL RN', 'TOTAL REV', 'TOTAL ADR'],
                    # 'Others':
                    ['REPORT DATE', 'BOOKED DATE', 'OTA RN', 'OTA REV', 'OTA ADR', 'B2B RN', 'B2B REV', 'B2B ADR', 
                    'TA RN', 'TA REV', 'TA ADR', 'GRTA RN', 'GRTA REV', 'GRTA ADR', 'CORP RN', 'CORP REV', 
                    'CORP ADR', 'GRCOR RN', 'GRCOR REV', 'GRCOR ADR', 'CORP_TEAM_A RN', 'CORP_TEAM_A REV', 
                    'CORP_TEAM_A ADR', 'CORP_TEAM_B RN', 'CORP_TEAM_B REV', 'CORP_TEAM_B ADR', 'CORP_TEAM_C RN', 
                    'CORP_TEAM_C REV', 'CORP_TEAM_C ADR', 'WIN RN', 'WIN REV', 'WIN ADR', 'HWS RN', 'HWS REV', 
                    'HWS ADR', 'FIT RN', 'FIT REV', 'FIT ADR', 'EXT RN', 'EXT REV', 'EXT ADR', 'DIS RN', 'DIS REV', 
                    'DIS ADR', 'REDM RN', 'REDM REV', 'REDM ADR', 'PRB RN', 'PRB REV', 'PRB ADR', 'BT RN', 'BT REV', 
                    'BT ADR', 'GOV RN', 'GOV REV', 'GOV ADR', 'COM RN', 'COM REV', 'COM ADR', 'HOU RN', 'HOU REV', 
                    'HOU ADR', 'STAFF RN', 'STAFF REV', 'STAFF ADR', 'FNF RN', 'FNF REV', 'FNF ADR', 'OTH RN', 'OTH REV', 
                    'OTH ADR', 'NON RN', 'NON REV', 'NON ADR', 'TOTAL RN', 'TOTAL REV', 'TOTAL ADR'],

            
            'Arbour':
                    ['REPORT DATE', 'BOOKED DATE', 'OTA RN', 'OTA REV', 'OTA ADR', 'B2B RN', 'B2B REV', 'B2B ADR', 'TA RN', 'TA REV', 
                    'TA ADR', 'GRTA RN', 'GRTA REV', 'GRTA ADR', 'CORP RN', 'CORP REV', 'CORP ADR', 'GRCOR RN', 'GRCOR REV', 
                    'GRCOR ADR', 'WIN RN', 'WIN REV', 'WIN ADR', 'HWS RN', 'HWS REV', 'HWS ADR', 'FIT RN', 'FIT REV', 
                    'FIT ADR', 'EXT RN', 'EXT REV', 'EXT ADR', 'REDM RN', 'REDM REV', 'REDM ADR', 'PRB RN', 'PRB REV', 
                    'PRB ADR', 'BT RN', 'BT REV', 'BT ADR', 'GOV RN', 'GOV REV', 'GOV ADR', 'GGOV RN', 'GGOV REV', 'GGOV ADR', 
                    'OTHER RN', 'OTHER REV', 'OTHER ADR', 'COM RN', 'COM REV', 'COM ADR', 'HOU RN', 'HOU REV', 'HOU ADR', 
                    'STAFF RN', 'STAFF REV', 'STAFF ADR', 'FNF RN', 'FNF REV', 'FNF ADR', 'NON RN', 'NON REV', 'NON ADR', 
                    'TOTAL RN', 'TOTAL REV', 'TOTAL ADR']
}



### cleaning function
def my_cleaner(fileList,headers):
    modFiles = []
    for i, files in enumerate(fileList):
        #loading sheet
        df = fileList[i]    
        if df.columns[0] == 'REPORT DATE':
            #if formatted
            # st.header('formatted type uploaded')
            df['REPORT DATE'] = df['REPORT DATE'].dt.date
            df['BOOKED DATE'] = df['BOOKED DATE'].dt.date
            df.columns = headers
            modFiles.append(df)
            
        else:
            df=df.drop(index = df.shape[0]-1)
            #dropping top 2 column
            df=df.drop([0,1])
            #print(df.iloc[:,0]) #FOR CALLING COLUMN BY POSITION
            #clean date for booked date
            df.iloc[:,0] = df.iloc[:,0].apply(lambda x: datetime.strptime(x, '%a %d/%m/%Y').date())
            #inserting report date tracking
            df.insert(0,'REPORT DATE',df.iloc[0][0])
            df.columns = headers####
            modFiles.append(df)

    #dup date collection with loop
    checker =[]    
    for j, frames in enumerate(modFiles):
        uniqueFrames=frames['REPORT DATE'].unique()
        for k, uniqueDates in enumerate(uniqueFrames):
            checker.append(uniqueDates)
    #DUPLICATE CHECK
    dup = {x for x in checker if checker.count(x) > 1}
    #duplicate warning message
    if len(dup) > 0:
        st.write("WARNING: found duplicate dates for", dup)

    
    #concat excel
    excl_merged = pd.concat(modFiles, ignore_index=True)
    #sorting concatted df into report date order to override wrong uploads
    excl_merged= excl_merged.sort_values(by='REPORT DATE')

    return excl_merged

### HEADERS
st.set_page_config(page_title = 'Report Cleaner')
st.title('PU Report Cleaner 2')


#### REPORT BUILDER####
st.subheader('Please Upload Excel File1')
###header selection cuz arbour somehow is different
headersOption = st.selectbox(
    'Select Hotel',
    ('Arbour', 'Others'))

st.write('Selected Hotel:', headersOption)

### this is what is going to be called
# headersDict[option]
fileList = [] # list to append
st.subheader('Please Upload Excel Files')
uploaded_files = st.file_uploader("Choose a XLSX file",type = 'xlsx', accept_multiple_files=True)
for uploaded_file in uploaded_files:
    df = pd.read_excel(uploaded_file, engine = 'openpyxl')
    #st.write("filename:", uploaded_file.name)
    fileList.append(df)
st.markdown('---')

#INITIATE OPTION WHEN FILE IS UPLOADED
if len(fileList)>0:
    headers = headersDict[headersOption]
    fileCleaned = my_cleaner(fileList,headers)


    
#DOWNLOADER

#write pd via memory  
    in_memory_fp = io.BytesIO()
    fileCleaned.to_excel(in_memory_fp, index=False)
    #creating a list of what dates are inputted
    dfUnique = fileCleaned['REPORT DATE'].unique()
    #reporting dates included
    st.write("dates included:", dfUnique)
    #resetting fileList variable to clear sort-of cache
    fileList=[]
    st.download_button(
        label="Download Excel workbook",
        data=in_memory_fp.getvalue(),
        file_name="workbook.xlsx",
        mime="application/vnd.ms-excel"
    )

