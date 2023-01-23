import openpyxl
import matplotlib as mpl
import pandas as pd
import numpy as np
import streamlit as st
import base64
import io
import plotly.express as px
import os
import glob
from openpyxl import load_workbook
from datetime import datetime
import datetime
import xlsxwriter
from io import BytesIO
import seaborn as sns
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter


#all the defs

def cleaner (df):
    # drop total row at the bottom
    df=df.drop(df.shape[0]-1)
    # move row 0 to col name and dropping the old row 0
    df.columns = list(df.iloc[0])
    df = df.drop([0])
    
    #turn datetime into string and set index
    df.index = pd.to_datetime(df['Date'], format='%a %d/%m/%Y')
    df.pop('Date') # removes old date column as we set index already
    
    #renaming RT to be RT_RoomTOTAL
    popList = ['TLRM','AVL','OCC%','OO/OI'] #not room types
    colList = df.columns.tolist() #all the columns

    for items in popList:
        colList.remove(items) #room types will be returned

    #string formatting
    for count, headers in enumerate (colList):
        newHeaders = headers.replace('(','').replace(')','').replace(' ','_')
        colList[count] = newHeaders
    
    #column list rearragements
    standardCol = ['TLRM', 'AVL','OCC%'] #head of list
    for items in colList:
        standardCol.append(items) #append body
    standardCol.append('OO/OI') #append tail
    df.columns = standardCol #standardCol is done and applied - pull if needed

    
    #setting values as numeric
    for cols in df.columns:
        df[cols] = pd.to_numeric(df[cols], errors='coerce')
    
    #creating a room list with names_total:total
    rtColNames = {}
    for _, rtLables in enumerate (colList):
        _,roomCount = rtLables.split('_')
        rtColNames.update({rtLables:int(roomCount)})
    
    #we also need dates to select from
    dates = df.index.tolist()
    
    return(df,rtColNames,dates)

# df2,rtColNames,dates = cleaner(df2)


def dateSlicer(df,startDate,endDate):
    dfSliced =df.sort_index().loc[startDate : endDate,:]
    return(dfSliced)

def genReport (df,rtColNames):
    
    #dynamically counts days inputted - can be over month
    daysInMonth = df.shape[0]
    
    #sum column and create rows
    df.loc['Total_Left'] = df.sum(numeric_only = True)

    #for RT columns only we multiply it by daysInMonths = capacicty
    for _, keys in enumerate(rtColNames):
        df.loc['Total_Capacity',keys] = rtColNames[keys]*daysInMonth
    
    #roomLeft/roomCapacity = % room left /// one minus that for occ room sold
    df.loc['Room_Type_Occ'] = (100*(1-(df.loc['Total_Left']/df.loc['Total_Capacity']))).round(1)#added rounding for excel ease
        
    #adding occ
    df.at['Room_Type_Occ','OCC%'] = (100*(1-(df.at['Total_Left','AVL']/df.at['Total_Left','TLRM']))).round(1)#added rounding for excel ease
    
    #adding total RS
    df.loc['Total_RS'] = df.loc['Total_Capacity']-df.loc['Total_Left']
    
    #adding portion sold
    df.loc['Portion_Sold'] = (100*(df.loc['Total_Capacity',[rtNames for rtNames in rtColNames]]/df.loc['Total_Left',[rtNames for rtNames in rtColNames]].sum())).round(1)#added rounding for excel ease
    
    
    #dropping rows that was used in calc as it is no longer needed
    df = df.drop('Total_Left')
    df = df.drop('Total_Capacity')
    df = df.drop('Total_RS')

    return(df)

def formattingA (df):
    # selecting each columns
    #formatting back ground gradients
    
    
    # for _, keys in enumerate(rtColNames):
    # df.background_gradient(axis=None,
    #                                     cmap='YlOrRd_r',
    #                                     subset=(keys),
    #                                     vmin=0, vmax=rtColNames[keys]) 


    # df.text_gradient(axis=None, low=0, high=10)  


    df.background_gradient(axis=None,
                              cmap='RdPu',
                              subset=(['Room_Type_Occ'],[rtNames for rtNames in rtColNames]),
                              vmin=0, vmax=100) 
    df.background_gradient(axis=None,
                              cmap='RdPu',
                              subset=(['Portion_Sold'],[rtNames for rtNames in rtColNames]),
                              vmin=0, vmax=100) 
    
    df.format('{:,.0f}')

    func = lambda s: s.strftime("%Y-%m-%d %A") if isinstance(s, pd.Timestamp) else s
    df.format_index(func)
    return(df)
    




st.title('Room Type Report')
st.subheader('Please Upload Excel Files')
fileList = [] # list to append
uploaded_files = st.file_uploader("Choose a XLSX file",type = 'xlsx', accept_multiple_files=True)
for uploaded_file in uploaded_files:
    df = pd.read_excel(uploaded_file, engine = 'openpyxl')
    #st.write("filename:", uploaded_file.name)
    fileList.append(df)

    #cleaner
    df2,rtColNames,dates = cleaner(fileList[0])


    #date choices

    def dateChoice (df):
            #column item choices:
        cols = df.index.tolist()
        newCols = []

        for item in cols:
            if item not in newCols:
                newCols.append(item)
        newCols.sort()
        #newCols is the variable for column headers
        return (newCols)

        #creating RD column choices
    RDCols = dateChoice(df2)

    reportDateRange = st.date_input(
        'Select Report Date Range', 
        (RDCols[0],RDCols[len(RDCols)-1]),
        min_value=RDCols[0], 
        max_value=RDCols[len(RDCols)-1],         
        )

    #date slicer
    dfSliced = dateSlicer(df2,reportDateRange[0], reportDateRange[1])
    #gen report
    dfReport = genReport(dfSliced,rtColNames)
    #colour report
    dfRT_Colour = dfReport.style.pipe(formattingA)
    
    
    # 
    st.dataframe(dfRT_Colour)

    #download stuff
    #working with excel sheet in memory

    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine="xlsxwriter",
                        datetime_format='mmm d yyyy ddd',
                        date_format='mmmm dd yyyy') as writer:
    #with pd.ExcelWriter(buffer, mode = 'a', engine="openpyxl") as writer:
        dfRT_Colour.to_excel(writer, sheet_name="Sheet1", index=True, freeze_panes=(1,1))
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        #size indexing
        worksheet.set_column('A:A', 22)
        writer.save()

        st.download_button(
            label="Download Excel",
            data=buffer,
            file_name="RT_Report.xlsx",
            mime="application/vnd.ms-excel",
        )


