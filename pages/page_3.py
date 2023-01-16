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




def cleanTable (df,chooseSeg,chooseData,reportDateRange,bookedDateRange):
    df=df.set_index(['REPORT DATE', 'BOOKED DATE'])
    idx = df.columns.str.split(' ', expand=True)
    idx = idx.tolist()
    df.columns=pd.MultiIndex.from_tuples(idx)
    df=df.unstack(level = 1).swaplevel(2,0,axis=1).swaplevel(1,2,axis=1).sort_index(axis=1)
    df=df.transpose()


    #copy of original is dfO
    dfO = df.copy()
    df=df.diff(axis=1)

    #OPTIONS
    segSlice = chooseSeg
    dataSlice = chooseData #if all return :
    startRD = reportDateRange[0]
    endRD = reportDateRange[1]
    startBD = bookedDateRange[0]
    endBD = bookedDateRange[1]



    #df copies
    dfO=dfO.loc[pd.IndexSlice[startBD : endBD,segSlice,dataSlice],startRD: endRD]
    df = df.loc[pd.IndexSlice[startBD : endBD,segSlice,dataSlice],startRD: endRD]

    #more cleaning
    dfF = pd.concat([df,dfO],keys = ['diff', 'actual'], axis = 1,)
    dfF =dfF.swaplevel(1,0,axis=1)
    dfF =dfF.sort_index(axis=1)
    dfF =dfF.unstack(level = [1,2])
    dfF =dfF.reindex(['REV', 'RN', 'ADR'], axis=1, level=3)

    return (dfF)


def colChoice (df):
        #column item choices:
    cols = []
    for i in enumerate(df.columns.tolist()):
        k = i[1].split()
        cols.append(k[0])
    newCols = []

    for item in cols:
        if item not in newCols:
            newCols.append(item)

    newCols.remove('REPORT')
    newCols.remove('BOOKED')
    #newCols is the variable for column headers
    return (newCols)


def dateChoice (df,dateType):
        #column item choices:
    cols = df[dateType].tolist()
    newCols = []

    for item in cols:
        if item not in newCols:
            newCols.append(item)
    newCols.sort()
    #newCols is the variable for column headers
    return (newCols)



st.title('PICK UP REPORT ver2')
file = st.file_uploader("Upload an Excel file", type="xlsx")
if file is not None:
    df = pd.read_excel(file, engine = 'openpyxl')

    #creating SEG column choices
    segCols = colChoice(df)
    #moving major segments up to top for ease
    segCols.insert(0, segCols.pop(segCols.index('HWS')))
    segCols.insert(0, segCols.pop(segCols.index('CORP')))
    segCols.insert(0, segCols.pop(segCols.index('TA')))
    segCols.insert(0, segCols.pop(segCols.index('OTA')))
    segCols.insert(0, segCols.pop(segCols.index('TOTAL')))
    #selection for SEG
    chooseSeg = st.multiselect(
        'Choose Segments',
        options = segCols)


    #choose booking data
    chooseData = st.multiselect(
    'Choose Booking Data',
    ['REV', 'RN' , 'ADR'])


    #creating RD column choices
    RDCols = dateChoice(df,'REPORT DATE')

    reportDateRange = st.date_input(
        'Select Report Date Range', 
        (RDCols[0],RDCols[len(RDCols)-1]),
        min_value=RDCols[0], 
        max_value=RDCols[len(RDCols)-1],         
        )
    #creating BD column choices
    BDCols = dateChoice(df,'BOOKED DATE')

    bookedDateRange = st.date_input(
        'Select Report Date Range', 
        (BDCols[0],BDCols[len(BDCols)-1]),
        min_value=BDCols[0], 
        max_value=BDCols[len(BDCols)-1],         
        ) 

    #data clean and filter
    dfCleaned = cleanTable(df,chooseSeg,chooseData,reportDateRange,bookedDateRange)
    st.dataframe(dfCleaned)
    choosePUorActual = st.selectbox(
        'diff / actual / actual|diff',
        ['diff', 'actual' , 'actual|diff'])

    def puOrActual (df,choice):
        df = df.filter(regex=choice,axis=1)
        return df

    dfCleaned = puOrActual(dfCleaned,choosePUorActual)

    # formatting stuff
    def formattingA (dfS,dfF):
        #colouring formats for each data field
        cmREV = sns.diverging_palette(0, 255,s=50, l=50, sep=1, n=9, as_cmap=True)
        cmRN  = sns.diverging_palette(0, 255,s=50, l=50, sep=1, n=9, as_cmap=True)
        cmADR = sns.diverging_palette(0, 255,s=50, l=50, sep=1, n=9, as_cmap=True)
        
        #max data setting
        dfDiffM = dfF.filter(regex='diff',axis=1)
        REVMax = np.abs(dfDiffM.filter(regex='REV',axis=1)).max().max()
        RNMax  = np.abs(dfDiffM.filter(regex='RN',axis=1)).max().max()
        ADRMax = np.abs(dfDiffM.filter(regex='ADR',axis=1)).max().max()

        # selecting each columns
        #formatting back ground gradients
        dfS.background_gradient(cmap=cmREV, 
                            subset=[c for c in dfS.columns if c[1] in ('diff') and c[3] in ('REV')],
                            vmin=-REVMax, vmax=REVMax) 
        
        dfS.background_gradient(cmap=cmRN,
                                subset=[c for c in dfS.columns if c[1] in ('diff') and c[3] in ('RN')],
                                vmin=-RNMax, vmax=RNMax)
            
        dfS.background_gradient(cmap=cmADR, 
                            subset=[c for c in dfS.columns if c[1] in ('diff') and c[3] in ('ADR')],
                            vmin=-ADRMax, vmax=ADRMax)
        
        
        #formatting number formats
        dfS.format('{:,.0f}', subset=[c for c in dfS.columns if c[3] in ('REV')]) 
        dfS.format('{:,.0f}', subset=[c for c in dfS.columns if c[3] in ('RN') ]) 
        dfS.format('{:,.1f}', subset=[c for c in dfS.columns if c[3] in ('ADR')]) 

        #date time format
        dfS.format_index(lambda v: v.strftime("%Y-%m-%d %A"))
        dfS.format_index({0:lambda v: v.strftime("%Y-%m-%d %A")},axis=1)
        return dfS

    dfCleaned = dfCleaned.style.pipe(formattingA,dfCleaned)


    #download stuff
    #working with excel sheet in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        dfCleaned.to_excel(writer, sheet_name="Sheet1", index=True)
        workbook = writer.book  
        

        writer.save()

        st.download_button(
            label="Download Excel",
            data=buffer,
            file_name="df1.xlsx",
            mime="application/vnd.ms-excel",
        )




    # towrite = io.BytesIO()
    # downloaded_file = dfCleaned.to_excel(towrite, encoding='utf-8', index=True, header=True) # write to BytesIO buffer
    # towrite.seek(0)  # reset pointer
    # b64 = base64.b64encode(towrite.read()).decode() 
    # linko= f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="myfilename.xlsx">Download excel file</a>'
    # st.markdown(linko, unsafe_allow_html=True)