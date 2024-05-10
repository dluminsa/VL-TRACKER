import streamlit as st 
import pandas as pd
import os
import glob
from openpyxl import * #load_workbook
from openpyxl.styles import *
import numpy as np
cola, colb = st.columns([1,1])

st.success('WELCOME, this app was developed by Dr. Luminsa Desire, for any concern, reach out to him at desireluminsa@gmail.com')
st.markdown('**Follow these simple instructions before you proceed:**')
cola,colb = st.columns([1,2])
cola.markdown('**You need two extracts; *a CPHL extract and an emr extract* for this facility**')
colb.markdown('**In the emr extract:**')
colb.markdown('Rename the **HIV Clinic NO.** column to **A**')
colb.markdown('Rename the **VIral load results** column to **RE**')
colb.markdown('Rename the **Viral load Obs date** column to **VOB**')
cola.markdown('Rename the **RETURN VISIT DATE** column to **RD**')
 
col1,col2 = st.columns([1,1])
cphl = col1.file_uploader(label='Upload your CPHL extract here')
col1.write('The cphl extract must be in csv form')
emr = col2.file_uploader(label='Upload your EMR extract here')
col2.write('The emr extract must be in xlsx form')

extcphl = None
extemr = None
if cphl is not None and emr is not None:
    # Get the file name
    filecphl = cphl.name
    extcphl = os.path.splitext(filecphl)[1]
    fileemr = emr.name
    extemr = os.path.splitext(fileemr)[1]

df =None
df2 = None

df = None
if cphl is not None and emr is not None:
    if extcphl !='.csv':  # Compare with '.csv'
        st.write('Use a csv file generated from CPHL')
        st.stop()
    else:
        df = pd.read_csv(cphl)
        if extemr != '.xlsx':
            st.write('First save the emr extract as an xlsx file')
            st.stop()
        else:
            dfc = pd.read_excel(emr)
            cphlcolumns = ['facility', 'ART-NUMERIC', 'art_number', 'date_collected', 'Dyear',
        'Dmonth', 'Dday', 'result_numeric', ]
            colcphl = df.columns.to_list()
            for column in cphlcolumns:
                if column not in colcphl:
                    st.write (f' ERROR !!! {column} is not in this CPHL extract')
                    st.stop()
                #print( 'REACH OUT TO YOUR TEAM LEAD FOR THE RIGHT EXTRACT')
            #print('kindly check the table above to see how to rename the columns')
            emrcolumns= ['A', 'RE', 'RD','VOB']
            colemr = dfc.columns.to_list()
            for column in emrcolumns:
                if column not in colemr:
                    st.write (f' ERROR !!! {column} is not in this EMR extract')
                    st.stop()
                else:
                    df[['ART-NUMERIC', 'Dyear','Dmonth', 'Dday']] = df[['ART-NUMERIC', 'Dyear','Dmonth', 'Dday']].apply(pd.to_numeric, errors='coerce')
                    df = df[((df['Dyear']==2024) | ((df['Dyear']==2023) & (df['Dmonth']>6)))].copy()
                    dfc['ART'] = dfc['A'].replace('[^0-9]','', regex = True)
                    dfc = dfc.dropna(subset=['ART'])
                    dfc['VOD'] = dfc['VOB']
                    dfc['RT'] = dfc['RD']
                    dfc['VOD'] = dfc['VOD'].astype(str)
                    dfc['RT'] = dfc['RT'].astype(str)
                    dfc['RD'] = dfc['RD'].astype(str)
                    dfc['VOB'] = dfc['VOB'].astype(str)
                    dfc['VOD'] = dfc['VOD'].str.replace('/', '*')
                    dfc['VOD'] = dfc['VOD'].str.replace('-', '*')
                    dfc['VOD'] = dfc['VOD'].str.replace('00:00:00', '')
                    dfc['RT'] = dfc['RT'].str.replace('/', '*')
                    dfc['RT'] = dfc['RT'].str.replace('-', '*')
                    dfc['RT'] = dfc['RT'].str.replace('00:00:00', '')
                    dfc['RD'] = dfc['RD'].str.replace('00:00:00', '')
                    dfc['VOB'] = dfc['VOB'].str.replace('00:00:00', '')
                    dfc[['VOyear', 'VOmonth', 'VOday']] = dfc['VOD'].str.split('*', expand = True)
                    dfc[['Ryear', 'Rmonth', 'Rday']] = dfc['RT'].str.split('*', expand = True)
                    dfc[['VOyear', 'VOmonth', 'VOday']] =dfc[['VOyear', 'VOmonth', 'VOday']].apply(pd.to_numeric, errors = 'coerce')
                    dfc[['Ryear', 'Rmonth', 'Rday']] =dfc[['Ryear', 'Rmonth', 'Rday']].apply(pd.to_numeric, errors = 'coerce')
                    dfc['Ryear'] = dfc['Ryear'].fillna(2022)
                    dfc['VOyear'] = dfc['VOyear'].fillna(2022)
                    e = dfc[dfc['Ryear']>31].copy()     
                    f = dfc[dfc['Ryear']<32].copy()
                    f = f.rename(columns={'Ryear': 'Rday1', 'Rday': 'Ryear'})
                    f = f.rename(columns={'Rday1': 'Rday'})
                    dfc = pd.concat([e,f])
                    a = dfc[dfc['VOyear']>31].copy()
                    b = dfc[dfc['VOyear']<32].copy()
                    b = b.rename(columns={'VOyear': 'VOday1', 'VOday': 'VOyear'})
                    b = b.rename(columns={'VOday1': 'VOday'})
                    dfc = pd.concat([a,b])
                    def NEW (x, y):
                        if x ==2024:
                                return 'NEW'
                        elif x == 2023 and y > 5:
                            return 'NEW'
                        else:
                            return 'OLD'  
                    dfc['RESULTS'] = dfc.apply(lambda row: NEW(row['VOyear'], row['VOmonth']), axis=1)
                    dfc[['Ryear', 'Rmonth', 'Rday']] =dfc[['Ryear', 'Rmonth', 'Rday']].apply(pd.to_numeric, errors = 'coerce')
                    dfd = dfc[dfc['Ryear']>2024].copy()
                    dfe = dfc[dfc['Ryear']==2024].copy()
                    dfe = dfe[((dfe['Rmonth']>3) | ((dfe['Rmonth']==3) & (dfe['Rday']>3)))].copy()
                    dfc = pd.concat([dfd,dfe])
                    dfj = dfc[dfc['RESULTS']== 'OLD']
                    dfj = dfj[['ART','A', 'RE', 'RD', 'VOB']]
                    dfj['ART'] = pd.to_numeric(dfj['ART'])
                    df[ 'ART-NUMERIC'] = pd.to_numeric(df[ 'ART-NUMERIC'])
                    dfk = dfj[dfj['ART'].isin(df[ 'ART-NUMERIC'])]
                    dfh = df[df['ART-NUMERIC'].isin(dfj[ 'ART'])]
                    dfh = dfh.rename(columns={'ART-NUMERIC':'ART'})
                    dft = pd.merge(dfk, dfh, on = 'ART', how= 'left')
                    def comp(a,b):
                        if a == b:
                            return 'SAME'
                        else:
                            return 'DIFFERENT'
                    dft[['result_numeric','RE']] = dft[['result_numeric','RE']].apply(pd.to_numeric, errors='coerce')
                    dft['COMPARE'] = dft.apply( lambda row: comp(row['RE'], row['result_numeric']), axis=1)
                    dft = dft.rename(columns = {'RD': 'RETURN-DATE', 'RE':'EMR-RESULTS', 'A':'ART-NO', 'VOB':'VL_Obs_date'})
                    dft['date_collected'] =  dft['date_collected'].astype(str)
                    dft['date_collected'] =  dft['date_collected'].str.replace('*', '-')
                    dft = dft[['ART-NO', 'RETURN-DATE','EMR-RESULTS', 'VL_Obs_date','art_number','result_numeric','date_collected', 'COMPARE']]
                    
    if df is not None and df2 is not None: 
        a = dft.shape[0]
        st.markdown(f'I see over **{a}** results at CPHL that are not yet entered into EMR')
                    


    # st.success('Analysis done')
    if st.button('DOWNLOAD FILE'):
        wb = Workbook()
        ws = wb.active

        # Convert DataFrame to Excel
        for r_idx, row in enumerate(dft.iterrows(), start=1):
            for c_idx, value in enumerate(row[1], start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        ws.insert_rows(0,2)

        blue = PatternFill(fill_type = 'solid', start_color = 'C8CDCD')
        # ws.column_dimensions['H'].width = 14

        for num in range (1, ws.max_row+1):
            for letter in ['E', 'F', 'G', 'H']:
                ws[f'{letter}{num}'].font = Font(b= True, i = True)
                ws[f'{letter}{num}'].font = Font(b= True, i = True)
                ws[f'{letter}{num}'].fill = blue
                ws[f'{letter}{num}'].border = Border(top = Side(style = 'thin', color ='000000'),
                                                    right = Side(style = 'thin', color ='000000'),
                                                    left = Side(style = 'thin', color ='000000'),
                                                    bottom = Side(style = 'thin', color ='000000'))
        ws['B1'] ='EMR DETAILS'
        ws['F1'] = 'CPHL DETAILS'
        ws['A2'] = 'ART-NO'
        ws['B2'] = 'RETUR VISIT DATE'
        ws['C2'] = 'EMR VL RESULTS'
        ws['D2'] = 'EMR VL DATE' 
        ws['E2'] = 'ART NO'
        ws['F2']  = 'CPHL RESULTS'
        ws['G2'] = 'CPHL DATE'
        ws['H2'] = 'COMPARE'

        letters = ['B', 'C', 'D','F','G','H']
        for letter in letters:
            ws.column_dimensions[letter].width =15

        ran = np.random.rand()*0.1
        rand = round(ran,3)

        file_path = os.path.join(os.path.expanduser('~'), 'Downloads', f'missing_results {rand}.xlsx')

        # Save the workbook
        wb.save(file_path)
        #col1, col2 = st.columns([1,2])
        st.markdown(f'**YOUR FILE HAS BEEN DOWNLOADED AS missing_results {rand} IN YOUR DOWNLOAD FOLDER**')
        #col2.success(f'File saved as missing_results {rand} in your download folder ')
    if df is not None:
        dfu = dft.set_index('ART-NO')
        st.write(dfu.head(3))

