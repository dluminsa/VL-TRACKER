import streamlit as st 
import pandas as pd
import os
import numpy as np
from openpyxl import *
from openpyxl.styles import *
from openpyxl.worksheet.datavalidation import DataValidation
import time

st.set_page_config(
    page_title = 'VL TRACKER'
)
st.success('WELCOME, this app was developed by Dr. Luminsa Desire, for any concern, reach out to him at desireluminsa@gmail.com')
current_time = time.localtime()
week = time.strftime("%U", current_time)


st.markdown(f"******* **REMINDER!! we are currently in week {week}** *****")

st.markdown('**FIRST RENAME THESE COLUMNS BEFORE YOU PROCEED:**')
col1, col2, col3 = st.columns([1,1,1])
col1.markdown('Rename the **HIV CLINIC NO.** column to **A**')
col1.markdown('Rename the **ART START DATE** column to **AS**')
col2.markdown('Rename the **RETURN VISIT DATE** column to **RD**')
col2.markdown('Rename the **RETURN VISIT DATE_1** column to **RD1**')
col3.markdown('Rename the **HIV VIRAL LOAD DATE** column to **VD**')
col3.markdown('Rename the **HIV VIRAL LOAD DATE_1** column to **VD1**')

file = st.file_uploader(label='Upload your EMR extract here')

ext = None
if file is not None:
    # Get the file name
    fileN = file.name
    ext = os.path.basename(fileN).split('.')[1]
df = None
if file is not None:
    if ext !='xlsx':
        st.write('Unsupported file format, first save the excel as xlsx and try again')
        st.stop()
    else:
        df = pd.read_excel(file)
        st.write('Excel accepted')

    if df is not None:
        columns = ['A', 'AS', 'VD', 'VD1', 'RD', 'RD1']
        cols = df.columns.to_list()
        if not all(column in cols for column in columns):
            missing_columns = [column for column in columns if column not in cols]
            for column in missing_columns:
                st.markdown(f' **ERROR !!! {column} is not in the file uploaded**')
                st.markdown('**First rename all the columns as guided above**')
                st.stop()
        else:
            # Convert 'A' column to string and create 'ART' column with numeric part
            df['A'] = df['A'].astype(str)
            df['ART'] = df['A'].str.replace('[^0-9]', '', regex=True)
            df['ART'] = pd.to_numeric(df['ART'], errors= 'coerce')
            df = df[df['ART']>0]
            df.dropna(subset='ART', inplace=True)
            #st.write(df.shape[0])
            df = df.copy()
            #CONVERTING DATES TO STRINGS
            df[['AS', 'RD', 'VD','VD1','RD1']] = df[['AS', 'RD', 'VD', 'VD1','RD1']].astype(str)

            df['AS'] = df['AS'].str.replace('/', '*')
            df['RD'] = df['RD'].str.replace('/', '*')
            df['VD'] = df['VD'].str.replace('/', '*')
            #df['LD'] = df['LD'].str.replace('/', '*')
            df['RD1'] = df['RD1'].str.replace('/', '*')
            df['VD1'] = df['VD1'].str.replace('/', '*')

            df['AS'] = df['AS'].str.replace('-', '*')
            df['RD'] = df['RD'].str.replace('-', '*')
            df['VD'] = df['VD'].str.replace('-', '*')
            #df['LD'] = df['LD'].str.replace('-', '*')
            df['RD1'] = df['RD1'].str.replace('-', '*')
            df['VD1'] = df['VD1'].str.replace('-', '*')

            df['AS'] = df['AS'].str.replace('00:00:00', '')
            df['RD'] = df['RD'].str.replace('00:00:00', '')
            df['VD'] = df['VD'].str.replace('00:00:00', '')
            #df['LD'] = df['LD'].str.replace('00:00:00', '')
            df['RD1'] = df['RD1'].str.replace('00:00:00', '')
            df['VD1'] = df['VD1'].str.replace('00:00:00', '')

            df[['Ayear', 'Amonth', 'Aday']] = df['AS'].str.split('*', expand = True)
            df[['Ryear', 'Rmonth', 'Rday']] = df['RD'].str.split('*', expand = True)
            df[['Vyear', 'Vmonth', 'Vday']] = df['VD'].str.split('*', expand = True)
            #df[['Lyear', 'Lmonth', 'Lday']] = df['LD'].str.split('*', expand = True)
            df[['RD1year', 'RD1month', 'RD1day']] = df['RD1'].str.split('*', expand = True)
            df[['VD1year', 'VD1month', 'VD1day']] = df['VD1'].str.split('*', expand = True)

            #SORTING THE VIRAL LOAD YEARS
            df[['Vyear', 'Vmonth', 'Vday']] =df[['Vyear', 'Vmonth', 'Vday']].apply(pd.to_numeric, errors = 'coerce') 
            df['Vyear'] = df['Vyear'].fillna(2022)
            a = df[df['Vyear']>31].copy()
            b = df[df['Vyear']<32].copy()
            b = b.rename(columns={'Vyear': 'Vday1', 'Vday': 'Vyear'})
            b = b.rename(columns={'Vday1': 'Vday'})
            df = pd.concat([a,b])
            dfa = df.shape[0]

            #SORTING THE VIRAL LOAD DATE1 YEARS
            df[['VD1year', 'VD1month', 'VD1day']] =df[['VD1year', 'VD1month', 'VD1day']].apply(pd.to_numeric, errors = 'coerce')
            df['VD1year'] = df['VD1year'].fillna(2022)
            a = df[df['VD1year']>31].copy()
            b = df[df['VD1year']<32].copy()
            b = b.rename(columns={'VD1year': 'VD1day1', 'VD1day': 'VD1year'})
            b = b.rename(columns={'VD1day1': 'VD1day'})
            df = pd.concat([a,b])
            dfb = df.shape[0]
            #SORTING THE RETURN VISIT DATE YEARS
            df[['Rday', 'Ryear']] = df[['Rday', 'Ryear']].apply(pd.to_numeric, errors='coerce')
            df['Ryear'] = df['Ryear'].fillna(2022)
            a = df[df['Ryear']>31].copy()
            b = df[df['Ryear']<32].copy()
            b = b.rename(columns={'Ryear': 'Rday1', 'Rday': 'Ryear'})
            b = b.rename(columns={'Rday1': 'Rday'})

            df = pd.concat([a,b])
            dfc = df.shape[0]
            #SORTING THE RETURN VISIT DATE1 YEAR
            df[['RD1day', 'RD1year']] = df[['RD1day', 'RD1year']].apply(pd.to_numeric, errors='coerce')
            df['RD1year'] = df['RD1year'].fillna(2022)
            a = df[df['RD1year']>31].copy()
            b = df[df['RD1year']<32].copy()
            b = b.rename(columns={'RD1year': 'RD1day1', 'RD1day': 'RD1year'})
            b = b.rename(columns={'RD1day1': 'RD1day'})
            df = pd.concat([a,b])

            dfd = df.shape[0]
            #SORTING THE ART START YEARS
            df[['Ayear', 'Amonth', 'Aday']] =df[['Ayear', 'Amonth', 'Aday']].apply(pd.to_numeric, errors = 'coerce')
            df['Ayear'] = df['Ayear'].fillna(2022)
            a = df[df['Ayear']>31].copy()
            b = df[df['Ayear']<32].copy()
            b = b.rename(columns={'Ayear': 'Aday1', 'Aday': 'Ayear'})
            b = b.rename(columns={'Aday1': 'Aday'})
            df = pd.concat([a,b])
            dfe = df.shape[0]
            #print(dfa, dfb, dfc, dfd, dfe)

            #SORTING THE TX_NEW

            def Eligible(x):
                if x <2024:
                    return ('ELLIGIBLE')
                else:
                    return('NOT')

            #FOR THOSE DUE
            def Due(x, y):
                if x ==2024:
                    if y < 4:
                        return 'ALREADY'
                    elif y > 3:
                        return 'BLED'     
                elif x == 2023:
                    if y > 6:
                        return 'ALREADY'
                    else:
                        return 'DUE'
                else:
                    return('DUE')

            #FOR WEEKS
            def week (c,a,b):
                if c ==2024:
                    if a == 4:
                        if 1 <= b <=7:
                            return 14
                        elif 8<= b <= 14:
                            return 15
                        elif 15 <= b <=21:
                            return 16
                        elif 22<= b <= 28:
                            return 17
                        elif 29 <= b <=30:
                            return 18
                    elif a == 5:
                        if 1 <= b <=5:
                            return 18
                        elif 6<= b <= 12:
                            return 19
                        elif 13 <= b <= 19:
                            return 20
                        elif 20 <= b <=26:
                            return 21
                        elif 27 <= b <= 31:
                            return 22
                    elif a == 6:
                        if 1 <= b <= 2:
                            return 22
                        elif 3 <= b <= 9:
                            return 23
                        elif 10 <= b <= 16:
                            return 24
                        elif 17 <= b <= 23:
                            return 25
                        elif 24 <= b <= 30:
                            return 26
                    else:
                        return None
                else:
                    return None
                    
            #APPLYING THE WEEK FORMULA ON VL DATES
            df[['Vday', 'Vyear','Vmonth']] = df[['Vday', 'Vyear', 'Vmonth']].apply(pd.to_numeric, errors='coerce')
            df['BWEEK'] = df.apply(lambda wee: week(wee['Vyear'],wee['Vmonth'], wee['Vday']), axis=1)
            #APPLYIN THE DUE FORMULA ON VL DATES
            df['DUE'] = df.apply(lambda due: Due(due['Vyear'], due['Vmonth']), axis=1)
            #APPLYING THE DUE FORMUA TO VD1 DATES
            df[['VD1day', 'VD1year','VD1month']] = df[['VD1day', 'VD1year', 'VD1month']].apply(pd.to_numeric, errors='coerce')
            df['DUE1'] = df.apply(lambda due: Due(due['VD1year'], due['VD1month']), axis=1)
            #APPLYING THE WEEK FORMULA TO RETURN VISIT DATES
            df[['Rday', 'Ryear','Rmonth']] = df[['Rday', 'Ryear', 'Rmonth']].apply(pd.to_numeric, errors='coerce')
            df['WEEK'] = df.apply(lambda row: week(row['Ryear'], row['Rmonth'], row['Rday']),axis=1)
            #APPLYING THE WEEK FORMULA TO RETURN VISIT DATES 1
            df[['RD1day', 'RD1year','RD1month']] = df[['RD1day', 'RD1year', 'RD1month']].apply(pd.to_numeric, errors='coerce')
            df['WEEK1'] = df.apply(lambda row: week(row['RD1year'], row['RD1month'], row['RD1day']),axis=1)
            #APPLYING THE FORMULA TO RULE OUT TX NEW
            df['Ayear'] = pd.to_numeric(df['Ayear'], errors= 'coerce')
            #APPLYING THE FORMULA TO RULE OUT TX NEW
            df['Ayear'] = pd.to_numeric(df['Ayear'], errors= 'coerce')
            df['ELL'] = df['Ayear'].apply(Eligible)
            #RECONSTRUCTING RETURN VISIT DATES
            df['Ryear'] = pd.to_numeric(df['Ryear'], errors= 'coerce')
            dfa = df[df['Ryear']==2024].copy()
            df['RD1year'] = pd.to_numeric(df['RD1year'], errors= 'coerce')
            dfb = df[df['RD1year']==2024].copy()

            dfa[['Rday', 'Ryear','Rmonth']] = dfa[['Rday', 'Ryear', 'Rmonth']].apply(pd.to_numeric, errors='coerce')
            CURRa = dfa[((dfa['Rmonth']>3) | ((dfa['Rmonth']==3) & (dfa['Rday'] >3)))].copy()
            dfb[['RD1day', 'RD1year','RD1month']] = dfb[['RD1day', 'RD1year', 'RD1month']].apply(pd.to_numeric, errors='coerce')
            CURRb = dfb[((dfb['RD1month']>3) | ((dfb['RD1month']==3) & (dfb['RD1day'] >3)))].copy()
            CURRb = CURRb.drop(columns=['WEEK'])
            CURRa = CURRa.drop(columns=['WEEK1'])
            CURRb = CURRb.rename(columns={'WEEK1': 'WEEK'})
            dfcurr = pd.concat([CURRa, CURRb])
            dfcurr['ELL'] = dfcurr['ELL'].astype(str)
            dfcurr = dfcurr[dfcurr['ELL'] == 'ELLIGIBLE'].copy()
            #COMPUTING WEEKLY BLEEDS
            #CHOOSE BLEEDS DONE IN THE QUARTER
            df[['Vday', 'Vyear','Vmonth']] = df[['Vday', 'Vyear', 'Vmonth']].apply(pd.to_numeric, errors='coerce')
            dfv = df[(df['Vyear']==2024) & (df['Vmonth']>3)].copy()
            dfv['DUE'] = dfv['DUE'].astype(str)
            dfv['DUE1'] = dfv['DUE1'].astype(str)
            #REBLEEDS
            REBLED = dfv[((dfv['DUE'] == 'BLED')& (dfv['DUE1'] == 'ALREADY'))].copy()
            REBLED['STATUS'] = np.nan
            REBLED['STATUS'] = REBLED['STATUS'].fillna('REBLED')
            REBLED = dfv[((dfv['DUE'] == 'BLED')& (dfv['DUE1'] == 'ALREADY'))].copy()
            REBLED['STATUS'] = np.nan
            REBLED['STATUS'] = REBLED['STATUS'].fillna('REBLED')
            REBLEDpivo = pd.pivot_table(REBLED, index= 'BWEEK', values='A', aggfunc = 'count')
            REBLEDF = REBLEDpivo.reset_index()
            REBLEDF = REBLEDF.rename(columns={'A': 'REBLED'})
            #BLEEDS, original dataframe should be dfv
            BLED = dfv[((dfv['DUE'] == 'BLED')& (dfv['DUE1'] == 'DUE'))].copy()
            BLED['STATUS'] = np.nan
            BLED['STATUS'] = BLED['STATUS'].fillna('BLED')
            BLEDpivo = pd.pivot_table(BLED, index= 'BWEEK', values='A', aggfunc = 'count')
            BLEDF = BLEDpivo.reset_index()
            BLEDF = BLEDF.rename(columns={'A': 'BLED'})
            BLEDF = BLEDF.rename(columns={'A': 'BLED'})
            #TOTAL BLEEDS
            TOTAL = pd.pivot_table(dfv, index= 'BWEEK', values='A', aggfunc = 'count')
            TOTAL = TOTAL.reset_index()   
            TOTAL = TOTAL.rename(columns={'A':'TOTAL BLEEDS'}) 
            #####FIRST PIVOT FROM BLED, REBLED, AND TOTAL BLEEDS   
            BLEDF['BWEEK'] = pd.to_numeric(BLEDF['BWEEK'], errors='coerce')     
            REBLEDF['BWEEK'] = pd.to_numeric(REBLEDF['BWEEK'], errors='coerce')
            dfy = pd.merge(BLEDF, REBLEDF, on = 'BWEEK', how ='outer')
            TOTAL['BWEEK'] = pd.to_numeric(TOTAL['BWEEK'], errors='coerce')
            dfy['BWEEK'] = pd.to_numeric(dfy['BWEEK'], errors='coerce')
            dfg = pd.merge(dfy, TOTAL, on = 'BWEEK', how='outer')
            dfg = dfg.rename(columns = {'BWEEK': 'WEEK'})
            dfg = dfg.set_index('WEEK')
            weekly = dfg.copy()
            #BLEEDING VS APPOINTMENT
            APPT = dfcurr[['A', 'WEEK', 'DUE']].copy()
            APPT['DUE'] = APPT['DUE'].astype(str)
            NOT_DUE = APPT[APPT['DUE']=='ALREADY']
            NOTDUE = pd.pivot_table(NOT_DUE, index='WEEK', values='A', aggfunc = 'count')
            NOTD = NOTDUE.reset_index()
            NOTD = NOTD.rename(columns={'A': 'NOT DUE'})
            DUE = APPT[APPT['DUE']=='DUE']
            DU = pd.pivot_table(DUE, index='WEEK', values='A', aggfunc = 'count')
            DU = DU.reset_index()
            DU = DU.rename(columns={'A': 'DUE'})
            BLED = APPT[APPT['DUE']=='BLED']
            BLE = pd.pivot_table(BLED, index='WEEK', values='A', aggfunc = 'count')
            BLE = BLE.reset_index()
            BLE = BLE.rename(columns={'A': 'BLED'})
            BLE['WEEK'] = pd.to_numeric(BLE['WEEK'], errors='coerce')
            NOTD['WEEK'] = pd.to_numeric(NOTD['WEEK'], errors='coerce')
            first = pd.merge(NOTD, BLE, on ='WEEK', how= 'outer')
            first['WEEK'] = pd.to_numeric(first['WEEK'], errors='coerce')
            first = pd.merge(NOTD, BLE, on ='WEEK', how= 'outer')
            DU['WEEK'] = pd.to_numeric(DU['WEEK'], errors='coerce')
            first['WEEK'] = pd.to_numeric(first['WEEK'], errors='coerce')
            second = pd.merge(first, DU, on ='WEEK', how= 'outer')
            appot = pd.pivot_table(dfcurr, index='WEEK', values='A', aggfunc='count')
            appte = appot.reset_index()
            appte = appte.rename(columns={'A': 'ON APPT'})
            APPTSUM = pd.merge(appte, second, on= 'WEEK', how ='outer')
            APPTSUM = APPTSUM.set_index('WEEK')
            #DETERMINING TX_CURR
            #MODIFY HERE TO INCLUDE 2025 LATER
            CURR = df.copy()
            CURR[['Rday', 'Ryear','Rmonth']] = CURR[['Rday', 'Ryear', 'Rmonth']].apply(pd.to_numeric, errors='coerce')
            CURR = CURR[CURR['Ryear']==2024].copy()
            CURR = CURR[((CURR['Rmonth']>3) | ((CURR['Rmonth']==3) & (CURR['Rday'] >3)))].copy()
            CUR = CURR.copy()
            a = CUR.shape[0]
            CUR['ELL'] = CUR['ELL'].astype(str)
            #COUNT THOSE NOT ELLIGIBLE
            f = CUR[CUR['ELL'] =='NOT'].shape[0]
            CURB = CUR[CUR['ELL'] =='ELLIGIBLE'].copy()
            b = CURB[CURB['DUE']== 'ALREADY'].shape[0]
            c = CURB[CURB['DUE']=='BLED'].shape[0]
            d = CURB[CURB['DUE']=='DUE'].shape[0]
            E =  b + c + f 
            G = int((E/a)*100)
            H = int((a*0.95)- E)
            data = {'TX_CURR' : [a],
            'NOT DUE' : [E],
            'VL COV' : [G],
            'BALANCE TO 95%' : [H],
            'DUE FOR VL' : [d]}
            PERFORMANCE = pd.DataFrame(data)
            PERFORMANCE = PERFORMANCE.set_index('TX_CURR')
            CURB = CURB.rename(columns = {'DUE' : 'VL STATUS'})
            linelist = CURB[['A', 'AS', 'RD', 'Ryear', 'Rmonth', 'Rday', 'VD', 'VL STATUS']].copy()
            linelist['AS'] = linelist['AS'].astype(str)
            linelist['AS'] = linelist['AS'].str.replace('*', '/')
            linelist['AS'] = linelist['AS'].str.replace('NaT', '')
            linelist['RD'] = linelist['RD'].astype(str)
            linelist['RD'] = linelist['RD'].str.replace('*', '/')
            linelist['RD'] = linelist['RD'].str.replace('NaT', '')
            linelist['VD'] = linelist['VD'].astype(str)
            linelist['VD'] = linelist['VD'].str.replace('*', '/')
            linelist['VD'] = linelist['VD'].str.replace('NaT', '')
            linelist['VL STATUS'] = linelist['VL STATUS'].astype(str)
            linelist = linelist[linelist['VL STATUS']== 'DUE']
            linelist = linelist.rename(columns = {'A': 'ART-NO', 'RD': 'RETURN DATE', 'VD': 'VIRAL LOAD DATE', 'AS':'ART START'})
            linelist[['Ryear', 'Rmonth', 'Rday']] = linelist[['Ryear', 'Rmonth', 'Rday']].apply(pd.to_numeric, errors='coerce')
            linelist = linelist.sort_values(by = ['Ryear', 'Rmonth', 'Rday'])
            bymonth = pd.pivot_table(linelist, index='Rmonth', values= 'ART-NO', aggfunc='count')
            bymonth = bymonth.reset_index()
            bymonth = bymonth.rename(columns={'ART-NO': 'DUE PER MONTH', 'Rmonth':'MONTH'})
            bymonth['MONTH'] = bymonth['MONTH'].astype(int)
            bymonth = bymonth.set_index('MONTH')
            bymonth = bymonth.transpose()
            st.write('**PERFORMANCE**')
            st.write(PERFORMANCE)
            st.write('**No. Not BLED IN EACH MONTH ON APPOINTMENT**')
            st.table(bymonth)
            weekly = weekly.rename(columns={'A':'REBLED'})
            cola, colb = st.columns([1,1])
            cola.write('**BLEEDS  DONE PER WEEK**')
            cola.write(weekly)
            colb.write('OF THOSE ON APPOINTMENT, HOW MANY HAVE BEEN BLED')
            colb.write(APPTSUM)
            st.markdown('**Sample linelist**')
            st.write(linelist.head(10))
    

if df is not None:
    def download_weekly(df):
        st.write(f"<h6>WEEKLY BLEEDS</h6>", unsafe_allow_html=True)

        if df is not None:
            dft = weekly.copy()
            csv_data = dft.to_csv(index=True)

                    # Create a download button for each facility

            st.download_button(
                        label="WEEKLY",
                        data=csv_data,
                        file_name=f"WEEKLY.csv",
                        mime="text/csv"
                    )


    def main():
        # Call the download functions
        download_weekly(df)


    if __name__ == "__main__":
        main()
if df is not None:
    if st.button('DOWNLOAD CURRENT LINELIST'):
        wb = Workbook()
        ws = wb.active

        # Convert DataFrame to Excel
        for r_idx, row in enumerate(linelist.iterrows(), start=1):
            for c_idx, value in enumerate(row[1], start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        ws.insert_rows(0)

        blue = PatternFill(fill_type = 'solid', start_color = 'C8CDCD')
        # ws.column_dimensions['H'].width = 14

        for num in range (1, ws.max_row+1):
            for letter in ['G', 'H']:
                ws[f'{letter}{num}'].font = Font(b= True, i = True)
                ws[f'{letter}{num}'].font = Font(b= True, i = True)
                ws[f'{letter}{num}'].fill = blue
                ws[f'{letter}{num}'].border = Border(top = Side(style = 'thin', color ='000000'),
                                                    right = Side(style = 'thin', color ='000000'),
                                                    left = Side(style = 'thin', color ='000000'),
                                                    bottom = Side(style = 'thin', color ='000000'))
        ws['A1'] ='ART NO.'
        ws['B1'] = 'ART START DATE'
        ws['C1'] = 'RETURN VISIT DATE'
        ws['D1'] = 'Ryear'
        ws['E1'] = 'Rmonth'
        ws['F1'] = 'Rday'
        ws['G1'] = 'VIRAL LOAD DATE' 
        ws['H1'] = 'VL STATUS'
        ws['I1']  = 'WAS CLIENT BLED (Y/N)'
        ws['J1'] = 'IF YES, HAS EMR BEEN UPDATED (Y/N)'


        letters = ['B', 'C', 'G']
        for letter in letters:
            ws.column_dimensions[letter].width =17

        ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 12
        ws['I1'].alignment = Alignment(wrap_text=True)
        ws.column_dimensions['J'].width = 20
        ws['J1'].alignment = Alignment(wrap_text=True)

        ran = np.random.rand()*0.1
        rand = round(ran,3)

        file_path = os.path.join(os.path.expanduser('~'), 'Downloads', f'VL LINELIST {rand}.xlsx')

        # Save the workbook
        wb.save(file_path)
        st.success('YOUR FILE HAS BEEN DOWNLOADED AS VL LINELIST {rand} IN YOUR DOWNLOAD FOLDER')
        























 
