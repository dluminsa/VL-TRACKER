import streamlit as st 
import pandas as pd
import os
import random
import numpy as np
import gspread
from openpyxl import Workbook
from pathlib import Path
import traceback
import time
from datetime import datetime, date
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
#from openpyxl import * #load_workbook
#from openpyxl.styles import *
# from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


SEMBABULE = {'Ssembabule HC IV':2321,'Kyabi HC III':536,'Ntuusi HC IV':968, 'Lwemiyaga HC III':1048,
            'Makoole HC II':252,'Mateete HC III':2367, 'Lwebitakuli Gvt HC III':607,'Ntete HC II':87,'Sembabule Kabaale HC II':77}

BUKOMANSIMBI = {'Butenga HC IV':1477,'Mirambi HC III':330,'Kagoggo HC II': 93,'Kisojjo HC II GOVT':77,'Bigasa HC III':842,
            'Kitanda HC III':326,"St. Mary'S Maternity Home HC III": 80,'Kingangazzi HC II':204
              }
KALUNGU = {'Lukaya Health Care Center-Uganda Cares HC II': 3676, 
           'Bukulula HC IV': 1212, 'Kalungu Kabaale HC II GOVT': 117,'Kalungu HC III': 797,'Kalungu Kasambya HC III GOVT':367,
           'Kiragga HC III':184,'Kiti HC III':186,'Kyamulibwa Gvt HC III':406,'Lukaya HC III':724,'MRC Kyamulibwa HC II':482}

LYANTONDE ={'Kabatema HC II':118,
           'Kabayanda HC II':95,'Kaliiro HC III':485,'Kasagama HC III':477,
           'Kinuuka  HC III':315,'Lyakajura HC II': 415,'Lyantonde Hospital':4190,'Mpumudde HC III':470}


MASAKA_CITY ={'Bukoto HC III':454,
             'Kiyumba HC IV':888,'Masaka Police HC II':418,'Mpugwe HC III':317,'Nyendo HC II':359,'TASO Masaka CLINIC':8093}

MASAKA_DISTRICT ={'Bukakata HC III':619,'Bukeeri HC III':357,'Masaka Buwunga HC III GOVT':341,'Kamulegu HC III':601,'Kyannamukaaka HC IV':1346}

MPIGI ={'Bunjako HC III':517,'Buwama HC III':2162,'Mpigi   HC IV':3124,'Bujuuko HC III':334,
       'Sekiwunga HC III':346,'Nsamu-Kyali HC III':201,'Butoolo HC III':833,'Buyiga HC II':155,'Kampiringisa HC III':266,
       'Ggolo HC III':309,'Kituntu HC III':333,'Dona Medical Centre HC III':151,'Nindye HC III':270,'Muduuma HC III':766,
       'Nabyewanga HC II':135,'Bukasa HC II':54,'Fiduga HC III':24,'Kiringente Epi Centre HC II':77,'St. Elizabeth Kibanga Ihu HC III':37}

BUTAMBALA ={'Bulo HC III': 737,
           'Butambala Epi Centre HC III GOVT':212,'Gombe HOSPITAL': 3411,'Kitimba HC III': 222,'Kyabadaza HC III':417,'Ngando HC III':417
          }

KYOTERA ={'Kabira HC III GOVT':1220,
          'Kabuwoko Gvt HC III':231,'Kakuuto HC IV':2672,'Kalisizo Hospital':4108,'Kasasa HC III': 870,'Kasaali HC III': 1232,
          'Kasensero HC II':1308,'Kayanja HC II Lwankoni GOVT':108,'Kirumba  HC III':326,'Kyebe HC III':707,'Lwankoni HC III':270,
          'Mutukula HC III':542,'Mitukula HC III':1108,'Nabigasa HC III':872,'Rakai Health Sciences Program CLINIC':2642}

RAKAI = {'Buyamba HC III':898,
         'Byakabanda HC III':183,'Kacheera HC III':476,'Kibaale HC II GOVT':567,'Kibanda HC III':309,'Kimuli HC III':595,'Kifamba HC III':327,
         'Kyalulangira HC III':369,'Lwamaggwa Gvt HC III':981,'Lwanda HC III':774,'Rakai Hospital':3138,'Rakai Kiziba HC II GOVT':408}

GOMBA= {'Buyanja  HC II (Gomba)': 113,'Gomba Kanoni HC III GOVT': 1295,'Kifampa HC III': 879,'Kisozi HC III GOVT':392,
        'Kyai HC III': 375,'Maddu HC IV': 2049,'Mamba HC II':317,'Mpenja HC III': 401,'Ngomanene HC II': 121}

WAKISO= { 'Bulondo HC III':317,'Busawamanze HC III':302,'Buwambo HC IV':960,
        'COMMUNITY HEALTH PLAN UGANDA':665,'Ggwatiro Nursing Home HC III':326,'Gombe (Wakiso) HC II':20,
         'Kabubbu HC II':682,'Kasangati HC IV':3067,'Kawanda HC III':986,'Kira HC III':1506,'Kiziba HC III':523,'Mende HC III':253,
         'Nabutiti HC III':209,'Nabweru HC III':1612,'Namayumba HC IV':1974,'Namulonge HC III':473,'Nansana HC II':121,
         'Nassolo Wamala HC II':209,'Triam Medical Centre CLINIC-NR':243,'Ttikalu HC III':421,'Wakiso Banda HC II GOVT':44,
         'Wakiso Epi Centre HC III GOVT':602,'Wakiso HC IV':3736,'Wakiso Kasozi HC III GOVT':232,'Watubba HC III':542,'Kakiri HC III':938}

KALANGALA= {'Bubeke HC III': 611,'Bufumira HC III': 405,'Bukasa HC IV': 1029, 'Bwendero HC III':1007,'Jaana HC II':13,
           'Kachanga Island HC II':219,'Kalangala HC IV':1443, 'Kasekulo HC II': 6,'Lujjabwa Island HC II': 345,'Lulamba HC III': 647,
           'Mazinga HC III': 524,'Mugoye HC III': 1131,'Mulabana HC II': 16,'Ssese Islands African Aids Project (SIAA':20}  

LWENGO = {'Katovu HC III':470, 'Kiwangala HC IV': 1623, 
         'Kyazanga HC IV': 2048,'Kyetume HC III': 535, 'Lwengo HC IV': 1455, 'Lwengo Kinoni Govt HC III':2265,'Nanywa HC III':498,
         }

ENTEBBE = {'Bussi HC III': 237, 'Bweyogerere HC III': 969, 'BUNAMWAYA H-C II':30,'JCRC (Wakiso)':13376,'Kasenge H-C II':65, 'Kajjansi HC III':1962, 'Kasanje HC III': 823,
'Kigungu HC III':630, 'Kirinya H-C II':53, 
           'Kyengera HC III':620, 'Lufuka Valley HC III': 233, 'Mildmay Uganda HOSPITAL':14651, 'Mutundwe HC II':50,'Mutungo HC II':94, 'Nakawuka HC III':1068, 'Nalugala HC II':69,
           'Ndejje HC IV':2124, 'Nsangi HC III':2713, 'Seguku HC II':96, 'TASO Entebbe CLINIC' :6357, 'Wagagai HC IV': 524,'ZINGA HC II':260,'Kasoozo H-C III':33,'Katabi H-C III':123,
           'Kimwanyi H-C III':25, 'Kireka H-C II':61, 'KYENGEZA H-C II': 12, 'LUBBE H-C II':13, 'MAGANJO  H-C II':35, 'MAGOGGO H-C II': 18, 'Matugga H-C III':73,
           'Migadde H-C II':18, 'Namugongo Fund For Special Children': 606, 'NSAGGU H-C II':31,
           'Nurture Africa H-C III':2405, 'Kitala HC II':165
}

filea = r'ALL.csv'
dfd = pd.read_csv(filea)
# districts = ['BUKOMANSIMBI', 'BUTAMBALA','ENTEBBE HUB', 'GOMBA','KALANGALA','KALUNGU', 'KYOTERA', 
#              'LYANTONDE','LWENGO','MASAKA_CITY', 'MASAKA_DISTRICT', 'MPIGI', 'RAKAI', 'SEMBABULE', 'WAKISO HUB']

districts = dfd['DISTRICT'].unique()
  # st.write('BEING UPDATED')
  # st.stop()

st.success('WELCOME, this app was developed by Dr. Luminsa Desire, for any concern, reach out to him at desireluminsa@gmail.com')


file = st.file_uploader('Upload your CPHL extract here')
if 'dist' not in st.session_state:
    st.session_state.dist = False

ext = None
if file is not None:
    # Get the file name
    fileN = file.name
    ext = os.path.splitext(fileN)[1]

df = None
if file is not None: 
    if ext == '.csv':  # Compare with '.csv'
        df = pd.read_csv(file)
    else:
        st.write('This may not be a CPHL extract, it must be in CSV form.')
    # Display DataFrame
    if df is not None:# and district is not None:
        df['facility'] =  df['facility'].str.replace('/', '-')
        #df['facility'] =  df['facility'].str.replace('Kinoni Welfare Medical Centre CLINIC', 'KINONI')
        df['facility'] =  df['facility'].str.replace('Mukwano Medical Centre CLINIC', 'Lukaya HC III')
        df['facility'] =  df['facility'].str.replace('St. Francis Maternity Home HC II', 'Lukaya HC III')
        df['facility'] =  df['facility'].str.replace('Teguzibirwa Dom Clinic', 'Lukaya HC III')
        
        dist = df['facility'].unique()

        if 'Ssembabule HC IV' in dist:
            st.write('**This extract is for BUKOMANSIMBI, SEMBABULE AND KALUNGU**')
            st.write('WHICH OF THE THREE DO YOU WANT TO ANALYSE?')
            district = st.radio('**CHOOSE ONE DISTRICT**', options=['BUKOMANSIMBI', 'KALUNGU', 'SEMBABULE'], index=None, horizontal=True)
        elif 'Bulo HC III'in dist:
            st.write('**This extract is for BUTAMBALA**')
            district = 'BUTAMBALA'
        elif 'Kifampa HC III' in dist:
            st.write('**This extract is for GOMBA**')
            district = 'GOMBA'
        elif 'Bubeke HC III' in dist:
            st.write('**This extract is for KALANGALA**')
            district = 'KALANGALA'
        elif 'Kakuuto HC IV' in dist:
            st.write('**This extract is for BOTH KYOTERA AND RAKAI**')
            st.write('WHICH OF THE TWO DO YOU WANT TO ANALYSE?')
            district = st.radio('**CHOOSE ONE DISTRICT**', options=['KYOTERA', 'RAKAI'], index=None, horizontal=True)
        elif 'Katovu HC III' in dist:
            st.write('**This extract is for LWENGO**')
            district = 'LWENGO'
        elif 'Kabatema HC II' in dist:
            st.write('**This extract is for LYANTONDE**')
            district = 'LYANTONDE'
        elif 'Kiyumba HC IV' in dist:
            st.write('**This extract is for BOTH MASAKA CITY AND MASAKA DISTRICT**')
            st.write('WHICH OF THE TWO DO YOU WANT TO ANALYSE?')
            district = st.radio('**CHOOSE ONE DISTRICT**', options=['MASAKA CITY', 'MASAKA DISTRICT'], index=None, horizontal=True)        
        elif 'Buwama HC III' in dist:
            st.write('**This extract is for MPIGI**')
            district = 'MPIGI'
        elif 'Bulondo HC III' in dist:
            st.write('**This extract is for WAKISO HUB**')
            district = 'WAKISO'
        elif 'Bussi HC III' in dist:
            st.write('**This extract is for ENTEBBE HUB**')
            district = 'WAKISO'
        else:
            st.write("**I can't determine the origin of this extract, kindly choose a district from below**")
            district = st.selectbox('Select a district:', districts, index=None)
         

        # if district  == 'BUKOMANSIMBI':
        #     fac = pd.DataFrame(list(BUKOMANSIMBI.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'SEMBABULE':
        #     fac = pd.DataFrame(list(SEMBABULE.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'MASAKA_CITY':
        #     fac = pd.DataFrame(list(MASAKA_CITY.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'MASAKA_DISTRICT':
        #     fac = pd.DataFrame(list(MASAKA_DISTRICT.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'KALUNGU':
        #     fac = pd.DataFrame(list(KALUNGU.items()), columns=['facility', 'Q3CURR'])
        # elif district == 'MPIGI':
        #     fac = pd.DataFrame(list(MPIGI.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'BUTAMBALA':
        #     fac = pd.DataFrame(list(BUTAMBALA.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'GOMBA':
        #     fac = pd.DataFrame(list(GOMBA.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'KYOTERA':
        #     fac = pd.DataFrame(list(KYOTERA.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'RAKAI':
        #     fac = pd.DataFrame(list(RAKAI.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'KALANGALA':
        #     fac = pd.DataFrame(list(KALANGALA.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'LYANTONDE':
        #     fac = pd.DataFrame(list(LYANTONDE.items()), columns=['facility', 'Q3CURR'])
        # elif district  == 'LWENGO':
        #     fac = pd.DataFrame(list(LWENGO.items()), columns=['facility', 'Q3CURR'])
        # elif district == 'WAKISO HUB':
        #     fac = pd.DataFrame(list(WAKISO.items()), columns=['facility', 'Q3CURR'])
        # elif district == 'ENTEBBE HUB':
        #     fac = pd.DataFrame(list(ENTEBBE.items()), columns=['facility', 'Q3CURR'])
        # else:
        #     st.write('NO DISTRICT CHOSEN')
            #print('NO DISTRICT CHOSEN'
            
        if district:
            fac = dfd[dfd['DISTRICT']==district].copy()
            facilities = fac['facility'].unique().tolist()
            facextr = df['facility'].unique().tolist()
            st.session_state.dist = False
           # emrcolumns= ['A', 'RE', 'VOB']
        
            for facility in facilities:
                # if facility not in facextr:
                #     st.write (f'**THIS EXTRACT DOES NOT HAVE FACILITIES IN {district}**')
                #     st.write('**You either uploaded a wrong exract or chose a wrong district, please try again!!**')
                #     st.stop()
                # else:
                    facilitys = fac['facility'].unique().tolist()
                    df = df[df['facility'].isin(facilitys)].copy()
                    df['ART'] = df['art_number'].replace('[^0-9]','',regex=True)
                    df['dCOL'] = df['date_collected'].astype(str)
                    
                    df['dCOL'] = df['dCOL'].str.replace('/', '*')
                    df['dCOL'] = df['dCOL'].str.replace('-', '*')
                    #df['dCOL'] = df['dCOL'].str.replace('/', '*')
                    
                  
                    
                    df[['Dyear', 'Dmonth', 'Dday']] = df['dCOL'].str.split('*', expand=True)
                    
                    df[['Dyear', 'Dmonth', 'Dday']]= df[['Dyear', 'Dmonth', 'Dday']].apply(pd.to_numeric, errors='coerce')
                    
                    df['Dyear'] = df['Dyear'].fillna(2022)
                    a = df[df['Dyear']>31].copy()
                    b = df[df['Dyear']<32].copy()
                    b = b.rename(columns={'Dyear': 'Dday1', 'Dday': 'Dyear'})
                    b = b.rename(columns={'Dday1': 'Dday'})
                    df = pd.concat([a,b])
                    # df['Dyear'] = df['Dyear'].astype(str)
                    # df['Dyear'] = df['Dyear'].str.replace('24', '2024', regex=False)
                    
                    df[['Dyear', 'Dmonth', 'Dday']]= df[['Dyear', 'Dmonth', 'Dday']].apply(pd.to_numeric, errors='coerce')
                    df['Dyear'] = df['Dyear'].replace(24, 2024, regex=False)
                    df = df[df['Dyear']>=2024].copy() #| ((df['Dyear']==2023) & (df['Dmonth']>9)))].copy()
                    df = df.sort_values(by= ['Dyear', 'Dmonth', 'Dday'], ascending=False)
                    dfhigh = df.copy()

                    def Viremia (x):
                        if 0<= x <= 200:
                            return 'Suppressed'
                        elif 201 <= x <= 399:
                            return 'LLV'
                        elif x == 400:
                            return 'suppressed'
                        elif 401 <= x <= 999:
                            return 'LLV'
                        elif x >= 1000:
                            return 'HLV'
                        else:
                            return None
                    
                    df['result_numeric'] = pd.to_numeric(df['result_numeric'],errors='coerce')
                    df['SUP']= df['result_numeric'].apply(Viremia)
                    #factys = dfd[dfd['DISTRICT']==district].copy()
                    facilities = df['facility'].unique()
                    dfdups = df.copy()
                    dfa = []
                    for facility in facilities:
                        dfs = df[df['facility']==facility]
                        if dfs.empty:
                           continue
                        dfs = dfs.sort_values(by= ['Dyear', 'Dmonth', 'Dday'], ascending=False)
                        dfs['ART'] =  pd.to_numeric(dfs['ART'], errors='coerce') 
                        dfs = dfs.drop_duplicates(subset='ART', keep='first')
                        dfs =dfs[['facility','ART','art_number','date_collected','Dyear', 'Dmonth', 'Dday','result_numeric','SUP']]
                        name = f'{facility}'
                        dfa.append(dfs)
                    #st.write(dfa[0])
                    dy = pd.concat(dfa) 
                
                    dfnodups = dy.copy()  
                    pivot = pd.pivot_table(dy, index='facility', values='ART', aggfunc='count')
                    dta = pivot.reset_index()
                    dta = dta.rename(columns={'ART':'BLEEDS'}) 
                    dy['SUP'] = dy['SUP'].astype(str)
                    NS = dy[(dy['SUP']== 'HLV') | (dy['SUP']=='LLV')].copy()
                    NS[['Dyear', 'Dmonth']] = NS[['Dyear', 'Dmonth']].apply(pd.to_numeric, errors= 'coerce')
                    NS = NS[NS['Dyear']==2024].copy()#| ((NS['Dyear']==2023) & (NS['Dmonth']>9)))]
                    HLV = NS[(NS['SUP']== 'HLV')].copy()
                    LLV = NS[(NS['SUP']== 'LLV')].copy()
                    pivo = pd.pivot_table(HLV, index='facility', values='ART', aggfunc='count')
                    dtb = pivo.reset_index()
                    dtb = dtb.rename(columns={'ART':'HLVs'})
                    piv = pd.pivot_table(LLV, index='facility', values='ART', aggfunc='count')
                    dtc = piv.reset_index()
                    dtc = dtc.rename(columns={'ART':'LLVs'})
                    dfa = pd.merge(fac,dta, on = 'facility', how = 'left')
                    dfb = pd.merge(dfa,dtb, on = 'facility', how = 'left')
                    dfc = pd.merge(dfb,dtc, on = 'facility', how = 'left')
                    #st.write(dfc)
                    #file = r"C:\Users\Desire Lumisa\Desktop\New folder (2)\THISBP.csv"
                    dfc[['Q3CURR', 'BLEEDS', 'HLVs', 'LLVs']] = dfc[['Q3CURR', 'BLEEDS', 'HLVs', 'LLVs']].apply(pd.to_numeric, errors='coerce')
                    dfc['VL COV'] = (dfc['BLEEDS']*100)/ (dfc['Q3CURR'])
                    dfc = dfc.dropna(subset=['VL COV'])
                    dfc['VL COV'] = dfc['VL COV'].astype(int)
                    dfc['BALANCE'] = (dfc['Q3CURR']*0.95)-(dfc['BLEEDS'])
                    dfc['BALANCE'] = dfc['BALANCE'].astype(int)
                    def achieve (v):
                        if v < 0:
                            return 0
                        else:
                            return v
                    dfc['BALANCE TO 95%'] = dfc['BALANCE'].apply(achieve)
                    dfc = dfc[['facility', 'Q3CURR', 'BLEEDS','VL COV','BALANCE TO 95%', 'HLVs', 'LLVs']]                                     
if df is not None and district is not None: 
        dfq =dfc.reset_index().copy()
        dfq = dfq[['facility', 'Q3CURR', 'BLEEDS','VL COV','BALANCE TO 95%', 'HLVs', 'LLVs']].copy()
        dfq['Q3CURR'] = dfq['Q3CURR'].astype(int)
        r = dfq['Q3CURR'].sum()
        #st.write(f'{r}, hello')
        t = dfq['BLEEDS'].sum()
        y = dfq['BALANCE TO 95%'].sum()
        u = dfq['HLVs'].sum()
        i = dfq['LLVs'].sum()
        o = int((t*100)/r)
        # #dfc= dfq.copy()
        dfq.loc[len(dfq), 'facility'] = 'TOTAL'
        #st.write(dfc)
        dfq.loc[len(dfq)-1, 'Q3CURR'] = int(r)
        dfq.loc[len(dfq)-1, 'BLEEDS'] = t
        dfq.loc[len(dfq)-1, 'VL COV'] = o
        dfq.loc[len(dfq)-1, 'BALANCE TO 95%'] = y
        dfq.loc[len(dfq)-1, 'HLVs'] = u
        dfq.loc[len(dfq)-1, 'LLVs'] = i
if df is not None and district is not None:           
        dfe = dfq.set_index('facility')
        dfe = dfe.sort_values(by = ['Q3CURR'])#, ascending=False)
        #with st.expander(f'**CLICK HERE TO VIEW VL COV FOR {district}**'):
        st.markdown(f'**VL COVERAGE FOR {district}**')
        #dfe = dfe.drop(columns=['index'])
        st.write(dfe)     
if df is not None and district is not None:       
       # if st.button('DOWNLOAD FILE FOR VL COVERAGE ', key='active'):
                wb = Workbook()
                ws = wb.active
                # Convert DataFrame to Excel
                for r_idx, row in enumerate(dfe.iterrows(), start=1):
                    for c_idx, value in enumerate(row[1], start=1):
                                ws.cell(row=r_idx, column=c_idx, value=value)

                ws.insert_rows(0)
                ws['A1'] = 'FACILITY'
                ws['B1'] = 'Q3 CURR'
                ws['C1'] = 'BLEEDS'
                ws['D1'] = 'VL COV'
                ws['E1'] = 'BALANCE TO 95%'
                ws['F1'] = 'HLVs'
                ws['G1'] = 'LLVs'
                
                max_row = ws.max_row
                ws.cell(row=max_row, column=1).alignment = Alignment(horizontal = 'center')
    
                ws.column_dimensions['A'].width = 20
                ws.column_dimensions['E'].width = 17
                ws.column_dimensions['B'].width = 10
                ws.column_dimensions['D'].width = 10

                #letters = ['A1', 'B1', 'C1', 'D1']
                

                letter = 'D'
                red = PatternFill(fill_type = 'solid', start_color = 'ff0000')
                yellow = PatternFill(fill_type = 'solid', start_color = 'ffff00')
                green = PatternFill(fill_type = 'solid', start_color = '04AA6D')

                for num in range(2, ws.max_row +1):
                    ws[f'{letter}{num}'].alignment = Alignment(horizontal='center')
                    if ws[f'{letter}{num}'].value <85:
                        ws[f'{letter}{num}'].fill = red
                    elif ws[f'{letter}{num}'].value <95:
                        ws[f'{letter}{num}'].fill = yellow   
                    else:
                        ws[f'{letter}{num}'].fill = green
                        ws[f'{letter}{num}'].border = Border(top= Side(style = 'thick'),
                                                        left= Side(style = 'thick'),
                                                        right= Side(style = 'thick'),
                                                        bottom= Side(style = 'thick')) 
                    
                blue = PatternFill(fill_type = 'solid', start_color = '80F5F5')
                letter = ['A1', 'B1', 'C1', 'D1','E1','F1','G1']
                for each in letter:
                    ws[f'{each}'].font = Font(b= True, i = True)
                    ws[f'{each}'].fill = blue
                    ws[f'{each}'].border = Border(top = Side(style = 'thin', color ='000000'),
                                                            right = Side(style = 'thin', color ='000000'),
                                                            left = Side(style = 'thin', color ='000000'),
                                                            bottom = Side(style = 'thin', color ='000000'))
              
                grey = PatternFill(fill_type = 'solid', start_color = 'ECF1F1')
                letter = ['A', 'B', 'C','E','F','G']
                for each in letter:
                    ws[f'{each}{max_row}'].font = Font(b= True, i = True)
                    ws[f'{each}{max_row}'].fill = grey
                    ws[f'{each}{max_row}'].border = Border(top = Side(style = 'thin', color ='000000'),
                                                            right = Side(style = 'thin', color ='000000'),
                                                            left = Side(style = 'thin', color ='000000'),
                                                            bottom = Side(style = 'thin', color ='000000'))



                ws.sheet_view.ShowGridLines = False        

                ran = random.random()
                rand = round(ran,2)
                file_path = os.path.join(os.path.expanduser('~'), 'Downloads', f'{district}VL_COV {rand}.xlsx')
                directory = os.path.dirname(file_path)
                Path(directory).mkdir(parents=True, exist_ok=True)

                  # Save the workbook
                wb.save(file_path)
                # Serve the file for download
                with open(file_path, 'rb') as f:
                      file_contents = f.read()           
                st.download_button(label=f'DONLOAD VL COV FOR {district} ', data=file_contents,file_name=f' {district} VL COV {rand}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
             
            
if df is not None and district is not None:
        def download_with_duplicates(df):
            st.write(f"<h6>CSV FILES for {district} WITH NO DUPLICATES</h6>", unsafe_allow_html=True)

            if df is not None and district is not None:
                dft = dfnodups.copy()
                uniques = dft['facility'].unique()

                # Create an expander to contain the download buttons
                with st.expander(f"Download files for {district} Facilities without duplicates"):
                     for facility in uniques:
                        dfs = dft[dft['facility'] == facility]
                        dfs = dfs[['facility', 'ART', 'art_number', 'date_collected', 'Dyear', 'Dmonth', 'Dday', 'result_numeric']]
                        csv_data = dfs.to_csv(index=False)

                        # Create a download button for each facility
                        st.download_button(
                            label=f"Download CSV for {facility} without duplicates",
                            data=csv_data,
                            file_name=f"{facility}_data_without_duplicates.csv",
                            mime="text/csv"
                        )
if df is not None and district is not None:
            dfw = dfhigh.copy()
            dfw['RDO'] = dfw['date_collected'].astype(str)
            dfw['result_numeric'] = pd.to_numeric(dfw['result_numeric'], errors='coerce')
            nsups = dfw[dfw['result_numeric']>999].copy()
            sups = dfw[dfw['result_numeric']<1000].copy()
            sups[['Dyear', 'Dmonth']] = sups[['Dyear', 'Dmonth']].apply(pd.to_numeric, errors='coerce')
            
            sups = sups.sort_values(by = ['Dyear', 'Dmonth'], ascending = True)
            firsts =[]
            nsups['facility'] = nsups['facility'].astype(str)
            sups['facility'] = sups['facility'].astype(str)
            
            sups['ART'] = pd.to_numeric(sups['ART'], errors='coerce')
            nsups['ART'] = pd.to_numeric(nsups['ART'], errors='coerce')

            nsups[['Dyear', 'Dmonth', 'Dday']] = nsups[['Dyear', 'Dmonth', 'Dday']].apply(pd.to_numeric, errors='coerce')
            nsups = nsups.sort_values(by = ['Dyear', 'Dmonth', 'Dday'], ascending=False)

            dus =[]
            
            for facility in facilities:
                nsups['facility'] = nsups['facility'].astype(str)
                dfq = nsups[nsups['facility']==facility].copy()
                dfq[['Dyear', 'Dmonth', 'Dday']] = dfq[['Dyear', 'Dmonth', 'Dday']].apply(pd.to_numeric, errors='coerce')
                dfq = dfq.sort_values(by = ['Dyear', 'Dmonth', 'Dday'], ascending=False)
                dfq['ART'] = pd.to_numeric(dfq['ART'],errors='coerce')
                dfz = dfq[dfq.duplicated(subset=['ART'], keep='last')]
                dus.append(dfz)
            dups = pd.concat(dus)
           
            dupsa =[]
            for facility in facilities:
                dups['facility'] = dups['facility'].astype(str)
                dfx = dups[dups['facility']==facility].copy()
                dfx['ART'] = pd.to_numeric(dfx['ART'],errors='coerce')
                dfy = dfx.drop_duplicates(subset=['ART'], keep='first')
                dupsa.append(dfy)
            dups = pd.concat(dupsa)
            
            dups['REBLED'] = np.nan
            dups['REBLED'] = dups['REBLED'].fillna('RN')
            
            notd = []
            #NS WHO ARE NOT DUPS
            for facility in facilities:
                nsups['facility'] = nsups['facility'].astype(str)
                dups['facility'] = dups['facility'].astype(str)
                dfx = nsups[nsups['facility']==facility].copy()
                dfm = dups[dups['facility']==facility].copy()
                        
                dfx['ART'] = pd.to_numeric(dfx['ART'],errors='coerce')
                dfm['ART'] = pd.to_numeric(dfm['ART'],errors='coerce')
                #dfy = dfx[~dfx.duplicated(subset=['ART'])]#, keep='first')]
                dfy = dfx[~dfx['ART'].isin(dfm['ART'])]        
                notd.append(dfy)
            notdups =pd.concat(notd)
            #ppp = notdups.copy()

            nodups = []
            for facility in facilities:
                sups['facility'] = sups['facility'].astype(str)
                dfx = sups[sups['facility']==facility].copy()
                dfx['ART'] = pd.to_numeric(dfx['ART'],errors='coerce')
                dfy = dfx.drop_duplicates(subset=['ART'], keep='first')
                nodups.append(dfy)
            sups = pd.concat(nodups)
            
            dfj = []
            for facility in facilities:
                sups['facility'] = sups['facility'].astype(str)
                dfa = sups[sups['facility'] == facility].copy()
                
                notdups['facility'] = notdups['facility'].astype(str)
                dfb = notdups[notdups['facility'] == facility].copy()
                
                dfa['ART'] = pd.to_numeric(dfa['ART'],errors='coerce')
                dfb['ART'] = pd.to_numeric(dfb['ART'],errors='coerce')
                dfy = pd.merge(dfa, dfb, on = 'ART', how= 'right')
                dfj.append(dfy)
            dfa = pd.concat(dfj)
            
            fna =dfa[dfa['RDO_x'].isnull()].copy()
            dn =dfa[~dfa['RDO_x'].isnull()].copy()
            dn[['Dyear_x', 'Dyear_y']] = dn[['Dyear_x', 'Dyear_y']].apply(pd.to_numeric, errors ='coerce')
            dn['YEAR'] = dn['Dyear_x']-dn['Dyear_y']
            dn[['Dmonth_x', 'Dmonth_y']] = dn[['Dmonth_x', 'Dmonth_y']].apply(pd.to_numeric, errors ='coerce')
            dn['MONTH'] = dn['Dmonth_x']- dn['Dmonth_y']
            dn['YEAR'] = pd.to_numeric(dn['YEAR'], errors='coerce')
            dn['MONTH'] = pd.to_numeric(dn['MONTH'], errors='coerce')
            fnb = dn[dn['YEAR']<0].copy()
            rsa = dn[dn['YEAR']>0].copy()
            fnc = dn[((dn['YEAR']==0)& (dn['MONTH'] <0))].copy()
            rsb = dn[((dn['YEAR']==0)& (dn['MONTH'] >0))].copy()
            fnd = dn[((dn['YEAR']==0)& (dn['MONTH'] ==0))].copy()
            fn = pd.concat([fna,fnb, fnc,fnd])
            rs = pd.concat([rsa,rsb])
            fn['REBLED'] = np.nan
            fn['REBLED'] = fn['REBLED'].fillna('FN')
            rs['REBLED'] = np.nan
            rs['REBLED'] = rs['REBLED'].fillna('RS')
            dfa = pd.concat([fn,rs])
            dfa = dfa.rename(columns= ({'art_number_y':'art_number', 'facility_y': 'facility', 'date_collected_y':'date_collected', 'result_numeric_y':'result_numeric',
                         'RDO_y':'RDO', 'Dyear_y':'Dyear', 'Dmonth_y':'Dmonth', 'Dday_y':'Dday', 'SUP_y':'SUP'}))
            dfa = dfa[['ART', 'art_number','facility', 'date_collected', 'result_numeric', 'RDO', 'Dyear','Dmonth', 'Dday', 'REBLED']].copy()
            dfsupd = pd.concat([dfa, dups])
            dfsupd['Dmonth'] = pd.to_numeric(dfsupd['Dmonth'], errors='coerce')
            dfsupa = dfsupd[dfsupd['Dmonth']<7].copy()
            dfsupb = dfsupd[dfsupd['Dmonth']>6].copy()
            dfsupa['REBLED'] = dfsupa['REBLED'].astype(str)
            dfsupa['DUE'] = dfsupa['REBLED'].str.replace('FN', 'DUE')
            
            dfsupb['DUE'] = np.nan
            dfsupb['DUE'] = dfsupb['DUE'].fillna('NOT')
            dfsupd = pd.concat([dfsupa,dfsupb], axis=0)
            dfsupd = dfsupd[['facility', 'ART', 'art_number', 'date_collected', 'result_numeric', 'REBLED','DUE']]
            

    # Prepare the credentials dictionary
secrets = st.secrets["connections"]["gsheets"]
credentials_info = {
        "type": secrets["type"],
        "project_id": secrets["project_id"],
        "private_key_id": secrets["private_key_id"],
        "private_key": secrets["private_key"],
        "client_email": secrets["client_email"],
        "client_id": secrets["client_id"],
        "auth_uri": secrets["auth_uri"],
        "token_uri": secrets["token_uri"],
        "auth_provider_x509_cert_url": secrets["auth_provider_x509_cert_url"],
        "client_x509_cert_url": secrets["client_x509_cert_url"]
    }
current_time = time.localtime()
week = time.strftime("%V", current_time)
week = int(week)-39
if df is not None and district is not None: 
        try:
            # Define the scopes needed for your application
            scopes = ["https://www.googleapis.com/auth/spreadsheets",
                        "https://www.googleapis.com/auth/drive"]              
                 
            credentials = Credentials.from_service_account_info(credentials_info, scopes=scopes)
                    
            # Authorize and access Google Sheets
            client = gspread.authorize(credentials)
                    
            # Open the Google Sheet by URL
            spreadsheetu = "https://docs.google.com/spreadsheets/d/1oXx9PN_Io9rkA-6p-bJHf29XNyw_fojupTzxtJAXPx8/edit?gid=1448429519#gid=1448429519"     
            spreadsheet = client.open_by_url(spreadsheetu)
            sheet1 = spreadsheet.worksheet("NS")
        except Exception as e:
                    # Log the error message
            st.write(f"CHECK: {e}")
            st.write(traceback.format_exc())
            st.write("COULDN'T CONNECT TO GOOGLE SHEET, TRY AGAIN")
            st.stop()
if df is not None and district is not None:                     
            if not st.session_state.dist:
                    try:
                        facys = dfsupd['facility'].unique()
                        for facility in facys:
                                    row1 = []
                                    row1.append(district)
                                    row1.append(facility)
                                    row1.append(week)
                                    dfsupd['facility'] = dfsupd['facility'].astype(str)
                                    dfk = dfsupd[dfsupd['facility']==facility].copy()
                                    dfk['REBLED'] = dfk['REBLED'].astype(str)
                                    sup = dfk[dfk['REBLED']=='RS']
                                    su = sup.shape[0]
                                    row1.append(su)
                                    
                                    notdue = dfk[dfk['DUE']=='NOT'].copy()
                                    notdue['REBLED'] = notdue['REBLED'].astype(str)
                                    notfn = notdue[notdue['REBLED'] =='FN']
                                    nofn = notfn.shape[0]
                                    row1.append(nofn)
                                    notrn = notdue[notdue['REBLED'] =='RN']
                                    norn = notrn.shape[0]
                                    row1.append(norn)
                                    
                                    duedue = dfk[dfk['DUE']=='DUE'].copy()
                                    duedue['REBLED'] = duedue['REBLED'].astype(str)
                                    duefn = duedue[duedue['REBLED'] =='FN']
                                    dufn = duefn.shape[0]
                                    row1.append(dufn)
                                    duern = duedue[duedue['REBLED'] =='RN']
                                    durn = duern.shape[0]
                                    row1.append(durn)
                                    sheet1.append_row(row1, value_input_option='RAW')          
                        st.session_state.dist = True
# st.success('Your data above has been submitted')
        except Exception as e:
            # Print the error message
            st.write(f"ERROR: {e}")
            st.stop()  # Stop the Streamlit app here to let the user manually retry     
if df is not None and district is not None:       
        def download_without_duplicates(df):
            st.write(f"<h6>CSV FILES for NS IN {district}</h6>", unsafe_allow_html=True)

            if df is not None and district is not None:
                dft = dfsupd.copy()
                #dft = ppp.copy()
                #dft = dft.rename(columns = {'facility_y': 'facility'})
                uniques = dft['facility'].unique()
                # Create an expander to contain the download buttons
                with st.expander(f"DOWNLOAD NON SUPPRESORS FOR {district})"):
                    for facility in uniques:
                        dfs = dft[dft['facility'] == facility]
                        dfs = dfs[['facility', 'ART', 'art_number', 'date_collected', 'result_numeric', 'REBLED','DUE']]
                        csv_data = dfs.to_csv(index=False)

                        # Create a download button for each facility
                        st.download_button(
                            label=f"Download NS for {facility}",
                            data=csv_data,
                            file_name=f"{facility}_NS.csv",
                            mime="text/csv"
                        )

        def main():
            # Call the download functions
            download_with_duplicates(df)
            download_without_duplicates(df)

        if __name__ == "__main__":
            main()

