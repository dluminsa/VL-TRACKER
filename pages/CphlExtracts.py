import streamlit as st 
import pandas as pd
import os
import random
#import numpy as np
from openpyxl import Workbook
from pathlib import Path
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
         'Kyazanga HC IV': 2048,'Kyetume HC III': 535, 'Lwengo HC IV': 1455, 'KINONI':2265,'Nanywa HC III':498,
         }

ENTEBBE = {'Bussi HC III': 237, 'Bweyogerere HC III': 969, 'BUNAMWAYA H-C II':30,'JCRC (Wakiso)':13376,'Kasenge H-C II':65, 'Kajjansi HC III':1962, 'Kasanje HC III': 823,
'Kigungu HC III':630, 'Kirinya H-C II':53, 
           'Kyengera HC III':620, 'Lufuka Valley HC III': 233, 'Mildmay Uganda HOSPITAL':14651, 'Mutundwe HC II':50,'Mutungo HC II':94, 'Nakawuka HC III':1068, 'Nalugala HC II':69,
           'Ndejje HC IV':2124, 'Nsangi HC III':2713, 'Seguku HC II':96, 'TASO Entebbe CLINIC' :6357, 'Wagagai HC IV': 524,'ZINGA HC II':260,'Kasoozo H-C III':33,'Katabi H-C III':123,
           'Kimwanyi H-C III':25, 'Kireka H-C II':61, 'KYENGEZA H-C II': 12, 'LUBBE H-C II':13, 'MAGANJO  H-C II':35, 'MAGOGGO H-C II': 18, 'Matugga H-C III':73,
           'Migadde H-C II':18, 'Namugongo Fund For Special Children': 606, 'NSAGGU H-C II':31, 'Nurture Africa H-C III':2405, 'Kitala HC II':165
}


districts = ['BUKOMANSIMBI', 'BUTAMBALA','ENTEBBE HUB', 'GOMBA','KALANGALA','KALUNGU', 'KYOTERA', 
             'LYANTONDE','LWENGO','MASAKA CITY', 'MASAKA DISTRICT', 'MPIGI', 'RAKAI', 'SEMBABULE', 'WAKISO HUB']
#st.write('BEING UPDATED')
#st.stop()

st.success('WELCOME, this app was developed by Dr. Luminsa Desire, for any concern, reach out to him at desireluminsa@gmail.com')
district = st.selectbox('Select a district:', districts, index=None)

file = st.file_uploader('Upload your CPHL extract here')


ext = None
if file is not None:
    # Get the file name
    fileN = file.name
    ext = os.path.splitext(fileN)[1]

df = None
if file and district is not None:
    if ext == '.csv':  # Compare with '.csv'
        df = pd.read_csv(file)
    else:
        st.write('This may not be a CPHL extract, it must be in CSV form.')
    # Display DataFrame
    if df is not None and district is not None:
        df['facility'] =  df['facility'].str.replace('/', '-')
        df['facility'] =  df['facility'].str.replace('Kinoni Welfare Medical Centre CLINIC', 'KINONI')
        df['facility'] =  df['facility'].str.replace('Mukwano Medical Centre CLINIC', 'Lukaya HC III')
        df['facility'] =  df['facility'].str.replace('St. Francis Maternity Home HC II', 'Lukaya HC III')
        df['facility'] =  df['facility'].str.replace('Teguzibirwa Dom Clinic', 'Lukaya HC III')
        if district  == 'BUKOMANSIMBI':
            fac = pd.DataFrame(list(BUKOMANSIMBI.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'SEMBABULE':
            fac = pd.DataFrame(list(SEMBABULE.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'MASAKA CITY':
            fac = pd.DataFrame(list(MASAKA_CITY.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'MASAKA DISTRICT':
            fac = pd.DataFrame(list(MASAKA_DISTRICT.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'KALUNGU':
            fac = pd.DataFrame(list(KALUNGU.items()), columns=['facility', 'Q2CURR'])
        elif district == 'MPIGI':
            fac = pd.DataFrame(list(MPIGI.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'BUTAMBALA':
            fac = pd.DataFrame(list(BUTAMBALA.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'GOMBA':
            fac = pd.DataFrame(list(GOMBA.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'KYOTERA':
            fac = pd.DataFrame(list(KYOTERA.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'RAKAI':
            fac = pd.DataFrame(list(RAKAI.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'KALANGALA':
            fac = pd.DataFrame(list(KALANGALA.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'LYANTONDE':
            fac = pd.DataFrame(list(LYANTONDE.items()), columns=['facility', 'Q2CURR'])
        elif district  == 'LWENGO':
            fac = pd.DataFrame(list(LWENGO.items()), columns=['facility', 'Q2CURR'])
        elif district == 'WAKISO HUB':
            fac = pd.DataFrame(list(WAKISO.items()), columns=['facility', 'Q2CURR'])
        elif district == 'ENTEBBE HUB':
            fac = pd.DataFrame(list(ENTEBBE.items()), columns=['facility', 'Q2CURR'])
        else:
            print('NO DISTRICT CHOSEN')

        coldist = fac['facility'].unique().tolist()
        colextr = df['facility'].unique().tolist()
        emrcolumns= ['A', 'RE', 'VOB']
        
        for column in coldist:
            if column not in colextr:
                st.write (f'**THIS EXTRACT DOES NOT HAVE FACILITIES IN {district}**')
                st.write('**You either uploaded a wrong exract or chose a wrong district, please try again!!**')
                st.stop()
            else:
                columns = fac['facility'].unique().tolist()
                df = df[df['facility'].isin(columns)].copy()
                df['ART-NUMERIC'] = df['art_number'].replace('[^0-9]','',regex=True)
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
                df[['Dyear', 'Dmonth', 'Dday']]= df[['Dyear', 'Dmonth', 'Dday']].apply(pd.to_numeric, errors='coerce')
                df = df[((df['Dyear']==2024) | ((df['Dyear']==2023) & (df['Dmonth']>9)))].copy()
                df = df.sort_values(by= ['Dyear', 'Dmonth', 'Dday'], ascending=False)

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
                facilities = df['facility'].unique()
                dfdups = df.copy()
                dfa = []
                for facility in facilities:
                    dfs = df[df['facility']==facility]
                    dfs = dfs.sort_values(by= ['Dyear', 'Dmonth', 'Dday'], ascending=False)
                    dfs['ART-NUMERIC'] =  pd.to_numeric(dfs['ART-NUMERIC'], errors='coerce') 
                    dfs = dfs.drop_duplicates(subset='ART-NUMERIC', keep='first')
                    dfs =dfs[['facility','ART-NUMERIC','art_number','date_collected','Dyear', 'Dmonth', 'Dday','result_numeric','SUP']]
                    name = f'{facility}'
                    dfa.append(dfs)
                dy = pd.concat(dfa) 
            
                dfnodups = dy.copy()  
                pivot = pd.pivot_table(dy, index='facility', values='ART-NUMERIC', aggfunc='count')
                dta = pivot.reset_index()
                dta = dta.rename(columns={'ART-NUMERIC':'BLEEDS'}) 
                dy['SUP'] = dy['SUP'].astype(str)
                NS = dy[(dy['SUP']== 'HLV') | (dy['SUP']=='LLV')].copy()
                NS[['Dyear', 'Dmonth']] = NS[['Dyear', 'Dmonth']].apply(pd.to_numeric, errors= 'coerce')
                NS = NS[((NS['Dyear']==2024)| ((NS['Dyear']==2023) & (NS['Dmonth']>9)))]
                HLV = NS[(NS['SUP']== 'HLV')].copy()
                LLV = NS[(NS['SUP']== 'LLV')].copy()
                pivo = pd.pivot_table(HLV, index='facility', values='ART-NUMERIC', aggfunc='count')
                dtb = pivo.reset_index()
                dtb = dtb.rename(columns={'ART-NUMERIC':'HLVs'})
                piv = pd.pivot_table(LLV, index='facility', values='ART-NUMERIC', aggfunc='count')
                dtc = piv.reset_index()
                dtc = dtc.rename(columns={'ART-NUMERIC':'LLVs'})
                dfa = pd.merge(fac,dta, on = 'facility', how = 'left')
                dfb = pd.merge(dfa,dtb, on = 'facility', how = 'left')
                dfc = pd.merge(dfb,dtc, on = 'facility', how = 'left')
            
                #file = r"C:\Users\Desire Lumisa\Desktop\New folder (2)\THISBP.csv"
                dfc[['Q2CURR', 'BLEEDS', 'HLVs', 'LLVs']] = dfc[['Q2CURR', 'BLEEDS', 'HLVs', 'LLVs']].apply(pd.to_numeric, errors='coerce')
                dfc['VL COV'] = (dfc['BLEEDS']*100)/ (dfc['Q2CURR'])
                dfc['VL COV'] = dfc['VL COV'].astype(int)
                dfc['BALANCE'] = (dfc['Q2CURR']*0.95)-(dfc['BLEEDS'])
                dfc['BALANCE'] = dfc['BALANCE'].astype(int)
                def achieve (v):
                    if v < 0:
                        return 0
                    else:
                        return v
                dfc['BALANCE TO 95%'] = dfc['BALANCE'].apply(achieve)
                dfc = dfc[['facility', 'Q2CURR', 'BLEEDS','VL COV','BALANCE TO 95%', 'HLVs', 'LLVs']]

                r = dfc['Q2CURR'].sum()
                t = dfc['BLEEDS'].sum()
                y = dfc['BALANCE TO 95%'].sum()
                u = dfc['HLVs'].sum()
                i = dfc['LLVs'].sum()
                o = int((t*100)/r)

                dfc.loc[len(dfc), 'facility'] = 'TOTAL'
                dfc.loc[len(dfc)-1, 'Q2CURR'] = r
                dfc.loc[len(dfc)-1, 'BLEEDS'] = t
                dfc.loc[len(dfc)-1, 'VL COV'] = o
                dfc.loc[len(dfc)-1, 'BALANCE TO 95%'] = y
                dfc.loc[len(dfc)-1, 'HLVs'] = u
                dfc.loc[len(dfc)-1, 'LLVs'] = i
    if df is not None:           
        dfe = dfc.set_index('facility')            
        st.write(dfe.head(2))     
    if df is not None:        
       # if st.button('DOWNLOAD FILE FOR VL COVERAGE ', key='active'):
                wb = Workbook()
                ws = wb.active
                # Convert DataFrame to Excel
                for r_idx, row in enumerate(dfc.iterrows(), start=1):
                    for c_idx, value in enumerate(row[1], start=1):
                                ws.cell(row=r_idx, column=c_idx, value=value)

                ws.insert_rows(0)
                ws['A1'] = 'FACILITY'
                ws['B1'] = 'Q2 CURR'
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
                        dfs = dfs[['facility', 'ART-NUMERIC', 'art_number', 'date_collected', 'Dyear', 'Dmonth', 'Dday', 'result_numeric']]
                        csv_data = dfs.to_csv(index=False)

                        # Create a download button for each facility
                        st.download_button(
                            label=f"Download CSV for {facility} without duplicates",
                            data=csv_data,
                            file_name=f"{facility}_data_without_duplicates.csv",
                            mime="text/csv"
                        )

        def download_without_duplicates(df):
            st.write(f"<h6>CSV FILES for {district} WITH DUPLICATES</h6>", unsafe_allow_html=True)

            if df is not None and district is not None:
                dft = dfdups.copy()
                uniques = dft['facility'].unique()

                # Create an expander to contain the download buttons
                with st.expander(f"Download files for {district} Facilities with duplicates)"):
                    for facility in uniques:
                        dfs = dft[dft['facility'] == facility]
                        dfs = dfs[['facility', 'ART-NUMERIC', 'art_number', 'date_collected', 'Dyear', 'Dmonth', 'Dday', 'result_numeric']]
                        csv_data = dfs.to_csv(index=False)

                        # Create a download button for each facility
                        st.download_button(
                            label=f"Download CSV for {facility} with duplicates",
                            data=csv_data,
                            file_name=f"{facility}_data_with_duplicates.csv",
                            mime="text/csv"
                        )

        def main():
            # Call the download functions
            download_with_duplicates(df)
            download_without_duplicates(df)

        if __name__ == "__main__":
            main()








