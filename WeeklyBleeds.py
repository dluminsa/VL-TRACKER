import pandas as pd 
import streamlit as st 
import os
import gspread
from pathlib import Path
import random
import plotly.express as px
import plotly.graph_objects as go
import traceback
import time
from streamlit_gsheets import GSheetsConnection
from datetime import datetime

st.set_page_config(
    page_title = 'NS TRACKER',
    page_icon =":bar_chart"
    )

#st.header('CODE UNDER MAINTENANCE, TRY AGAIN TOMORROW')
#st.stop()
cola,colb,colc = st.columns([1,3,1])
colb.subheader('TRACKING NS')

today = datetime.now()
todayd = today.strftime("%Y-%m-%d")# %H:%M")
wk = today.strftime("%V")
week = int(wk)-39
cola,colb = st.columns(2)
cola.write(f"**DATE TODAY:    {todayd}**")
colb.write(f"**CURRENT WEEK:    {week}**")
dd = int(week)
k = int(wk)


if 'tx' not in st.session_state:     
     try:
        #cola,colb= st.columns(2)
        conn = st.connection('gsheets', type=GSheetsConnection)
        exist = conn.read(worksheet= 'ALLNS', usecols=list(range(20)),ttl=5)
        tx = exist.dropna(how='all')
        st.session_state.tx = tx
     except:
         st.write("POOR NETWORK, COULDN'T CONNECT TO DELIVERY DATABASE")
         st.stop()
dfapt = st.session_state.tx.copy()
#######################FILTERS
if 'txa' not in st.session_state:     
     try:
        #cola,colb= st.columns(2)
        conn = st.connection('gsheets', type=GSheetsConnection)
        exist = conn.read(worksheet= 'SUP', usecols=list(range(12)),ttl=5)
        txa = exist.dropna(how='all')
        st.session_state.txa = txa
     except:
         st.write("POOR NETWORK, COULDN'T CONNECT TO DELIVERY DATABASE")
         st.stop()
dfr = st.session_state.txa.copy()
dfall = pd.read_csv('ALLNS.csv')
############################

clusters = dfall['CLUSTER'].unique()

#FILTERS
st.sidebar.subheader('**Filter from here**')
CLUSTER = st.sidebar.multiselect('CHOOSE A CLUSTER', clusters, key='a')

#create for the state
if not CLUSTER:
    dfr2 = dfr.copy()
    dfall2 = dfall.copy()
    dfapt2 = dfapt.copy() 
else:
    dfr2 = dfr[dfr['CLUSTER'].isin(CLUSTER)].copy()
    dfall2 = dfall[dfall['CLUSTER'].isin(CLUSTER)].copy()
    dfapt2 = dfapt[dfapt['CLUSTER'].isin(CLUSTER)].copy()
    
district = st.sidebar.multiselect('**CHOOSE A DISTRICT**', dfall2['DISTRICT'].unique(), key='d')
#create for the state
if not district:
    dfr3 = dfr2.copy()
    dfall3 = dfall2.copy()
    dfapt3 = dfapt2.copy() 
else:
    dfr3 = dfr2[dfr2['DISTRICT'].isin(district)].copy()
    dfall3 = dfall2[dfall2['DISTRICT'].isin(district)].copy()
    dfapt3 = dfapt2[dfapt2['DISTRICT'].isin(district)].copy()

facility = st.sidebar.multiselect('**CHOOSE A FACILITY**', dfall3['facility'].unique(), key='c')
if not facility:
    dfr4 = dfr3.copy()
    dfall4 = dfall3.copy()
    dfapt4 = dfapt3.copy() 
else:
    dfr4 = dfr3[dfr3['facility'].isin(facility)].copy()
    dfall4 = dfall3[dfall3['facility'].isin(facility)].copy()
    dfapt4 = dfapt3[dfapt3['facility'].isin(facility)].copy()


# Base DataFrame to filter
dfr = dfr4.copy()
dfall = dfall4.copy()
dfapt = dfapt.copy()
checkd = dfapt['DISTRICT'].unique()
st.write(checkd)

# Apply filters based on selected criteria
if CLUSTER:
    dfr = dfr[dfr['CLUSTER'].isin(CLUSTER)].copy()
    dfall = dfall[dfall['CLUSTER'].isin(CLUSTER)].copy()
    dfapt = dfapt[dfapt['CLUSTER'].isin(CLUSTER)].copy()


if district:
    dfr = dfr[dfr['DISTRICT'].isin(district)].copy()
    dfall = dfall[dfall['DISTRICT'].isin(district)].copy()
    dfapt = dfapt[dfapt['DISTRICT'].isin(district)].copy()

if facility:
    dfr = dfr[dfr['facility'].isin(facility)].copy()
    dfall = dfall[dfall['facility'].isin(facility)].copy()
    dfapt = dfapt[dfapt['facility'].isin(facility)].copy()
###FACILITIES THAT HAVEN'T UPLOADED EMR
s1 = dfall['facility'].unique()
s2 = dfapt['facility'].unique()
notemr = set(s1) - set(s2)

if facility:
    if facility not in s2:
        st.warning('**EMR EXTRACT FOR THIS FACILITY HAS NOT BEEN UPLOADED YET**')
        st.stop()
    else:
        pass

num = len(list(notemr))
if not facility:
    if num >0:
        if num >1:
            st.write(f'**{num} facilities have not uploaded their emr extarcts for tracking**')
            with st.expander('CLICK HERE TO SEE THEM'):
                st.write(notemr)
        if num ==1:
            st.write(f'**{num} facility has not uploaded their emr extarcts for tracking**')
            with st.expander('CLICK HERE TO SEE IT'):
                st.write(notemr)
    else:
        st.write('**ALL EMR EXTRACTS HAVE BEEN UPLOADED**')
##################NS THAT ARE DEAD
total = dfapt.shape[0]
dead = dfapt[dfapt['DD'].notna()]
dd = dead.shape[0]
########### REMAINING NS AFTER THE DEAD THEN TO
dfapt = dfapt[dfapt['DD'].isnull()].copy()
to = dfapt[dfapt['TO'].notna()]
totalto = to.shape[0]
dfapt = dfapt[dfapt['TO'].isnull()].copy()
#####ACTIVE
dfapt['Ryear'] = pd.to_numeric(dfapt['Ryear'], errors='coerce')
active = dfapt[dfapt['Ryear']==2025]
ager = active.copy()
ac = active.shape[0]
lost = dfapt[dfapt['Ryear']<2025]
los = lost.shape[0]
html_table = """
<h6><b><u style="color: purple;">QUICK SUMMARY</u></b></h6>
"""
st.markdown(html_table, unsafe_allow_html=True)

cola, colb,colc, cold, cole = st.columns(5)
cola.write('**TOTAL**')
colb.write('**ACTIVE**')
colc.write('**LTFU**')
cold.write('**T/O**')
cole.write('**DEAD**')

cola.write(f'**{total}**')
colb.write(f'**{ac}**')
colc.write(f'**{los}**')
cold.write(f'**{totalto}**')
cole.write(f'**{dd}**')
st.divider()
###########
html_table = """
<h6><b><u style="color: green;">REBLEEDING AMONGST THOSE THAT ARE ACTIVE</u></b></h6>
"""
st.markdown(html_table, unsafe_allow_html=True)
facz = active['facility'].unique()

dfnot = []
dfbled = []
for facilit in facz:
    dfa = active[active['facility']==facilit].copy()
    dfb = dfr[dfr['facility']==facilit].copy()

    dfa['ARTN'] = pd.to_numeric(dfa['ARTN'], errors = 'coerce')
    dfb['ART'] = pd.to_numeric(dfb['ART'], errors = 'coerce')

    dfbld = dfa[dfa['ARTN'].isin(dfb['ART'])]
    dfnt = dfa[~dfa['ARTN'].isin(dfb['ART'])]
    dfnot.append(dfnt)
    dfbled.append(dfbld)
if len(dfbled) == 0:
    bled = 0
else:
    dfbleds = pd.concat(dfbled)
    bled = dfbleds.shape[0]
dfnots = pd.concat(dfnot)

rebleds = dfbleds.shape[0]
dfnots[['Vmonth', 'Vyear']] = dfnots[['Vmonth', 'Vyear']].apply(pd.to_numeric, errors='coerce')
awr = dfnots[((dfnots['Vyear']>2024) | ((dfnots['Vyear'] == 2024) & (dfnots['Vmonth'] > 9)))].copy()
aw = awr.shape[0]
due = dfnots[((dfnots['Vyear']<2024) | ((dfnots['Vyear'] ==2024) & (dfnots['Vmonth'] <10)))].copy()
du = due.shape[0]

cola, colb,colc, cold = st.columns(4)
cola.write('**ACTIVE**')
colb.write('**REBLED (cphl)**')
colc.write('**AWR (emr)**')
cold.write('**DUE**')

cola.write(f'**{ac}**')
colb.write(f'**{bled}**')
colc.write(f'**{aw}**')
cold.write(f'**{du}**')
st.markdown('**KEY: AWR>> AWAITING RESULTS, HAS RECENT VL DATE IN EMR**')
st.divider()
html_table = """
<h6><b><u style="color: red;">FOR THOSE THAT ARE DUE, WHEN ARE THEY ON APPOINTMENT</u></b></h6>
"""
st.markdown(html_table, unsafe_allow_html=True)

today = datetime.now()
todayd = today.strftime("%Y-%m-%d")# %H:%M")

mon = today.strftime("%m")
mon = int(mon)

day = today.strftime("%d")
day = int(day)
wk = today.strftime("%V")
week = int(wk)-39

due[['Ryear', 'Rmonth', 'Rday', 'RWEEK']] = due[['Ryear', 'Rmonth', 'Rday', 'RWEEK']].apply(pd.to_numeric, errors='coerce')

tude = due[((due['Ryear']==2025) & (due['Rmonth']==mon) & (due['Rday']== day))].copy()

tud = tude.shape[0]

wiki = due[((due['Ryear']==2025) & (due['RWEEK']==week))].copy()
wik = wiki.shape[0]

jan = due[((due['Ryear']==2025) & (due['Rmonth']==1))].copy()
ja = jan.shape[0]

feb = due[((due['Ryear']==2025) & (due['Rmonth']==2))].copy()
fe = feb.shape[0]

marc = due[((due['Ryear']==2025) & (due['Rmonth']==3))].copy()
mar = marc.shape[0]

others = due[((due['Ryear']>2025) | ((due['Ryear']==2025) & (due['Rmonth']>4)))].copy()
other = others.shape[0]

cola, colb,colc, cold, cole, colf = st.columns(6)
cola.write('**TODAY**')
colb.write('**THIS WEEK**')
colc.write('**JAN**')
cold.write('**FEB**')
cole.write('**MARCH**')
colf.write('**OTHER Qtrs**')

cola.write(f'**{tud}**')
colb.write(f'**{wik}**')
colc.write(f'**{ja}**')
cold.write(f'**{fe}**')
cole.write(f'**{mar}**')
colf.write(f'**{other}**')

st.write('**DOWNLOADS**')
with st.expander('**CLICK HERE TO DOWNLOAD NS LINELIST ON APPOINTMENT**'):
    cola,colb = st.columns(2)
    optioniz = ['TODAY', 'THIS WEEK', 'JAN', 'FEB', 'MARCH', 'OTHER Qtrs']
    perd = colb.selectbox('**FILTER BY RETURN PERIOD**', optioniz, index=None)
    if perd =='TODAY':
        due = tude.copy()
    elif perd == 'THIS WEEK':
        due = wiki.copy()
    elif perd == 'JAN':
        due = jan.copy()
    elif perd == 'FEB':
        due = feb.copy()
    elif perd == 'MARCH':
        due = marc.copy()
    elif perd == 'OTHER Qtrs':
        due = others.copy()
    else:
        due = due.copy()
    
    due = due[['facility','ART','result_numeric', 'date_collected','RD', 'VD']].copy()
    due = due.rename(columns = {'RD': 'RETURN DATE', 'VD': 'VL DATE(EMR)'})
    due = due.reset_index()
    due = due.drop(columns = 'index')
    st.write(due.head(5))
    csv_data = due.to_csv(index=False)
    st.download_button(
                        label=" DOWNLOAD THIS DATA SET",
                        data=csv_data,
                        file_name="ACTIVITIES.csv",
                        mime="text/csv")
st.divider()
    
html_table = """
<h6><b><u style="color: purple;">FOR THOSE THAT WERE ON APPOINTMENT, LAST MONTH, HOW MANY WERE BLED</u></b></h6>
"""
st.markdown(html_table, unsafe_allow_html=True)
cola, colb,colc, cold, cole, colf = st.columns(6)
cola.write('**ON APPT**')
colb.write('**ATTENDED**')
colc.write('**MISSED**')
cold.write('**REBLED(cphl)**')
cole.write('**AWR(emr)**')
colf.write('**NOT BLED**')

cola.write(f'**7**')
colb.write(f'**0**')
colc.write(f'**0**')
cold.write(f'**0**')
cole.write(f'**0**')
colf.write(f'**0**')
with st.expander('**CLICK HERE TO DOWNLOAD NS LINELIST**'):
    st.write('BEING DEVELOPED')
st.divider()          
html_table = """
<h6><b><u style="color: blue;">AGE DISTRIBUTION FOR ACTIVE NS</u></b></h6>
"""
st.markdown(html_table, unsafe_allow_html=True)

ager = ager[['facility', 'ART', 'AG']].copy()
ager['AG'] = pd.to_numeric(ager['AG'], errors='coerce')

def band(x):
   if x < 10:
       return '0 to 9'
   elif x < 20:
       return '10 to 19'
   elif x < 30:
       return '20 to 29'
   elif x < 40:
       return '30 to 39'
   elif x < 50:
       return '40 to 49'
   else:
       return 'Above 50'
       

ager['BAND'] = ager['AG'].apply(band)
ager = ager.sort_values(by = 'AG')

fig = px.histogram(ager, x='BAND', text_auto=True,
                   title="Distribution of NS by their Age Bands",
                   labels={'BAND': 'Age Band', 'count': 'Number of IDs'})

# Customize layout
fig.update_layout(xaxis_title="Age Band",
                  yaxis_title="Count of NS",
                  bargap=0.1)

# Show the figure
st.plotly_chart(fig)
st.divider()        
html_table = """
<h6><b><u style="color: trend;">TREND LINE FOR REBLEEDING</u></b></h6>
"""
st.markdown(html_table, unsafe_allow_html=True)
         
            
