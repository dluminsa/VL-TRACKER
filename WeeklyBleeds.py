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
notemr = set(s1)- set(s2)

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
st.stop()
watervl = water.copy() 
dfvl = dftx.copy()
check = water.shape[0]
if check == 0:
    st.warning('***NO DATA FOR THE SELECTION MADE**')
    st.stop()
else:
    pass
#st.write(water.columns)
st.divider()


mostd = water.groupby('DISTRICT')['DUETOTAL'].sum()
mostf = water.groupby('FACILITY')['DUETOTAL'].sum()

####TOP3
topdis3 = mostd.nlargest(3)
topdis3 = topdis3.reset_index()
mostdis3 = ','.join(topdis3['DISTRICT'].unique())

topfas3 = mostf.nlargest(3)
topfas3 = topfas3.reset_index()
mostfas3 = ','.join(topfas3['FACILITY'].unique())

##TOP 2
topdis2 = mostd.nlargest(2)
topdis2 = topdis2.reset_index()
mostdis2 = ','.join(topdis2['DISTRICT'].unique())

##TOP 1
# topfas3 = mostf.nlargest(3)
# topfas3 = topfas3.reset_index()
# mostfas3 = ','.join(topfas3['FACILITY'].unique())



checkf = water['FACILITY'].nunique()
checkd = water['DISTRICT'].nunique()
if facility and not DISTRICT:
    pass
elif checkf <3:
    pass
elif checkd >3:
    st.success(f'**DISTRICS WITH MOST UNBLED NS {mostdis3}, MOST AFFECTED FACILITIES  ARE {mostfas3}**')
elif checkd ==2:
    st.success(f'**DISTRICS WITH MOST UNBLED NS ARE {mostdis2}, MOST AFFECTED FACILITIES ARE {mostfas3}**')
elif checkd ==1:
    st.success(f'**FACILITIES WITH THE MOST UNBLED NS ARE {mostfas3}**')

        
st.divider()
############################################################################################

BLED = int(water['SUPP'].sum()) + int(water['NOTRN'].sum())
NOT = int(water['DUERN'].sum()) + int(water['DUEFN'].sum())
     
BLED = int(BLED)
NOT = int(NOT)

labels = ['BLED', 'DUE']
values = [BLED, NOT]
# Specify custom colors
colors = ['darkblue', 'red']  # Colors for NO_MMD and MMD
# Create the 3D pie chart
figp = go.Figure(data=[go.Pie(
    labels=labels,
    values=values,
    hole=0.1,  # Creates a donut chart (0 for a full pie)
    textinfo='label+percent',  # Show labels and percentages
    pull=[0.1, 0],  # Slightly pull both slices for emphasis
    marker=dict(colors=colors)
)])

# Update layout for 3D effect
figp.update_traces(textposition='inside', textinfo='percent+label')
st.markdown(f'**{BLED} HAVE BEEN REBLED, {NOT} HAVE NOT**')
if facility and not DISTRICT:# and not CLUSTER:
    st.write(f'**SHOWING DATA FOR {facility} facility**')
st.plotly_chart(figp, use_container_width=True)
#####################ONLY SHOWS WHEN THERE ARE MANY FACILITIES OR DISTRTICTS
dist = water['DISTRICT'].nunique()
fact = water['FACILITY'].nunique()
distc = water['DISTRICT'].unique()
factc = water['FACILITY'].unique()

#QUICK SUMMARY
#TOTAL NS
cola,colb,colc,cold,cole,colf = st.columns([2,1,1,1,1,1])
if int(dist)>1:
   cola.write('**DISTRICT**')
   colb.write('**TOTAL**')
   colc.write('**BLED**')
   cold.write('**NOT_BLED**')
   cole.write("**SUPP'SSD**")
   colf.write(f'**NOT**')
   for distr in distc:
      watera = water[water['DISTRICT']==distr].copy()
      tot = int(watera['TOTAL'].sum())
      bled = int(watera['SUPP'].sum()) + int(watera['NOTRN'].sum())
      notbled = int(watera['DUERN'].sum()) + int(watera['DUEFN'].sum())
      sups = int(watera['SUPP'].sum())
      notsups = int(watera['NOTRN'].sum()) + int(watera['DUERN'].sum())
      cola.write(f'**{distr}**')
      colb.write(f'**{tot}**')
      colc.write(f'**{bled}**')
      cold.write(f'**{notbled}**')
      cole.write(f'**{sups}**')
      colf.write(f'**{notsups}**')
        
if int(dist)==1:
   cola.write('**FACILITY**')
   colb.write('**TOTAL**')
   colc.write('**BLED**')
   cold.write('**NOT_BLED**')
   cole.write("**SUPP'SSD**")
   colf.write(f'**NOT**')
   for facil in factc:
      watera = water[water['FACILITY']==facil].copy()
      tot = int(watera['TOTAL'].sum())
      bled = int(watera['SUPP'].sum()) + int(watera['NOTRN'].sum())
      notbled = int(watera['DUERN'].sum()) + int(watera['DUEFN'].sum())
      sups = int(watera['SUPP'].sum())
      notsups = int(watera['NOTRN'].sum()) + int(watera['DUERN'].sum()) 
      cola.write(f'**{facil}**')
      colb.write(f'**{tot}**')
      colc.write(f'**{bled}**')
      cold.write(f'**{notbled}**')
      cole.write(f'**{sups}**')
      colf.write(f'**{notsups}**') 

if int(dist) > 1:
    st.divider()
    x = []
    y = []
    water['DUETOTAL'] = pd.to_numeric(water['DUETOTAL'], errors='coerce')
    districts = water['DISTRICT'].unique()
    water = water.sort_values(by = ['DUETOTAL'], ascending = False)
    for each in districts:
        x.append(each)
        dist = water[water['DISTRICT']==each]['DUETOTAL'].sum()
        y.append(dist)   
 
    sorted_indices = sorted(range(len(y)), key=lambda i: y[i], reverse=True)
    x = [x[i] for i in sorted_indices]
    y = [y[i] for i in sorted_indices]
    num_bars = len(x)
    colors = [f'rgba({random.randint(0, 255)}, {random.randint(0, 255)}, {random.randint(0, 255)}, 0.7)' for _ in range(num_bars)]
    
    figd = go.Figure(data=[
        go.Bar(x=x, y=y, marker_color=colors)
    ])
    
    # Update layout
    figd.update_layout(
        title='NS DUE FOR REBLEEDING',
        xaxis_title='District',
        yaxis_title='TOTAL DUE FOR REBLEEDING',
        xaxis_tickangle=-45  # Optional: angle x-axis labels for better visibility
    )
    st.plotly_chart(figd)#, use_container_width=True)
elif int(fact) > 1:
    st.divider()
    x = []
    y = []
    water['DUETOTAL'] = pd.to_numeric(water['DUETOTAL'], errors='coerce')
    districts = water['FACILITY'].unique()
    water = water.sort_values(by = ['DUETOTAL'], ascending = False)
    for each in districts:
        x.append(each)
        dist = water[water['FACILITY']==each]['DUETOTAL'].sum()
        y.append(dist)   
 
    sorted_indices = sorted(range(len(y)), key=lambda i: y[i], reverse=True)
    x = [x[i] for i in sorted_indices]
    y = [y[i] for i in sorted_indices]
    num_bars = len(x)
    colors = [f'rgba({random.randint(0, 255)}, {random.randint(0, 255)}, {random.randint(0, 255)}, 0.7)' for _ in range(num_bars)]
    
    figd = go.Figure(data=[
        go.Bar(x=x, y=y, marker_color=colors)
    ])
    
    # Update layout
    distict = '.'.join(water['DISTRICT'].unique())
    figd.update_layout(
        title=f'TOTAL NS DUE FOR REBLEEDING IN {distict}',
        xaxis_title='FACILITIES',
        yaxis_title='TOTAL DUE',
        xaxis_tickangle=-45  # Optional: angle x-axis labels for better visibility
    )
    st.plotly_chart(figd)#, use_container_width=True)
    
else:
    pass    

#############################################################################################
#LINE GRAPHS
st.divider()
#TREND OF MISSED APOINTMENTS
st.success('**TRENDS IN CLIENTS DUE, REBLED SUPPRESSING AND NOT**')
dfq =dftx.copy()
dfq = dfq.rename(columns = {'DUETOTAL': 'TOTAL DUE', 'NOTRN': 'RN'})
grouped = dfq.groupby('WEEK').sum(numeric_only=True).reset_index()

melted = grouped.melt(id_vars=['WEEK'], value_vars=['TOTAL DUE', 'SUPP'],#'RN', 'SUPP'],
                            var_name='CATEGORIES', value_name='Total')

# melted = grouped.melt(id_vars=['SURGE'], value_vars=['TWO', 'THREE', 'FOUR'],
#                             var_name='INTERVAL', value_name='Total')

#melted2 = grouped.melt(id_vars=['SURGE'], value_vars=['RTT', 'TO','DEAD'],
       #                     var_name='INTERVAL', value_name='Total')
melted['WEEK'] = melted['WEEK'].astype(int)
melted['WEEK'] = melted['WEEK'].astype(str)
#melted2['SURGE'] = melted2['SURGE'].astype(int)
#melted2['SURGE'] = melted2['SURGE'].astype(str)

fig2 = px.line(melted, x='WEEK', y='Total', color='CATEGORIES', markers=True,color_discrete_sequence=['red','black'],
              title='REBLEEDING TRENDS', labels={'WEEK':'WEEK', 'Total': 'No. of clients', 'INTERVAL': 'CATEGORIES'})

#fig3 = px.line(melted2, x='SURGE', y='Total', color='INTERVAL', markers=True, color_discrete_sequence=['black','red', 'yellow'],
             # title='RTT VS TO VS DEAD', labels={'SURGE':'WEEK', 'Total': 'No. of clients', 'INTERVALS': 'VARIABLES'})

fig2.update_layout(
    width=800,  # Set the width of the plot
    height=400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
)
fig2.update_xaxes(type='category')
st.plotly_chart(fig2, use_container_width= True)

###############################

html_table = """
       <h4><b><u style="color: maroon;">VL SECTION</u></b></h4>
     """
cola,colb,colc = st.columns([1,2,1])
colb.markdown(html_table, unsafe_allow_html=True)

#PIE CHART
BLED = int(watervl['BLED'].sum()) #+ int(water['NOTRN'].sum())
NOT = int(watervl['Q3'].sum()) - int(water['BLED'].sum())
     
BLED = int(BLED)
NOT = int(NOT)

labels = ['BLED', 'DUE']
values = [BLED, NOT]
#st.divider()
col1, col2,col3 = st.columns([1,4,1])
# Values
colors = ['blue', 'red']
# Creating the pie chart with specified colors and hole
fig = go.Figure(data=[go.Pie(labels=labels, values=values, textinfo='label+value', 
                             insidetextorientation='radial', marker=dict(colors=colors), hole=0.4)])

# Updating the layout for better readability
fig.update_traces(textposition='inside', textfont_size=20)
fig.update_layout(title_text='VL COVERAGE AT CPHL', title_x=0.3)

col1, col2,col3 = st.columns([1,4,1])
with col2:
     st.plotly_chart(fig, use_container_width=True)
#
####TRACKING TXML
st.info('**TRENDS IN VL COVERAGE**')
grouped = dfvl.groupby('WEEK').sum(numeric_only=True).reset_index()
grouped['WEEK'] = grouped['WEEK'].astype(int)  # Ensure SURGE is integer
grouped['WEEK'] = grouped['WEEK'].astype(str)# Convert SURGE to string

# Create the line chart using Plotly Express
figM = px.line(grouped, 
               x='WEEK', 
               y='COV', 
               title='VL COVERAGE TRENDS', 
               labels={'WEEK': 'WEEK', 'COV': 'VL coverage'},
               markers=True)

# Update trace color to red
figM.update_traces(line=dict(color='red'))

# Update layout for better appearance
figM.update_layout(
    width=800,  # Set the width of the plot
    height=400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
)

# Set x-axis to categorical
figM.update_xaxes(type='category')

# Display the plot
st.plotly_chart(figM, use_container_width=True)
st.divider()
st.write(' ')
st.write(' ')
st.write(' ')
st.write(' ')
st.info('@ LUMINSA DESIRE')
