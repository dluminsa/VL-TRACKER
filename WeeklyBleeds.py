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

# st.set_page_config(
#     page_title = 'PROGRAM GROWTH',
#     page_icon =":bar_chart"
#     )

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
        exist = conn.read(worksheet= 'NS', usecols=list(range(8)),ttl=5)
        tx = exist.dropna(how='all')
        st.session_state.tx = tx
     except:
         st.write("POOR NETWORK, COULDN'T CONNECT TO DELIVERY DATABASE")
         st.stop()
dftx = st.session_state.tx.copy()
#st.write(dftx.columns)
#dftx[['DUEFN','SUPP', 'NOTFN', 'NOTRN', 'DUERN']] = dftx[['DUEFN','SUPP', 'NOTFN', 'NOTRN' 'DUERN']].apply(pd.to_numeric, errors='coerce')
dftx['DUETOTAL'] = dftx['DUEFN'] + dftx['DUERN']
dftx['TOTAL'] = dftx['DUEFN'] + dftx['DUERN'] +dftx['SUPP'] + dftx['NOTFN'] + dftx['NOTRN']
st.write(dftx)

#######################FILTERS
#
weeks = dftx['WEEK'].unique()

fac = dftx['FACILITY'].unique()
districts = dftx['DISTRICT'].unique()

#TO USE WHERE WEEKS ARE NOT NEEDED FOR TX
dfy = []
for every in fac:
    dff = dftx[dftx['FACILITY']== every]
    dff = dff.drop_duplicates(subset=['FACILITY'], keep = 'last')
    dfy.append(dff)
water = pd.concat(dfy)


#REMOVE DUPLICATES FROM TX SHEET # HOLD THIS IN SESSION LATER
dfs=[]   
for each in weeks:
    dftx['WEEK'] = pd.to_numeric(dftx['WEEK'], errors='coerce')
    dfa = dftx[dftx['WEEK']==each]
    dfa = dfa.drop_duplicates(subset=['FACILITY'], keep = 'last')
    dfs.append(dfa)
dftx = pd.concat(dfs)


#FILTERS
st.sidebar.subheader('**Filter from here**')
DISTRICT = st.sidebar.multiselect('CHOOSE A DISTRICT', districts, key='a')

#create for the state
if not DISTRICT:
    dftx2 = dftx.copy()
    water2 = water.copy() 
else:
    dftx['DISTRICT'] = dftx['DISTRICT'].astype(str)
    dftx2 = dftx[dftx['DISTRICT'].isin(DISTRICT)]
    
    water['DISTRICT'] = water['DISTRICT'].astype(str)
    water2 = water[water['DISTRICT'].isin(DISTRICT)]

facility = st.sidebar.multiselect('**CHOOSE A FACILITY**', dftx2['FACILITY'].unique(), key='c')
if not facility:
    dftx3 = dftx2.copy()
    water3 = water2.copy()
else:
    dftx3 = dftx2[dftx2['FACILITY'].isin(facility)].copy()
    water3 = water2[water2['FACILITY'].isin(facility)].copy()

# Base DataFrame to filter
dftx = dftx3.copy()
water = water3.copy()
# Apply filters based on selected criteria


if DISTRICT:
    water = water[water['DISTRICT'].isin(DISTRICT)].copy()
    dftx = dftx[dftx['DISTRICT'].isin(DISTRICT)].copy()

if facility:
    water = water[water['FACILITY'].isin(facility)].copy()
    dftx = dftx[dftx['FACILITY'].isin(facility)].copy()
    
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
if facility and not district:
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
st.write(BLED)
st.write(NOT)
labels = ['BLED', 'DUE']
values = [BLED, NOT]
# Specify custom colors
colors = ['darkblue', 'red']  # Colors for NO_MMD and MMD
# Create the 3D pie chart
figp = go.Figure(data=[go.Pie(
    labels=labels,
    values=values,
    hole=0.2,  # Creates a donut chart (0 for a full pie)
    textinfo='label+percent',  # Show labels and percentages
    pull=[0.1, 0],  # Slightly pull both slices for emphasis
    marker=dict(colors=colors)
)])

# Update layout for 3D effect
figp.update_traces(textposition='inside', textinfo='percent+label')
st.markdown(f'**{BLED} HAVE BEEN REBLED, {NOT} HAVE NOT**')
if facility and not district and not CLUSTER:
    st.write(f'**SHOWING DATA FOR {facility} facility**')
st.plotly_chart(figp, use_container_width=True)
#####################ONLY SHOWS WHEN THERE ARE MANY FACILITIES OR DISTRTICTS
dist = water['DISTRICT'].nunique()
fact = water['FACILITY'].nunique()
distc = water['DISTRICT'].unique()
factc = water['FACILITY'].unique()

#QUICK SUMMARY
#TOTAL NS
cola,colb,colc,cold,cole,colf = st.columns(6)
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

melted = grouped.melt(id_vars=['WEEK'], value_vars=['TOTAL DUE', 'RN', 'SUPP'],
                            var_name='CATEGORIES', value_name='Total')

# melted = grouped.melt(id_vars=['SURGE'], value_vars=['TWO', 'THREE', 'FOUR'],
#                             var_name='INTERVAL', value_name='Total')

#melted2 = grouped.melt(id_vars=['SURGE'], value_vars=['RTT', 'TO','DEAD'],
       #                     var_name='INTERVAL', value_name='Total')
melted['WEEK'] = melted['WEEK'].astype(int)
melted['WEEK'] = melted['WEEK'].astype(str)
#melted2['SURGE'] = melted2['SURGE'].astype(int)
#melted2['SURGE'] = melted2['SURGE'].astype(str)

fig2 = px.line(melted, x='WEEK', y='Total', color='CATEGORIES', markers=True,
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



fig3.update_layout(
    width=800,  # Set the width of the plot
    height = 400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
)
fig3.update_xaxes(type='category')
colx,coly = st.columns([2,1])
with colx:
    st.plotly_chart(fig2, use_container_width= True)

with coly:
    st.plotly_chart(fig3, use_container_width= True)
    #st.plotly_chart(fig3, use_container_width=True)
#############################################################################################
# #HIGHEST TXML 
st.divider()
highest = water[water['TWO']>100]

highest = highest.sort_values(by=['TWO'])#, ascending=False)
highesta = highest.shape[0]

highesty = water[water['TWO']<101]
highesty = highesty[highesty['TWO']>49]
highestb = highesty.sort_values(by=['TWO'], ascending=False)
highestb = highestb.shape[0]
# highestb = highest[highest['WEEK']==m]

coly, colu = st.columns(2)
with coly:
    if highesta ==0:
        st.warning('**FACILITY SELECTED IS NOT AMONG**')
        pass
    else:
        figa = px.bar(
        highest,
        x='TWO',
        y='FACILITY',
        orientation='h',
        title='FACILITIES WITH >100 MISSED APPTS',
        labels={'TWO': 'CLIENTS MISSED', 'FACILITY': 'Facility'}
            )
        figa.update_traces(marker_color='#be7869')
        st.plotly_chart(figa, use_container_width=True)
with colu:
    if highestb ==0:
        st.write('**FACILITY SELECTED IS NOT AMONG**')
        pass
    else:
        figa = px.bar(
        highesty,
        x='TWO',
        y='FACILITY',
        orientation='h',
        title='FACILITIES WITH 50-100 MISSED APPTS',
        labels={'TWO': 'CLIENTS MISSED', 'FACILITY': 'Facility'}
        )
        figa.update_traces(marker_color='green')
        st.plotly_chart(figa, use_container_width=True)
st.divider()
#MMD PERFORMANCE
#OF THOSE THAT ARE DUE, HOW MANY ARE OURS, HOW MANY ARE VISITOR

dftx[['M2','M3', 'M6']] = dftx[['M2','M3', 'M6']].apply(pd.to_numeric, errors='coerce')
M2 = water['M2'].sum()
M3 = water['M3'].sum()
M6 = water['M6'].sum()


# Creating the grouped bar chart
fig4 = go.Figure(data=[
    go.Bar(name='<3 MTHS', x=['<3 MTHS'], y=[M2], marker=dict(color='red')),
    go.Bar(name='3-5 MTHS', x=['3-5 MTHS'], y=[M3], marker=dict(color='green')),
    go.Bar(name='6+ MTHS', x=['6+ MTHS'], y=[M6], marker=dict(color='blue'))
])

# Setting the layout to have no gap between bars
fig4.update_layout(barmode='group', bargap=0, bargroupgap=0)

NO_MMD = M2
MMD = int(M3) + int(M6)

labels = ['NO MMD', 'MMD']
values = [NO_MMD, MMD]
# Specify custom colors
colors = ['DarkRed', 'purple']  # Colors for NO_MMD and MMD
# Create the 3D pie chart
figp = go.Figure(data=[go.Pie(
    labels=labels,
    values=values,
    hole=0.3,  # Creates a donut chart (0 for a full pie)
    textinfo='label+percent',  # Show labels and percentages
    pull=[0.1, 0],  # Slightly pull both slices for emphasis
    marker=dict(colors=colors)
)])

# Update layout for 3D effect
figp.update_traces(textposition='inside', textinfo='percent+label')

M2 = int(M2)
M3 = int(M3)
M6 = int(M6)
# Display the chart
st.success(f'**{M2} Clients were given < 3 Months, {M3} received between 4 to 5 five months, {M6} received 6+ MTHS**')
cola,colb = st.columns(2)
with cola:
    st.plotly_chart(fig4, use_container_width=True)

with colb:
    st.markdown('')
    st.markdown('')
    st.plotly_chart(figp, use_container_width=True)

st.divider()
##########################################################################
#######ONE YEAR COHORT
#filtered_df = filtered_df[filtered_df['WEEK']==k].copy()

total = wateryr['TOTAL'].sum()
newti = wateryr['TI'].sum()
newlydx = wateryr['ORIG'].sum()

newlos = wateryr['LOST'].sum()
newto  = wateryr['TO'].sum()
newdd = wateryr['DEAD'].sum()
active = wateryr['ACTIVE'].sum()

labels = ["NEWLY DX", "TIs", "TOTAL", 'LTFU',"TOs","DEAD", "ACTIVE"]
values = [newlydx, newti, total, -newlos, -newto,-newdd, active]
measure = ["absolute", "relative", "total", "relative", "relative", "relative","total"]
# Create the waterfall chart
figy = go.Figure(go.Waterfall(
    name="Waterfall",
    orientation="v",
    measure=measure,
    x=labels,
    textposition="outside",
    text=[f"{v}" for v in values],
    y=values
))

# Add titles and labels and adjust layout properties
figy.update_layout(
    title="ONE YEAR COHORT ANALYSIS",
    xaxis_title="Categories",
    yaxis_title="Values",
    showlegend=True,
    height=425,  # Adjust height to ensure the chart fits well
    margin=dict(l=20, r=20, t=60, b=20),  # Adjust margins to prevent clipping
    yaxis=dict(automargin=True)
)

st.plotly_chart(figy)
st.divider()
########################################################################################
#ONE YEAR PIE CHART
col1, col2 = st.columns(2)
pied = wateryr.copy()#[filtered_df['WEEK']==k]
#pied['LOST NEW'] = pied['ORIGINAL COHORT']- pied['ONE YEAR ACTIVE'] 
pied = pied[['LOST', 'ACTIVE']]
melted = pied.melt(var_name='Category', value_name='values')
fig = px.pie(melted, values= 'values', title='ONE YEAR RETENTION RATE', names='Category', hole=0.3,color='Category',  
             color_discrete_map={'LOST': 'red', 'ACTIVE': 'blue'} )
    #fig.update_traces(text = 'RETENTION', text_position='Outside')
grouped = dfyr2.groupby('SURGE').sum(numeric_only=True).reset_index()

melted = grouped.melt(id_vars=['SURGE'], value_vars=['ACTIVE', 'LOST'],
                            var_name='OUTCOME', value_name='Total')
melted['SURGE'] = melted['SURGE'].astype(int)
melted['SURGE'] = melted['SURGE'].astype(str)
colors = ['DarkRed', 'purple']

fig2 = px.line(melted, x='SURGE', y='Total', color='OUTCOME', markers=True,
              title='MISSED APPOINTMENTS', labels={'SURGE':'WEEK', 'Total': 'No. of clients', 'OUTCOME': 'OUTCOME'})

fig2.update_layout(
    width=800,  # Set the width of the plot
    height=400,  # Set the height of the plot
    xaxis=dict(showline=True, linewidth=1, linecolor='black'),  # Show x-axis line
    yaxis=dict(showline=True, linewidth=1, linecolor='black')   # Show y-axis line
    #marker=dict(colors=colors)
)


fig2.update_xaxes(type='category')
if pied.shape[0]==0:
    pass
else:
    with col1:
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.plotly_chart(fig2, use_container_width=True)

###############################
st.divider()
st.info('**EARLY RETENTION**')
st.write('**NOTE: The 9 months cohort will be the focus for 1 year cohort next quarter, the TX NEWs are the clients diagnosed this quarter**') 
pied = wateryr.copy()#[filtered_df['WEEK']==k]
#pied['LOST NEW'] = pied['ORIGINAL COHORT']- pied['ONE YEAR ACTIVE'] 
pied = pied[['LOSTS', 'ACTIVES']]
pied = pied.rename(columns={'LOSTS':'LOST', 'ACTIVES':'ACTIVE'})
melted = pied.melt(var_name='Category', value_name='values')
fig6 = px.pie(melted, values= 'values', title='6 MTHS', names='Category', hole=0.3,color='Category',  
             color_discrete_map={'LOST': 'red', 'ACTIVE': 'blue'} )

pied = waterly.copy()#[filtered_df['WEEK']==k]
#pied['LOST NEW'] = pied['ORIGINAL COHORT']- pied['ONE YEAR ACTIVE'] 
pied = pied[['LOSTT', 'ACTIVET']]
pied = pied.rename(columns={'LOSTT':'LOST', 'ACTIVET':'ACTIVE'})
melted = pied.melt(var_name='Category', value_name='values')
fig3 = px.pie(melted, values= 'values', title='3 MTHS', names='Category', hole=0.3,color='Category',  
             color_discrete_map={'LOST': 'red', 'ACTIVE': 'yellow'} )
#pied['LOST NEW'] = pied['ORIGINAL COHORT']- pied['ONE YEAR ACTIVE']
pied = waterly.copy() 
pied = pied[['LOSTO', 'ACTIVEO']]
pied = pied.rename(columns={'LOSTO':'LOST', 'ACTIVEO':'ACTIVE'})
melted = pied.melt(var_name='Category', value_name='values')
fig1 = px.pie(melted, values= 'values', title='TX NEWS', names='Category', hole=0.3,color='Category',  
             color_discrete_map={'LOST': 'red', 'ACTIVE': 'green'} )
pied = wateryr.copy() 
pied = pied[['LOSTN', 'ACTIVEN']]
pied = pied.rename(columns={'LOSTN':'LOST', 'ACTIVEN':'ACTIVE'})
melted = pied.melt(var_name='Category', value_name='values')
fig9 = px.pie(melted, values= 'values', title='9 MTHS', names='Category', hole=0.3,color='Category',  
             color_discrete_map={'LOST': 'red', 'ACTIVE': 'purple'} )

cola, colb, colc,cold = st.columns(4)
with cola:
    st.plotly_chart(fig9, use_container_width=True)
with colb:
    st.plotly_chart(fig6, use_container_width=True)
with colc:
    st.plotly_chart(fig3, use_container_width=True)
with cold:
    st.plotly_chart(fig1, use_container_width=True)
####TRACKING TXML
st.info('**TRENDS IN TXML (FOR CLIENTS THAT WERE REPORTED AS TXML LAST QUARTER)**')
grouped = dftx.groupby('SURGE').sum(numeric_only=True).reset_index()
grouped['SURGE'] = grouped['SURGE'].astype(int)  # Ensure SURGE is integer
grouped['SURGE'] = grouped['SURGE'].astype(str)# Convert SURGE to string

# Create the line chart using Plotly Express
figM = px.line(grouped, 
               x='SURGE', 
               y='TXML', 
               title='TXML FOR Q4', 
               labels={'SURGE': 'WEEK', 'TXML': 'No. of clients'},
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

html_table = """
<h4><b><u style="color: green;">CYCLE OF INTERUPTION AND RETURN TO ART (CIRA)</u></b></h4>
"""
st.markdown(html_table, unsafe_allow_html=True)

#LOST IN LESS THAN 3 MONTHS
lesl = watercira['L1'].sum() +  watercira['L10'].sum() + watercira['L20'].sum() + watercira['L30'].sum() + watercira['L40'].sum() +  watercira['L50'].sum() + watercira['LG50'].sum()
#LOST IN 3 to 5 MTHS
thrl = watercira['L13'].sum() +  watercira['L103'].sum() + watercira['L203'].sum() + watercira['L303'].sum() + watercira['L403'].sum() +  watercira['L503'].sum() + watercira['LG503'].sum()
#LOST IN 6 MTHS
sixl = watercira['L16'].sum() +  watercira['L106'].sum() + watercira['L206'].sum() + watercira['L306'].sum() + watercira['L406'].sum() +  watercira['L506'].sum() + watercira['LG506'].sum()

#ACTIVE THAN 3 MONTHS
lesA = watercira['A1'].sum() +  watercira['A10'].sum() + watercira['A20'].sum() + watercira['A30'].sum() + watercira['A40'].sum() +  watercira['A50'].sum() + watercira['AG50'].sum()
#ACTIVE IN 3 to 5 MTHS
thrA = watercira['A13'].sum() +  watercira['A103'].sum() + watercira['A203'].sum() + watercira['A303'].sum() + watercira['A403'].sum() +  watercira['A503'].sum() + watercira['AG503'].sum()
#ACTIVE IN 6 MTHS
sixA = watercira['A16'].sum() +  watercira['A106'].sum() + watercira['A206'].sum() + watercira['A306'].sum() + watercira['A406'].sum() +  watercira['A506'].sum() + watercira['AG506'].sum()

totallos = lesl + thrl + sixl
totalact = lesA + thrA + sixA
totalcira = totallos + totalact

# Creating the grouped bar chart
figC = go.Figure(data=[
    go.Bar(name='IIT(TOTAL)', x=['IIT(TOTAL)'], y=[totalcira], marker=dict(color='rgb(0, 71, 171)')),  # Cobalt Blue,
    go.Bar(name='RETURNED', x=['RETURNED'], y=[totalact], marker=dict(color='green'))
])

# Setting the layout to have no gap between bars
figC.update_layout(title = 'Returning clients to ART', barmode='group', bargap=0, bargroupgap=0)


###STACKED BAR CHART
totalact = lesA + thrA + sixA

# Define the values for the variables
a = lesA
b = thrA
c = sixA

#LESS THAN 1 YR
a1 = watercira['A1'].sum() + watercira['A13'].sum() + watercira['A16'].sum()
a2 = a1 + watercira['L1'].sum() + watercira['L13'].sum() + watercira['L16'].sum() 
if a2 ==0:
   a3 = 0
else:
   a3 = round(int((a1/a2)*100))
     
#LESS THAN 10 YRs
a21 = watercira['A10'].sum() + watercira['A103'].sum() + watercira['A106'].sum()
a22 = a21 + watercira['L10'].sum() + watercira['L103'].sum() + watercira['L106'].sum() 
if a22 ==0:
   a23 = 0
else:
   a23 = round(int((a21/a22)*100))

#LESS THAN 20 YRs
a31 = watercira['A20'].sum() + watercira['A203'].sum() + watercira['A206'].sum()
a32 = a31 + watercira['L20'].sum() + watercira['L203'].sum() + watercira['L206'].sum() 
if a32 ==0:
   a33 = 0
else:
   a33 = round(int((a31/a32)*100))

#LESS THAN 30 YRs
a41 = watercira['A30'].sum() + watercira['A303'].sum() + watercira['A306'].sum()
a42 = a41 + watercira['L30'].sum() + watercira['L303'].sum() + watercira['L306'].sum() 
if a42 ==0:
   a43 = 0
else:
   a43 = round(int((a41/a42)*100))

#LESS THAN 40 YRs
a51 = watercira['A40'].sum() + watercira['A403'].sum() + watercira['A406'].sum()
a52 = a51 + watercira['L40'].sum() + watercira['L403'].sum() + watercira['L406'].sum() 
if a52 ==0:
   a53 = 0
else:
   a53 = round(int((a51/a52)*100))
     
#LESS THAN 50 YRs
a61 = watercira['A50'].sum() + watercira['A503'].sum() + watercira['A506'].sum()
a62 = a61 + watercira['L50'].sum() + watercira['L503'].sum() + watercira['L506'].sum() 
if a62 ==0:
   a63 = 0
else:
   a63 = round(int((a61/a62)*100))

#GREATER THAN 50 YRs
a71 = watercira['AG50'].sum() + watercira['AG503'].sum() + watercira['AG506'].sum()
a72 = a71 + watercira['LG50'].sum() + watercira['LG503'].sum() + watercira['LG506'].sum() 
if a72 ==0:
   a73 = 0
else:
   a73 = round(int((a71/a72)*100))

# Create the stacked bar chart
figD = go.Figure()

# Add the bottom layer (a) in cobalt blue
figD.add_trace(go.Bar(
    name='<3 MONTHS',
    y=[a],
    x=['Variables'],
    marker_color='rgb(0, 71, 171)'  # Cobalt blue color
))

# Add the middle layer (b) in green
figD.add_trace(go.Bar(
    name='3-5 MONTHS',
    y=[b],
    x=['Variables'],
    marker_color='green'
))

# Add the top layer (c) in purple
figD.add_trace(go.Bar(
    name='6 + MONTHS',
    y=[c],
    x=['Variables'],
    marker_color='purple'
))

# Update the layout to make it a stacked bar chart
figD.update_layout(
    barmode='stack',
    title='Length of interruption before return',
    yaxis_title='Values'
)
cola, colb,colc, cold = st.columns([1,1,1,3])
cola.write('**AGE**')
colb.write('**Returned**')
colc.write('**IIT(Total)**')
cold.write('**Proportion CIRA Returned**')
cola.write('**<01**')
colb.write(f'**{int(a1)}**')
colc.write(f'**{int(a2)}**')
cold.write(f'**{int(a3)} %**')

cola.write('**1-9**')
colb.write(f'**{int(a21)}**')
colc.write(f'**{int(a22)}**')
cold.write(f'**{int(a23)} %**')

cola.write('**10-19**')
colb.write(f'**{int(a31)}**')
colc.write(f'**{int(a32)}**')
cold.write(f'**{int(a33)} %**')

cola.write('**20-29**')
colb.write(f'**{int(a41)}**')
colc.write(f'**{int(a42)}**')
cold.write(f'**{int(a43)} %**')

cola.write('**30-39**')
colb.write(f'**{int(a51)}**')
colc.write(f'**{int(a52)}**')
cold.write(f'**{int(a53)} %**')

cola.write('**40-49**')
colb.write(f'**{int(a61)}**')
colc.write(f'**{int(a62)}**')
cold.write(f'**{int(a63)} %**')

cola.write('**50+**')
colb.write(f'**{int(a71)}**')
colc.write(f'**{int(a72)}**')
cold.write(f'**{int(a73)} %**')
st.divider()

cola,colb = st.columns(2)
with cola:
    #st.markdown('**Returning clients to ART**')
    st.plotly_chart(figC, use_container_width=True)

with colb:
    #st.markdown('**Length of interruption before return**')
    st.plotly_chart(figD, use_container_width=True)

st.divider()

st.write('')
st.write('')
st.write('')
st.success('**CREATED BY Dr. LUMINSA DESIRE**')


