import streamlit as st 
import gspread
from oauth2client.service_account import ServiceAccountCredentials 


scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\Desire Lumisa\Desktop\GSPREAD\vl-tracker-423614-751fa3295fc1.json", scope)

#creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\Desire Lumisa\Desktop\GSPREAD\vl-tracker-423614-751fa3295fc1.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1oXx9PN_Io9rkA-6p-bJHf29XNyw_fojupTzxtJAXPx8/edit#gid=0')

#values = sheet.row_values(1)
#print(values)