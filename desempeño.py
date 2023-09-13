import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import numpy as np
from email.message import EmailMessage
from datetime import datetime
import ssl
import smtplib
import cred2

creds_path = r"Credenciales/cred.json"
st.set_page_config(
    page_title="DESEMPEÑO | NOMINA ",
    page_icon="💵",
    layout="wide")
# Definir las credenciales de la cuenta de servicio de Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
gc = gspread.authorize(creds)

hojas = ["DATOS","Respuestas de formulario 1", "EVA_SUP", ]