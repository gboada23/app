import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import os
import streamlit as st
import numpy as np

creds_path = r"Credenciales/cred.json"
st.set_page_config(
    page_title="INCIDENCIAS | NOMINA ",
    page_icon="💵",
    layout="wide")
# Definir las credenciales de la cuenta de servicio de Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
gc = gspread.authorize(creds)
sheets = ['ASIS_SUP', 'FIN_SEMANA', 'FIN_SUPERVISORES', 'JORNADAS_ESPECIALES', 'ESPECIALES_SUPERVISORES']
dfs = []

primer = gc.open("EVALUACIONES COLECTIVAS").worksheet("ASIS_OP").get_values("C:K")
headers = primer.pop(0)
df = pd.DataFrame(primer, columns=headers)

df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')

dfs.append(df)
for sheet in sheets:
    data = gc.open("EVALUACIONES COLECTIVAS").worksheet(sheet).get_values("B:E")
    headers2 = data.pop(0)
    df = pd.DataFrame(data, columns=headers2)
    ## convertir a datetime las columnas de cada DF
    df['FECHA'] = pd.to_datetime(df['FECHA'], format='%d/%m/%Y')
   

        
    dfs.append(df)

# SACAMOS CADA DF DE LA POSICION QUE SE ENCUENTRA EN LA LISTA
FIN_OP = dfs[2]
FIN_SUP = dfs[3]
ESPECIALES_OP = dfs[4]
ESPECIALES_SUP = dfs[5]

# OPERAMOS CADA DF

INCIDENCIAS = dfs[0].loc[:,["Date","Operario","Regional", "Agencia","Asistencia", "Novedad","SOPORTE","VER"]]
INCIDENCIAS["SOPORTE"].replace('', np.nan, inplace=True)
INCIDENCIAS.loc[INCIDENCIAS['SOPORTE'].isnull(), 'VER'] = ' '
INCIDENCIAS = INCIDENCIAS[(INCIDENCIAS["Asistencia"]!= "Asistente") & (INCIDENCIAS["Asistencia"]!= "Vacaciones")]
INCIDENCIAS.loc[INCIDENCIAS['Asistencia'] == 'Inasistente', 'VER'] = ' '

#INCIDENCIAS_SUP = ASIS_SUP.loc[:,["FECHA", "SUPERVISOR","COORDINADOR","ASISTENCIA","NOVEDAD"]]
INCIDENCIAS_SUP = dfs[1][dfs[1]["ASISTENCIA"]!="Asistencia"]
@st.cache_data()
def nomina(inicio, final): 
    inicio = pd.to_datetime(inicio).strftime('%d/%m/%Y')
    final = pd.to_datetime(final).strftime('%d/%m/%Y')
    incidencias = INCIDENCIAS[(INCIDENCIAS["Date"]  >= inicio) & (INCIDENCIAS["Date"]  <= final)] 
    incidencias_sup = INCIDENCIAS_SUP[(INCIDENCIAS_SUP["FECHA"] >=inicio) & (INCIDENCIAS_SUP["FECHA"] <=final)]
    fin_operarios = FIN_OP[(FIN_OP["FECHA"] >=inicio) & (FIN_OP["FECHA"] <=final)]
    fin_supervisores = FIN_SUP[(FIN_SUP["FECHA"] >=inicio) & (FIN_SUP["FECHA"] <=final)]
    especiales_operario = ESPECIALES_OP[(ESPECIALES_OP["FECHA"] >=inicio) & (ESPECIALES_OP["FECHA"] <=final)]
    especiales_supervisores = ESPECIALES_SUP[(ESPECIALES_SUP["FECHA"] >=inicio) & (ESPECIALES_SUP["FECHA"] <=final)]

    incidencias["Date"] = incidencias['Date'].dt.strftime('%d/%m/%Y').astype(str)
    incidencias_sup["FECHA"] = incidencias_sup["FECHA"].dt.strftime('%d/%m/%Y').astype(str)
    fin_operarios["FECHA"] = fin_operarios["FECHA"].dt.strftime('%d/%m/%Y').astype(str)
    fin_supervisores["FECHA"] = fin_supervisores["FECHA"].dt.strftime('%d/%m/%Y').astype(str)
    especiales_operario["FECHA"] = especiales_operario["FECHA"].dt.strftime('%d/%m/%Y').astype(str)
    especiales_supervisores["FECHA"] = especiales_supervisores["FECHA"].dt.strftime('%d/%m/%Y').astype(str)
    return incidencias, incidencias_sup, fin_operarios, fin_supervisores, especiales_operario, especiales_supervisores
# Definimos la aplicación de Streamlit

def main():
    st.title("Sistema de incidencias Multiservicios Serviplus")
    st.write("#")
    # Crea el formulario para ingresar las fechas de inicio y final
    col1, col2 = st.columns((3,3))
    inicio = col1.date_input('Fecha de inicio', value=pd.to_datetime('today') - pd.Timedelta(days=10))
    final = col2.date_input('Fecha de finalización',value=pd.to_datetime('today'))
    
        # Ejecutamos la función de nomina con las fechas ingresadas
    incidencias, incidencias_sup, fin_operarios, fin_supervisores, especiales_operario, especiales_supervisores = nomina(inicio, final)

        # Creamos un objeto de Pandas ExcelWriter para guardar los 6 dataframes en un mismo archivo de Excel
    with pd.ExcelWriter('Nomina.xlsx') as writer:
        incidencias.to_excel(writer, sheet_name='Incidencias',index=False)
        incidencias_sup.to_excel(writer, sheet_name='Incidencias Supervisores',index=False)
        fin_operarios.to_excel(writer, sheet_name='Fin de Semana Operarios',index=False)
        fin_supervisores.to_excel(writer, sheet_name='Fin de semana Supervisores',index=False)
        especiales_operario.to_excel(writer, sheet_name='Especiales Operarios',index=False)
        especiales_supervisores.to_excel(writer, sheet_name='Especiales Supervisores',index=False)
    
           
        # Ejecutamos la función de nómina con las fechas ingresadas desde streamlit
    incidencias, incidencias_sup, fin_operarios, fin_supervisores, especiales_operario, especiales_supervisores = nomina(inicio, final)
        
    col1, col2 = st.columns((4,4))
    with col1:
        st.write(f"{len(incidencias)} Incidencias de Operarios")
        st.dataframe(incidencias)
    with col2:
        st.write(f"{len(fin_operarios)} datos de fin de semana o feriado de Operarios")
        st.dataframe(fin_operarios)
    st.divider()
    col1, col2 = st.columns((4,4))
   
    with col1:
        st.write(f"{len(incidencias_sup)} Incidencias de Supervisores")
        st.dataframe(incidencias_sup)
    with col2:
        st.write(f"{len(fin_supervisores)} Incidencias Fin de semana o Feriado de supervisores")
        st.dataframe(fin_supervisores)
    st.divider()
    col1, col2 = st.columns((4,4))
    with col1:
        st.write(f"{len(especiales_operario)} Jornadas especiales Operarios")
        st.dataframe(especiales_operario)
    with col2:
        st.write(f"{len(especiales_supervisores)} Jornadas especiales Supervisores")
        st.dataframe(especiales_supervisores)
    col1, col2, col3 = st.columns((3,3,3))
    with col2:
        st.warning(f"Para descargar las incidencias del {inicio} hasta el {final} presiona el boton")
        with open('Nomina.xlsx', 'rb') as f:  
         bytes_data = f.read()
        col1, col2, col3 = st.columns(3)
        col2.download_button(label="Descargar Todo", data=bytes_data, file_name=f'Incidencias del {inicio} al {final}.xlsx', mime='application/vnd.ms-excel')     
    # Muestra un mensaje de éxito al usuario
    
# Ejecuta la aplicación de Streamlit
if __name__ == '__main__':
    main()
