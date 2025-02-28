import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
from pandas import ExcelWriter
import os

# Definir colores para los estados
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")    # Tarde / Exceso de almuerzo
green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # Almuerzo dentro del tiempo
blue_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")   # Horas extras

def procesar_excel(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        print(f"El archivo {ruta_archivo} no existe.")
        return
    
    # Determinar el motor de lectura
    engine = 'openpyxl' if ruta_archivo.endswith('.xlsx') else 'xlrd'

    # Leer archivo de Excel
    try:
        # Usar header=1 para leer los encabezados de la segunda fila
        df = pd.read_excel(ruta_archivo, engine=engine, header=1)
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return
    
    # Verificar las columnas disponibles
    print("Columnas en el archivo:", df.columns.tolist())
    
    # Filtrar las columnas necesarias
    df = df[['ID de Evento', 'Nombre', 'Apellido', 'Tiempo']]
    
    # Convertir la columna 'Tiempo' a formato datetime
    df['Tiempo'] = pd.to_datetime(df['Tiempo'])
    
    # Extraer la fecha y la hora
    df['Fecha'] = df['Tiempo'].dt.date
    df['Hora'] = df['Tiempo'].dt.time
    
    # Crear un nuevo DataFrame para la salida
    output_df = pd.DataFrame(columns=['ID', 'Nombre', 'Apellido', 'Fecha', 'Hora de Entrada', 'Hora de Almuerzo Salida', 'Hora de Almuerzo Entrada', 'Hora de Salida'])
    
    for fecha, grupo_fecha in df.groupby('Fecha'):
        for id_persona, grupo_persona in grupo_fecha.groupby('ID de Evento'):
            nombre = grupo_persona.iloc[0]['Nombre']
            apellido = grupo_persona.iloc[0]['Apellido']
            
            entrada = almuerzo_salida = almuerzo_entrada = salida = None
            
            for _, row in grupo_persona.iterrows():
                hora = row['Hora']
                
                if datetime.strptime('06:00:00', '%H:%M:%S').time() <= hora <= datetime.strptime('08:00:00', '%H:%M:%S').time():
                    entrada = hora
                elif datetime.strptime('12:00:00', '%H:%M:%S').time() <= hora <= datetime.strptime('14:59:00', '%H:%M:%S').time():
                    almuerzo_salida = hora
                elif almuerzo_salida and datetime.strptime('13:00:00', '%H:%M:%S').time() <= hora <= datetime.strptime('15:59:00', '%H:%M:%S').time():
                    almuerzo_entrada = hora
                elif datetime.strptime('17:00:00', '%H:%M:%S').time() <= hora <= datetime.strptime('18:00:00', '%H:%M:%S').time():
                    salida = hora
            
            output_df.loc[len(output_df)] = [id_persona, nombre, apellido, fecha, entrada, almuerzo_salida, almuerzo_entrada, salida]
    
    # Guardar el DataFrame en un nuevo archivo de Excel
    with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Asistencia')
        workbook = writer.book
        worksheet = writer.sheets['Asistencia']
        
        for row in worksheet.iter_rows(min_row=2, max_col=8, max_row=len(output_df)+1):
            try:
                # Revisar la hora de entrada
                if row[4].value and row[4].value > datetime.strptime('08:00:00', '%H:%M:%S').time():
                    row[4].fill = red_fill
                
                # Revisar el tiempo de almuerzo
                if row[5].value and row[6].value:
                    tiempo_almuerzo = (datetime.combine(datetime.min, row[6].value) - datetime.combine(datetime.min, row[5].value)).seconds / 3600
                    if tiempo_almuerzo > 1:
                        row[5].fill = red_fill
                        row[6].fill = red_fill
                    else:
                        row[5].fill = green_fill
                        row[6].fill = green_fill
                
                # Revisar la hora de salida
                if row[7].value and row[7].value > datetime.strptime('18:00:00', '%H:%M:%S').time():
                    row[7].fill = blue_fill
            except:
                pass  # Prevenir errores si hay valores vacíos

# Ruta del archivo de entrada
ruta_archivo = './excel/xd.xlsx'

# Llamar a la función correctamente
procesar_excel(ruta_archivo)