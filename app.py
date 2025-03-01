import pandas as pd
from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
import webbrowser
import threading
import os

app = Flask(__name__)

# Definir colores para los estados
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Tarde / Exceso de almuerzo
green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # Almuerzo dentro del tiempo
blue_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")  # Horas extras

def procesar_excel(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        print(f"El archivo {ruta_archivo} no existe.")
        return None

    engine = 'openpyxl' if ruta_archivo.endswith('.xlsx') else 'xlrd'
    
    try:
        df = pd.read_excel(ruta_archivo, engine=engine, header=1)
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return None
    
    df = df[['ID', 'Nombre', 'Apellido', 'Tiempo']]
    df['Tiempo'] = pd.to_datetime(df['Tiempo'])
    df['Fecha'] = df['Tiempo'].dt.date
    df['Hora'] = df['Tiempo'].dt.time
    
    output_df = pd.DataFrame(columns=[
        'ID', 'Nombre', 'Apellido', 'Fecha', 
        'Hora de Entrada', 'Hora de Almuerzo Salida', 
        'Hora de Almuerzo Entrada', 'Hora de Salida'
    ])

    for (id_persona, fecha), grupo in df.groupby(['ID', 'Fecha']):
        nombre = grupo.iloc[0]['Nombre']
        apellido = grupo.iloc[0]['Apellido']
        
        registros = grupo.sort_values('Tiempo')['Tiempo'].tolist()
        registros_filtrados = [registros[0]]

        for t in registros[1:]:
            if (t - registros_filtrados[-1]) > timedelta(minutes=3):
                registros_filtrados.append(t)
        
        entrada = almuerzo_salida = almuerzo_entrada = salida = None

        for i, t in enumerate(registros_filtrados):
            hora = t.time()
            if i == 0:
                entrada = hora
            elif i == len(registros_filtrados) - 1:
                salida = hora
            elif almuerzo_salida is None and (t - registros_filtrados[i-1]) > timedelta(minutes=45):
                almuerzo_salida = hora
            elif almuerzo_salida and almuerzo_entrada is None:
                almuerzo_entrada = hora

        output_df.loc[len(output_df)] = [
            id_persona, nombre, apellido, fecha, entrada, 
            almuerzo_salida, almuerzo_entrada, salida
        ]

    output_path = 'output.xlsx'

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Asistencia')
        workbook = writer.book
        worksheet = writer.sheets['Asistencia']

        # Ajustar el ancho de las columnas
        column_widths = {
            "A": 10,  # ID
            "B": 20,  # Nombre
            "C": 20,  # Apellido
            "D": 12,  # Fecha
            "E": 18,  # Hora de Entrada
            "F": 22,  # Hora de Almuerzo Salida
            "G": 22,  # Hora de Almuerzo Entrada
            "H": 18,  # Hora de Salida
        }
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

        # Aplicar colores según condiciones
        for row in worksheet.iter_rows(min_row=2, max_col=8, max_row=len(output_df)+1):
            try:
                cell_entrada = row[4]  # Hora de Entrada
                cell_almuerzo_salida = row[5]  # Almuerzo salida
                cell_almuerzo_entrada = row[6]  # Almuerzo entrada
                cell_salida = row[7]  # Hora de salida
                
                if cell_entrada.value:
                    hora_entrada = datetime.strptime(str(cell_entrada.value), '%H:%M:%S').time()
                    if hora_entrada > datetime.strptime('08:01:00', '%H:%M:%S').time():
                        cell_entrada.fill = red_fill
                
                if cell_almuerzo_salida.value and cell_almuerzo_entrada.value:
                    hora_almuerzo_salida = datetime.strptime(str(cell_almuerzo_salida.value), '%H:%M:%S').time()
                    hora_almuerzo_entrada = datetime.strptime(str(cell_almuerzo_entrada.value), '%H:%M:%S').time()
                    tiempo_almuerzo = (datetime.combine(datetime.min, hora_almuerzo_entrada) - datetime.combine(datetime.min, hora_almuerzo_salida)).seconds / 60 
                    if tiempo_almuerzo > 60:
                        cell_almuerzo_salida.fill = red_fill
                        cell_almuerzo_entrada.fill = red_fill
                    else:
                        cell_almuerzo_salida.fill = green_fill
                        cell_almuerzo_entrada.fill = green_fill
                
                if cell_salida.value:
                    hora_salida = datetime.strptime(str(cell_salida.value), '%H:%M:%S').time()
                    if hora_salida > datetime.strptime('19:00:00', '%H:%M:%S').time():
                        cell_salida.fill = blue_fill
            except Exception as e:
                print(f"Error al aplicar color en la fila {row[0].row}: {e}")

    return output_path

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        
        file_path = os.path.join("uploads", file.filename)
        file.save(file_path)
        
        output_path = procesar_excel(file_path)
        if output_path:
            return send_file(output_path, as_attachment=True)
    
    return render_template('upload.html')

if __name__ == '__main__':
    if not os.path.exists("uploads"):
        os.makedirs("uploads")
    app.run(host='0.0.0.0', port=5000, debug=True)

def abrir_navegador():
    webbrowser.open("http://127.0.0.1:5000/")  # URL de tu app Flask

if __name__ == '__main__':
    if not os.path.exists("uploads"):
        os.makedirs("uploads")

    threading.Timer(1.5, abrir_navegador).start()  # Espera 1.5 segundos antes de abrir el navegador
    app.run(host='0.0.0.0', port=5000, debug=True)