from flask import Flask, render_template, request, session, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'clave_secreta_excel_app'
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Crear directorio de uploads si no existe
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No se ha seleccionado ningún archivo'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No se ha seleccionado ningún archivo'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Leer archivo Excel
        try:
            df = pd.read_excel(filepath)
            
            # Guardar información en la sesión
            session['filepath'] = filepath
            
            # Intentar identificar las columnas por nombre o posición
            columns = df.columns.tolist()
            column_mapping = {}
            
            # Mapeo por posición (A=0, B=1, etc.)
            if len(df.columns) >= 23:  # Asegurarse de que hay suficientes columnas
                column_mapping = {
                    'MRU': df.columns[2],  # Columna C
                    'Canton': df.columns[5],  # Columna F
                    'Cuen': df.columns[8],  # Columna I
                    'Direccion': df.columns[9],  # Columna J
                    'Nombre_Cliente': df.columns[17],  # Columna R
                    'Identificacion': df.columns[18],  # Columna S
                    'Numero_Servicio': df.columns[19],  # Columna T
                    'Facturas_Pendientes': df.columns[20],  # Columna U
                    'Valor_Vencido': df.columns[22]  # Columna W
                }
            
            # Si no podemos mapear por posición, permitir al usuario elegir
            if not column_mapping:
                return jsonify({
                    'success': True,
                    'columns': columns
                })
            else:
                return jsonify({
                    'success': True,
                    'columns': columns,
                    'column_mapping': column_mapping
                })
            
        except Exception as e:
            return jsonify({'error': f'Error al procesar el archivo: {str(e)}'}), 500
    
    return jsonify({'error': 'Formato de archivo no permitido. Use archivos Excel (.xlsx, .xls)'}), 400

@app.route('/filter', methods=['POST'])
def filter_data():
    if 'filepath' not in session:
        return jsonify({'error': 'No hay archivo cargado'}), 400
    
    data = request.json
    mru_column = data.get('mru_column')
    canton_column = data.get('canton_column')
    mru_value = data.get('mru_value')
    canton_value = data.get('canton_value')
    
    output_columns = data.get('output_columns', {})
    
    if not mru_column or not canton_column:
        return jsonify({'error': 'Debe seleccionar las columnas de MRU y Cantón'}), 400
    
    try:
        df = pd.read_excel(session['filepath'])
        
        # Aplicar filtros si se proporcionan valores
        filtered_df = df
        if mru_value:
            filtered_df = filtered_df[filtered_df[mru_column].astype(str).str.contains(mru_value, case=False)]
        if canton_value:
            filtered_df = filtered_df[filtered_df[canton_column].astype(str).str.contains(canton_value, case=False)]
        
        # Calcular el total del valor vencido si la columna está en el output
        total_valor_vencido = None
        valor_column = output_columns.get('Valor_Vencido')
        if valor_column and valor_column in filtered_df.columns:
            # Convertir a numérico antes de sumar, ignorando errores
            total_valor_vencido = pd.to_numeric(filtered_df[valor_column], errors='coerce').sum()
        
        # Seleccionar solo las columnas de salida si se especifican
        if output_columns and all(col in df.columns for col in output_columns.values()):
            columns_to_show = list(output_columns.values())
            filtered_df = filtered_df[columns_to_show]
        
        # Convertir a diccionario para la respuesta JSON
        result = filtered_df.head(100).to_dict(orient='records')
        column_labels = {col: col for col in filtered_df.columns}
        
        # Si tenemos un mapeo de columnas, usarlo para las etiquetas
        if output_columns:
            column_labels = {v: k for k, v in output_columns.items()}
        
        return jsonify({
            'success': True,
            'total_rows': len(filtered_df),
            'data': result,
            'column_labels': column_labels,
            'total_valor_vencido': float(total_valor_vencido) if total_valor_vencido is not None else None
        })
    except Exception as e:
        return jsonify({'error': f'Error al filtrar datos: {str(e)}'}), 500

@app.route('/get_unique_values', methods=['POST'])
def get_unique_values():
    if 'filepath' not in session:
        return jsonify({'error': 'No hay archivo cargado'}), 400
    
    column_name = request.json.get('column_name')
    
    if not column_name:
        return jsonify({'error': 'Debe especificar una columna'}), 400
    
    try:
        df = pd.read_excel(session['filepath'])
        unique_values = df[column_name].astype(str).unique().tolist()
        
        return jsonify({
            'success': True,
            'values': unique_values
        })
    except Exception as e:
        return jsonify({'error': f'Error al obtener valores únicos: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)