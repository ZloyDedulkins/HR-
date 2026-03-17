from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    result_table = None
    if request.method == 'POST':
        file = request.files['file']
        if file:
            # Чтение Excel
            fired = pd.read_excel(file, sheet_name=0)
            staff = pd.read_excel(file, sheet_name=1)
            exclude = pd.read_excel(file, sheet_name=2)

            # Исключаем сотрудников
            fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])]

            # Группировка по подразделению
            result = fired_clean.groupby('подразделение').size().reset_index(name='Уволенные')
            result = result.merge(staff, on='подразделение', how='left')
            result['Текучесть %'] = (result['Уволенные'] / result['штат']) * 100
            result['Текучесть %'] = result['Текучесть %'].round(2)

            # Сохраняем в Excel в памяти
            output = BytesIO()
            result.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            return send_file(output,
                             download_name="turnover_result.xlsx",
                             as_attachment=True)
    return render_template('index.html')
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
