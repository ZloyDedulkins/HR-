from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    if request.method == 'POST':
        file = request.files['file']
        if file:
            fired = pd.read_excel(file, sheet_name=0)
            staff = pd.read_excel(file, sheet_name=1)
            exclude = pd.read_excel(file, sheet_name=2)

            fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])]
            result = fired_clean.groupby('подразделение').size().reset_index(name='Уволенные')
            result = result.merge(staff, on='подразделение', how='left')
            result['Текучесть %'] = (result['Уволенные'] / result['штат']) * 100
            result['Текучесть %'] = result['Текучесть %'].round(2)

    return render_template('index.html', result=result)
