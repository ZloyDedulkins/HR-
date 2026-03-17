from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    chart_labels = []
    chart_values = []
    error = None

    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            error = 'Файл не выбран. Попробуйте ещё раз.'
        else:
            try:
                workbook = pd.ExcelFile(file)
                fired = workbook.parse('Уволенные')
                staff = workbook.parse('Штатка')
                exclude = workbook.parse('Исключения')

                fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])]
                result_df = fired_clean.groupby('подразделение').size().reset_index(name='Уволенные')
                result_df = result_df.merge(staff, on='подразделение', how='left')
                result_df['Текучесть %'] = ((result_df['Уволенные'] / result_df['штат']) * 100).round(2)
                result_df = result_df.sort_values('Текучесть %', ascending=False)

                chart_labels = result_df['подразделение'].fillna('Без подразделения').astype(str).tolist()
                chart_values = result_df['Текучесть %'].fillna(0).tolist()
                result = result_df.to_dict(orient='records')
            except Exception as exc:
                error = f'Не удалось обработать файл: {exc}'

    return render_template(
        'index.html',
        result=result,
        chart_labels=chart_labels,
        chart_values=chart_values,
        error=error,
    )


if __name__ == '__main__':
    app.run(debug=True)
