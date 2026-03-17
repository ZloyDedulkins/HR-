from io import BytesIO

from flask import Flask, render_template, request, send_file, session
import pandas as pd

app = Flask(__name__)
app.secret_key = 'hr-dashboard-secret-key'

def is_adult_at_dismissal(row: pd.Series) -> bool:
    dismissal_date = pd.to_datetime(row.get('Дата увольнения'), errors='coerce', dayfirst=True)
    birth_date = pd.to_datetime(row.get('Дата рождения'), errors='coerce', dayfirst=True)

    if pd.isna(dismissal_date) or pd.isna(birth_date):
        return True

    age_years = dismissal_date.year - birth_date.year - (
        (dismissal_date.month, dismissal_date.day) < (birth_date.month, birth_date.day)
    )
    return age_years >= 18


@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    chart_labels = []
    chart_values = []
    summary = None
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

                fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])].copy()
                fired_clean = fired_clean[fired_clean.apply(is_adult_at_dismissal, axis=1)]
                result_df = fired_clean.groupby('подразделение').size().reset_index(name='Уволенные')
                result_df = result_df.merge(staff, on='подразделение', how='inner')
                result_df['Текучесть %'] = ((result_df['Уволенные'] / result_df['штат']) * 100).round(2)
                result_df['подразделение'] = result_df['подразделение'].fillna('Без подразделения').astype(str)
                result_df = result_df.sort_values('подразделение', ascending=True)

                summary = {
                    'total_fired': int(result_df['Уволенные'].sum()),
                    'avg_turnover': round(result_df['Текучесть %'].mean(), 2) if not result_df.empty else 0,
                    'total_staff': round(result_df['штат'].sum(), 2) if not result_df.empty else 0,
                }

                chart_labels = result_df['подразделение'].tolist()
                chart_values = result_df['Текучесть %'].fillna(0).tolist()
                result = result_df.to_dict(orient='records')
                session['result_records'] = result
            except Exception as exc:
                error = f'Не удалось обработать файл: {exc}'

    return render_template(
        'index.html',
        result=result,
        summary=summary,
        chart_labels=chart_labels,
        chart_values=chart_values,
        error=error,
    )


@app.route('/download-result')
def download_result():
    result_records = session.get('result_records')
    if not result_records:
        return 'Сначала рассчитайте текучесть, чтобы скачать файл.', 400

    result_df = pd.DataFrame(result_records)
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Итог')

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name='itog_tekuchesti.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    app.run(debug=True)
