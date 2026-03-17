from io import BytesIO

from flask import Flask, redirect, render_template, request, send_file, session, url_for
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


def build_full_result(file_obj) -> pd.DataFrame:
    workbook = pd.ExcelFile(file_obj)
    fired = workbook.parse('Уволенные')
    staff = workbook.parse('Штатка')
    exclude = workbook.parse('Исключения')

    fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])].copy()
    fired_clean = fired_clean[fired_clean.apply(is_adult_at_dismissal, axis=1)]

    fired_counts = fired_clean.groupby('подразделение').size().reset_index(name='Уволенные')

    staff_base = staff[['подразделение', 'штат']].copy()
    staff_base['подразделение'] = staff_base['подразделение'].fillna('Без подразделения').astype(str)

    fired_counts['подразделение'] = fired_counts['подразделение'].fillna('Без подразделения').astype(str)

    result_df = staff_base.merge(fired_counts, on='подразделение', how='left')
    result_df['Уволенные'] = result_df['Уволенные'].fillna(0).astype(int)
    result_df['штат'] = pd.to_numeric(result_df['штат'], errors='coerce').fillna(0)

    result_df['Текучесть %'] = 0.0
    non_zero_staff = result_df['штат'] != 0
    result_df.loc[non_zero_staff, 'Текучесть %'] = (
        (result_df.loc[non_zero_staff, 'Уволенные'] / result_df.loc[non_zero_staff, 'штат']) * 100
    ).round(2)

    result_df = result_df.sort_values('подразделение', ascending=True)
    return result_df


def apply_department_filter(result_df: pd.DataFrame, selected_departments: list[str]) -> pd.DataFrame:
    if not selected_departments:
        return result_df
    return result_df[result_df['подразделение'].isin(selected_departments)].copy()


@app.route('/', methods=['GET', 'POST'])
def index():
    error = None

    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            error = 'Файл не выбран. Попробуйте ещё раз.'
        else:
            try:
                full_result_df = build_full_result(file)
                full_records = full_result_df.to_dict(orient='records')
                session['full_result_records'] = full_records
                return redirect(url_for('index'))
            except Exception as exc:
                error = f'Не удалось обработать файл: {exc}'

    full_records = session.get('full_result_records')
    result = None
    summary = None
    chart_labels = []
    chart_values = []
    department_options = []
    selected_departments = request.args.getlist('departments')

    if full_records:
        full_result_df = pd.DataFrame(full_records)
        department_options = full_result_df['подразделение'].dropna().astype(str).sort_values().unique().tolist()
        filtered_df = apply_department_filter(full_result_df, selected_departments)

        summary = {
            'total_fired': int(filtered_df['Уволенные'].sum()),
            'avg_turnover': round(filtered_df['Текучесть %'].mean(), 2) if not filtered_df.empty else 0,
            'total_staff': round(filtered_df['штат'].sum(), 2) if not filtered_df.empty else 0,
        }

        chart_labels = filtered_df['подразделение'].tolist()
        chart_values = filtered_df['Текучесть %'].fillna(0).tolist()
        result = filtered_df.to_dict(orient='records')
        session['result_records'] = result

    return render_template(
        'index.html',
        result=result,
        summary=summary,
        chart_labels=chart_labels,
        chart_values=chart_values,
        error=error,
        department_options=department_options,
        selected_departments=selected_departments,
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
