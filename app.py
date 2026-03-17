from io import BytesIO
import re

from flask import Flask, redirect, render_template, request, send_file, session, url_for
import pandas as pd

app = Flask(__name__)
app.secret_key = 'hr-dashboard-secret-key'

AZS_POSITIONS = {
    'администратор',
    'бригадир-заправщик',
    'заместитель директора азс по направлению розничной торговли',
    'заправщик',
    'заправщик агнкс',
    'оператор-кассир',
    'старший оператор-кассир',
    'уборщик',
    'фармацевт',
}

MB_POSITIONS = {
    'директор кафе',
    'заместитель директора азс по направлению кафе',
    'кассир',
    'менеджер кафе',
    'повар',
    'работник торгового зала',
}


def normalize_text(value) -> str:
    if pd.isna(value):
        return ''
    return str(value).strip()


def normalize_for_compare(value) -> str:
    text = normalize_text(value).lower().replace('ё', 'е')
    text = re.sub(r'\s+', ' ', text)
    return text.strip(' ,.;:')


def normalize_column_name(value) -> str:
    return re.sub(r'[^a-zа-я0-9]+', '', normalize_for_compare(value))


def find_column(df: pd.DataFrame, variants: list[str]) -> str | None:
    normalized_columns = {normalize_column_name(col): col for col in df.columns}
    for variant in variants:
        found = normalized_columns.get(normalize_column_name(variant))
        if found:
            return found
    return None


def normalize_business_unit(value) -> str:
    normalized = normalize_for_compare(value)
    aliases = {
        'бк': 'БК',
        'азс': 'АЗС',
        'мб': 'МБ',
    }
    return aliases.get(normalized, normalize_text(value) or 'Не определен')


def determine_business_unit(department, position) -> str:
    department_str = normalize_text(department)
    position_str = normalize_for_compare(position)
    first_symbol = department_str[:1].upper()

    if first_symbol == 'М':
        return 'БК'

    if first_symbol.isdigit() or first_symbol == 'О':
        if position_str in AZS_POSITIONS:
            return 'АЗС'
        if position_str in MB_POSITIONS:
            return 'МБ'

    return 'Не определен'


def is_adult_at_dismissal(row: pd.Series) -> bool:
    dismissal_date = pd.to_datetime(row.get('Дата увольнения'), errors='coerce', dayfirst=True)
    birth_date = pd.to_datetime(row.get('Дата рождения'), errors='coerce', dayfirst=True)

    if pd.isna(dismissal_date) or pd.isna(birth_date):
        return True

    age_years = dismissal_date.year - birth_date.year - (
        (dismissal_date.month, dismissal_date.day) < (birth_date.month, birth_date.day)
    )
    return age_years >= 18


def build_staff_business_unit_maps(
    staff: pd.DataFrame,
    staff_business_unit_col: str | None,
    staff_position_col: str | None,
) -> tuple[dict[tuple[str, str], str], dict[str, str]]:
    if not staff_business_unit_col:
        return {}, {}

    work = staff.copy()
    work['подразделение'] = work['подразделение'].fillna('Без подразделения').astype(str)
    work['Бизнес-юнит'] = work[staff_business_unit_col].apply(normalize_business_unit)
    work['__position_norm'] = work[staff_position_col].apply(normalize_for_compare) if staff_position_col else ''

    dept_position_map: dict[tuple[str, str], str] = {}
    for _, row in work.iterrows():
        key = (row['подразделение'], row['__position_norm'])
        if key[1]:
            dept_position_map[key] = row['Бизнес-юнит']

    dept_map: dict[str, str] = {}
    for dep, dep_rows in work.groupby('подразделение'):
        unique_units = dep_rows['Бизнес-юнит'].dropna().astype(str).str.strip().unique().tolist()
        if len(unique_units) == 1:
            dept_map[dep] = unique_units[0]

    return dept_position_map, dept_map


def build_full_result(file_obj) -> pd.DataFrame:
    workbook = pd.ExcelFile(file_obj)
    fired = workbook.parse('Уволенные')
    staff = workbook.parse('Штатка')
    exclude = workbook.parse('Исключения')

    fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])].copy()
    fired_clean = fired_clean[fired_clean.apply(is_adult_at_dismissal, axis=1)]

    fired_position_col = find_column(fired_clean, ['должность'])

    staff_business_unit_col = find_column(staff, ['бизнес-юнит', 'бизнес юнит', 'бизнес-юнита', 'бизнес юнита'])
    staff_position_col = find_column(staff, ['должность'])
    staff_dept_position_map, staff_dept_map = build_staff_business_unit_maps(
        staff,
        staff_business_unit_col,
        staff_position_col,
    )

    fired_clean['подразделение'] = fired_clean['подразделение'].fillna('Без подразделения').astype(str)
    fired_clean['__position_norm'] = (
        fired_clean[fired_position_col].apply(normalize_for_compare) if fired_position_col else ''
    )
    fired_clean['Бизнес-юнит'] = fired_clean.apply(
        lambda row: determine_business_unit(row['подразделение'], row[fired_position_col] if fired_position_col else ''),
        axis=1,
    )
    fired_clean['Бизнес-юнит'] = fired_clean.apply(
        lambda row: (
            row['Бизнес-юнит']
            if row['Бизнес-юнит'] != 'Не определен'
            else staff_dept_position_map.get((row['подразделение'], row['__position_norm']))
            or staff_dept_map.get(row['подразделение'])
            or 'Не определен'
        ),
        axis=1,
    )
    fired_counts = fired_clean.groupby(['Бизнес-юнит', 'подразделение']).size().reset_index(name='Уволенные')

    staff_base = staff[['подразделение', 'штат']].copy()
    staff_base['подразделение'] = staff_base['подразделение'].fillna('Без подразделения').astype(str)
    if staff_business_unit_col:
        staff_base['Бизнес-юнит'] = staff[staff_business_unit_col].apply(normalize_business_unit)
    else:
        staff_base['Бизнес-юнит'] = staff.apply(
            lambda row: determine_business_unit(row['подразделение'], row[staff_position_col] if staff_position_col else ''),
            axis=1,
        )

    result_df = staff_base.merge(fired_counts, on=['Бизнес-юнит', 'подразделение'], how='left')
    result_df['Уволенные'] = result_df['Уволенные'].fillna(0).astype(int)
    result_df['штат'] = pd.to_numeric(result_df['штат'], errors='coerce').fillna(0)

    result_df['Текучесть %'] = 0.0
    non_zero_staff = result_df['штат'] != 0
    result_df.loc[non_zero_staff, 'Текучесть %'] = (
        (result_df.loc[non_zero_staff, 'Уволенные'] / result_df.loc[non_zero_staff, 'штат']) * 100
    ).round(2)

    result_df = result_df.sort_values(['Бизнес-юнит', 'подразделение'], ascending=True)
    return result_df


def apply_filters(
    result_df: pd.DataFrame,
    selected_departments: list[str],
    selected_business_units: list[str],
) -> pd.DataFrame:
    filtered = result_df.copy()

    if selected_departments:
        filtered = filtered[filtered['подразделение'].isin(selected_departments)]

    if selected_business_units:
        filtered = filtered[filtered['Бизнес-юнит'].isin(selected_business_units)]

    return filtered.copy()


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
    business_unit_options = []
    selected_departments = request.args.getlist('departments')
    selected_business_units = request.args.getlist('business_units')

    if full_records:
        full_result_df = pd.DataFrame(full_records)
        department_options = full_result_df['подразделение'].dropna().astype(str).sort_values().unique().tolist()
        business_unit_options = full_result_df['Бизнес-юнит'].dropna().astype(str).sort_values().unique().tolist()
        filtered_df = apply_filters(full_result_df, selected_departments, selected_business_units)

        summary = {
            'total_fired': int(filtered_df['Уволенные'].sum()),
            'avg_turnover': round(filtered_df['Текучесть %'].mean(), 2) if not filtered_df.empty else 0,
            'total_staff': round(filtered_df['штат'].sum(), 2) if not filtered_df.empty else 0,
        }

        chart_labels = [f"{bu} / {dep}" for bu, dep in zip(filtered_df['Бизнес-юнит'], filtered_df['подразделение'])]
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
        business_unit_options=business_unit_options,
        selected_departments=selected_departments,
        selected_business_units=selected_business_units,
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
