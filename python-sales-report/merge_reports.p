import pandas as pd
import re
from pathlib import Path

# ---------- 1. Настройки ----------
INPUT_DIR = Path('.')  # текущая папка
OUTPUT_FILE = 'сводная_итог.xlsx'

# ---------- 2. Исправление опечаток в товарах ----------
fix_names = {
    'акумуляторная отвертка': 'Аккумуляторная отвертка',
    'аккумуляторная отвертка': 'Аккумуляторная отвертка',
    'аккумуляторная овертка': 'Аккумуляторная отвертка',
    'Электроовертка': 'Электроотвертка',
    'Электроотвертка': 'Электроотвертка',
    'молоток': 'Молоток',
    'пила': 'Пила',
    'отвертка': 'Отвертка',
    'отвертка крестообразная': 'Отвертка крестообразная',
    'отвертка крестобразная': 'Отвертка крестообразная',
    'индикатор анпряжения': 'Индикатор напряжения',
    'индикатор напряжения': 'Индикатор напряжения',
    'кернер': 'Кернер',
    'фонарик на лобный': 'Фонарик налобный',
    'фонарик налобный': 'Фонарик налобный',
    'Чехол сумка для инструментов': 'Чехол-сумка для инструментов',
    'сверло по керамограниту': 'Сверло По Керамограниту',
    'Лобзик': 'Лобзик',
}

# ---------- 3. Очистка чисел ----------
def clean_number(value):
    if pd.isna(value):
        return None
    s = str(value)
    s = re.sub(r'[^\d\-.,]', '', s)
    s = s.replace(',', '.')
    if s.count('.') > 1:
        parts = s.split('.')
        s = parts[0] + '.' + ''.join(parts[1:])
    try:
        return float(s)
    except ValueError:
        return None

# ---------- 4. Определение квартала по дате ----------
def get_quarter_from_date(date_value):
    if pd.isna(date_value):
        return None
    try:
        s = str(date_value).strip()
        dt = pd.to_datetime(s, dayfirst=True, errors='raise')
        month = dt.month
        if 1 <= month <= 3:
            return '2026 Квартал 1'
        elif 4 <= month <= 6:
            return '2026 Квартал 2'
        elif 7 <= month <= 9:
            return '2026 Квартал 3'
        elif 10 <= month <= 12:
            return '2026 Квартал 4'
        else:
            return None
    except Exception:
        return None

# ---------- 5. Город и квартал из имени файла ----------
def parse_filename(filename):
    name = filename.lower()
    # город
    if 'мск' in name or 'москва' in name:
        city = 'Москва'
    elif 'спб' in name:
        city = 'СПБ'
    else:
        return None, None
    # квартал
    quarter = None
    for q in [1, 2, 3, 4]:
        if f'{q} квартал' in name:
            quarter = f'2026 Квартал {q}'
            break
    if quarter is None:
        return None, None
    return city, quarter

# ---------- 6. Сканируем папку ----------
all_rows = []
warnings = []

xlsx_files = [f for f in INPUT_DIR.glob('*.xlsx') if f.name != OUTPUT_FILE]
print(f'Найдено файлов: {len(xlsx_files)}')

for path in sorted(xlsx_files):
    city_default, quarter_default = parse_filename(path.name)
    if city_default is None:
        print(f'⚠️ Пропущен (не понял город/квартал): {path.name}')
        continue

    print(f'Обрабатываю: {path.name} → {city_default}, {quarter_default}')

    df = pd.read_excel(path, sheet_name='Лист1')
    df.columns = df.columns.str.strip()

    for idx, row in df.iterrows():
        if pd.isna(row.get('Товар')) and pd.isna(row.get('Цена')):
            continue
        if pd.isna(row.get('Товар')):
            continue

        city = row.get('Город')
        if pd.isna(city) or str(city).strip() == '':
            city = city_default
        else:
            city = str(city).strip()

        item = str(row.get('Товар')).strip()
        item = fix_names.get(item, item)

        qty = clean_number(row.get('Количество'))
        price = clean_number(row.get('Цена'))
        if qty is None or price is None:
            continue

        # дата: сохраняем как есть для проверки дублей
        date_raw = row.get('Дата')
        date_str = str(date_raw).strip() if not pd.isna(date_raw) else ''

        # проверка даты
        quarter_from_date = get_quarter_from_date(date_raw)
        if quarter_from_date is not None and quarter_from_date != quarter_default:
            warnings.append(
                f'{path.name}, строка {idx + 2}: дата "{date_str}" → {quarter_from_date}, '
                f'но файл помечен как {quarter_default}. Оставлен квартал из имени файла.'
            )

        total = qty * price

        all_rows.append({
            'Город': city,
            'товар': item,
            'Колличество': qty,
            'Цена': price,
            'сумма': total,
            'Квартал': quarter_default,
            'Дата_str': date_str,
        })

df_all = pd.DataFrame(all_rows)
print(f'\nВсего строк до удаления дублей: {len(df_all)}')

# ---------- 7. Удаление дублей ----------
df_clean = df_all.drop_duplicates(
    subset=['Дата_str', 'товар', 'Колличество', 'Цена', 'Город']
).copy()

removed = len(df_all) - len(df_clean)
print(f'Удалено дублей: {removed}')
print(f'Строк после удаления дублей: {len(df_clean)}')

# ---------- 8. Сводные ----------
pivot_sum = pd.pivot_table(
    df_clean,
    index=['Город', 'товар'],
    columns='Квартал',
    values='сумма',
    aggfunc='sum',
    fill_value=0
)

pivot_qty = pd.pivot_table(
    df_clean,
    index=['Город', 'товар'],
    columns='Квартал',
    values='Колличество',
    aggfunc='sum',
    fill_value=0
)

pivot_sum['Итог Сумма (рублей)'] = pivot_sum.sum(axis=1)
pivot_qty['Итог Сумма (количество)'] = pivot_qty.sum(axis=1)

# ---------- 9. Сборка итоговой таблицы ----------
result_rows = []
for city in pivot_sum.index.get_level_values('Город').unique():
    city_sum = pivot_sum.loc[city].sum()
    city_qty = pivot_qty.loc[city].sum()
    result_rows.append({
        'Город': city,
        'товар': '',
        **{f'{col} (руб)': city_sum.get(col, 0) for col in pivot_sum.columns},
        **{f'{col} (шт)': city_qty.get(col, 0) for col in pivot_qty.columns},
    })
    for (c, t) in pivot_sum.index:
        if c == city:
            result_rows.append({
                'Город': '',
                'товар': t,
                **{f'{col} (руб)': pivot_sum.loc[(c, t), col] for col in pivot_sum.columns},
                **{f'{col} (шт)': pivot_qty.loc[(c, t), col] for col in pivot_qty.columns},
            })

total_sum = pivot_sum.sum()
total_qty = pivot_qty.sum()
result_rows.append({
    'Город': 'Общий итог',
    'товар': '',
    **{f'{col} (руб)': total_sum.get(col, 0) for col in pivot_sum.columns},
    **{f'{col} (шт)': total_qty.get(col, 0) for col in pivot_qty.columns},
})

result = pd.DataFrame(result_rows)
# ---------- Форматирование Excel ----------
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Сначала сохраняем как обычно
result.to_excel(OUTPUT_FILE, index=False)

# Потом открываем и форматируем
wb = load_workbook(OUTPUT_FILE)
ws = wb.active

# Стили
thin = Side(border_style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
bold = Font(bold=True)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center')

# Заливка для итогов
fill_city = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')   # голубой
fill_total = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')  # жёлтый

# Ширина колонок
ws.column_dimensions['A'].width = 14  # Город
ws.column_dimensions['B'].width = 32  # товар
for col in range(3, ws.max_column + 1):
    ws.column_dimensions[get_column_letter(col)].width = 18

# Формат чисел для колонок с рублями и штуками
for row in range(2, ws.max_row + 1):
    for col in range(3, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        # числа с разделителем тысяч
        if isinstance(cell.value, (int, float)):
            cell.number_format = '#,##0.00'

# Обработка строк
for row in range(2, ws.max_row + 1):
    city_val = ws.cell(row=row, column=1).value
    item_val = ws.cell(row=row, column=2).value

    is_total_row = (city_val == 'Общий итог')
    is_city_row = (city_val and city_val != 'Общий итог')

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        cell.border = border
        cell.alignment = center if col >= 3 else left

        if is_total_row:
            cell.font = bold
            cell.fill = fill_total
        elif is_city_row:
            cell.font = bold
            cell.fill = fill_city

# Заголовки
for col in range(1, ws.max_column + 1):
    cell = ws.cell(row=1, column=col)
    cell.font = bold
    cell.alignment = center
    cell.border = border

# Закрепить первую строку
ws.freeze_panes = 'A2'

wb.save(OUTPUT_FILE)
print('✅ Форматирование применено.')
print(f'\n✅ Готово! Файл: {OUTPUT_FILE}')

# ---------- 10. Предупреждения по датам ----------
if warnings:
    print(f'\n⚠️ Найдено расхождений дат: {len(warnings)}')
    for w in warnings[:30]:
        print('  ' + w)
    if len(warnings) > 30:
        print(f'  ... и ещё {len(warnings) - 30}')
else:
    print('\n✅ Расхождений дат не найдено.')
