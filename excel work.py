import pandas as pd
from DfOperate import DfOperate

df = pd.read_excel('medical_diagnostic_devices_10000.xlsx')

status_mapping = {
    'OK': 'operational',
    'op': 'operational',
    'broken': 'faulty',
    'planned_installation': 'planned_installation',
    'maintenance_scheduled': 'maintenance_scheduled',
    'operational': 'operational',
    'faulty': 'faulty'
}

df['status'] = df['status'].astype(str).str.lower().str.strip()
df['status'] = df['status'].map(status_mapping).fillna(df['status'])

data = DfOperate(df)

table_filter_warranty = data.get_tables_by_warranty_status()
problem_clinics = data.find_problem_clinics()
calibration_statuses = data.get_calibration_statuses()
pivot_table_clinics = data.create_pivot_table()

with pd.ExcelWriter('tasks.xlsx', engine='openpyxl') as writer:
    table_filter_warranty['Гарантия истекла'].to_excel(writer, sheet_name='Гарантия истекла', index=False)
    table_filter_warranty['Истечет менее чем через месяц'].to_excel(writer,
                                                                    sheet_name='Гарантия менее месяца',
                                                                    index=False)
    table_filter_warranty['Истечет менее чем через полгода'].to_excel(writer,
                                                                      sheet_name='Гарантия менее полугода',
                                                                      index=False)
    table_filter_warranty['Истечет более чем через полгода'].to_excel(writer,
                                                                      sheet_name='Гарантия более полугода',
                                                                      index=False)
    problem_clinics.to_excel(writer, sheet_name='Наиболее проблемные клиники', index=False)
    calibration_statuses.to_excel(writer, sheet_name='Статусы калибровки', index=False)
    pivot_table_clinics.to_excel(writer, sheet_name='Сводная таблица по клиникам', index=False)
