import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import math
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def calculate_capacity(lot_size, parts_per_pu, working_days_per_week, luf_days):
    pu_per_lot = math.ceil(lot_size / parts_per_pu)
    daily_production_rate = lot_size / working_days_per_week
    daily_pu_need_per_luf = math.ceil(daily_production_rate / parts_per_pu)
    luf_need_of_pu = daily_pu_need_per_luf * luf_days
    total_packaging_units = luf_need_of_pu + (pu_per_lot * 2)
    
    return {
        'Lot size': lot_size,
        'Parts per Packaging Unit': parts_per_pu,
        'Working days per week': working_days_per_week,
        'LUF days': luf_days,
        'PU per lot': pu_per_lot,
        'Daily Production Rate': daily_production_rate,
        'Daily PU need per LUF': daily_pu_need_per_luf,
        'LUF need of PU': luf_need_of_pu,
        'Safety Stock PU': pu_per_lot * 2,
        'Total Packaging Units': total_packaging_units
    }

def create_dashboard(results):
    df = pd.DataFrame(list(results.items()), columns=['Metric', 'Value'])
    df = df[df['Metric'].isin(['LUF need of PU', 'Safety Stock PU', 'Total Packaging Units'])]
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    sns.barplot(x='Metric', y='Value', data=df, ax=ax1)
    ax1.set_title('Verpackungseinheiten Übersicht')
    ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45, ha='right')
    
    pie_data = df[df['Metric'].isin(['LUF need of PU', 'Safety Stock PU'])]
    ax2.pie(pie_data['Value'], labels=pie_data['Metric'], autopct='%1.1f%%')
    ax2.set_title('Verteilung der Verpackungseinheiten')
    
    plt.tight_layout()
    return fig

def create_sensitivity_analysis(base_params, variable_param, range_values):
    results = []
    for value in range_values:
        params = base_params.copy()
        params[variable_param] = value
        capacity = calculate_capacity(**params)
        results.append(capacity['Total Packaging Units'])
    
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(range_values, results)
    ax.set_xlabel(variable_param)
    ax.set_ylabel('Total Packaging Units')
    ax.set_title(f'Sensitivitätsanalyse: {variable_param}')
    return fig

def create_excel_report(results):
    df = pd.DataFrame(list(results.items()), columns=['Metric', 'Value'])
    
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Capacity Calculation"
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    wb.save(output)
    return output.getvalue()

st.title('Dynamische Behälterkreislauf Simulation')

st.sidebar.header('Eingabeparameter')
lot_size = st.sidebar.slider('Losgröße', min_value=100, max_value=10000, value=1000, step=100)
parts_per_pu = st.sidebar.slider('Teile pro Verpackungseinheit', min_value=1, max_value=100, value=10)
working_days_per_week = st.sidebar.slider('Arbeitstage pro Woche', min_value=1, max_value=7, value=5)
luf_days = st.sidebar.slider('LUF-Tage', min_value=1, max_value=30, value=5)

base_params = {
    'lot_size': lot_size,
    'parts_per_pu': parts_per_pu,
    'working_days_per_week': working_days_per_week,
    'luf_days': luf_days
}

results = calculate_capacity(**base_params)

st.header('Berechnete Ergebnisse')
for key, value in results.items():
    st.write(f"{key}: {value}")

st.header('Dashboard')
fig = create_dashboard(results)
st.pyplot(fig)

st.header('Sensitivitätsanalyse')
param_to_analyze = st.selectbox('Parameter für Sensitivitätsanalyse', 
                                ['lot_size', 'parts_per_pu', 'working_days_per_week', 'luf_days'])
range_min = st.number_input(f'Minimaler Wert für {param_to_analyze}', value=1)
range_max = st.number_input(f'Maximaler Wert für {param_to_analyze}', value=100)
range_step = st.number_input(f'Schrittweite für {param_to_analyze}', value=1)

range_values = list(range(range_min, range_max + 1, range_step))
sensitivity_fig = create_sensitivity_analysis(base_params, param_to_analyze, range_values)
st.pyplot(sensitivity_fig)

st.header('Szenarienvergleich')
st.write('Passen Sie die Parameter an und klicken Sie auf "Szenario hinzufügen", um verschiedene Szenarien zu vergleichen.')

if 'scenarios' not in st.session_state:
    st.session_state.scenarios = []

if st.button('Szenario hinzufügen'):
    st.session_state.scenarios.append(calculate_capacity(**base_params))

if st.session_state.scenarios:
    df_scenarios = pd.DataFrame(st.session_state.scenarios)
    st.write(df_scenarios)

    if len(st.session_state.scenarios) > 1:
        fig, ax = plt.subplots(figsize=(10, 6))
        df_scenarios['Total Packaging Units'].plot(kind='bar', ax=ax)
        ax.set_xlabel('Szenario')
        ax.set_ylabel('Total Packaging Units')
        ax.set_title('Vergleich der Szenarien')
        st.pyplot(fig)

excel_report = create_excel_report(results)
st.download_button(
    label="Excel-Bericht herunterladen",
    data=excel_report,
    file_name="capacity_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)