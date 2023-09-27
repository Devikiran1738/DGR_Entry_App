from pathlib import Path
import PySimpleGUI as sg
import pandas as pd
# Add some color to the window
sg.theme('DarkTeal9')
current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data_Entry.xlsx'
# Load the data if the file exists, if not, create a new DataFrame
if EXCEL_FILE.exists():
    df = pd.read_excel(EXCEL_FILE)
else:
    df = pd.DataFrame()
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Site_Name', size=(15,1)), sg.InputText(key='Site_Name')],
    [sg.Text('Capacity', size=(15,1)), sg.InputText(key='Capacity')],
    [sg.Text('Day_Generation', size=(15,1)), sg.InputText(key='Day_Generation')],
    [sg.Text('Specific_Yield', size=(15,1)), sg.InputText(key='Specific_Yield')],
    [sg.Text('Irradiance', size=(15,1)), sg.InputText(key='Irradiance')],
    [sg.Text('Plant_Start_Time', size=(15,1)), sg.InputText(key='Plant_Start_Time')],
    [sg.Text('Plant_Stop_Time', size=(15,1)), sg.InputText(key='Plant_Stop_Time')],
    [sg.Text('Grid_Failure',size=(15,1)),sg.InputText(key='Grid_Failure')],
    [sg.Text('Downtime',size=(15,1)),sg.InputText(key='Downtime')],
    [sg.Text('Remarks',size=(15,1)),sg.InputText(key='Remarks')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]
window = sg.Window('Simple data entry form', layout)
def clear_input():
    for key in values:
        window[key]('')
    return None
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)  # This will create the file if it doesn't exist
        sg.popup('Data saved!')
        clear_input()
window.close()
