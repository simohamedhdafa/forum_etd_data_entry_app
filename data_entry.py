import PySimpleGUI as sg
import pandas as pd
from os import getcwd, path

# Add some color to the window
sg.theme('DarkTeal9')
# handle exception FileNoteFoundException !
FILE_NAME = 'Data_Entry.xlsx'
EXCEL_FILE = FILE_NAME

# if fileNoteFoundError
if not path.exists(getcwd()+'\\'+FILE_NAME):
    writer = pd.ExcelWriter(FILE_NAME, engine='xlsxwriter')
    writer.save()

df = pd.read_excel(EXCEL_FILE)

layout = [
    #[sg.Listbox(values = sg.theme_list(), size =(20, 12), key ='-LIST-', enable_events = True)],
    [sg.Text('Veuillez renseigner le formulaire suivant:')],
    [sg.Text('Nom de famille', size=(15,1)), sg.InputText(key='Nom')],
    [sg.Text('Prénom', size=(15,1)), sg.InputText(key='Prenom')],
    # calendar
    [sg.Text('Date de naissance', key='datenaissance')],
    [sg.Input(key='birthdate', size=(20,1)), sg.CalendarButton('choisir une date', close_when_date_chosen=True,  target='birthdate', location=(0,0), no_titlebar=False, )],
    # /calendar
    [sg.Text('Ville de résidence', size=(15,1)), sg.InputText(key='Ville')],
    [sg.Radio('Sciences mathématiques', "RADIO1", default=True), 
        sg.Radio('Sciences physiques', "RADIO1"), 
        sg.Radio('Sciences de la vie et de la terre', "RADIO1")],
    [sg.Text('Lycée', size=(15,1)), sg.InputText(key='lycee')],
    [sg.Text('Choisir une date de concours', size=(15,1)), sg.Combo(['14-05-2022', 
                                                    '28-05-2022', 
                                                    '25-06-2022',
                                                    '02-07-2022',
                                                    '09-07-2022',
                                                    '16-07-2022',
                                                    '23-07-2022'], key='Dates concours')],
    [sg.Text('Tel', size=(15,1)), sg.InputText(key='tel')],
    [sg.Text('email', size=(15,1)), sg.InputText(key='email')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Forum de l\'étudiant à Casablanca.', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    #sg.theme(values['-LIST-'][0])
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    elif event == 'Clear':
        clear_input()
    elif event == 'Submit':
        #print(values)
        #break
        values['tel'] = '\''+values['tel']+'\''
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('ajouté avec succès!')
        clear_input()
window.close()
