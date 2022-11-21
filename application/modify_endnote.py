import PySimpleGUI as sg
import os.path

import rispy
import pandas as pd


def load_file(filename, encoding='utf-8'):
    '''Get a file path name and encoding and returns pandas DataFrame'''
    with open(filename, encoding=encoding) as file:
        ris_file = rispy.load(file)
        df = pd.DataFrame(ris_file)
        return df


def get_rows(dataframe, included_word=[], excluded_word=[]):
    '''Get a dataframe and return a list of rows that include/exclude words in keywords '''
    row_list = []
    for index, row in dataframe.iterrows():
        keywords = row['keywords']
        try:
            if included_word != [] and excluded_word != []:
                included_check = all(
                    elem in keywords for elem in included_word)
                excluded_check = any(
                    elem in keywords for elem in excluded_word)
                if (included_check and (excluded_check == False)):
                    row_list.append(index)
            else:
                if included_word != []:
                    check = all(elem in keywords for elem in included_word)
                    if(check):
                        row_list.append(index)
                if excluded_word != []:
                    check = any(elem in keywords for elem in excluded_word)
                    if(check == False):
                        row_list.append(index)
        except:
            print(f'row with index of:{index} has no keywords')
    return row_list


def save_as_excel(dataframe, row_list, name):
    new_df = dataframe.iloc[row_list]
    new_df.to_excel(f'{name}.xlsx')

# First the window layout in 2 columns


file_list_column = [
    [
        sg.Text('فولدر فایل ها'),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse('انتخاب'),
    ],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
        )
    ],
]

# For now will only show the name of the file that was chosen
image_viewer_column = [
    [sg.Text('یک فایل از طرف چپ انتخاب کنید')],
    [sg.Text(size=(40, 1), key="-TOUT-")],
    [sg.In(size=(25, 1), enable_events=True, visible=False, key="-INCLUDED_WORD-"),
     sg.Checkbox('', key='-IN_CHECK-', default=False, enable_events=True),
     sg.Text('(کلمه یا کلمه هایی که میخوای داخل منبع باشن رو بنویس(چند کلمه رو با یه اسپیس فاصله بنویس')],
    [sg.In(size=(25, 1), enable_events=True, visible=False, key="-EXCLUDED_WORD-"), sg.Checkbox('', key='-EX_CHECK-', default=False,
                                                                                                enable_events=True), sg.Text('(کلمه یا کلمه هایی که میخوای داخل منبع نباشن رو بنویس(چند کلمه رو با یه اسپیس فاصله بنویس')],
    [sg.In(size=(25, 1), enable_events=True, key="-NAME-"),
        sg.Button('تمام', key="-OUT-"), sg.Text('اسم فایل خروجی رو اینجا بنویس')],
]

# ----------- FULL layout ----------
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeparator(),
        sg.Column(image_viewer_column)
    ]
]

window = sg.Window('برنامه محمدی', layout)

# Run the event Loop
while True:
    event, values = window.read()
    print(values)

    if event == "EXIT" or event == sg.WIN_CLOSED:
        break
    # Folder name was filled in, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values['-FOLDER-']
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []

        fname = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            and f.lower().endswith(('.ris'))
        ]
        window['-FILE LIST-'].update(fname)

    elif event == '-FILE LIST-':  # A file was chosen from the list below
        try:
            filename = os.path.join(
                values['-FOLDER-'], values['-FILE LIST-'][0]
            )
            window['-TOUT-'].update(filename.split('/')[-1])
        except:
            pass

    if values["-IN_CHECK-"]:
        window.Element('-INCLUDED_WORD-').Update(visible=True)
    else:
        window.Element('-INCLUDED_WORD-').Update(visible=False)
    if values["-EX_CHECK-"]:
        window.Element('-EXCLUDED_WORD-').Update(visible=True)
    else:
        window.Element('-EXCLUDED_WORD-').Update(visible=False)

    if event == '-OUT-':
        try:
            name = os.path.join(
                values['-FOLDER-'], values['-NAME-']
            )
            name = name.split('\\')
            name = '/'.join(name)
            if values["-IN_CHECK-"]:
                included_word = values['-INCLUDED_WORD-'].split(' ')
            else:
                included_word = []
            if values["-EX_CHECK-"]:
                excluded_word = values["-EXCLUDED_WORD-"].split(' ')
            else:
                excluded_word = []

            row_list = get_rows(load_file(filename),
                                included_word, excluded_word)
            save_as_excel(load_file(filename), row_list, name)
        except:
            pass
window.close()
