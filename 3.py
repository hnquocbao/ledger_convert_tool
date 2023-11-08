import PySimpleGUI as sg
import os
from app import converter

sg.theme("DarkTeal2")

layout = [
    [sg.T("")],
    [sg.Text("Choose a file: "),
     sg.Input(key="-IN-"),
     sg.FileBrowse(file_types=(("ALL Excel Files", "*.xlsx"), ("ALL Files", "*.*"), ))],
    [sg.Button("Submit")],
]

window = sg.Window('My File Browser', layout, size=(600,150))

filename = ""
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break
    elif event == "Submit":
        filename = values['-IN-']
        if filename:
            converter(filename)
            sg.Popup('Done.')
        break
window.close()
print(filename)