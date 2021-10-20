#!/usr/bin/env python
# import pandas as pd
import os
import PySimpleGUI as sg
# import sqlite3
from classes import *
from functions import *

"""
    SKUs Quality Control Application
"""


sg.theme('Python')
sg.set_options(element_padding=(0, 0))

# ------ Menu Definition ------ #
menu_def = [
    ['Import', ['Clusters::IC', 'SKUs', ['Monthly::ISM', 'Bimonthly - Food::ISF', 'Bimonthly - Non Food::ISN'], 'Outlets::IO']],
    ['Export', ['Clusters::EC', 'SKUs', ['SKUs::ES', 'SKUs - Analysis::ESA'], 'Outlets', ['Missing::EOM','Atypicals:EOA']]],
    ['Delete', ['Clusters::DC', 'SKUs::DS']],
    ['Console', ['Clear::CC']],
    ['About', ['Importing', ['Clusters::AC', 'SKUs::AS', 'Outlets::AO'], '---', 'SKU Quality Control Application::AA', ]],
    ['Exit', ['Exit Application::EA']],
]

# ------ GUI Defintion ------ #
layout = [
    [sg.MenubarCustom(
        menu_def,
        tearoff=False,
        bar_background_color='#f0f0f0',
        background_color='#f0f0f0',
        text_color='black',)],
    [sg.Multiline(size=(100,20),
        pad=0,
        background_color='black', 
        text_color='white', 
        write_only=True, 
        # reroute_stdout=True, 
        reroute_cprint=True,
        auto_refresh=True,
        autoscroll = True,
        default_text=f'''Click the options under the "About" section of the menu to read details on how to use the application.\n\n\n''',
        font=('Courier', 9),
        key='-ML-',)],
]

window = sg.Window("SKU Quality Control Application",
                    layout,
                    default_element_size=(12, 1),
                    grab_anywhere=True,
                    default_button_element_size=(12, 1),
                    finalize=True,
                    enable_close_attempted_event=True)

cp = mycprint(window)
iom = IOManager(window)
db = DbManager(window)

test_files_dir = os.path.join(os.getcwd(), 'test_files')
test_files_dir = r'C:\ANASTASIS\Python\My_Projects\QualityControl\test_files'
clusters_file = os.path.join(test_files_dir, 'clusters', 'cluster file.xlsx')
skus_file = os.path.join(test_files_dir, 'skus', 'Monthly-Juice.xlsx')
print(os.getcwd())
while 1:
    event, values = window.read()
    # print(f'Event={event}')
    # MENU "Exit" or Click X -> Close App
    title = 'WARNING!'
    message = 'Clicking "Yes" WILL TERMINATE the application.'
    if (event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event.endswith('::EA')) and popup_yes_no(title, message):
        break
    # MENU "About" -> Print info messages
    if '::A' in event:
        if event.endswith('::AA'):
            message = APP_INFO_MESSAGE
        elif event.endswith('::AC'):
            message = CLUSTERS_IMPORT_MESSAGE
        elif event.endswith('::AS'):
            message = SKUS_IMPORT_MESSAGE
        elif event.endswith('::AO'):
            message = OUTLETS_IMPORT_MESSAGE
        cp(message, c=INFO_OUTPUT_FORMAT)
    # MENU "Delete" -> Delete table rows
    elif '::D' in event:
        if event.endswith('::DC'):
            message = f'Clicking "OK" WILL DELETE all rows from "clusters" database table.'
            if popup_yes_no(title, message):
                db.delete_table_rows('clusters')
        if event.endswith('::DS'):
            message = f'Clicking "OK" WILL DELETE all rows from "skus" database table.'
            if popup_yes_no(title, message):
                db.delete_table_rows('skus')
    # MENU "Console" -> Clear Console
    elif event.endswith('::CC'):
        cp('', u=True)
    # MENU "Import" -> Clusters
    elif event.endswith('::IC'):
        # clusters_filename = sg.popup_get_file('Select Clusters file', no_window=True, file_types=(('Excel files',"*.xlsx"), ('Excel files',"*.xls")))
        if 1 > 0: # clusters_filename:
            # db.update_clusters()
            iom.parse_clusters(clusters_file)
    # MENU "Import" -> SKUs
    elif '::IS' in event:
        # selected_sku_type = event[-1]
        # sku_file_type = {'M':'(MONTHLY)', 'F':'(BIMONTHLY - FOOD)', '':'(BIMONTHLY - NON FOOD)'}
        # skus_filename = sg.popup_get_file(f'Select SKUs file {sku_file_type[selected_sku_type]}', file_types=(('Excel files',"*.xlsx"), ('Excel files',"*.xls")))
        if 1 > 0: # skus_filename:
            iom.parse_skus(skus_file)
    # MENU "Import" -> Outlets
    elif event.endswith('::IO'):
        outlets_filename = sg.popup_get_file('Select Outlets file', no_window=True, file_types=(('Excel files',"*.xlsx"), ('Excel files',"*.xls")))
        if 1 > 0 or outlets_filename:
            pass
window.close()