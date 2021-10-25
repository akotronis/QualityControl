#!/usr/bin/env python
from line_profiler import LineProfiler
from datetime import datetime
import os
import pandas as pd
import PySimpleGUI as sg
import tempfile
import webbrowser
from classes import *
from functions import *

"""
    SKUs Quality Control Application
"""


sg.set_options(element_padding=(0, 0))

# ------ Menu Definition ------ #
menu_def = [
    ['&Import', ['&Clusters::IC', '&SKUs', ['&Monthly::ISM', 'Bimonthly - &Food::ISF', 'Bimonthly - &Non Food::ISN'], '&Outlets', ['&Monthly::IOM', 'Bimonthly - &Food::IOF', 'Bimonthly - &Non Food::ION']]],
    ['&Export', ['&Clusters::EC', '&SKUs', ['&SKUs::ES', 'SKU &Analysis::ESA'], '&Outlets', ['&Missing::EOM','&Atypicals::EOA'], '---', '&Total Report::ET']],
    ['&Delete', ['&Clusters::DC', '&SKUs::DS']],
    ['&Console', ['&Clear::CC']],
    ['&About', ['&SKU Quality Control Application::AA', '---', '&Importing', ['&Clusters::AC', '&SKUs::AS', '&Outlets::AO'], '---', 'Read the Docs::AD']],
    ['E&xit', ['E&xit Application::EA']],
]

# ------ GUI Defintion ------ #
layout = [
    [sg.Menu(
        menu_def,
        tearoff=False,
        # bar_background_color='#f0f0f0',
        # background_color='#f0f0f0',
        text_color='black',
        key='-MN-')],
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


##################### TESTS #####################
# test_files_dir = os.path.join(os.path.abspath(os.pardir), 'test_files')
# clusters_file = os.path.join(test_files_dir, 'clusters', 'cluster file.xlsx')
# skus_file = os.path.join(test_files_dir, 'skus', 'Monthly-Juice.xlsx')
#################################################

##################### MAIN CLASSES ####################
CLUSTERS_TO_1 = False
cp = mycprint(window)
db = DbManager(window)
iom = IOManager(window, db)
#######################################################

while 1:
    toggle_menu(menu_def, window, True)
    event, values = window.read()
    # MENU "Exit" or Click X -> Close App
    title = 'WARNING!'
    message = 'This will close the application.'
    if (event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event.endswith('::EA')) and popup_yes_no(title, message):
        break
    toggle_menu(menu_def, window, False)
    # MENU "Import" -> Clusters
    if event.endswith('::IC'):
        clusters_file = sg.popup_get_file('Select Clusters file', no_window=True, file_types=(('Excel files',"*.xlsx"), ('Excel files',"*.xls")))
        if clusters_file:
            clusters_df = iom.parse_file(clusters_file, 'clusters')
            if clusters_df is not None:
                db.update_table('clusters', clusters_df.values.tolist())
    # MENU "Import" -> SKUs or Outlets
    elif '::IS' in event or '::IO' in event:
        clusters_is_empty = db.table_is_empty('clusters')
        if clusters_is_empty:
            cp('"clusters" data base table is empty. Please import clusters', c=WARNING_OUTPUT_FORMAT, l=True)
            continue
        if '::IO' in event:
            skus_is_empty = db.table_is_empty('skus')
            if skus_is_empty:
                cp('"skus" data base table is empty. Please import skus', c=WARNING_OUTPUT_FORMAT, l=True)
                continue
        file_type = 'skus' if '::IS' in event else 'outlets'
        selected_ptype = event[-1]
        input_file_ptype = {'M':'(MONTHLY)', 'F':'(BIMONTHLY - FOOD)', 'N':'(BIMONTHLY - NON FOOD)'}
        input_file = sg.popup_get_file(f'Select {file_type.upper()} file {input_file_ptype[selected_ptype]}', title=f'Select {file_type.upper()}', file_types=(('CSV files',"*.csv"), ('Excel files',"*.xlsx"), ('Excel files',"*.xls")))
        if not input_file:
            continue
        input_file_ptype = {'M':'Mountly', 'F':'Food', 'N':'Non_Food'}
        imported_ptype = input_file_ptype[selected_ptype]
        parse_output = iom.parse_file(input_file, file_type, imported_ptype)
        input_file_ptype = {'M':1, 'F':2, 'N':3}
        imported_ptype_num = input_file_ptype[selected_ptype]
        clusters_df = db.table_to_df('clusters')
        if parse_output is None or clusters_df is None:
            continue
        analysis_class = Analysis(window, CLUSTERS_TO_1)
        # MENU "Import" -> SKUs
        if file_type == 'skus':
            skus_df, sku_ids_names_dict = parse_output
            sku_ids = sku_ids_names_dict.keys()
            # Check is any sku is already in database
            already_in = db.skus_already_in_database(sku_ids, imported_ptype_num)
            if already_in:
                continue
            # Update sku database table
            date_now = datetime.now().date()
            _sku_file = os.path.split(input_file)[1]
            sku_ids = sku_ids_names_dict.keys()
            values = [[sku_id, imported_ptype_num, sku_name, _sku_file, date_now] for sku_id, sku_name in sku_ids_names_dict.items()]
            db.update_table('skus', values)
            #########################################################
            ################### Line Profile code ###################
            #########################################################
            try:
                lp = LineProfiler()
                lp_wrapper = lp(analysis_class.sku_analysis)
                sku_df_dict = lp_wrapper(clusters_df, skus_df, sku_ids, imported_ptype)
                with open('profile_output.txt', 'w') as f:
                    lp.print_stats(f)
            except:
                sku_df_dict = analysis_class.sku_analysis(clusters_df, skus_df, sku_ids, imported_ptype)
            #########################################################
            #########################################################
            #########################################################
            if sku_df_dict is not None:
                db.update_table('analysis', sku_df_dict=sku_df_dict)
        else:
            # MENU "Import" -> Outlets
            if not db.table_count('skus', imported_ptype_num):
                cp(f'No entries in "skus" database table for the selected period type', c=WARNING_OUTPUT_FORMAT)
                continue
            outlets_df = parse_output
            analysis_df = db.table_to_df('analysis')
            #########################################################
            ################### Line Profile code ###################
            #########################################################
            try:
                lp = LineProfiler()
                lp_wrapper = lp(analysis_class.outlet_analysis)
                missing_atypicals_per_outlet = lp_wrapper(clusters_df, analysis_df, outlets_df, imported_ptype)
                with open('profile_output.txt', 'w') as f:
                    lp.print_stats(f)
            except:
                missing_atypicals_per_outlet = analysis_class.outlet_analysis(clusters_df, analysis_df, outlets_df, imported_ptype)
            #########################################################
            #########################################################
            #########################################################
            if missing_atypicals_per_outlet is not None:
                db.update_table('atypicals', missing_atypicals_per_outlet=missing_atypicals_per_outlet)
                db.update_table('missing', missing_atypicals_per_outlet=missing_atypicals_per_outlet)
    # MENU "Export -> Clusters
    elif event.endswith('::EC'):
        iom.export_files('clusters')
    # MENU "Export -> SKUs
    elif event.endswith('::ES'):
        iom.export_files('skus')
    # MENU "Export -> SKUs Analysis
    elif event.endswith('::ESA'):
        iom.export_files('analysis')
    # MENU "Export -> Outlets Missing
    elif event.endswith('::EOM'):
        iom.export_files('missing')
    # MENU "Export -> Outlets Atypicals
    elif event.endswith('::EOA'):
        iom.export_files('atypicals')
    elif event.endswith('::ET'):
        iom.export_total()
    # MENU "Delete" -> Delete table rows
    elif '::D' in event:
        table_name = 'clusters' if event.endswith('::DC') else 'skus'
        table_count = db.table_count(table_name)
        message = f'This WILL DELETE all rows ({table_count})\nfrom "{table_name}" database table.'
        if event.endswith('::DC'):
            if popup_yes_no(title, message):
                db.delete_table_rows('clusters')
        if event.endswith('::DS'):
            if popup_yes_no(title, message):
                db.delete_table_rows('skus')
                db.delete_table_rows('missing')
                db.delete_table_rows('atypicals')
    # MENU "Console" -> Clear Console
    elif event.endswith('::CC'):
        cp('', u=True)
    # MENU "About" -> Print info messages
    elif '::A' in event:
        if not event.endswith('::AD'):
            if event.endswith('::AA'):
                message = APP_INFO_MESSAGE
            elif event.endswith('::AC'):
                message = CLUSTERS_IMPORT_MESSAGE
            elif event.endswith('::AS'):
                message = SKUS_IMPORT_MESSAGE
            elif event.endswith('::AO'):
                message = OUTLETS_IMPORT_MESSAGE
            cp(message, c=INFO_OUTPUT_FORMAT, l=True)
        else:
            html = DOCS
            with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html') as f:
                url = f'file://{f.name}'
                f.write(html)
            webbrowser.open(url)
window.close()