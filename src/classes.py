import contextlib
# import numpy as np
import os
import pandas as pd
import sqlite3
from subprocess import call
import time

from functions import *

class DbManager():

    def __init__(self, window):
        self.db_filename = os.path.join(os.getcwd(), 'db.sqlite3')
        # Method to print on app console
        self.cp = mycprint(window)

    def table_exists(self, table_name):
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            # Use context manager to auto commit or rollback
            with _con as con:
                self.cp(f'Checking for table "{table_name}"\n...')
                sql = f"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='{table_name}'"
                return con.execute(sql).fetchone()[0]

    def create_clusters_tables(self):
        columns = ['c1', 'c2']
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                try:
                    # Clusters Table
                    sql = f'CREATE TABLE IF NOT EXISTS clusters({", ".join(columns)})'
                    con.execute(sql)
                except:
                    self.cp('Error while creating "clusters" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)

    def delete_table_rows(self, table_name):
        if not self.table_exists(table_name):
            self.cp(f'"{table_name}" database table does not exist.', c=ERROR_OUTPUT_FORMAT)
        else:
            # Use context manager to auto close connection
            with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
                # Use context manager to auto commit or rollback
                with _con as con:
                    try:
                        db_entries = con.execute(f'SELECT COUNT(*) FROM {table_name}').fetchone()[0]
                        if not db_entries:
                            message = f'0 rows deleted from "{table_name}" database table'
                        else:
                            start = time.time()
                            con.execute(f'DELETE FROM {table_name}')
                            duration = timer(start, time.time())
                            message = f'{db_entries} rows succesfully deleted from "{table_name}" database table in {duration}!'
                        self.cp(message, c=SUCCESS_OUTPUT_FORMAT)
                    except:
                        self.cp(f'Error while clearing "{table_name}" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)

    def update_clusters(self, values=[]):
        self.delete_table_rows('clusters')
        col_num = 50
        row_num = 10_000
        columns = [f'c{i}' for i in range(1, col_num+1)]
        values = [col_num*(i,) for i in range(1,row_num+1)]

        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                try:
                    # Clusters Table
                    sql = f'CREATE TABLE IF NOT EXISTS clusters({", ".join(columns)})'
                    con.execute(sql)
                except:
                    self.cp('Error while updating "clusters" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)
            with _con as con:
                try:
                    # Insert values
                    self.cp(f'Inserting {len(values)} rows in "clusters" database table ... please wait ...', l=False)
                    start = time.time()
                    con.executemany(f'INSERT INTO clusters VALUES ({", ".join(col_num*["?"])})', values)
                    db_entries = con.execute('SELECT COUNT(*) FROM clusters').fetchone()[0]
                    duration = timer(start, time.time())
                    self.cp(f'"clusters" database table succesfully updated in {duration}! Current rows: {db_entries}', c=SUCCESS_OUTPUT_FORMAT)
                except:
                    self.cp('Error while updating "clusters" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)
    

class IOManager():
    
    def __init__(self, window):
        # Method to print on app console
        self.cp = mycprint(window)

    def delete_files(self, *args):
        for f in args:
            try:
                os.remove(f)
            except OSError:
                pass

    def parse_clusters(self, filename):
        start = time.time()
        self.cp(f'Parsing "clusters" file ... please wait ...')
        c = SUCCESS_OUTPUT_FORMAT
        df_columns = ['IDOutlet', 'Cluster_Mountly', 'Cluster_food', 'Cluster_non_Food']
        try:
            xl_fl = pd.ExcelFile(filename)
            df = xl_fl.parse(xl_fl.sheet_names[0], usecols=df_columns,
                             dtype={'IDOutler':'int64', 'Cluster_Mountly':'Int8',
                                    'Cluster_food':'Int8', 'Cluster_non_Food':'Int8'})
            duration = timer(start, time.time())
            message = f'"clusters" file parsed in {duration}!'
            return df
        except:
            message = f'Error while parsing "cluster" file. Please check your input.'
            c = ERROR_OUTPUT_FORMAT
        finally:
            self.cp(message, c=c, l=False)

    def parse_skus(self, filename):
        start = time.time()
        self.cp(f'Parsing "skus" file ... please wait ...')
        c = SUCCESS_OUTPUT_FORMAT
        df_columns = [
                'IDOutlet',
                'PeriodType',
                # 'OutletType',
                # 'Area',
                'PeriodName',
                # 'Sales NU' column was replaced by 'Purch', so if any variable's name contains 'sales'
                # it now means 'purchases'
                'Purch',
                # 'SKU Name',
                'IDProduct',
                'IDBrand',
                'IDGoods',
            ]
        column_types = {'PeriodType':'category', 'PeriodName':'category',
                        'IDOutlet':'int16', 'Purch':'Int16',
                        'IDProduct':'Int16', 'IDBrand':'Int16', 'IDGoods':'Int16',}
        try:
            _, extension = os.path.splitext(filename)
            if extension in ['.xlsx', '.xls']:
                self.cp('Converting excel to csv ... please wait ...', l=False)
                try:
                    vbscript = os.path.join(os.getcwd(), 'ExcelToCsv.vbs')
                    csv_filename = os.path.join(os.getcwd(), f'{int(time.time())}.csv')
                    with open(vbscript, 'w') as f:
                        f.write(VB_EXCEL_TO_CSV)
                    call(['cscript.exe', vbscript, filename, csv_filename, '1'])
                    self.cp('Reading csv file ... please wait ...', l=False)
                    df = pd.read_csv(csv_filename, usecols=df_columns, dtype=column_types, sep=None)
                except:
                    self.cp('Cannot convert to csv or read converted file. Reading excel file ... please wait ...', l=False)
                    xl_fl = pd.ExcelFile(filename)
                    df = xl_fl.parse(xl_fl.sheet_names[0], usecols=df_columns, dtype=column_types)
                finally:
                    self.delete_files(vbscript)
                    self.delete_files(csv_filename)
            elif extension == '.csv':
                df = pd.read_csv(filename, usecols=df_columns, dtype=column_types, sep=None)
            duration = timer(start, time.time())
            message = f'"skus" file parsed in {duration}!'
            return df
        except:
            message = f'Error while parsing "skus" file. Please check your input.'
            c = ERROR_OUTPUT_FORMAT
        finally:
            self.cp(message, c=c, l=False)
        
        
        


    def parse_file(self, xl_file_to_parse, file_type='sku'):
        '''
        file_types = sku, cluster, outlet
        We assume the the file that is imported has at least the columns below (df_columns) in one sheet.
        This is the sheet that is parsed
        '''
        
        start = time.time()
        if file_type == 'sku':
            df_columns = [
                'IDOutlet',
                'PeriodType',
                'OutletType',
                'Area',
                'PeriodName',
                # 'Sales NU' column was replaced by 'Purch', so if any variable's name contains 'sales'
                # it now means 'purchases'
                'Purch',
                'SKU Name',
                'IDProduct',
                'IDBrand',
                'IDGoods',
            ]        
        elif file_type == 'cluster':
            df_columns = [
                'IDOutlet',
                'Cluster_Mountly',
                'Cluster_food',
                'Cluster_non_Food',
            ]
        elif file_type == 'outlet':
            df_columns = [
                'IDOutlet',
                'PeriodType',
                # 'IDPeriod',
                'IDProduct',
                'IDBrand',
                'IDGoods',
                'LMPurch',
                'Purch',
            ]
        xl_fl = pd.ExcelFile(xl_file_to_parse)
        sheet_names = xl_fl.sheet_names
        if len(sheet_names) != 1:
            raise
        for sheet_name in sheet_names:
            one_row_df = xl_fl.parse(sheet_name, nrows=1)
            sheet_columns = one_row_df.columns
            if all([c in sheet_columns for c in df_columns]):
                df = xl_fl.parse(sheet_name)[df_columns]
                if file_type in ['sku', 'outlet']:
                    df = df.rename(
                            columns={'IDProduct': 'ID Product',
                                    'IDBrand': 'ID Brand',
                                    'IDGoods': 'ID SKU'}
                                    )
                if file_type == 'cluster':
                    if (not df['IDOutlet'].is_unique) or (df['IDOutlet'].isnull().sum() > 0):
                        raise
                    duration = timer(start, time.time())
                    message = f'"{file_type}" file parsed in {duration}!'
                    self.cp(message, c=SUCCESS_OUTPUT_FORMAT)
                return df
            else:
                raise
        