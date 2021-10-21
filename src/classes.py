import contextlib
import numpy as np
import os
import pandas as pd
import sqlite3
from subprocess import call
import time

from functions import *

class SkuAnalysis():

    def __init__(self, window):
        self.cp = mycprint(window)
    
    def ordered_time_periods(self, years, df_per_vals, period_type=None):    
        order_by_type = {
            'Food':['Feb-Mar','Apr-May','Jun-Jul','Aug-Sep','Oct-Nov','Dec-Jan'],
            'Mountly':['Jan-Feb','Feb-Mar','Mar-Apr','Apr-May','May-Jun','Jun-Jul','Jul-Aug','Aug-Sep','Sep-Oct','Oct-Nov','Nov-Dec','Dec-Jan'],
            'Non_Food':['Jan-Feb','Mar-Apr','May-Jun','Jul-Aug','Sep-Oct','Nov-Dec']
        }
        ordered_periods = []
        for year in years:
            first_periods = order_by_type[period_type][:-1]
            last_period = order_by_type[period_type][-1]
            for period in first_periods:
                ordered_periods.append('{} {}'.format(period, year))
            ordered_periods.append('{} {}'.format(last_period, year+1))
        first_idx = [i for i,v in enumerate(ordered_periods) if v in df_per_vals][0]
        last_idx = len(ordered_periods) - [i for i,v in enumerate(ordered_periods[::-1]) if v in df_per_vals][0]
        return ordered_periods[first_idx:last_idx]
        
    def perform_analysis(self, clusters_df, skus_df, skus, period_type):
        ''' Returns dict of type {sku_name_1:sku_df_1, sku_name_2:sku_df_2, ...}'''
        
        start = time.time()
        # Needed columns
        columns_kept = ['IDOutlet', 'PeriodName', 'Purch']

        sku_report = {}
        for sku in skus:
            self.cp(f'sku: {sku} START')
            product, brand, _sku = map(int, sku.split('-'))
            mask = (skus_df['IDProduct'] == product) & (skus_df['IDBrand'] == brand) & (skus_df['IDGoods'] == _sku)
            sku_df = skus_df[mask]
            sku_df = sku_df[columns_kept]
            
            # Locate years that appear in PeriodName
            period_name_unique_values = skus_df['PeriodName'].unique()
            years = [int(p.split()[-1]) for p in period_name_unique_values]
            min_year = min(years)-1
            max_year = max(years)+1
            years = [y for y in range(min_year, max_year+1)]
            
            # Make the ordered periods for the columns
            all_ordered_time_periods = self.ordered_time_periods(years, period_name_unique_values, period_type)
            
            # Pivot the table. Transform PeriodName from single column to multiple columns with the sales as values
            pivoted_index = ['IDOutlet'] #, 'PeriodType', 'OutletType', 'Area',]
            pivoted_values = ['Purch'] #, 'PeriodType', 'OutletType', 'Area',]
            sku_df_pivoted = sku_df.pivot_table(index=pivoted_index, columns=['PeriodName'], values=pivoted_values).reset_index()
            first_cols = [item[0] for item in sku_df_pivoted.columns if not item[-1]]
            sku_df_pivoted.columns = [item[0] if not item[-1] else item[-1] for item in sku_df_pivoted.columns]
            
            # Extend columns with potentially missing periods
            columns = first_cols + all_ordered_time_periods
            sku_df_pivoted = sku_df_pivoted.reindex(columns, axis='columns')
            
            # Remove rows with only 0 or 1 non-missing sales values, since they will produce only NaN Diffs
            sku_df_pivoted = sku_df_pivoted[
                sku_df_pivoted[all_ordered_time_periods].count(axis='columns') > 1
            ]
            # Fill sales in periods with missing values
            sku_df_pivoted[all_ordered_time_periods] = sku_df_pivoted[all_ordered_time_periods].fillna(0)

            # If there are not at least two columns with positive sales, continue to next sku, since there can be no Diffs
            positive_sales = [sku_df_pivoted[c].sum() > 0 for c in all_ordered_time_periods]
            if not (sum(positive_sales) > 1):
                continue

            # Drop trailing columns with no sales
            first_period_with_sales_ind = positive_sales.index(True)
            last_period_with_sales_ind = len(positive_sales) - positive_sales[::-1].index(True)
            cleaned_period_columns = all_ordered_time_periods[first_period_with_sales_ind:last_period_with_sales_ind]
            cleaned_columns = first_cols + cleaned_period_columns
            sku_df_pivoted = sku_df_pivoted[cleaned_columns]

            # Create cluster column
            cluster_dict = {'Food':'Cluster_food', 'Non_Food':'Cluster_non_Food', 'Mountly':'Cluster_Mountly'}
            cluster_column = sku_df_pivoted['IDOutlet'].map(clusters_df.set_index('IDOutlet')[cluster_dict[period_type]])
            # Replace NaN with 1.0, when 1. outlet is not in cluster file,
            #                         or 2. outlet exists in cluster file but has missing value on the specific period type
            cluster_column = cluster_column.fillna(1.0)
            if 'cluster' not in sku_df_pivoted.columns:
                location = sku_df_pivoted.columns.get_loc(cleaned_period_columns[0])
                sku_df_pivoted.insert(location, 'cluster', cluster_column)
                first_cols.append('cluster')
            
            # Create dfs with period differences
            sku_df_with_diffs = sku_df_pivoted.copy()
            diff_columns = ['Diff_{}'.format(i+1) for i in range(len(cleaned_period_columns)-1)]
            sku_df_with_diffs[diff_columns] = pd.DataFrame(sku_df_with_diffs[cleaned_period_columns].apply(lambda r:pd.Series(diffs(r)), axis='columns'))
            
            # Convert wide to long format
            sku_df_with_diffs_long = pd.wide_to_long(sku_df_with_diffs, ['Diff_'], i='IDOutlet', j='Diff_Order')
            sku_df_with_diffs_long = sku_df_with_diffs_long.reset_index()
            sku_df_with_diffs_long = sku_df_with_diffs_long.rename(columns={'Diff_': 'Diff_Values'})
            columns = first_cols + ['Diff_Values']
            sku_df_with_diffs_long = sku_df_with_diffs_long[columns]
        
            # Remove NaN Diff_Values
            sku_df_with_diffs_long_clean = sku_df_with_diffs_long[sku_df_with_diffs_long['Diff_Values'].notnull()]

            ###################### MAKE ALL CLUSTERS = 1 ######################
            sku_df_with_diffs_long_clean = sku_df_with_diffs_long_clean.copy()
            sku_df_with_diffs_long_clean.loc[:, 'cluster'] = 1
            ###################################################################
        
            # Group by cluster
            grouped_per_cluster = sku_df_with_diffs_long_clean[['cluster', 'Diff_Values']].groupby('cluster')
            
            # https://stackoverflow.com/questions/19894939/calculate-arbitrary-percentile-on-pandas-groupby
            # https://pandas.pydata.org/pandas-docs/stable/user_guide/groupby.html?highlight=filter
            # Coefficient of Variation (CV)
            # https://en.wikipedia.org/wiki/Coefficient_of_variation

            report_df = grouped_per_cluster.agg(N=('Diff_Values','count'),
                                                min_Diff=('Diff_Values', np.min),
                                                max_Diff=('Diff_Values', np.max),
                                                mean_Diff=('Diff_Values', np.mean),
                                                std_Diff=('Diff_Values', np.std),
                                                CV_Diff=('Diff_Values', lambda x: np.std(x, ddof=1) / np.mean(x) if np.mean(x) else np.nan),
                                                perc90_Diff=('Diff_Values', lambda x: x.quantile(0.9)),
                                                perc95_Diff=('Diff_Values', lambda x: x.quantile(0.95)),
                                                perc99_Diff=('Diff_Values', lambda x: x.quantile(0.99)),
                                                ).reset_index()

            report_df = report_df[report_df['CV_Diff'].notnull()]
            report_df['cluster'] = report_df['cluster'].astype(int)

            # REMOVE ROWS WITH 'N' < limit
            limit = 10
            report_df = report_df.loc[report_df['N'] >= limit]

            sku_report[sku] = report_df
            self.cp(f'sku: {sku} END {timer(start, time.time())}')

        return sku_report

##############################################################################################################################
##############################################################################################################################

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
    
##############################################################################################################################
##############################################################################################################################

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

    def excel_to_csv(self, vbscript, excel_filename, csv_filename):
        with open(vbscript, 'w') as f:
            f.write(VB_EXCEL_TO_CSV)
        call(['cscript.exe', vbscript, excel_filename, csv_filename, '1'])
        self.cp('Reading csv file ... please wait ...', l=False)

    def parse_clusters(self, filename):
        start = time.time()
        self.cp(f'Parsing "clusters" file ... please wait ...')
        c = SUCCESS_OUTPUT_FORMAT
        df_columns = ['IDOutlet', 'Cluster_Mountly', 'Cluster_food', 'Cluster_non_Food']
        column_types = {'IDOutler':'int64', 'Cluster_Mountly':'Int8',
                        'Cluster_food':'Int8', 'Cluster_non_Food':'Int8'}
        try:
            xl_fl = pd.ExcelFile(filename)
            df = xl_fl.parse(xl_fl.sheet_names[0], usecols=df_columns)#,  dtype=column_types)
            if not df['IDOutlet'].is_unique or df['IDOutlet'].isnull().any():
                raise
            duration = timer(start, time.time())
            message = f'"clusters" file parsed in {duration}! {len(df)} unique outlets.'
            return df
        except:
            message = f'Error while parsing "cluster" file. Please check your input.'
            c = ERROR_OUTPUT_FORMAT
        finally:
            self.cp(message, c=c, l=False)

    def parse_skus(self, filename, selected_sku_type):
        start = time.time()
        self.cp(f'Parsing "skus" file ... please wait ...')
        c = SUCCESS_OUTPUT_FORMAT
        df_columns = [
                'IDOutlet',
                'PeriodType',
                'PeriodName',
                # 'Sales NU' column was replaced by 'Purch', so if any variable's name contains 'sales'
                # it now means 'purchases'
                'Purch',
                'SKU Name',
                'IDProduct',
                'IDBrand',
                'IDGoods',
            ]
        column_types = {'IDOutlet':'int16', 'PeriodType':'category', 'PeriodName':'category',
                         'Purch':'Int16', 'IDProduct':'Int16', 'IDBrand':'Int16', 'IDGoods':'Int16',}
        try:
            _, extension = os.path.splitext(filename)
            if extension in ['.xlsx', '.xls']:
                self.cp('Converting excel to csv ... please wait ...', l=False)
                try:
                    vbscript = os.path.join(os.getcwd(), 'ExcelToCsv.vbs')
                    csv_filename = os.path.join(os.getcwd(), f'{int(time.time())}.csv')
                    self.excel_to_csv(vbscript,  filename, csv_filename)
                    df = pd.read_csv(csv_filename, usecols=df_columns, sep=None)#, dtype=column_types)
                except:
                    self.cp('Cannot convert to csv or read converted file. Reading excel file ... please wait ...', l=False)
                    xl_fl = pd.ExcelFile(filename)
                    df = xl_fl.parse(xl_fl.sheet_names[0], usecols=df_columns)#, dtype=column_types)
                finally:
                    self.delete_files(vbscript)
                    self.delete_files(csv_filename)
            elif extension == '.csv':
                df = pd.read_csv(filename, usecols=df_columns, sep=None)#, dtype=column_types)
            if df['PeriodType'].nunique(dropna=False) != 1 or df['PeriodType'].unique()[0] != selected_sku_type:
                raise
            id_cols = ['IDProduct','IDBrand','IDGoods']
            id_cols_name = id_cols + ['SKU Name']
            unique_skus_df = df[id_cols_name].drop_duplicates(subset=id_cols)
            sku_ids_names = dict(zip(join_columns(unique_skus_df, id_cols), unique_skus_df['SKU Name'].apply(lambda x:' '.join(x.split()))))
            duration = timer(start, time.time())
            message = f'"skus" file parsed in {duration}! {len(unique_skus_df)} unique skus.'
            return df, sku_ids_names
        except:
            message = f'Error while parsing "skus" file.\n- Please check your input and \n- Make sure imported file matches the sku type you selected.'
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
        