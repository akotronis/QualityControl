import contextlib
import numpy as np
import os
import pandas as pd
import sqlite3
from subprocess import call
import time

from constants import WARNING_OUTPUT_FORMAT, ERROR_OUTPUT_FORMAT, SUCCESS_OUTPUT_FORMAT, INFO_OUTPUT_FORMAT, VB_EXCEL_TO_CSV, \
                      CREATE_CLUSTERS_SQL, CREATE_SKUS_SQL, CREATE_ANALYSIS_SQL, CREATE_ATYPICALS_SQL, CREATE_MISSING_SQL
from functions import *

class Analysis():

    def __init__(self, window, clusters_to_1=False, counts_lower_bound=0):
        self.cp = mycprint(window)
        self.clusters_to_1 = clusters_to_1
        self.counts_lower_bound = counts_lower_bound
    
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
        
    def sku_analysis(self, clusters_df, skus_df, skus, period_type):
        ''' Returns dict of type {sku_name_1:sku_df_1, sku_name_2:sku_df_2, ...}'''

        self.cp(f'Analysing SKUs ({period_type[0]}) ... please wait ...')
        start = time.time()
        # Needed columns
        columns_kept = ['IDOutlet', 'PeriodName', 'Purch']

        sku_report = {}
        try:
            c = SUCCESS_OUTPUT_FORMAT
            for sku in progress_bar(f' Analysing SKUs ({period_type[0]}) '.center(30, '='), skus, title='', orientation='v', keep_on_top=False, grab_anywhere=True, no_titlebar=True, no_button=True):
                # self.cp(f'sku: {sku} START')
                # Filter skus dataframe for specfic sku
                product, brand, _sku = map(int, sku.split('-'))
                mask = (skus_df['IDProduct'] == product) & (skus_df['IDBrand'] == brand) & (skus_df['IDGoods'] == _sku)
                sku_df = skus_df[mask]
                sku_df = sku_df[columns_kept]
                
                # Locate years that appear in PeriodName
                period_name_unique_values = set(skus_df['PeriodName'].values)
                years = [int(p.split()[-1]) for p in period_name_unique_values]
                min_year = min(years)-1
                max_year = max(years)+1
                years = [y for y in range(min_year, max_year+1)]
                
                # Make the ordered periods for the columns
                all_ordered_time_periods = self.ordered_time_periods(years, period_name_unique_values, period_type)
                
                # Pivot the table. Transform PeriodName from single column to multiple columns with the sales as values
                pivoted_index = ['IDOutlet']
                pivoted_values = ['Purch']
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
                cluster_dict = {'Mountly':'mountly', 'Food':'food', 'Non_Food':'non_food'}
                cluster_column = sku_df_pivoted['IDOutlet'].map(clusters_df.set_index('id_outlet')[cluster_dict[period_type]])
                # Replace NaN with 1.0, when 1. outlet is not in cluster file,
                #                         or 2. outlet exists in cluster file but has missing value on the specific period type
                cluster_column = cluster_column.fillna(1.0)
                if 'cluster' not in sku_df_pivoted.columns:
                    location = sku_df_pivoted.columns.get_loc(cleaned_period_columns[0])
                    sku_df_pivoted.insert(location, 'cluster', cluster_column)
                    first_cols.append('cluster')
                
                # Create dfs with period differences
                sku_df_with_diffs = sku_df_pivoted.reset_index(drop=True) #.copy()
                diff_columns = ['Diff_{}'.format(i+1) for i in range(len(cleaned_period_columns)-1)]
                # Slower way
                # sku_df_with_diffs[diff_columns] = sku_df_with_diffs[cleaned_period_columns].apply(lambda r:pd.Series(diffs_old(r)), axis='columns')
                # Faster way
                sku_df_with_diffs[diff_columns] = diffs(sku_df_with_diffs[cleaned_period_columns])

                # Convert wide to long format
                sku_df_with_diffs_long = pd.wide_to_long(sku_df_with_diffs, ['Diff_'], i='IDOutlet', j='Diff_Order')
                sku_df_with_diffs_long = sku_df_with_diffs_long.reset_index()
                sku_df_with_diffs_long = sku_df_with_diffs_long.rename(columns={'Diff_': 'Diff_Values'})
                columns = first_cols + ['Diff_Values']
                sku_df_with_diffs_long = sku_df_with_diffs_long[columns]
            
                # Remove NaN Diff_Values
                sku_df_with_diffs_long_clean = sku_df_with_diffs_long[sku_df_with_diffs_long['Diff_Values'].notnull()]

                ###################### MAKE ALL CLUSTERS = 1 ######################
                # sku_df_with_diffs_long_clean = sku_df_with_diffs_long_clean.copy()
                if self.clusters_to_1:
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
                if self.counts_lower_bound:
                    report_df = report_df.loc[report_df['N'] >= self.counts_lower_bound]

                sku_report[sku] = report_df
                # self.cp(f'sku: {sku} END {timer(start, time.time())}')
            message = f'Sku analysis finished in {timer(start, time.time())}!'
            return sku_report
        except Exception as e:
            c = ERROR_OUTPUT_FORMAT
            message = 'Error while analysing SKUs'
            print(e)
        finally:
            self.cp(message, c=c)
        
    def outlet_analysis(self, clusters_df, analysis_df, outlets_df, period_type):
        self.cp(f'Analysing Outlets ({period_type[0]}) (preprocessing data) ... please wait ...')
        start = time.time()
        try:
            c1 = SUCCESS_OUTPUT_FORMAT
            c2 = INFO_OUTPUT_FORMAT
            outlets_df = outlets_df.loc[(outlets_df['LMPurch'] > 0) & (outlets_df['Purch'] > 0)]
            # Calculate D=y*ln(y/y')+y'-y where y=LMPurch and y'=Purch
            # Slower way
            # outlets_df['diff'] = outlets_df[['LMPurch', 'Purch']].apply(lambda r:pd.Series(diffs_old(r)), axis='columns')
            # Faster way
            outlets_df['diff'] = diffs(outlets_df[['LMPurch', 'Purch']])
            cluster_dict = {'Mountly':'mountly', 'Food':'food', 'Non_Food':'non_food'}
            cluster_column = outlets_df['IDOutlet'].map(clusters_df.set_index('id_outlet')[cluster_dict[period_type]])
            # Replace NaN with 1.0, when 1. outlet is not in cluster file,
            #                         or 2. outlet exists in cluster file but has missing value on the specific period type
            cluster_column = cluster_column.fillna(1.0)
            outlets_df['cluster'] = cluster_column.astype(int)
            ###################################################################
            ###################### MAKE ALL CLUSTERS = 1 ######################
            if self.clusters_to_1:
                outlets_df['cluster'] = 1
            ###################################################################
            ###################################################################
            outlets_df['PeriodType'] = outlets_df['PeriodType'].map({'Mountly':1, 'Food':2, 'Non_Food':3})
            sku_columns = ['IDProduct', 'IDBrand', 'IDGoods']
            sku_cols_with_ptype = sku_columns + ['PeriodType']
            # Make sku_id column from four ID columns: Product, Brand, SKU and PeriodType
            outlets_df['sku_id'] = join_columns(outlets_df, sku_cols_with_ptype)
            # Make sku_id column from four ID columns: Product, Brand, SKU and PeriodType
            analysis_df['sku_id'] = join_columns(analysis_df, ['sku_id', 'period_type'])
            # Merge outlets with analysis
            outlets_df = outlets_df.merge(analysis_df, on=['sku_id', 'cluster'], how='left')
            outlets_df = outlets_df.loc[~(outlets_df['diff'] <= outlets_df['perc90_diff'])]
            
            # Initialize missing_atypicals_per_outlet dict as below
            # missing_atypicals_per_outlet = {outlet_id: {'missing': list of dicts,
            #                                             'atypicals': list of dicts}}
            missing_atypicals_per_outlet = {}
            # Group outlets_df per outlet
            grouped_per_outlet = outlets_df.groupby('IDOutlet')
            for outlet_id, outlet_df in progress_bar(f' Analysing Outlets ({period_type[0]}) '.center(30, '='), grouped_per_outlet, title='', orientation='v', keep_on_top=False, grab_anywhere=True, no_titlebar=True, no_button=True):
                missing_atypicals_per_outlet[outlet_id] = {'missing':[], 'atypicals':[]}
                # NOT Found (Product, Brand, SKU, PeriodType, cluster) in analysis table
                df_not_found = outlet_df.loc[outlet_df['mean_diff'].isnull()]
                if not df_not_found.empty:
                    # Create mssing objects
                    missing_line_dicts = [{h:v for h,v in zip(df_not_found.columns, line)} for line in df_not_found.values]
                    # Add missing SKUS of specific outlet to the "missing" list containing all the missing SKUS
                    missing_atypicals_per_outlet[outlet_id]['missing'] = missing_line_dicts
                ############## Perform analysis on df with FOUND skus ##############
                # Found (Product, Brand, SKU, PeriodType, cluster) in analysis table
                df = outlet_df.loc[outlet_df['mean_diff'].notnull()]
                # df is NOT empty means we have atypical values
                if not df.empty:
                    df['stars'] = df[['diff', 'perc90_diff', 'perc95_diff', 'perc99_diff']].apply(
                        lambda r:'***' if not (r['diff'] <= r['perc99_diff']) else ('**' if not (r['diff'] <= r['perc95_diff'])  else '*'),
                        axis='columns'
                    )
                    # Make HRH value with newton's method:
                    # Solving f(x)=mean_diff where f(x)=y*ln(y/x)+x-y, y=LMPurch
                    df['proposed_purch_1'] = df[['LMPurch', 'mean_diff']].apply(
                        lambda r:newton(
                                    lambda x: r['LMPurch']*log(r['LMPurch']/x)+x-r['LMPurch']-r['mean_diff'],
                                    lambda x: -r['LMPurch']/x + 1,
                                    # Approximating the solution that is < than LMPurch, so we give an initial value < LMPurch
                                    r['LMPurch'] / 2,
                                ), axis='columns')            
                    df['proposed_purch_2'] = df[['LMPurch', 'mean_diff']].apply(
                        lambda r:newton(
                                    lambda x: r['LMPurch']*log(r['LMPurch']/x)+x-r['LMPurch']-r['mean_diff'],
                                    lambda x: -r['LMPurch']/x + 1,
                                    # Approximating the solution that is > than LMPurch, so we give an initial value > LMPurch
                                    r['LMPurch'] + 1,
                                ), axis='columns')
                    atypicals_line_dicts = [{h:v for h,v in zip(df.columns, line)} for line in df.values]
                    # Add atypicals of specific outlet to the "atypicals" list containing all the atypicals
                    missing_atypicals_per_outlet[outlet_id]['atypicals'] = atypicals_line_dicts
            message1 = f'Outlet analysis finished in {timer(start, time.time())}!'
            message2 = '\n'.join([f"Store {k}: {len(v['atypicals'])} atypical, {len(v['missing'])} missing SKU(s) found"
                                for k,v in missing_atypicals_per_outlet.items()])
            return missing_atypicals_per_outlet
        except Exception as e:
            c1 = ERROR_OUTPUT_FORMAT
            message1 = 'Error while analysing Outlets'
            message2 = ''
            print(e)
        finally:
            self.cp(message1, c=c1)
            self.cp(message2, c=c2)

##############################################################################################################################
##############################################################################################################################

class DbManager():

    def __init__(self, window):
        self.db_filename = os.path.join(os.getcwd(), 'db.sqlite3')
        # Method to print on app console
        self.cp = mycprint(window)
        self.create_tables()

    def connection_is_open(self, con):
        try:
            con.cursor()
            return True
        except Exception as ex:
            return False
    
    def database_exists(self):
        return 'db.sqlite3' in os.listdir(os.getcwd())

    def create_tables(self):
        # Use context manager to auto close connection
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            # Use context manager to auto commit or rollback
            with _con as con:
                for sql  in [CREATE_CLUSTERS_SQL, CREATE_SKUS_SQL, CREATE_ANALYSIS_SQL, CREATE_ATYPICALS_SQL, CREATE_MISSING_SQL]:
                    con.execute(sql)

    def table_exists(self, table_name):
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                self.cp(f'Checking for table "{table_name}" ...')
                sql = f"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='{table_name}'"
                return con.execute(sql).fetchone()[0]

    def table_count(self, table_name, period_type=False, report=False):
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                sql_table_name = table_name
                if table_name == 'analysis' and report:
                    sql_table_name = f'''(SELECT skus.period_type
                                          FROM analysis LEFT JOIN skus ON analysis.sku_id=skus.id)'''
                where = ''
                if period_type:
                    where = f' WHERE period_type={period_type}'
                
                sql = f"SELECT COUNT(*) FROM {sql_table_name}{where}"
                return con.execute(sql).fetchone()[0]

    def table_is_empty(self, table_name):
        return not self.table_count(table_name)

    def skus_already_in_database(self, sku_ids, selected_period_type):
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                sku_ids_list = ', '.join([f"'{sku_id}'" for sku_id in sku_ids])
                sql = f"SELECT COUNT(*) FROM skus WHERE period_type='{selected_period_type}' AND sku_id IN ({sku_ids_list})"
                already_in = con.execute(sql).fetchone()[0]
                if already_in:
                    self.cp(f'{already_in} skus are already in database. Please check your input', c=WARNING_OUTPUT_FORMAT)
                return already_in

    def delete_table_rows(self, table_name):
        if not self.table_exists(table_name):
            self.cp(f'"{table_name}" database table does not exist.', c=ERROR_OUTPUT_FORMAT)
        else:
            with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
                with _con as con:
                    try:
                        db_entries = con.execute(f'SELECT COUNT(*) FROM {table_name}').fetchone()[0]
                        if table_name == 'skus':
                            analysis_entries = con.execute(f'SELECT COUNT(*) FROM analysis').fetchone()[0]
                        if not db_entries:
                            message = f'0 rows deleted from "{table_name}" database table'
                        else:
                            start = time.time()
                            con.execute(f'DELETE FROM {table_name}')
                            duration = timer(start, time.time())
                            message = f'{db_entries} rows succesfully deleted from "{table_name}" database table in {duration}!'
                            if table_name == 'skus':
                                con.execute(f'DELETE FROM analysis')
                                duration = timer(start, time.time())
                                message += f'\n{analysis_entries} rows succesfully deleted from "analysis" database table in {duration}!'
                        self.cp(message, c=SUCCESS_OUTPUT_FORMAT)
                    except Exception as e:
                        self.cp(f'Error while clearing "{table_name}" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)
                    finally:
                        pass

    def update_table(self, table_name, values=[], sku_df_dict=None, missing_atypicals_per_outlet=None):
        # Clear table
        if table_name not in ['skus', 'analysis']:
            self.delete_table_rows(table_name)
        c = SUCCESS_OUTPUT_FORMAT
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                try:
                    # Get database table columns
                    db_columns = [c[1] for c in con.execute(f'PRAGMA table_info({table_name})').fetchall()[1:]]
                    # Make values for tables with foreign key (analysis, atypicals)
                    if table_name =='analysis':
                        values = []
                        for sku_id, df in sku_df_dict.items():
                            sku_row_id = con.execute(f'SELECT id from skus WHERE sku_id="{sku_id}"').fetchone()[0]
                            rows = [row + [sku_row_id] for row in df.values.tolist()]
                            values.extend(rows)
                    elif table_name == 'missing':
                        values = []
                        for outlet_id, miss_atyp_dict in missing_atypicals_per_outlet.items():
                            for row_dict in miss_atyp_dict['missing']:
                                values.append(['-'.join(row_dict['sku_id'].split('-')[:-1]), row_dict['PeriodType'], row_dict['cluster'],
                                            outlet_id, row_dict['LMPurch'], row_dict['Purch']])
                    elif table_name == 'atypicals':
                        values = []
                        for outlet_id, miss_atyp_dict in missing_atypicals_per_outlet.items():
                            for row_dict in miss_atyp_dict['atypicals']:
                                sku_id = '-'.join(row_dict['sku_id'].split('-')[:-1])
                                period_type = row_dict['PeriodType']
                                cluster = row_dict['cluster']
                                sql = f'''SELECT id FROM
                                            (SELECT analysis.id, skus.sku_id, period_type, cluster
                                                FROM analysis JOIN skus ON analysis.sku_id=skus.id)
                                            WHERE sku_id="{sku_id}" AND period_type={period_type} AND cluster={cluster}
                                        '''
                                analysis_id = con.execute(sql).fetchone()
                                # analysis_id = analysis_id[0]
                                if analysis_id is not None:
                                    analysis_id = analysis_id[0]
                                else:
                                    print(f'{sku_id}, ptype:{period_type}, cluster:{cluster}')
                                row = [outlet_id, row_dict['LMPurch'], row_dict['Purch'], row_dict['stars'],
                                        row_dict['proposed_purch_1'], row_dict['proposed_purch_2'], analysis_id]
                                values.append(row)
                    if values:
                        # Insert values
                        self.cp(f'Inserting {len(values)} rows in "{table_name}" database table ... please wait ...')
                        start = time.time()
                        con.executemany(f'INSERT INTO {table_name}({", ".join(db_columns)}) VALUES ({", ".join(len(values[0]) * ["?"])})', values)
                        db_entries = con.execute(f'SELECT COUNT(*) FROM {table_name}').fetchone()[0]
                        duration = timer(start, time.time())
                        message = f'"{table_name}" database table succesfully updated in {duration}! Current rows: {db_entries}'
                        
                    else:
                        message = f'No values to insert in "{table_name}" database table'
                        c = WARNING_OUTPUT_FORMAT
                    self.cp(message, c=c)
                except:
                    self.cp(f'Error while updating "{table_name}" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)

    def table_to_df(self, table_name, export=False):
        with contextlib.closing(sqlite3.connect(self.db_filename)) as _con:
            with _con as con:
                try:
                    if table_name in ['clusters', 'skus', 'missing']:
                        db_columns = [c[1] for c in con.execute(f'PRAGMA table_info({table_name})').fetchall()[1:]]
                        df = pd.read_sql_query(f"SELECT {','.join(db_columns)} FROM {table_name}", con)
                    elif table_name == 'analysis':
                        count = 'count,' if export else ''
                        if not export:
                            sql = f'''SELECT skus.sku_id, skus.period_type, cluster, skus.sku_name, {count}mean_diff, perc90_diff,
                                        perc95_diff,perc99_diff
                                        FROM analysis LEFT JOIN skus ON analysis.sku_id=skus.id''' 
                        else:
                            sql = f'''SELECT skus.sku_id, skus.period_type, cluster, skus.sku_name, {count}mean_diff, perc90_diff,
                                        perc95_diff,perc99_diff
                                        FROM skus LEFT JOIN analysis ON skus.id=analysis.sku_id
                                        UNION
                                      SELECT skus.sku_id, skus.period_type, cluster, skus.sku_name, {count}mean_diff, perc90_diff,
                                        perc95_diff,perc99_diff
                                        FROM analysis LEFT JOIN skus ON skus.id=analysis.sku_id'''
                        df = pd.read_sql_query(sql, con)
                    elif table_name == 'atypicals':
                        sql = f'''SELECT oa.id_outlet, oa.cluster, skus.period_type, skus.sku_id, skus.sku_name, oa.lm_purch, 
                                    oa.purch, oa.stars, oa.proposed_purch_1, oa.proposed_purch_2
                                    FROM (atypicals LEFT JOIN analysis ON atypicals.analysis_id=analysis.id) AS oa
                                    LEFT JOIN skus on oa.sku_id=skus.id'''
                        df = pd.read_sql_query(sql, con)
                    if table_name != 'clusters':
                        added_columns = ['id_product', 'id_brand', 'id_sku']
                        df[added_columns] = df['sku_id'].str.split('-', expand=True).astype(int)
                        if export:
                            df = df.drop('sku_id', axis='columns')
                        id_cols = [c for c in df.columns if 'id_' in c]
                        other_cols = [c for c in df.columns if c not in id_cols]
                        df = df[id_cols + other_cols]
                    return df
                except:
                    self.cp(f'Error while fetching "{table_name}" database table. Please try again.', c=ERROR_OUTPUT_FORMAT)
                
##############################################################################################################################
##############################################################################################################################

class IOManager():
    
    def __init__(self, window, db):
        # Method to print on app console
        self.cp = mycprint(window)
        self.db = db

    def delete_files(self, *args):
        for f in args:
            try:
                os.remove(f)
            except OSError:
                pass

    def excel_to_csv(self, vbscript, excel_filename, csv_filename):
        with open(vbscript, 'w') as f:
            f.write(VB_EXCEL_TO_CSV)
        # call(['cscript.exe', vbscript, excel_filename, csv_filename, '1'])
        subprocess_call(['cscript.exe', vbscript, excel_filename, csv_filename, '1'])
        self.cp('Reading csv file ... please wait ...')

    def export_files(self, table_name, popup=True):
        table_is_empty = self.db.table_is_empty(table_name)
        table_name_message = {'clusters':'clusters','skus':'skus','analysis':'skus','atypicals':'outlets','missing':'outlets'}
        if table_is_empty:
            self.cp(f'"{table_name}" data base table is empty. Please import {table_name_message[table_name]}', c=WARNING_OUTPUT_FORMAT, l=True)
        else:
            try:
                c = SUCCESS_OUTPUT_FORMAT
                if popup:
                    default_export_name = table_name.title()
                    table_file = sg.popup_get_file('', save_as=True, no_window=True,
                                                    initial_folder=os.getcwd(),
                                                    default_extension='.xlsx',
                                                    default_path=f'{default_export_name}Export.xlsx',
                                                    file_types=(('Excel files',"*.xlsx"),))
                if not popup or table_file:
                    start = time.time()
                    table_df = self.db.table_to_df(table_name, export=True)
                    if 'period_type' in table_df.columns:
                        table_df['period_type'] = table_df['period_type'].map({1:'Mountly', 2:'Food', 3:'Non_Food'})
                if popup and table_file:
                    table_df.to_excel(table_file, index=False, freeze_panes=(1,0))
                    duration = timer(start, time.time())
                    message = f'{os.path.split(table_file)[1]} succesfully exported in {duration}!'
                elif not popup:
                    return table_df
            except:
                c = ERROR_OUTPUT_FORMAT
                message = 'Error while exporting file. Make sure a file with the same name is not open'
            finally:
                if popup and table_file:
                    self.cp(message, c=c)

    def export_total(self):
        clusters_df = self.export_files('clusters', popup=False)
        skus_df = self.export_files('skus', popup=False)
        analysis_df = self.export_files('analysis', popup=False)
        outlets_df = self.export_files('atypicals', popup=False)
        missing_df = self.export_files('missing', popup=False)

    def parse_file(self, filename, file_type=None, imported_ptype=None):
        start = time.time()
        period_type_txt = f'period type: {imported_ptype}, ' if imported_ptype else ''
        self.cp(f'Selected {period_type_txt}file name: {os.path.split(filename)[1]}', c=INFO_OUTPUT_FORMAT, l=True)
        self.cp(f'Parsing "{file_type}" file ... please wait ...')
        c = SUCCESS_OUTPUT_FORMAT
        if file_type == 'clusters':
            df_columns = ['IDOutlet','Cluster_Mountly','Cluster_food','Cluster_non_Food']
        elif file_type == 'skus':
            df_columns = ['IDOutlet','PeriodType','PeriodName','Purch','SKU Name',
                          'IDProduct','IDBrand','IDGoods']
        elif file_type == 'outlets':
            df_columns = ['IDOutlet','PeriodType','PeriodName','LMPurch','Purch',
                          'IDProduct','IDBrand','IDGoods']
        try:
            _, extension = os.path.splitext(filename)
            if extension in ['.xlsx', '.xls']:
                self.cp('Converting excel to csv ... please wait ...')
                try:
                    vbscript = os.path.join(os.getcwd(), 'ExcelToCsv.vbs')
                    csv_filename = os.path.join(os.getcwd(), f'{int(time.time())}.csv')
                    self.excel_to_csv(vbscript,  filename, csv_filename)
                    df = pd.read_csv(csv_filename, usecols=df_columns, sep=None, engine='python')#, dtype=column_types)
                except:
                    self.cp('Cannot convert to csv or read converted file. Reading excel file ... please wait ...', c=WARNING_OUTPUT_FORMAT)
                    xl_fl = pd.ExcelFile(filename)
                    df = xl_fl.parse(xl_fl.sheet_names[0], usecols=df_columns)#, dtype=column_types)
                finally:
                    self.delete_files(vbscript)
                    self.delete_files(csv_filename)
            elif extension == '.csv':
                df = pd.read_csv(filename, usecols=df_columns, sep=None, engine='python')#, dtype=column_types)
            # input data validation
            if file_type == 'clusters':
                if not df['IDOutlet'].is_unique or df['IDOutlet'].isnull().any():
                    raise
            else:
                if df['PeriodType'].nunique(dropna=False) != 1 or df['PeriodType'].unique()[0] != imported_ptype:
                    raise
                if df[['IDOutlet','IDProduct','IDBrand','IDGoods']].applymap(lambda x: not x > 0).any().any():
                    raise
            if file_type == 'skus':
                if df[['PeriodName','SKU Name']].isnull().any().any():
                    raise
            if file_type == 'outlets':
                if df['PeriodName'].nunique(dropna=False) != 1 or df['PeriodName'].isnull().any():
                    raise
            duration = timer(start, time.time())
            if file_type == 'skus':
                id_cols = ['IDProduct','IDBrand','IDGoods']
                id_cols_name = id_cols + ['SKU Name']
                unique_skus_df = df[id_cols_name].drop_duplicates(subset=id_cols)
                sku_ids_names_dict = dict(zip(join_columns(unique_skus_df, id_cols), unique_skus_df['SKU Name'].apply(lambda x:' '.join(x.split()))))
                message = f'"skus" file parsed in {duration}! {len(unique_skus_df)} unique skus'
                return df, sku_ids_names_dict
            else:
                message = f'"{file_type}" file parsed in {duration}!'
                if file_type == 'clusters':
                    message += f' {len(df)} unique outlets'
                else:
                    df = df.drop('PeriodName', axis='columns')
                return df
        except:
            message = f'Error while parsing "{file_type}" file.\n- Check your input'
            if file_type != 'clusters':
                message += '\n- Make sure imported file period type matches the period type you selected'
            if file_type == 'outlets':
                message += '\n- Make sure imported file has unique period name'
            c = ERROR_OUTPUT_FORMAT
        finally:
            self.cp(message, c=c)
        
