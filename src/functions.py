from math import log
import numpy as np
import pandas as pd
import PySimpleGUI as sg


APP_INFO_MESSAGE = 3*'''App info message'''
CLUSTERS_IMPORT_MESSAGE = f'''{"---------- IMPORTING A CLUSTERS FILE ----------".center(100)}

- Input must be an excel file with the following format:
  1) Must have a single sheet
  2) Columns must be the below:
     - IDOutlet (Positive integer, non-missing, unique values)
     - Cluster_Mountly (Positive integer, missing allowed)
     - Cluster_food (Positive integer, missing allowed)
     - Cluster_non_Food (Positive integer, missing allowed)
     Please note the spaces, if any, and the case of the characters (lowercase, uppercase)
  3) Columns must be on the first line, data from second to the end and no merged cells are allowed
- If a new file is submitted, all existing clusters WILL BE DELETED and replaced by the new ones

'''
SKUS_IMPORT_MESSAGE = f'''{"---------- IMPORTING AN SKUs FILE ----------".center(100)}

- Input must be an excel or csv file (csv will be much faster) with the following format:
  1) Must have a single sheet
  2) Columns must be the below:
     - IDOutlet (Positive integer, non-missing)
     - PeriodType (String, non-missing, case sensitive, unique value)
       - Allowed values: Mountly, or Food, or Non_Food, according to what type you selected
     - PeriodName (String, non-missing, case sensitive). Values should be of the form:
        a) Month year [e.g. Feb 2019] or of the form
        b) Month1-Month2 year [e.g. Feb-Mar 2019]
        Months values should be one of: Jan,Feb,Mar,Apr,Jun,Jul,Aug,Sep,Oct,Nov,Dec.
        Allowed PeriodName Month format per PeriodType:
        a) Mountly: Jan-Feb,Feb-Mar,Mar-Apr,Apr-May,May-Jun,Jun-Jul,Jul-Aug,Aug-Sep,
                    Sep-Oct,Oct-Nov,Nov-Dec,Dec-Jan
        b) Food: Feb-Mar,Apr-May,Jun-Jul,Aug-Sep,Oct-Nov,Dec-Jan 
        c) Non_Food: Jan-Feb,Mar-Apr,May-Jun,Jul-Aug,Sep-Oct,Nov-Dec
     - Purch (Numeric, missing allowed)
     - SKU Name (String, non-missing)
     - IDProduct (Positive integer, non-missing)
     - IDBrand (Positive integer, non-missing)
     - IDGoods (Positive integer, non-missing)
     Please note the spaces, if any, and the case of the characters (lowercase, uppercase)
  3) Columns must be on the first line, data from second to the end and no merged cells are allowed
- Files that don't have the above format will be rejected and won't update the database
- Files that contain skus that already exist in the database for the selected period type will be
  rejected and won't update the database
- You can select Export skus from the menu to check the skus that are currently
  in the database
- Make sure that the clusters database is up to date before uploading the sku files

'''
OUTLETS_IMPORT_MESSAGE = f'''{"---------- IMPORTING AN OUTLETS FILE ----------".center(100)}

- Input must be an excel file with the following format:
  1) Must have a single sheet
  2) Columns must be the below:
     - IDOutlet (Positive integer, non-missing)
     - PeriodType (String, non-missing, case sensitive, unique value)
       - Allowed values: Mountly, or Food, or Non_Food
     - PeriodName (String, non-missing, case sensitive, unique value)
     - IDProduct (Positive integer, non-missing)
     - IDBrand (Positive integer, non-missing)
     - IDGoods (Positive integer, non-missing)
     - LMPurch (Numeric, missing allowed)
     - Purch (Numeric, missing allowed)
     Please note the spaces, if any, and the case of the characters (lowercase, uppercase)
  3) Columns must be on the first line, data from second to the end and no merged cells are allowed
- Files that don't have the above format will be rejected and won't update the database
- Make sure that the clusters database is up to date before uploading the outlet file
'''
WARNING_OUTPUT_FORMAT = 'black on orange'
ERROR_OUTPUT_FORMAT = 'black on red'
SUCCESS_OUTPUT_FORMAT = 'white on green'
INFO_OUTPUT_FORMAT = 'white on blue'

VB_EXCEL_TO_CSV = '''
if WScript.Arguments.Count < 3 Then
    WScript.Echo "Please specify the source and the destination files. Usage: ExcelToCsv <xls/xlsx source file> <csv destination file> <worksheet number (starts at 1)>"
    Wscript.Quit
End If
csv_format = 6
Set objFSO = CreateObject("Scripting.FileSystemObject")
src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
worksheet_number = CInt(WScript.Arguments.Item(2))
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)
oBook.Worksheets(worksheet_number).Activate
oBook.SaveAs dest_file, csv_format
oBook.Close False
oExcel.Quit
'''

CREATE_CLUSTERS_SQL = '''
CREATE TABLE IF NOT EXISTS "clusters" (
	"id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
	"id_outlet" integer unsigned NOT NULL UNIQUE CHECK ("id_outlet" >= 0),
	"mountly" integer unsigned NULL CHECK ("mountly" >= 0),
	"food" integer unsigned NULL CHECK ("food" >= 0),
	"non_food" integer unsigned NULL CHECK ("non_food" >= 0)
)
'''

CREATE_SKUS_SQL = '''
CREATE TABLE IF NOT EXISTS "skus" (
	"id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
    "sku_id" varchar(100) NOT NULL,
    "period_type" integer unsigned NULL CHECK ("period_type" >= 0),
	"sku_name" varchar(500) NOT NULL,
	"sku_file_name" varchar(500) NOT NULL,
	"imported_date" date NOT NULL
)
'''

CREATE_ANALYSIS_SQL = '''
CREATE TABLE IF NOT EXISTS "analysis" (
	"id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
	"cluster" integer unsigned NULL CHECK ("cluster" >= 0),
	"count" integer unsigned NOT NULL CHECK ("count" >= 0),
	"min_diff" real NULL,
	"max_diff" real NULL,
	"mean_diff" real NULL,
	"std_diff" real NULL,
	"cv_diff" real NULL,
	"perc90_diff" real NULL,
	"perc95_diff" real NULL,
	"perc99_diff" real NULL,
	"sku_id" integer NULL REFERENCES "skus" ("id") DEFERRABLE INITIALLY DEFERRED
)
'''
CREATE_ATYPICALS_SQL = '''
CREATE TABLE IF NOT EXISTS "atypicals" (
	"id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
    "id_outlet" integer unsigned NULL CHECK ("id_outlet" >= 0),
	"lm_purch" real NULL,
	"purch" real NULL,
	"stars" varchar(100) NULL,
	"proposed_purch_1" real NULL,
	"proposed_purch_2" real NULL,
	"analysis_id" integer NULL REFERENCES "analysis" ("id") DEFERRABLE INITIALLY DEFERRED
)
'''

CREATE_MISSING_SQL = '''
CREATE TABLE IF NOT EXISTS "missing" (
	"id" integer NOT NULL PRIMARY KEY AUTOINCREMENT,
    "sku_id" varchar(100) NOT NULL,
	"period_type" integer unsigned NULL CHECK ("period_type" >= 0),
	"cluster" integer unsigned NULL CHECK ("cluster" >= 0),
	"id_outlet" integer unsigned NULL CHECK ("id_outlet" >= 0),
	"lm_purch" real NULL,
	"purch" real NULL
)
'''
DOCS = '''
Documentation
'''
#######################################################################################################
#######################################################################################################

def mycprint(w):
    '''A closure that carries the window object to use from printing to the app console'''
    def cprint(*args, **kwargs):
        message = args[0]
        ML_KEY = '-ML-'
        if kwargs.get('u'):
            w[ML_KEY].update('')
            if not message.strip():
                return cprint
        cp = sg.cprint
        # If l!= False, print a line ABOVE the output message
        if kwargs.get('l'):
            cp(100*'=')
        if kwargs.get('cons'):
            print(message)
        kwargs = {k:v for k,v in kwargs.items() if k not in ['w','u','cons','l']}
        cp(*args, **kwargs)
    return cprint


def timer(start, end):
    '''Input start and end in seconds.
       Returns a string of the form ..h:..m:..s
    '''
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    if not (hours or minutes):
        txt, vals = "{:05.2f}sec", (seconds,)
    elif not hours:
        txt, vals = "{}min:{:05.2f}sec", (int(minutes),seconds)
    else:
        txt, vals = "{}hour:{}min:{:05.2f}sec", (int(hours),int(minutes),seconds)
    return txt.format(*vals)


def join_columns(df, col_list, sep='-'):
    new_col = df[col_list].apply(lambda r:sep.join([str(x) for x in r]), axis='columns')
    return new_col


def diffs_old(r):
    output = []
    for i,v in enumerate(r):
        if i < len(r) - 1:
            if r[i] == 0 or r[i+1] == 0:
                item = np.nan
            else:
                item = r[i] * log(r[i] / r[i+1]) + r[i+1] - r[i]
            output.append(item)
    return output

def diffs(df):
    diff_columns = ['Diff_{}'.format(i+1) for i in range(len(df.columns)-1)]
    values = []
    for _, row in df.iterrows():
        values.append(diffs_old(row.values))
    df_result = pd.DataFrame({col:lst_v for col, lst_v in zip(diff_columns, list(zip(*values)))}, index=df.index)
    return df_result


def newton(f, Df, x0, epsilon=1e-6, max_iter=100):
    # https://www.math.ubc.ca/~pwalls/math-python/roots-optimization/newton/
    '''Approximate solution of f(x)=0 by Newton's method.
    Parameters
    ----------
    f : function
        Function for which we are searching for a solution f(x)=0.
    Df : function
        Derivative of f(x).
    x0 : number
        Initial guess for a solution f(x)=0.
    epsilon : number
        Stopping criteria is abs(f(x)) < epsilon.
    max_iter : integer
        Maximum number of iterations of Newton's method.
    Returns
    -------
    xn : number
        Implement Newton's method: compute the linear approximation
        of f(x) at xn and find x intercept by the formula
            x = xn - f(xn)/Df(xn)
        Continue until abs(f(xn)) < epsilon and return xn.
        If Df(xn) == 0, return None. If the number of iterations
        exceeds max_iter, then return None.
    Examples
    --------
    >>> f = lambda x: x**2 - x - 1
    >>> Df = lambda x: 2*x - 1
    >>> newton(f,Df,1,1e-8,10)
    Found solution after 5 iterations.
    1.618033988749989
    '''
    xn = x0
    try:
        for n in range(0,max_iter):
            fxn = f(xn)
            if abs(fxn) < epsilon:
                # print('Found solution after',n,'iterations.')
                return xn
            Dfxn = Df(xn)
            if Dfxn == 0:
                # print('Zero derivative. No solution found.')
                return None
            xn = xn - fxn/Dfxn
        # print('Exceeded maximum iterations. No solution found.')
        return None
    except:
        return None


def progress_bar(key, iterable, *args, title='', **kwargs):
    """
    Takes your iterable and adds a progress meter onto it
    :param key: Progress Meter key
    :param iterable: your iterable
    :param args: To be shown in one line progress meter
    :param title: Title shown in meter window
    :param kwargs: Other arguments to pass to one_line_progress_meter
    :return:
    """
    sg.set_options(element_padding=((5, 5),(5, 5)))
    sg.one_line_progress_meter(title, 0, len(iterable), key, *args, **kwargs)
    for i, val in enumerate(iterable):
        yield val
        if not sg.one_line_progress_meter(title, i+1, len(iterable), key, *args, **kwargs):
            break

def toggle_menu(menu_def, window, enable=True):
    for item in menu_def:
        if item[0].startswith('!') and enable:
            item[0] = item[0][1:]
        if not item[0].startswith('!') and not enable:
            item[0] = '!' + item[0]
    window['-MN-'].Update(menu_def)
    return menu_def

def popup_yes_no(title='', message=''):
    '''A confirmation popup
    '''
    message += '\n\nConfirm?\n'
    sg.set_options(auto_size_buttons=False)
    layout = [
        [sg.Text(message, auto_size_text=True, justification='center')],
        [sg.Button('Yes'), sg.Button('No')]
    ]
    window = sg.Window(title, layout, finalize=True, modal=True, element_justification='c', size=(300,150))
    event, values = window.read()
    window.close()
    return event == 'Yes'
