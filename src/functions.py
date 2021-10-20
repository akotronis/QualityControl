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

- Input must be an excel file with the following format:
  1) Must have a single sheet
  2) Columns must be the below:
     - IDOutlet (Positive integer, non-missing)
     - PeriodType (String, non-missing, case sensitive, unique value)
       - Allowed values: Mountly, or Food, or Non_Food, according to what type you selected
     - OutletType (String)
     - Area (String)
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
- You can select Export skus from the select action page to check the skus that are currently
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
        # If l!= False, print a line ABOVR the output message
        if kwargs.get('l') != False:
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
        txt, vals = "{:05.2f}s", (seconds,)
    elif not hours:
        txt, vals = "{}m:{:05.2f}s", (int(minutes),seconds)
    else:
        txt, vals = "{}h:{}m:{:05.2f}s", (int(hours),int(minutes),seconds)
    return txt.format(*vals)


def popup_yes_no(title='', message=''):
    '''A confirmation popup
    '''
    message += '\n\nConfirm?\n'
    sg.set_options(auto_size_buttons=False)
    layout = [
        [sg.Text(message, auto_size_text=True, justification='center')],
        [sg.Button('Yes'), sg.Button('No')]
    ]
    window = sg.Window(title, layout, finalize=True, modal=True, element_justification='c')
    event, values = window.read()
    window.close()
    return event == 'Yes'



# def popup(w, messages=None):
#     w.disappear()
#     if isinstance(messages, str):
#         messages = [messages]
#     sg.popup(*messages, grab_anywhere=True, title='Info',)
#     w.reappear()