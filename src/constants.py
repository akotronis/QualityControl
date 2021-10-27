APP_INFO_MESSAGE = f'''{"---------- About SKU Quality Control Application ----------".center(100)}

The purpose of SKU Analysis application is to investigate the atypical purchases that are recorded
in one audit period.
The process of finding atypical purchases reaches the SKU level per store of the Emrc Retail Audit sample.

{" Please read the docs for further instructions ".center(100)}
'''


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
HTML_DOCS = '''
Documentation
'''
#######################################################################################################
#######################################################################################################
