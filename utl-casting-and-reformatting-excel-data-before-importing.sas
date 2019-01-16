Casting and reformatting excel data before importing

  EXAMPLES (Using passthru SQL to cleanup excel sheet - tip of the iceberg)

     1. Is the column numeric or character and how many values are numeric or character
     2. Convert numerics to character
     3. Convert character ro numeric
     4. Extract hour from datetime, year and month from date
     5. Substring character string (good to split into 1024 segments)
     6. Set flag when height > 65


github
https://tinyurl.com/y9c7cjsa
https://github.com/rogerjdeangelis/-utl-casting-and-reformatting-excel-data-before-importing

SAS Forum
https://tinyurl.com/ybky4wtw
https://communities.sas.com/t5/SAS-Programming/how-to-import-a-excel-with-format-into-sas/m-p/527721

Other excel rpos

https://tinyurl.com/ybnm6azh
https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=

macros
https://tinyurl.com/y9nfugth
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories


INPUT
=====

* Make Excel Sheet


%utlfkil(d:/xls/class.xlsx);
libname xel "d:/xls/class.xlsx";
data class ;
   format dateSAS mmddyy10. datetimeSAS datetime15.0 age $2.;
   set sashelp.class(obs=7 rename=(age=agen weight=weightn));
   age=put(agen,2.);
   weight=put(weightn,5.1);
   if mod(_n_,3)=0 then age="AA";
   dateSAS=today();
   datetimeSAS=datetime();
   drop agen weightn;
run;quit;
libname xel clear;


 WORKBOOK d:/xls/class.xlsx

                                                                  MIXED        NUMERIC     CHARACTER
   CHARACTER    EXCEL DATE    EXCEL DATETIME        CHARCATER      TYPE                      TYPE
  +----------------------------------------------------------------------------------------------------+
  |     A      |    B       |         C            |    D       |     E      |    F       |    G       |
  +------------=------------=----------------------=------------=------------=------------=------------+
1 | NAME       | DATESAS    |    DATETIMESAS       |   SEX      |    AGE     |  HEiGHT    |  WEIGHT    |
  +------------+------------+----------------------+------------+------------+------------+------------+
2 | ALFRED     | 01/16/2019 |    16JAN19:14:52     |    M       |    14      |    69      |  112.5     |
  +------------+------------+----------------------+------------+------------+------------+------------+
   ...
  +------------+------------+----------------------+------------+------------+------------+------------+
N | WILLIAM    |    M       |    15        15      |    M       |    AA      |   66.5     |  112       |
  +------------+------------+----------------------+-----------+-------------+------------+------------+

[CLASS]


SOLUTIONS
=========

----------------------------------------------------------------------------------
1. Is the column numeric or character and how many values are numeric or character
----------------------------------------------------------------------------------

proc sql dquote=ansi;
   connect to excel (Path="d:\xls\class.xlsx");
     select * from connection to Excel
         (
          Select
               count(*) as numRows
              ,count(*) + sum(isnumeric(age)) as age_character
              ,-1*sum(isnumeric(age)) as age_numeric
          from
               [class]
         );
     disconnect from Excel;
quit;

Two rows aare character and 5 are numeric

                     age_
    numRows     character  age_numeric
    -------     ---------  -----------
          7             2            5


--------------------------------
2. Convert numerics to character
--------------------------------

proc sql dquote=ansi;
   connect to excel (Path="d:\xls\class.xlsx");
     create table classCast as
     select * from connection to Excel
         (
          Select
               name
              ,format(height,'###.0') as height
          from
               [class]
         );
     disconnect from Excel;
quit;

     Variables in Creation Order

#    Variable    Type     Len    Format

1    NAME        Char     255    $255.
2    HEIGHT      Char    1024    $1024.
3    WEIGHT      Char    1024    $1024.

WORK.CLASSCAST total obs=7

Obs    NAME       HEIGHT    WEIGHT

 1     Alfred      69.0     112.5
 2     Alice       56.5     84.0

 You should run utl_optlen macro


%utl_optlen(inp=classcast,out=classcat);

  Variables in Creation Order

 Variable    Type    Len

 NAME        Char      7
 HEIGHT      Char      4
 WEIGHT      Char      5


--------------------------------
3. Convert character to numeric
--------------------------------

proc sql dquote=ansi;
   connect to excel (Path="d:\xls\class.xlsx");
     create table classCast as
     select * from connection to Excel
         (
          Select
               name
               ,CDbl(weight) as weight
          from
               [class]
         );
     disconnect from Excel;
quit;

  Variables in Creation Order

#    Variable    Type    Len

1    NAME        Char    255
2    WEIGHT      Num       8   ** converted from character to numeric;


--------------------------------------------------------
4. Extract hour from datetime, year and month from date
--------------------------------------------------------

proc sql dquote=ansi;
   connect to excel (Path="d:\xls\class.xlsx");
     create table classDates as
     select * from connection to Excel
         (
          Select
               name
               ,month(datesas) as monthx
               ,day(datesas) as dayx
               ,year(datesas) as yearx
               ,hour(datetimeSAS) as hourx
          from
               [class]
         );
     disconnect from Excel;
quit;

Up to 40 obs WORK.CLASSDATES total obs=7

Obs    NAME       MONTHX    DAYX    YEARX    HOURX

 1     Alfred        1       16      2019      14
 2     Alice         1       16      2019      14
 3     Barbara       1       16      2019      14
 4     Carol         1       16      2019      14
 5     Henry         1       16      2019      14
 6     James         1       16      2019      14
 7     Jane          1       16      2019      14


----------------------------------------------------------------------
5. Substring EXCEL character string (good to split into 1024 segments)
----------------------------------------------------------------------

proc sql dquote=ansi;
   connect to excel (Path="d:\xls\class.xlsx");
     create table classFlg as
     select * from connection to Excel
         (
          Select
               height
               ,IIF(Height>65, 1, 0) as flg
          from
               [class]
         );
     disconnect from Excel;
quit;


WORK.CLASSFLG total obs=7

Obs    HEIGHT    FLG

 1      69.0      1
 2      56.5      0
 3      65.3      1
 4      62.8      0
 5      63.5      0
 6      57.3      0
 7      59.8      0


 _ __ ___  ___   ___  __ _| |
| '_ ` _ \/ __| / __|/ _` | |
| | | | | \__ \ \__ \ (_| | |
|_| |_| |_|___/ |___/\__, |_|
                        |_|
;


https://ss64.com/access/

a
  Abs             The absolute value of a number (nore negative sn).
 .AddMenu         Add a custom menu bar/shortcut bar.
 .AddNew          Add a new record to a recordset.
 .ApplyFilter     Apply a filter clause to a table, form, or report.
  Array           Create an Array.
  Asc             The Ascii code of a character.
  AscW            The Unicode of a character.
  Atn             Display the ArcTan of an angle.
  Avg (SQL)       Average.
b
 .Beep (DoCmd)    Sound a tone.
 .BrowseTo(DoCmd) Navate between objects.
c
  Call            Call a procedure.
 .CancelEvent (DoCmd) Cancel an event.
 .CancelUpdate    Cancel recordset changes.
  Case            If Then Else.
  CBool           Convert to boolean.
  CByte           Convert to byte.
  CCur            Convert to currency (number)
  CDate           Convert to Date.
  CVDate          Convert to Date.
  CDbl            Convert to Double (number)
  CDec            Convert to Decimal (number)
  Choose          Return a value from a list based on position.
  ChDir           Change the current directory or folder.
  ChDrive         Change the current drive.
  Chr             Return a character based on an ASCII code.
 .ClearMacroError (DoCmd) Clear MacroError.
 .Close (DoCmd)           Close a form/report/window.
 .CloseDatabase (DoCmd)   Close the database.
  CInt                    Convert to Integer (number)
  CLng                    Convert to Long (number)
  Command                 Return command line option string.
 .CopyDatabaseFile(DoCmd) Copy to an SQL .mdf file.
 .CopyObject (DoCmd)      Copy an Access database object.
  Cos                     Display Cosine of an angle.
  Count (SQL)             Count records.
  CSng             Convert to Single (number.)
  CStr             Convert to String.
  CurDir           Return the current path.
  CurrentDb        Return an object variable for the current database.
  CurrentUser      Return the current user.
  CVar             Convert to a Variant.
d
  Date             The current date.
  DateAdd          Add a time interval to a date.
  DateDiff         The time difference between two dates.
  DatePart         Return part of a given date.
  DateSerial       Return a date given a year, month, and day.
  DateValue        Convert a string to a date.
  DAvg             Average from a set of records.
  Day              Return the day of the month.
  DCount           Count the number of records in a table/query.
  Delete (SQL)          Delete records.
 .DeleteObject (DoCmd)  Delete an object.
  DeleteSetting         Delete a value from the users registry
 .DoMenuItem (DoCmd)    Display a menu or toolbar command.
  DFirst           The first value from a set of records.
  Dir              List the files in a folder.
  DLast            The last value from a set of records.
  DLookup          Get the value of a particular field.
  DMax             Return the maximum value from a set of records.
  DMin             Return the minimum value from a set of records.
  DoEvents         Allow the operating system to process other events.
  DStDev           Estimate Standard deviation for domain (subset of records)
  DStDevP          Estimate Standard deviation for population (subset of records)
  DSum             Return the sum of values from a set of records.
  DVar             Estimate variance for domain (subset of records)
  DVarP            Estimate variance for population (subset of records)
e
 .Echo             Turn screen updating on or off.
  Environ          Return the value of an OS environment variable.
  EOF              End of file input.
  Error            Return the error message for an error No.
  Eval             Evaluate an expression.
  Execute(SQL/VBA) Execute a procedure or run SQL.
  Exp              Exponential e raised to the nth power.
f
  FileDateTime      Filename last modified date/time.
  FileLen           The size of a file in bytes.
 .FindFirst/Last/Next/Previous Record.
 .FindRecord(DoCmd) Find a specific record.
  First (SQL)       Return the first value from a query.
  Fix               Return the integer portion of a number.
  For               Loop.
  Format            Format a Number/Date/Time.
  FreeFile          The next file No. available to open.
  From              Specify the table(s) to be used in an .
  FV                Future Value of an annuity.
g
  GetAllSettings    List the settings saved in the registry.
  GetAttr           Get file/folder attributes.
  GetObject         Return a reference to an ActiveX object
  GetSetting        Retrieve a value from the users registry.
  form.GoToPage     Move to a page on specific form.
 .GoToRecord (DoCmd)Move to a specific record in a dataset.
h
  Hex               Convert a number to Hex.
  Hour              Return the hour of the day.
 .Hourglass (DoCmd) Display the hourglass icon.
  HyperlinkPart     Return information about data stored as a hyperlink.
i
  If Then Else      If-Then-Else
  IIf               If-Then-Else function.
  Input             Return characters from a file.
  InputBox          Prompt for user input.
  Insert (SQL)      Add records to a table (append query).
  InStr             Return the position of one string within another.
  InstrRev          Return the position of one string within another.
  Int               Return the integer portion of a number.
  IPmt              Interest payment for an annuity
  IsArray           Test if an expression is an array
  IsDate            Test if an expression is a date.
  IsEmpty           Test if an expression is Empty (unassned).
  IsError           Test if an expression is returning an error.
  IsMissing         Test if a missing expression.
  IsNull            Test for a NULL expression or Zero Length string.
  IsNumeric         Test for a valid Number.
  IsObject          Test if an expression is an Object.
L
  Last (SQL)        Return the last value from a query.
  LBound            Return the smallest subscript from an array.
  LCase             Convert a string to lower-case.
  Left              Extract a substring from a string.
  Len               Return the length of a string.
  LoadPicture       Load a picture into an ActiveX control.
  Loc               The current position within an open file.
 .LockNavationPane(DoCmd) Lock the Navation Pane.
  LOF               The length of a file opened with Open()
  Log               Return the natural logarithm of a number.
  LTrim             Remove leading spaces from a string.
m
  Max (SQL)         Return the maximum value from a query.
 .Maximize (DoCmd)  Enlarge the active window.
  Mid               Extract a substring from a string.
  Min (SQL)         Return the minimum value from a query.
 .Minimize (DoCmd)  Minimise a window.
  Minute            Return the minute of the hour.
  MkDir             Create directory.
  Month             Return the month for a given date.
  MonthName         Return  a string representing the month.
 .Move              Move through a Recordset.
 .MoveFirst/Last/Next/Previous Record
 .MoveSize (DoCmd)  Move or Resize a Window.
  MsgBox            Display a message in a dialogue box.
n
  Next              Continue a for loop.
  Now               Return the current date and time.
  Nz                Detect a NULL value or a Zero Length string.
o
  Oct               Convert an integer to Octal.
  OnClick, OnOpen   Events.
 .OpenForm (DoCmd)  Open a form.
 .OpenQuery (DoCmd) Open a .
 .OpenRecordset         Create a new Recordset.
 .OpenReport (DoCmd)    Open a report.
 .OutputTo (DoCmd)      Export to a Text/CSV/Spreadsheet file.
p
  Partition (SQL)       Locate a number within a range.
 .PrintOut (DoCmd)      Print the active object (form/report etc.)
q
  Quit                  Quit Microsoft Access
r
 .RefreshRecord (DoCmd) Refresh the data in a form.
 .Rename (DoCmd)        Rename an object.
 .RepaintObject (DoCmd) Complete any pending screen updates.
  Replace               Replace a sequence of characters in a string.
 .Re               Re the data in a form or a control.
 .Restore (DoCmd)       Restore a maximized or minimized window.
  RGB                   Convert an RGB color to a number.
  Rht                 Extract a substring from a string.
  Rnd                   Generate a random number.
  Round                 Round a number to n decimal places.
  RTrim                 Remove trailing spaces from a string.
 .RunCommand            Run an Access menu or toolbar command.
 .RunDataMacro (DoCmd)  Run a named data macro.
 .RunMacro (DoCmd)      Run a macro.
 .RunSavedImportExport (DoCmd) Run a saved import or export specification.
 .RunSQL (DoCmd)        Run an SQL .
s
 .Save (DoCmd)          Save a database object.
  SaveSetting           Store a value in the users registry
 .SearchForRecord(DoCmd) Search for a specific record.
  Second                Return the seconds of the minute.
  Seek                  The position within a file opened with Open.
  Select (SQL)          Retrieve data from one or more tables or queries.
  Select Into (SQL)     Make-table .
  Select-Sub (SQL) Sub.
 .SelectObject (DoCmd)  Select a specific database object.
 .SendObject (DoCmd)    Send an email with a database object attached.
  SendKeys              Send keystrokes to the active window.
  SetAttr               Set the attributes of a file.
 .SetDisplayedCategories (DoCmd)  Change Navation Pane display options.
 .SetFilter (DoCmd)     Apply a filter to the records being displayed.
  SetFocus              Move focus to a specified field or control.
 .SetMenuItem (DoCmd)   Set the state of menubar items (enabled /checked)
 .SetOrderBy (DoCmd)    Apply a sort to the active datasheet, form or report.
 .SetParameter (DoCmd)  Set a parameter before opening a Form or Report.
 .SetWarnings (DoCmd)   Turn system messages on or off.
  Sgn                   Return the sn of a number.
 .ShowAllRecords(DoCmd) Remove any applied filter.
 .ShowToolbar (DoCmd)   Display or hide a custom toolbar.
  Shell                 Run an executable program.
  Sin                   Display Sine of an angle.
  SLN                   Straht Line Depreciation.
  Space                 Return a number of spaces.
  Sqr                   Return the square root of a number.
  StDev (SQL)           Estimate the standard deviation for a population.
  Str                   Return a string representation of a number.
  StrComp               Compare two strings.
  StrConv               Convert a string to Upper/lower case or Unicode.
  String                Repeat a character n times.
  Sum (SQL)             Add up the values in a  result set.
  Switch                Return one of several values.
  SysCmd                Display a progress meter.
t
  Top 1 *               Get first rpw
  Tan                   Display Tangent of an angle.
  Time                  Return the current system time.
  Timer                 Return a number (single) of seconds since midnht.
  TimeSerial            Return a time given an hour, minute, and second.
  TimeValue             Convert a string to a Time.
 .TransferDatabase (DoCmd)      Import or export data to/from another database.
 .TransferSharePointList(DoCmd) Import or link data from a SharePoint Foundation site.
 .TransferSpreadsheet (DoCmd)   Import or export data to/from a spreadsheet file.
 .TransferSQLDatabase (DoCmd)   Copy an entire SQL Server database.
 .TransferText (DoCmd)          Import or export data to/from a text file.
  Transform (SQL)       Create a crosstab .
  Trim                  Remove leading and trailing spaces from a string.
  TypeName              Return the data type of a variable.
u
  UBound                Return the largest subscript from an array.
  UCase                 Convert a string to upper-case.
  Undo                  Undo the last data edit.
  Union (SQL)           Combine the results of two SQL queries.
  Update (SQL)          Update existing field values in a table.
 .Update                Save a recordset.
v
  Val                   Extract a numeric value from a string.
  Var (SQL)             Estimate variance for sample (all records)
  VarP (SQL)            Estimate variance for population (all records)
  VarType               Return a number indicating the data type of a variable.
w
  Weekday               Return the weekday (1-7) from a date.
  WeekdayName           Return the day of the week.
y
  Year                  Return the year for a given date.




