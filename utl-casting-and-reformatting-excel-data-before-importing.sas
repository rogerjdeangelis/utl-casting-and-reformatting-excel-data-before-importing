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

