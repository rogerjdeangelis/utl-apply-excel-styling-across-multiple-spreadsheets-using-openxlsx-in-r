# utl-apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r
Apply Excel Styling Across Multiple Spreadsheets using openxlsx in R 
    %let pgm=utl-apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r;

    Apply Excel Styling Across Multiple Spreadsheets using openxlsx in R


     PROBLEM

        Create 3 sheets with  solution by https://stackoverflow.com/users/10466439/jaskeil

             1. Header (background blue foreground letters white
                blue_white <- createStyle(fgFill = "#03a9f4", fontColour = "white");

                Declare these style we want to use for columns
                pct        <- createStyle(numFmt = "PERCENTAGE");
                currency   <- createStyle(numFmt = "CURRENCY");
                round2     <- createStyle(numFmt = "0.000000");

             2. Add style
                addStyle(wb, .x, currency, cols = 3, rows = 2:nrow(.y))

             3, Create sheetnames
                sheets <- seq_along(dat_grouped);

                [2023-09-29]
                [2023-09-29]
                [2023-09-29]

             Note
                 1. Existing formats can be queried.
                 2, Another method is to create an empty excel template and fill the data in later.

    github
    https://tinyurl.com/yj5a55bp
    https://github.com/rogerjdeangelis/utl-apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r

    Stackoverflow
    https://tinyurl.com/2p8mczs9
    https://stackoverflow.com/questions/77203075/apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r

    Related repo
    https://tinyurl.com/58e2r94h
    https://github.com/rogerjdeangelis/utl-preserving-excel-formatting-when-writing-to-an-existing-worksheet
    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    data sd1.have ;
    retain DATE LOGICAL CURRENCY ACCOUNTING HLINK PERCENTAGE TINYNUMBER;
    format DATE date10.
    TINYNUMBER E15.10;
    informat
    DATE date10.
    LOGICAL 8.
    CURRENCY 8.
    ACCOUNTING 8.
    HLINK $27.
    PERCENTAGE 8.
    TINYNUMBER E15.10
    ;input
    DATE LOGICAL CURRENCY ACCOUNTING HLINK PERCENTAGE TINYNUMBER;
    cards4;
    29SEP2023 1 -2 -2 https://CRAN.R-project.org/ -1 5.151492E-10
    28SEP2023 0 -1 -1 https://CRAN.R-project.org/ -0.5 1.020065E-10
    27SEP2023 1 0 0 https://CRAN.R-project.org/ 0 5.234703E-10
    ;;;;
    run;quit;



    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* INPUT                                                                                                                  */
    /*                                                                                                                        */
    /* SD1.HAVE total obs=5                                                                                                   */
    /*                                                                                                                        */
    /* Obs      DATE       LOGICAL    CURRENCY    ACCOUNTING               HLINK               PERCENTAGE    TINYNUMBER       */
    /*                                                                                                                        */
    /*  1     29SEP2023       1          -2           -2        https://CRAN.R-project.org/       -1.0       1.15149E-06      */
    /*  2     28SEP2023       0          -1           -1        https://CRAN.R-project.org/       -0.5       1.02007E-06      */
    /*  3     27SEP2023       1           0            0        https://CRAN.R-project.org/        0.0       1.23470E-06      */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /*  OUTPUT EXCEL  ( 3 sheets)  Workbook d:/xls/want.xlsx                                                                  */
    /*                                                                                                                        */
    /*  Header row has a blue background and white letters                                                                    */                                                   */
    /*                                                                                                                        */
    /*    [A1]                                                                                                                */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /*    |     A      |    B       |       C      |    D     |            E                |    F       |       G        |   */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /* 1  |   DATE     | LOGICAL    | CURRENCY   | ACCOUNTING |  HLINK                      | PERCENTAGE | TINYNUMBER     |   */
    /*    +------------+------------+------------+------------+----------------------------+------------+-----------------+   */
    /* 2  |09/27/2023  |    0       |   $2.00    |     2      | https://CRAN.R-project.org/ |  100.00%   | 0.000000000908 |   */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /*                                                                                                                        */
    /*    [2023-09-27] ==> sheet 1                                                                                            */
    /*                                                                                                                        */
    /*    [A1]                                                                                                                */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /*    |     A      |    B       |       C      |    D     |            E                |    F       |       G        |   */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /* 1  |   DATE     | LOGICAL    | CURRENCY   | ACCOUNTING |  HLINK                      | PERCENTAGE | TINYNUMBER     |   */
    /*    +------------+------------+------------+------------+----------------------------+------------+-----------------+   */
    /* 2  |09/28/2023  |    1       |  -$1.00    |    -1      | https://CRAN.R-project.org/ |  -50.00%   | 0.000000000052 |   */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /*                                                                                                                        */
    /*    [2023-09-28] ==> sheet 2                                                                                            */
    /*                                                                                                                        */
    /*    [A1]                                                                                                                */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /*    |     A      |    B       |       C      |    D     |            E                |    F       |       G        |   */
    /*    +----9----------------------------------------------------------------------------------------------------------+   */
    /* 1  |   DATE     | LOGICAL    | CURRENCY   | ACCOUNTING |  HLINK                      | PERCENTAGE | TINYNUMBER     |   */
    /*    +------------+-------- ---+------------+------------+----------------------------+------------+-----------------+   */
    /* 2  |09/29/2023  |    0       |   $2.00    |     2      | https://CRAN.R-project.org/ | -100.00%   | 0.000000000908 |   */
    /*    +---------------------------------------------------------------------------------------------------------------+   */
    /*                                                                                                                        */
    /*    [2023-09-29] ==> sheet 3                                                                                             */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    %utlfkil(d:/xls/want.xlsx);

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    proc r;
    export data=sd1.have r=have;
    submit;
    library(dplyr);
    library(openxlsx);
    library(purrr);

    dat_grouped <- have %>% split(~ DATE);

    blue_white <- createStyle(fgFill = "#03a9f4", fontColour = "white");
    pct <- createStyle(numFmt = "PERCENTAGE");
    currency <- createStyle(numFmt = "CURRENCY");
    round2 <- createStyle(numFmt = "0.000000000000");

    wb <- write.xlsx(dat_grouped, "d:/xls/want.xlsx");
    sheets <- seq_along(dat_grouped);

    walk2(sheets, dat_grouped, ~ addStyle(wb, .x, pct, cols = 6, rows = 2:nrow(.y)));
    walk2(sheets, dat_grouped, ~ addStyle(wb, .x, currency, cols = 3, rows = 2:nrow(.y)));
    walk2(sheets, dat_grouped, ~ addStyle(wb, .x, round2, cols = 7, rows = 2:nrow(.y)));
    walk(sheets, ~ addStyle(wb, .x, blue_white, row = 1, cols = seq_len(ncol(have))));
    walk(sheets, ~ setColWidths(wb, .x, cols = seq_len(ncol(have)), widths = "auto"));
    saveWorkbook(wb, "d:/xls/want.xlsx", overwrite = TRUE);
    endsubmit;
    run;quit;
    ');

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
