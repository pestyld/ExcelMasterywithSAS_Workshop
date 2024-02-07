/**********************************************************************************/
/* WORKSHEET 2 - Employee List                                                    */
/**********************************************************************************/
/* REQUIREMENT: The 01_excel_setup program needs to be run first                  */
/**********************************************************************************/


/**********************/
/* OUTPUT TO EXCEL    */
/**********************/
ods excel file="&outpath/worksheet_02.xlsx";
          

ods excel options(
			  sheet_name = 'Employee List'           /*<--- specifies the name for the next worksheet */
			  sheet_interval = 'NOW'                 /*<--- create new worksheet  */
              autofilter = 'ALL'                     /*<--- turns on filtering for specified columns in the worksheet */
			  flow = "TABLES"                        /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
              absolute_row_height = '30'             /*<--- specifies the row heights */
              absolute_column_width = '30,10,35,15,35,20,20,20,20,20,20'   /*<--- specifies the column widths */
			  frozen_headers = '3'                   /*<--- specifies that headers can scroll or not scroll with the scroll bar */
              frozen_rowheaders = '2'                /*<--- specifies if the row headers are on the left scroll when the table data scrolls */
			  embedded_titles = 'ON'                 /*<--- specifies whether titles should appear in the worksheet */
	  	   );	



title &titleFmt 'List of Employee Information';
proc print data = work.emp_info_all noobs label;
	id NAME EMPNO;
	var DIVISION JOBCODE TITLE SALARY HDATE EMPYOS STATUS GENDER EDLEV LOCATION PHONE ROOM;
run;
title;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;