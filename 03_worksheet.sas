/**********************************************************************************/
/* CREATE WORKSHEET 1-3                                                           */
/**********************************************************************************/
/* REQUIREMENT: The 01_explore_and_prepare_data.sas program needs to be run first */
/**********************************************************************************/


/******************************************/
/* ATTRIBUTES SETUP                       */
/******************************************/
%let titleFmt = height=16pt justify=left;
%let sasBlue = ;
%let sasDarkBlue =;


/*******************************************************/
/* UPDATES TO EXCEL OUTPUT                             */
/*******************************************************/
ods excel file="&path/excel_output/02_dev_excel.xlsx";


/***************/
/* WORKSHEET 1 */
/***************/
ods excel options(sheet_name='Overview');   /*<--- specifies the name for the next worksheet */

proc odstext;
	h1 'Viewing Employee Data';
	p 'Description: This report analyzes abc';
	list;
		item 'Employee List: The employee list worksheet does xyz';
		item 'XYZ: Describes xyz';
	end;
run;


/***************/
/* WORKSHEET 2 */
/***************/
ods excel options(
			  sheet_name='Employee List'           /*<--- specifies the name for the next worksheet */
              autofilter='ALL'                     /*<--- turns on filtering for specified columns in the worksheet */
			  flow="TABLES"                        /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
              absolute_row_height='30'             /*<--- specifies the row heights */
              absolute_column_width='30,20,10,35,0,35,0'   /*<--- specifies the column widths */
			  frozen_headers='2'                   /*<--- specifies that headers can scroll or not scroll with the scroll bar */
              frozen_rowheaders='3'                /*<--- specifies if the row headers are on the left scroll when the table data scrolls */
			  embedded_titles='ON'                 /*<--- specifies whether titles should appear in the worksheet */
	  	   );	


title height=14pt justify=left 'List of Employee Information';
proc print data=work.emp_info_all noobs label;
	ID NAME IDNUM EMPNO;
	var DIVISION JOBCODE TITLE SALARY HDATE EMPYOS STATUS GENDER EDLEV LOCATION PHONE ROOM;
run;



/***************/
/* WORKSHEET 3 */
/***************/
proc print data=work.emp_info_all(obs=10);
run;

title &titleFmt "Total Head Count by Location";
proc sgpie data=work.emp_info_all;
	donut LOCATION / 
		dataskin = none
		holevalue holelabel='Headcount'
		datalabelloc=callout;
run;

title &titleFmt "Total Yearly Cost by Division";
proc sgpie data=work.emp_info_all;
	donut DIVISION /  
		response=SALARY
		dataskin = none
        maxslices=6
		holevalue holelabel='xxx'
		datalabelloc=callout;
run;


ods excel close;