/******************************************/
/* CREATE WORKSHEET 4 - Salary by Division*/
/******************************************/


/*******************************************************/
/* EXCEL OUTPUT                                        */
/*******************************************************/
ods excel file="&path/excel_output/04_dev_excel.xlsx";


/***********************************/
/* WORKSHEET 1 - TABLE OF CONTENTS */
/***********************************/
ods excel options(sheet_name='Overview');   /*<--- specifies the name for the next worksheet */

proc odstext;
	h1 'Viewing Employee Data';
	p 'Description: This report analyzes abc';
	list;
		item 'Employee List: The employee list worksheet does xyz';
		item 'XYZ: Describes xyz';
	end;
run;


/*******************************/
/* WORKSHEET 2 - EMPLOYEE LIST */
/*******************************/
ods excel options(
			  sheet_name='Employee List'           /*<--- specifies the name for the next worksheet */
              autofilter='ALL'                     /*<--- turns on filtering for specified columns in the worksheet */
			  flow="TABLES"                        /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
              absolute_row_height='30'             /*<--- specifies the row heights */
              absolute_column_width='30,20,10,35,0,35,0'   /*<--- specifies the column widths */
			  frozen_headers='on'                  /*<--- specifies that headers can scroll or not scroll with the scroll bar */
              frozen_rowheaders='3'                /*<--- sspecifies if the row headers are on the left scroll when the table data scrolls */
	  	   );	

proc print data=work.emp_info_all noobs label;
	ID NAME IDNUM EMPNO;
	var DIVISION JOBCODE TITLE SALARY HDATE EMPYOS STATUS GENDER EDLEV LOCATION PHONE ROOM;
run;


/************************************/
/* WORKSHEET 3 - SALARY BY DIVISION */
/************************************/
ods excel options(
			  sheet_name='Salary by Division'     /*<--- specifies the name for the next worksheet */
              sheet_interval='NONE'
              suppress_bylines='ON'
              absolute_column_width='0'          /*<--- specifies the column widths */
	          autofilter='ON'                    /*<--- turns on filtering for specified columns in the worksheet */
			  flow="TABLES"                      /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
			  frozen_headers='OFF'               /*<--- specifies that headers can scroll or not scroll with the scroll bar */
              frozen_rowheaders='OFF'            /*<--- specifies if the row headers are on the left scroll when the table data scrolls */
	  	   );	

ods graphics / width=8in height=6in;
title &titleFmt "Total Salary by Division";
proc sgplot data=work.emp_info_all 
		    noborder;
	hbar DIVISION / 
		response=SALARY  
        categoryorder=respdesc
		nooutline
		datalabel
	;
	xaxis display=none;
	yaxis display=(NOLABEL NOTICKS);
	yaxistable Salary / stat=freq label='Total Employees' labeljustify=center valuejustify=center;
	yaxistable Salary / stat=mean label='Mean Salary' labeljustify=center valuejustify=center;
run;
ods graphics / reset;


proc report data=work.emp_info_all;
	/* Group by DIVISION */
	by Division;

	/* Columns to include in the report */
	columns("DIVISION SUMMARY" DIVISION NAME TITLE SALARY);

	/* Specify what each column does in the report */
	define DIVISION / group order=internal;
	define SALARY / analysis sum order=internal;

	/* Break after each group summarization */
	break after DIVISION / summarize suppress style={background=lightgray fontsize=10pt fontweight=bold};
run;

ods excel close;