/**********************************************************************************/
/* WORKSHEET 4 - Division Analysis                                                */
/**********************************************************************************/
/* REQUIREMENT: The 01_excel_setup program needs to be run first                  */
/**********************************************************************************/


/**********************/
/* OUTPUT TO EXCEL    */
/**********************/
ods excel file="&outpath/worksheet_04.xlsx";

ods excel options(
			  sheet_name='Division Analysis'     /*<--- specifies the name for the next worksheet */
              sheet_interval='NONE'              /*<--- create a new worksheet */
              suppress_bylines='ON'              /*<--- remove by lines from PROC REPORT by groups */
              absolute_column_width='0'          /*<--- reset the column widths */
			  flow="TABLES"                      /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
			  frozen_headers='OFF'               /*<--- turn off frozen headers */
              frozen_rowheaders='OFF'            /*<--- turn off row headers */
	  	   );	



/**************************/
/* Salary by Division     */
/**************************/
ods graphics / width=8in;

title &titleFmt "Total Salary by Division";
proc sgplot data=work.emp_info_all 
		    noborder;
	hbar DIVISION / 
		response=SALARY
        categoryorder=respdesc
		nooutline
		datalabel 
		barwidth=.7
		fillattrs=(color=&sasBlue)
	;
	xaxis display=none;
	yaxis display=(NOLABEL NOTICKS);
	yaxistable DIVISION / stat=freq label='Total Employees' labeljustify=center valuejustify=center;
	yaxistable SALARY / stat=mean label='Mean Salary' labeljustify=center valuejustify=center;
run;

ods graphics / reset=width;


/**********************************/
/* List of Employees by Division  */
/**********************************/
proc odstext;
	p 'Detailed Employee Information for Each Division' / style=[color=&sasDarkBlue fontsize=&ws_title_text];
quit;


proc report data=work.emp_info_all;
	/* Group by DIVISION */
	by Division;
	
	/* Columns to include in the report */
	columns("DIVISION SUMMARY" DIVISION NAME TITLE SALARY EMP_STATUS);

	/* Specify what each column does in the report */
	define DIVISION / group order;
	define SALARY / analysis sum;

	/* Break after each group (DIVISION) summarization */
	break after DIVISION / summarize 
	                       suppress 
	                       style={background=lightgray fontsize=10pt fontweight=bold};
run;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;