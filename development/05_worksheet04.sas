/*********************************************************************************
 WORKSHEET 4 - Division Analysis                                              
**********************************************************************************
 REQUIREMENT: The 00_config and 01_prepare_data programs need to be run.                  
*********************************************************************************/


/**********************/
/* OUTPUT TO EXCEL    */
/**********************/
ods excel file = "&dev_outpath/worksheet_04.xlsx";

ods excel options(
			  sheet_name = 'Division Analysis'       /*<--- specifies the name for the next worksheet */
              sheet_interval = 'NONE'                /*<--- create a new worksheet */
              suppress_bylines = 'ON'                /*<--- remove by lines from PROC REPORT by groups */
              absolute_column_width = '0'            /*<--- reset the column widths */
              row_heights = '30,20,20,20,20,20,20'   /*<--- specifies the height of the row using positional parameters */
			  flow = "TABLES"                        /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
			  frozen_headers = 'OFF'                 /*<--- turn off frozen headers */
              frozen_rowheaders = 'OFF'              /*<--- turn off row headers */
	  	   );	
		   
/* Use PNG images */
ods graphics / imagefmt=png;


/**************************/
/* WORKSHEET TITLE        */
/**************************/
%worksheet_title('Division Analysis');


/******************************/
/* Salary by Division Visual  */
/******************************/
ods graphics / width=8in;

title &titleFmt "Total Salary by Division";
proc sgplot data=work.emp_info_all 
		    noborder;
	hbar DIVISION / 
		response=SALARY
        categoryorder=respdesc
		nooutline
		datalabel datalabelattrs=(color = &sasDarkGray size=9.5pt)
		barwidth=.7
		fillattrs = (color = &sasBlue);
	xaxis display = none;
	yaxis display = (NOLABEL NOTICKS) valueattrs = (color = &sasDarkGray);
	yaxistable DIVISION / 
		stat = freq 
		label = 'Total Employees' labeljustify=center labelattrs = (color = &sasDarkGray size = 11pt)
		valuejustify = center valueattrs = (color = &sasDarkGray size = 9.5pt);
	yaxistable SALARY / 
		stat = mean 
		label = 'Mean Salary' labeljustify = center labelattrs = (color = &sasDarkGray size = 11pt) 
		valuejustify = center valueattrs = (color = &sasDarkGray size = 9.5pt);
	format SALARY dollar16.;
run;

ods graphics / reset = width;
title;


/**********************************/
/* List of Employees by Division  */
/**********************************/
%worksheet_title('Detailed Employee Information for Each Division');


proc report data = work.emp_info_all;
	/* Group by DIVISION */
	by Division;
	
	/* Columns to include in the report */
	columns("DIVISION SUMMARY" DIVISION NAME TITLE SALARY EMP_STATUS);

	/* Specify what each column does in the report */
	define DIVISION / group order;
	define SALARY / analysis sum;
	
	/* Color the row of employees on leave in red */
	compute EMP_STATUS;
		length color $100;
		if EMP_STATUS = 'On Leave' then do;
			color = "style={backgroundcolor=&sasRed}";
		end;
		else do;
			color = 'style={backgroundcolor=white}';
		end;
		call define(_row_, 'style', color);
	endcomp;

/* 	Break after each group (DIVISION) for summarization */
	break after DIVISION / summarize 
	                       suppress 
	                       style={fontsize = 11pt fontweight = bold};
run;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;