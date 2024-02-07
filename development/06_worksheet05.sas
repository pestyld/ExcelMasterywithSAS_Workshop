/**********************************************************************************/
/* WORKSHEET 5 - Employee Leave                                                   */
/**********************************************************************************/
/* REQUIREMENT: The 01_explore_and_prepare_data.sas program needs to be run first */
/**********************************************************************************/


/*******************************************************/
/* EXCEL OUTPUT                                        */
/*******************************************************/
ods excel file="&outpath/worksheet_05.xlsx";


ods excel options(sheet_name='Employees on Leave'
                  sheet_interval='NONE'
                  autofilter='ALL'
				  flow="TABLES"
                  absolute_row_height='20');

proc print data=work.emp_leave noobs;
	var NAME IDNUM EMPNO DIVISION JOBCODE TITLE LVNOTES LVBEGDTE LVENDDTE PRPERCNT LVSALARY;
run;

proc sgplot data=work.emp_leave
            noborder;
	hbar DIVISION / categoryorder=respdesc;
run;
proc means data=work.emp_leave sum;
	class DIVISION;
	var LVSALARY;
quit;

proc sgplot data=work.emp_leave
            noborder;
	hbar LVNOTES / 
		categoryorder=respdesc;
run;

proc sgplot data=work.emp_leave
            noborder;
	highlow y=NAME low=LVBEGDTE high=LVENDDTE /
		type = bar
		barwidth = .5
		nooutline
		group = DIVISION 
;
run;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;