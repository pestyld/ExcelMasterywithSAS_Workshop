/******************************************/
/* CREATE WORKSHEET N - Employee Leave    */
/******************************************/

ods excel file="&path/final_output/04_dev_excel.xlsx";

ods excel options(sheet_name='Employees on Leave'
                  autofilter='ALL'
				  flow="TABLES"
                  absolute_row_height='20'
				  sheet_interval='NONE');

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
	hbar LVNOTES / categoryorder=respdesc;
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

ods excel close;
