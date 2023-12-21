/*************************************/
/* 01_EXPLORE AND PREPARE DATA       */
/*************************************/


/**********************************************/
/* 1. DOWNLOAD AND OPEN THE EXCEL WORKBOOK    */
/**********************************************/
/* Workshop folder > data > emp_info.xlsx     */


/**********************************************/
/* 2. MAKE A CONNECTION TO THE EXCEL WORKBOOK */
/**********************************************/

/* Store the path of the root folder */
%let fileName =  /%scan(&_sasprogramfile,-1,'/');
%let path = %sysfunc(tranwrd(&_sasprogramfile, &fileName,));


/* Make a connection to the emp_info.xlsx workbook */
libname xl xlsx "&path/data/emp_info.xlsx";



/********************************/
/* 3. PREVIEW TABLES            */
/********************************/

/* Macro will preview 5 rows and view the column metadata */
%macro table_preview(table, n=5);
	/* Preview table */
	title height=18pt color=Blue "TABLE: %upcase(&table)";
	title2 color=Blue "--------------------------------------------------------------------------------------------------------------------------------------------------------------";
	proc print data=&table(obs=&n);
	run;
	title; 

	/* Column metadata */
	ods select Position;
	proc contents data=&table varnum;
	run;
%mend;

%table_preview(xl.empinfo)
%table_preview(xl.jobcodes)
%table_preview(xl.salary)
%table_preview(xl.leave)



/*********************************/
/* 3. EXPLORE DATA               */
/*********************************/
/* Number of rows in each table */
proc sql;
	select count(*) as EmpInfoTable from xl.empinfo;
	select count(*) as JobCodesTable from xl.JobCodes;
	select count(*) as salaryTable from xl.salary;
	select count(*) as leaveTable from xl.leave;
quit;


/* Look for duplicate lookup values in the xl.empinfo worksheet */
proc freq data=xl.empinfo order=freq noprint;
	tables IDNUM / out=work.empinfo_dups(where=(Count>1));
run;


/* Look for duplicate lookup values in the xl.salary worksheet */
proc freq data=xl.salary order=freq noprint;
	tables IDNUM / out=work.salary_dups(where=(Count>1));
run;


/* Look for duplicate lookup values in the xl.jobcodes worksheet */
proc freq data=xl.jobcodes order=freq noprint;
	tables jobcode / out=work.jobcodes_dups(where=(Count>1));
run;



/********************************/
/* 3. PREPARE TABLES            */
/********************************/

/* Remove duplicates jobcodes from the xl.jobcodes worksheet */
proc sort data=xl.jobcodes 
		  out=work.jobcodes 
		  nodupkey;
	by jobcode;
run;


/* Perform a join of the following tables */
/* - xl.empinfo    */
/* - work.jobcodes */
/* - xl.salary     */
/* - xl.leave      */

proc sql;
create table work.emp_info_all as
	select emp.NAME label='Name',
		   emp.IDNUM label='Identification Number' format=SSN11.,
		   emp.EMPNO label='Employee Number',
		   emp.DIVISION label='Division',
		   emp.JOBCODE label='Job Code',
           job.TITLE label='Job Title',
		   sal.SALARY label='Salary' format=dollar20.2,
		   emp.HDATE label='Hire Date',
		   floor(yrdif(emp.HDATE,mdy(1,1,2023),'age')) as EMPYOS label='Years of Service',
		   sal.ENDDATE label='End Date',
		   emp.STATUS label='Status',
		   emp.GENDER label='Employee Gender',
		   emp.EDLEV label='Education Level',
		   emp.LOCATION label='Office Location',
		   emp.PHONE label='Extension Number',
		   emp.ROOM label='Office Location',
		   leave.LVTYPE label='Type of Leave',
           leave.LVNOTES label='Notes About Leave',
		   leave.LVBEGDTE label='Leave Begin Date', 
		   leave.LVENDDTE label='Leave End Date', 
           leave.PRPERCNT label='Payroll Percentage',
		   case
				when leave.LVTYPE is null then . 
				else sal.SALARY * leave.PRPERCNT
		   end as LVSALARY label='Leave Salary' format=dollar20.2
	from xl.empinfo as emp 
    	left join work.jobcodes as job on emp.jobcode = job.jobcode
	    left join xl.salary as sal on emp.idnum = sal.idnum
	    left join xl.leave as leave on emp.idnum = leave.idnum
	order by Division, Name;

	/* Count number of rows in final table */
	select count(*) as NumRows_Final_Table from work.emp_info_all;

	/* View final table */
	select * from work.emp_info_all;
quit;



/* Create ACTIVE and LEAVE tables */
data work.emp_active(drop=LVTYPE LVNOTES LVBEGDTE LVENDDTE PRPERCNT LVSALARY) 
     work.emp_leave;
	set work.emp_info_all;
	if LVTYPE = '' then output work.emp_active;
		else output work.emp_leave;
run;
proc print data=work.emp_active(obs=5);
run;
proc print data=work.emp_leave(obs=5);
run;