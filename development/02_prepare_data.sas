/*************************************/
/* 01 PREPARE DATA                   */
/*************************************/

/*************************************/
/* 1. CONNECT TO THE EXCEL WORKBOOK  */
/*************************************/
/* Make a connection to the emp_info.xlsx workbook */
libname xl xlsx "&path/data/2024M04_emp_info_raw.xlsx";



/********************************/
/* 1. REMOVE DUPLICATES         */
/********************************/
/* Remove duplicates jobcodes from the xl.jobcodes worksheet and create a SAS table */
proc sort data=xl.jobcodes     /* Read from Excel */
		  out=work.jobcodes    /* Create SAS table with no duplicates */
		  nodupkey;
	by jobcode;
run;



/********************************/
/* 2. PREPARE FINAL TABLE       */
/********************************/
/* Perform a join of the following tables */
/* - xl.empinfo    */
/* - work.jobcodes */
/* - xl.salary     */
/* - xl.leave      */

proc sql;
create table work.emp_info_all as
	select emp.NAME label='Name',
		   put(emp.IDNUM,SSN11.) as IDNUM label='Identification Number',
		   emp.EMPNO label='Employee Number',
		   emp.DIVISION label='Division',
		   emp.JOBCODE label='Job Code',
           job.TITLE label='Job Title',
		   sal.SALARY label='Salary' format=dollar20.2,
		   emp.HDATE label='Hire Date',
		   floor(yrdif(emp.HDATE,mdy(1,1,2023),'age')) as EMPYOS label='Years of Service',
		   emp.STATUS label='Status',
		   emp.GENDER label='Employee Gender',
		   emp.EDLEV label='Education Level',
		   emp.LOCATION label='Office Location',
		   emp.PHONE label='Extension Number',
		   emp.ROOM label='Office Location',
		   emp.TODAYS_DATE label="Today's Date",
		   /* Create column for active and on leave */
		   case
		   		when leave.LVTYPE is null then 'Active'
		   		else 'On Leave'
		   end as EMP_STATUS label='Employee Status',
		   leave.LVTYPE label='Type of Leave',
           leave.LVNOTES label='Notes About Leave',
		   leave.LVBEGDTE label='Leave Begin Date', 
		   leave.LVENDDTE label='Leave End Date',
		   leave.LVENDDTE - leave.LVBEGDTE as LVDAYS label='Leave Days',
           leave.PRPERCNT label='Payroll Percentage',
           /* Create new salary for employees on leave */
		   case
				when leave.LVTYPE is null then . 
				else sal.SALARY * leave.PRPERCNT
		   end as LVSALARY label='Leave Salary' format=dollar20.2
	from xl.empinfo as emp 
    	left join work.jobcodes as job on emp.jobcode = job.jobcode
	    left join xl.salary as sal on emp.idnum = sal.idnum
	    left join xl.leave as leave on emp.idnum = leave.idnum
	order by Division, Salary;


	/* Count number of rows in final table */
	select count(*) as NumRows_Final_Table from work.emp_info_all;


	/* View final table */
	title "FINAL EMP_INFO_ALL TABLE";
	select * from work.emp_info_all(obs=10);
quit;



/****************************************/
/* 3. PREPARE ACTIVE AND LEAVE TABLES   */
/****************************************/
/* Create ACTIVE and LEAVE tables */
data work.emp_active(drop=LVTYPE LVNOTES LVBEGDTE LVENDDTE PRPERCNT LVSALARY LVDAYS) 
     work.emp_leave;
	set work.emp_info_all;
	if LVTYPE = '' then output work.emp_active;
	else output work.emp_leave;
run;


/* Sort leave data */
proc sort data=work.emp_leave;
	by descending LVBEGDTE;
run;


/* Preview tables */
title "PREVIEW WORK.EMP_ACTIVE TABLE";
proc print data=work.emp_active(obs=5);
run;
title;

title "PREVIEW WORK.EMP_LEAVE TABLE";
proc print data=work.emp_leave(obs=5);
run;
title;