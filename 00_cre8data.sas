/*************************************************/
/* CREATE DATA FOR THE WORKSHOP                  */
/*************************************************/
/* - Uses the sample sampsio library             */
/* - Should be available in your SAS enviornment */
/* - Creates 4 tables in your WORK library       */
/*************************************************/

/************************************************************/
/* REQUIREMENT: Set the path to your workshop root folder   */
/************************************************************/
* Current folder. SAS program must be saved to the location *; 
* Works in SAS Studio and SAS Enterprise Guide *;
%let fileName =  %scan(&_sasprogramfile,-1,'/');
%let path = %sysfunc(tranwrd(&_sasprogramfile, &fileName,));



/********************************/
/*  CREATE DATA                 */
/********************************/
data work.jobcodes;
	set sampsio.jobcodes;
run;

data work.empinfo;
	set sampsio.empinfo;
	drop STATUS	EDLEV;
run;

data work.leave;
	set sampsio.leave;
run;


data work.salary;
	set sampsio.salary;
run;

/********************************/
/* PREVIEW TABLES               */
/********************************/
%macro head(table, n=5);
	title "TABLE: %upcase(&table)";
	proc print data=&table(obs=&n);
	run;
%mend;

%head(work.empinfo)
%head(work.jobcodes)
%head(work.leave)
%head(work.salary)

proc contents data=work.salary;
run;

proc print data=work.leave;
run;

proc sql;
create table emp_info_final as
	select emp.*, 
           job.TITLE,
           sal.SALARY, sal.ENDDATE
	from work.empinfo as emp 
    	left join work.jobcodes as job on emp.jobcode = job.jobcode
	    left join work.salary as sal on emp.idnum = sal.idnum;
quit;

proc print data=emp_info_final;
run;
