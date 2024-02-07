/*************************************************/
/* CREATE DATA FOR THE WORKSHOP                  */
/*************************************************/
/* - Uses the sample sampsio library             */
/* - Should be available in your SAS enviornment */
/* - Creates the emp_info.xlsx file in your path */
/*************************************************/

/*******************************************************************************************/
/* REQUIREMENT: Set the path to your workshop root folder to create the data if necessary  */
/*******************************************************************************************/
%let data_path = C:/Users/pestyl/OneDrive - SAS/github repos/ExcelMasterywithSAS_Workshop/data;



/*******************************************************************/
/*  CREATE EMP_INFO.XLSX FILE IN THE DATA FOLDER                   */
/*******************************************************************/
/* Use the SAMPSIO default SAS library to create the XLSX workbook */
/* - Modify the years to more current years                        */
/* - Increase salary by 30%                                        */
/*******************************************************************/

%let currMonth = %sysfunc(today(), yymm.);
%put &=currMonth;

libname xlout xlsx "&data_path/2024M04_emp_info_raw.xlsx";

data xlout.empinfo;
	set sampsio.empinfo;
	call streaminit(11);
	/* Add years to make data more current */
	addYears=rand('uniform',20,32);
	HDATE=mdy(month(HDATE),day(HDATE),YEAR(HDATE)+addYears);
	BIRTHDAY=mdy(month(BIRTHDAY), day(BIRTHDAY), year(BIRTHDAY)+addYears);
	/* Specify a constant for today's date */
	TODAYS_DATE=mdy(04,01,2024);
	/* Fix a missing HDATE and SALARY */
	if HDATE = . then HDATE = mdy(10,10,2020);
	format TODAYS_DATE date9.;
	drop addYears;
run;

data xlout.salary;
	set sampsio.salary;
	SALARY = SALARY * 1.40;
run;

data xlout.jobcodes;
	set sampsio.jobcodes;
run;

data xlout.leave;
	set sampsio.leave;
	call streaminit(1);
	LEAVEDAYS=round(rand('uniform',14,120));
	LVBEGDTE=mdy(rand('uniform',1,4),day(LVBEGDTE),2024);	
	LVENDDTE=LVBEGDTE+LEAVEDAYS;
	drop LEAVEDAYS;
run;

libname xlout clear;




/********************************/
/*  DELETE BAK FILE             */
/********************************/
* If the bak files exists, delete using the FDELETE function *;
* Reference the Excel workbook .bak file *;
filename f_bak "&data_path/2024M04_emp_info_raw.xlsx.bak";
data _null_;
	if fexist('f_bak') = 1 then do;
		f_bak_exists = fdelete('f_bak');
		put 'NOTE: The file was deleted';
	end;
run;