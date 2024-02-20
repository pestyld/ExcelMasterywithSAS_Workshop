/**********************************************************************************************************
 CREATE DATA FOR THE WORKSHOP                  
***********************************************************************************************************
 - Uses the sample sampsio library             
 - Should be available in your SAS enviornment
 - Creates the <YYYY>M<MM>_emp_info_raw.xlsx workbook in your data folder with the current year and month
***********************************************************************************************************/

/*****************************************************************************
 REQUIREMENT: Set the path to where you want to create the Excel workbook  
*****************************************************************************/
%let data_path = C:/Users/pestyl/OneDrive - SAS/github repos/ExcelMasterywithSAS_Workshop/data;



/******************************************************************
  CREATE EMP_INFO.XLSX WORKBOOK IN THE SPECIFIED FOLDER          
*******************************************************************
 Use the SAMPSIO default SAS library to create the XLSX workbook 
 - Modifies the years to more current years                        
 - Increase salary by 30%                                        
 - Modifies column metadata                                       
/*******************************************************************/
/* Dynamically store today's year and month */
%let currMonthYear = %sysfunc(today(), YYMM.);
%put &=currMonthYear;


/* Create Excel workbook: <YYYY><M<NN>_emp_info_raw.xlsx using the macro above to name the file */
libname xlout xlsx "&data_path/&currMonthYear._emp_info_raw.xlsx";


/* Create the Empinfo worksheet */
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

/* Create the salary worksheet */
data xlout.salary;
	set sampsio.salary;
	SALARY = SALARY * 1.40;
run;

/* Create the jobcodes worksheet */
data xlout.jobcodes;
	set sampsio.jobcodes;
run;

/* Create the leave worksheet */
data xlout.leave;
	set sampsio.leave;
	call streaminit(1);
	LEAVEDAYS=round(rand('uniform',14,120));
	LVBEGDTE=mdy(rand('uniform',1,4),day(LVBEGDTE),2024);	
	LVENDDTE=LVBEGDTE+LEAVEDAYS;
	drop LEAVEDAYS NAME;
run;

/* Close the libname */
libname xlout clear;



/**********************************************************
 DELETES BAK FILE
 - Writing to an Excel workbook creates a .bak file. 
   This script deletes that extra file
**********************************************************/
* If the bak files exists, delete using the FDELETE function *;
* Reference the Excel workbook .bak file *;
filename f_bak "&data_path/&currMonthYear._emp_info_raw.xlsx.bak";
data _null_;
	if fexist('f_bak') = 1 then do;
		f_bak_exists = fdelete('f_bak');
		put 'NOTE: The file was deleted';
	end;
run;