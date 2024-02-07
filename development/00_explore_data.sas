/*************************************/
/* 00 EXPLORE DATA                   */
/*************************************/

/**********************************************/
/* 1. DOWNLOAD AND OPEN THE EXCEL WORKBOOK    */
/**********************************************/
/* Main folder > data > emp_info.xlsx         */
/**********************************************/



/**********************************************/
/* 2. MAKE A CONNECTION TO THE EXCEL WORKBOOK */
/**********************************************/

/* REQUIREMENT: Specify the path to your main workshop folder */
%let path = C:/Users/pestyl/OneDrive - SAS/github repos/ExcelMasterywithSAS_Workshop;

/* Make a connection to the emp_info.xlsx workbook */
libname xl xlsx "&path/data/2024M04_emp_info_raw.xlsx";



/********************************/
/* 3. PREVIEW TABLES            */
/********************************/

/* Macro will preview 5 rows and view the column metadata of the data each worksheet */
%macro table_preview(table, n = 5);

	/* Preview table */
	title height = 18pt color = Blue "TABLE: %upcase(&table)";
	title2 color = Blue "----------------------------------------------------------------------------------------------------------------";
	proc print data = &table(obs = &n);
	run;
	title; 

	/* Column metadata */
	ods select Position;
	proc contents data = &table varnum;
	run;
%mend;

%table_preview(xl.empinfo)
%table_preview(xl.jobcodes)
%table_preview(xl.salary)
%table_preview(xl.leave)
;



/*****************************************************/
/* 4. VIEW NUMBER OF ROWS IN EACH SPREADSHEET TABLE  */
/*****************************************************/
proc sql;
	select count(*) as EmpInfoTable_n_rows from xl.empinfo;
	select count(*) as JobCodesTable_n_rows from xl.JobCodes;
	select count(*) as salaryTable_n_rows from xl.salary;
	select count(*) as leaveTable_n_rows from xl.leave;
quit;



/*****************************************/
/* 5. LOOK FOR DUPLICATE ROWS            */
/*****************************************/
proc sql;
/* a. Look for duplicate lookup values in the xl.empinfo worksheet */
title height = 14pt color = Blue "Search for duplicate Employee ID numbers: EMPINFO";
	select IDNUM, count(*) as IDNUM_XL_EMPINFO
	from xl.empinfo
	group by IDNUM
	having calculated IDNUM_XL_EMPINFO > 1;

/* b. Look for duplicate lookup values in the xl.salary worksheet */
title height = 14pt color = Blue "Search for duplicate Employee ID numbers: SALARY";
	select IDNUM, count(*) as IDNUM_XL_SALARY
	from xl.salary
	group by IDNUM
	having calculated IDNUM_XL_SALARY > 1;

/* c. Look for duplicate lookup values in the xl.jobcodes worksheet */
title height = 14pt color = Blue "Search for duplicate Employee ID numbers: JOBCODES";
	select JOBCODE, count(*) as IDNUM_XL_JOBCODES
	from xl.jobcodes
	group by JOBCODE
	having calculated IDNUM_XL_JOBCODES > 1;
quit;
title;


libname xl clear;

/*****************************************/
/* Continue your data exploration        */
/*****************************************/