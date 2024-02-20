/*********************************************************************************
 WORKSHEET 5 - Employee Leave                                                
**********************************************************************************
 REQUIREMENT: The 00_config and 01_prepare_data programs need to be run.                  
*********************************************************************************/


/*******************************************************/
/* EXCEL OUTPUT                                        */
/*******************************************************/
ods excel file = "&dev_outpath/worksheet_05.xlsx";


ods excel options(
				sheet_name = 'Employee Leave'
                sheet_interval = 'NONE'
                autofilter = 'ALL'
                row_heights = '30,20,20,20,20,20,20' 
				flow = "TABLES"
			);

/* Use PNG images */
ods graphics / imagefmt=png;


/**************************/
/* WORKSHEET TITLE        */
/**************************/
%worksheet_title('Employees on Leave Overview');


/******************************/
/* EMPLOYEE LEAVE LIST        */
/******************************/
proc print data = work.emp_leave noobs;
	var NAME IDNUM EMPNO DIVISION JOBCODE TITLE LVNOTES LVBEGDTE LVENDDTE PRPERCNT LVSALARY;
run;


/*************************************************/
/* PERCENTAGE OF EMPLOYEES ON LEAVE BY DIVISION  */
/*************************************************/
/* Create an aggregated table with the total emplopees, total on leave, and percentage of employees on leave by division */
proc sql;
create table pct_leave_division as
	select DIVISION,
	       count(*) as Total_Employees label = 'Total Employees',                             
		   sum(ifn(EMP_STATUS = 'On Leave',1,0)) as Total_Leave label = 'Total on Leave',   
		   calculated Total_Leave/calculated Total_Employees as PctLeave 
		   		format = percent7. label = 'Percentage on Leave'
		from work.emp_info_all
		group by DIVISION
		having calculated PctLeave > 0
		order by PctLeave desc;
quit;


/* Create the plot */
ods graphics / height = 5in width = 8in;
title &titleFmt "Percentage of Employees on Leave by Division";
proc sgplot data = work.pct_leave_division
            noborder;
	vbar DIVISION / 
		response = PctLeave 
		categoryorder = respdesc
		nooutline 
		barwidth=.4
		fillattrs = (color = &sasRed);
	yaxis labelattrs = (color = &sasDarkGray) 
	      valueattrs = (color = &sasDarkGray);
	xaxis display = (nolabel)
	      colorbands = even colorbandsattrs = (color = &sasSlate transparency = .95);
	xaxistable Total_Employees Total_Leave / 
		valueattrs = (color = &sasDarkGray size = 10pt) 
		labelattrs = (color = &sasDarkGray size = 10pt weight = bold)
		pad = (top = 5px bottom = 5px) 
		position = top location = inside;
run;
title;
ods graphics / reset = height reset = width;


/******************************/
/* EMPLOYEE LEAVE TIMELINE    */
/******************************/
ods graphics / width = 8in height = 7in;
title &titleFmt "Employees Leave Timeline";
proc sgplot data = work.emp_leave
            noborder;
	highlow y = NAME low=LVBEGDTE high = LVENDDTE /
		type = bar barwidth = .3
		fillattrs = (color=&sasRed)
		nooutline
		highlabel = LVENDDTE 
		labelattrs = (color=&sasDarkBlue size=10pt)
	;
	xaxis display = (nolabel) 
	      valueattrs =(color = &sasDarkGray) valuesformat = monyy7.;
	yaxis display = (nolabel)
	      valueattrs=(color = &sasDarkGray) 
	      colorbands = even colorbandsattrs = (transparency = 0.95 color = &sasSlate);
	/* Adding reference line for today's date. Using a static value here to always show the line */
	/* The label is custom and will always use the date the program was run to show how to make this dynamic */
	refline '01APR2024'd / 
		label = "Date: &currDate" labelattrs = (color = &sasDarkBlue weight = bold size = 11pt)
		axis = x 
		lineattrs = (color = &sasDarkGray);
run;
title;
ods graphics / reset = width reset = height;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;