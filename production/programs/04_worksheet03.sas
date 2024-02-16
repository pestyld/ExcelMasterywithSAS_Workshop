/*********************************************************************************
 WORKSHEET 3  - Company Overview                                              
**********************************************************************************
 REQUIREMENT: The 00_config and 01_prepare_data programs need to be run.                  
**********************************************************************************
 NOTE: The SGPIE procedure is a preproduction feature, which means that it has  
       not been fully developed, tested, or documented.                         
**********************************************************************************/


/**********************/
/* OUTPUT TO EXCEL    */
/**********************/
ods excel options(
			  sheet_name = 'Company Overview'        /*<--- specifies the name for the next worksheet */
			  sheet_interval = 'NOW'                 /*<--- new worksheet */
              autofilter = 'NONE'                    /*<--- turns on filtering for specified columns in the worksheet */
              flow = "TABLES"                        /*<--- specifies that a designated Worksheet area enables Wrap Text and disables newline character insertion */
              row_heights = '20,20,20,20,20,20,20'   /*<--- specifies the height of the row using positional parameters */
              absolute_column_width = '40, 15'       /*<--- specifies the column widths */
			  frozen_headers = 'OFF'                 /*<--- specifies that headers can scroll or not scroll with the scroll bar */
              frozen_rowheaders = 'OFF'              /*<--- specifies if the row headers are on the left scroll when the table data scrolls */
	  	   );	
		   
/* Use PNG images */
ods graphics / imagefmt=png;
	

/**************************/
/* WORKSHEET TITLE        */
/**************************/
%worksheet_title('Employee Company Overview Information');



/**************************/
/* HEAD COUNT BY LOCATION */
/**************************/
title &titleFmt "Total Head Count by Location";
proc sgpie data = work.emp_info_all;
	donut LOCATION / 
		holevalue holevalueattrs= (color = &sasDarkBlue)
		holelabel = 'Headcount' holelabelattrs = (color = &sasDarkBlue)
		datalabelattrs = (color = &sasDarkBlue)
		datalabelloc = callout 
		maxslices = 4
		otherfillattrs = (color = &sasSlate);
	styleattrs datacolors = (&sasTeal &sasBlue &sasDarkBlue);
run;
title;

proc sql;
	select LOCATION, count(*) as TotalEmp label = 'Total Employees'
		from work.emp_info_all
		group by LOCATION
		order by TotalEmp desc;
quit;



/**************************/
/* TOTAL ACTIVE EMPLOYEES */
/**************************/
title &titleFmt "Total Active Employees";
proc sgpie data = work.emp_info_all;
	donut EMP_STATUS / 
        holevalue holevalueattrs = (color = &sasDarkBlue)
		holelabel = 'Headcount' holelabelattrs = (color = &sasDarkBlue)
		datalabelattrs = (color = &sasDarkBlue)
		datalabelloc = callout;
	styleattrs datacolors = (&sasGreen &sasRed);
run;
title;

proc sql;
	select EMP_STATUS, count(*) as TotalEmp label = 'Total Employees'
		from work.emp_info_all
		group by EMP_STATUS
		order by TotalEmp desc;
quit;



/*********************************/
/* YEARS OF SERVICE BY DIVISION  */
/*********************************/
title &titleFmt "Average Years of Service (YOS) by Division";
ods graphics / width = 8in;
proc sgplot data = work.emp_info_all
            noborder;
	hbar DIVISION /
		response = EMPYOS
		stat = MEAN 
		categoryorder=respdesc
		fillattrs = (color = &sasBlue)
		nooutline;
	yaxis display = (nolabel);
quit;
ods graphics / reset=width;
title;

proc sql;
	select DIVISION, 
	       round(mean(EMPYOS)) as MeanYOS label = 'Avg YOS'
		from work.emp_info_all
		group by DIVISION
		order by MeanYOS desc;
run;


/***************************/
/* EMPLOYEES HIRED BY YEAR */
/***************************/
/* Find the range of hired years and store in macro variables */
proc sql noprint;
	select min(year(HDATE)) as MinYear, 
	       max(year(HDATE)) as MaxYear
		into :MinYear trimmed, 
		     :MaxYear trimmed
		from work.emp_info_all;
quit;
%put &=MinYear &=MaxYear;

/* Create a dummy table to get all years in range even if one didn't hire */
data work.all_years;
	TotalHired  =  0;
	do Year = &MinYear to &MaxYear;
		output;
	end;
run;

/* Join dummy table and a summarized table of hired by year */
proc sql;
create table work.hired_by_year as
	select ay.Year, 
	       coalesce(sum.TotalHired, ay.TotalHired) as TotalHired 
	from work.all_years as ay 
		left join (select year(HDATE) as Year, count(*) as TotalHired
				    from work.emp_info_all
	                group by calculated Year) as sum
	    on ay.Year  =  sum.Year;
quit;

/* Plot the data */
title &titleFmt "Employees Hired by Year";
ods graphics / width = 8in;
proc sgplot data = work.hired_by_year
            noborder;
	vbar Year /
		response = TotalHired
		nooutline
		datalabel datalabelattrs = (color = &sasDarkBlue size=10pt)
		fillattrs = (color = &sasBlue);
	yaxis display = NONE;
	xaxis display = (nolabel);
run;
ods graphics / reset=width;
title;