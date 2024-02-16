/******************************************/
/* CONFIGURATION                          */
/******************************************/

/**********************************************/
/* FOLDER PATHS                               */
/**********************************************/
/* Set the path to main folder > data folder that contains the raw Excel data */
%let data_path = &folder_path./data;

/* Set the path to main folder > production folder to run all of the SAS programs */
%let production_path = &folder_path./production;

/* Set the path to main folder > final_project > final_excel_workbook to output the final Excel file */
%let outpath = &production_path./output;



/******************************************/
/* MACRO VARIABLES                        */
/******************************************/
/* Dynamically store today's year and month */
%let currMonthYear = %sysfunc(today(), YYMM.);
%put &=currMonthYear;

/* Dynamically store today's date */
%let currDate = %sysfunc(today(), WEEKDATE.);
%put &=currDate;


/* Create custom company colors for the Excel report */
%let sasBlue = CX0766D1;
%let sasDarkBlue = CX032954;
%let sasYellow = CXFFCC33;
%let sasRed = CXF24949;
%let sasTeal = CX3ADBE6;
%let sasGreen = CX36D982;
%let sasSlate = CX7E889A;

/* Font sizes */
%let ws_title_text = 20pt; /* Font size for the ODS text for top of each worksheet */


/* Add additional styles */
%let titleFmt = height=16pt justify=left color=&sasDarkBlue; /* Title format for all procedures */



/***********************/
/* CUSTOM MACRO        */
/***********************/
/* Worksheet title and section text */
%macro worksheet_title(title_string);
	proc odstext;
		p &title_string / style = [color = &sasDarkBlue fontsize = &ws_title_text];
	quit;
%mend;