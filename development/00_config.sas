/*************************************/
/* 00 CONFIG                         */
/*************************************/

/***********************************************/
/* REQUIREMENT: SPECIFY MAIN FOLDER PATH ONLY  */
/***********************************************/
/* REQUIREMENT: Specify the path to your main workshop folder */
%let path = C:/Users/pestyl/OneDrive - SAS/github repos/ExcelMasterywithSAS_Workshop;



/***********************************************/
/* SET FOLDER PATHS                            */
/***********************************************/
/* Data path */
%let data_path = &path./data;

/* Program path */
%let program_path = &path./development;

/* Specify Excel file output folder */
%let outpath = &program_path./output;



/******************************************/
/* CUSTOM ATTRIBUTES SETUP                */
/******************************************/

/* Dynamically store today's year and month */
%let currMonthYear = %sysfunc(today(), YYMM.);
%put &=currMonthYear;

%let currDate = %sysfunc(today(), WEEKDATE.);
%put &=currDate;


/* Create custom company colors */
%let sasBlue = CX0766D1;
%let sasDarkBlue = CX032954;
%let sasYellow = CXFFCC33;
%let sasRed = CXF24949;
%let sasTeal = CX3ADBE6;
%let sasGreen = CX36D982;
%let sasSlate = CX7E889A;
%let sasDarkGray = CX222222;

/* Font sizes */
%let ws_title_text = 20pt;


/* Add additional styles */
%let titleFmt = height=16pt justify=left color=&sasDarkBlue;



/***********************/
/* CUSTOM MACRO        */
/***********************/
%macro worksheet_title(title_string);
	proc odstext;
		p &title_string / style = [color = &sasDarkBlue fontsize = &ws_title_text];
	quit;
%mend;