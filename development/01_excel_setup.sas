/***********************************************/
/* REQUIREMENT: SPECIFY MAIN FOLDER PATH ONLY  */
/***********************************************/

/* REQUIREMENT: Specify the path to your main workshop folder */
%let folder_path = C:/Users/pestyl/OneDrive - SAS/github repos/ExcelMasterywithSAS_Workshop;



/* Data path */
%let data_path = &folder_path./data;

/* Program path */
%let program_path = &folder_path./development;

/* Specify Excel file output folder */
%let outpath = &program_path./excel_output;



/******************************************/
/* CUSTOM ATTRIBUTES SETUP                */
/******************************************/

/* Dynamically store today's year and month */
%let currMonth = %sysfunc(today(), YYMM.);
%put &=currMonth;

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

/* Font sizes */
%let ws_title_text = 20pt;


/* Add additional styles */
%let titleFmt = height=16pt justify=left color=&sasDarkBlue;