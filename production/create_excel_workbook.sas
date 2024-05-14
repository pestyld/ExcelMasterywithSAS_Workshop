/***********************************************************************************
 PROGRAM DESCRIPTION: Create Monthly Company HR Excel Report                      
************************************************************************************
 SUMMARY: This project automates the analysis of HR data by identifying the most 
          recent HR emp_info_raw.xlsx file based on the current year and month. It then 
          proceeds to execute all available SAS programs in the designated folder, consolidating 
          the results into a comprehensive Excel workbook providing insights into our 
          company's HR metrics.
 CREATED BY: Peter S                                                             
 DATE: 03/01/2024                                                                  
************************************************************************************
 REQUIRED INPUT DATA                                                              
************************************************************************************
 1. <YYYY>M<MM>_emp_info_raw.xlsx : main folder > data                                
    - Excel file containing all HR data from our systems. The default file    
      name is <YYYY>M<MM>_emp_info_raw for the year and month the data was extracted.                                                         
***********************************************************************************
 REQUIRED PROGRAM FILES                                                           
***********************************************************************************
 - 00_config.sas       - Creates macro variables and programs used in the Excel output.
 - 01_prepare_data.sas - Prepares the <YYYY>M<MM>_emp_info_raw.xlsx file into SAS tables. 
 - 02_worksheet01.sas  - Creates the first worksheet with a workbook overview.
 - 03_worksheet02.sas  - Creates the second worksheet with a list of all employees.
 - 04_worksheet03.sas  - Creates the third worksheet with company overview information.
 - 05_worksheet04.sas  - Creates the fourth worksheet with division analysis.
 - 06_worksheet05.sas  - Creates the fifth worksheet with employee leave overview.
***********************************************************************************
* OUTPUT FILES                                                                     
***********************************************************************************
 1. <YYYY>M<MM>_HR_REPORT.xlsx
   - Location: main folder > production > output 
   - Dynamically names the workbook using the month and year the program was run.
***********************************************************************************
  REQUIREMENT: SPECIFY MAIN FOLDER PATH                                       
***********************************************************************************/
/* Specify the path to your main workshop folder */
%let folder_path = /* Enter path to main workshop folder */;



/**************************************************/
/* CONFIGURATION FILE SETTINGS                    */
/**************************************************/
%include "&folder_path/production/programs/00_config.sas";


/**********************/
/* PREPARE DATA       */
/**********************/
/* Program will read from the raw Excel data (main folder > data) and prepare the tables in the WORK library */
%include "&production_path./programs/01_prepare_data.sas";


/**************************************************/
/* CREATE EXCEL WORKBOOK                          */
/**************************************************/
/* Close all output destinations */
ods _all_ close;

/* Use PNG images */
ods graphics / imagefmt=png;

/* Write to Excel and create a workbook dynamically named <YYYY>M<MM>_HR_REPORT_FINAL.XLSX */
ods excel file="&outpath/&currMonthYear._HR_REPORT_FINAL.xlsx" 
          options(sheet_interval ="NONE");


/* Run the following programs. Each program creats a worksheet in the workbook */
%include "&production_path./programs/02_worksheet01.sas";
%include "&production_path./programs/03_worksheet02.sas";
%include "&production_path./programs/04_worksheet03.sas";
%include "&production_path./programs/05_worksheet04.sas";
%include "&production_path./programs/06_worksheet05.sas";


/* Close and write to Excel */
ods excel close;

/* Reopen HTML destination */
ods html;