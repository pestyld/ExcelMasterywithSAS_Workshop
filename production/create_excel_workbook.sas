/***********************************************************************************
 PROGRAM DESCRIPTION: Create Monthly Company HR Excel Report                      
************************************************************************************
 SUMMARY: Program prepares and uploads data for the third party legal data report.
          Report can be found in All Reports -> Legal -> Third Party Reports.    
          Report identifies all requested third party legal data. Program executes
          as a job every day at 6AM (EST).                                        
 CREATED BY: Peter S                                                              
 DATE: 04/1/2024                                                                  
************************************************************************************
 REQUIRED INPUT DATA                                                              
************************************************************************************
 1. 2024M04_emp_info_raw.xlsx : main folder > data                                
    - Excel file contains all of the HR data from our systems. The default file    
      name of the file is <YYYY>M<MM>_emp_info_raw for the year and month the      
      data was extracted.                                                         
***********************************************************************************
 REQUIRED PROGRAM FILES                                                           
***********************************************************************************
 - 
 - 
 - 
 - 
 -
***********************************************************************************
* OUTPUT FILES                                                                     
***********************************************************************************
 1. 2024M02_HR_REPORT : main folder > final_project > final_excel_workbook        
    - test Prepared data of ServiceNow third party data. Used to append to Midas  
            tickets. Table is used temporarily (session scope)                    
***********************************************************************************
  REQUIREMENT: SPECIFY MAIN FOLDER PATH                                       
***********************************************************************************/
/* Specify the path to your main workshop folder */
%let folder_path = C:/Users/pestyl/OneDrive - SAS/github repos/ExcelMasterywithSAS_Workshop;


/**************************************************/
/* CUSTOM SETTINGS                                */
/**************************************************/
%include "&production_path./programs/00_config.sas";


/**********************/
/* PREPARE DATA       */
/**********************/
/* Program will read from the raw Excel data (main folder > data) and prepare the final tables in the WORK library */
%include "&production_path./programs/01_prepare_data.sas";



/**************************************************/
/* CREATE EXCEL WORKBOOK                          */
/**************************************************/
/* Close all output destinations */
ods _all_ close;

/* Use PNG images */
ods graphics / imagefmt=png;

/* Write to Excel and create a **DYNAMIC WORKBOOK NAME**  */
ods excel file="&outpath/&currMonthYear._HR_REPORT_FINAL.xlsx" 
          options(sheet_interval ="NONE");


/* Run the following programs. Each program creats a worksheet */
%include "&production_path./programs/02_worksheet01.sas";
%include "&production_path./programs/03_worksheet02.sas";
%include "&production_path./programs/04_worksheet03.sas";
%include "&production_path./programs/05_worksheet04.sas";
%include "&production_path./programs/06_worksheet05.sas";


/* Close and write to Excel */
ods excel close;

/* Reopen HTML destination */
ods html;