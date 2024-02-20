# Excel Mastery with SAS: Unlocking Data Insights and Dynamic Reporting

## Requirements
- Access to SAS9 or SAS Viya

## Description
Join us for an immersive experience where data mastery meets reporting excellence! Discover the art of seamlessly reading and writing Excel data with SAS, unlocking a realm of possibilities. From Excel data ingestion to crafting dynamic reports in Excel using SAS, this session empowers you to navigate the convergence of SAS and Excel. Learn efficient SAS techniques to harness the power of both tools, ensuring a seamless workflow for comprehensive data analysis and reporting. Elevate your skills and gain the expertise needed to unleash the full potential of SAS in creating impactful Excel reports. 

## Setup
1. Open the **date** > **cre8data.sas** program. Modify the **data_path** macro variable to reference your **data** folder. Save and run the program. This will create an Excel workbook name using the current year and month *(YYYYMMM_emp_info_raw.xlsx)*. The program in this project will use the current year and month to try and reference the most recent Excel workbook. You will have to create that workbook with this script.
2. You will also have to modify **path** macro variable in the **development** > **00_config.sas** program to reference the location of your main project folder.
3. You will also have to modify **folder_path** macro variable in the **production** > **create_excel_workbook.sas** program to reference the location of your main project folder.

## Folder descriptions

### data 
- The **data** folder contains the program to create the raw data (**cre8data.sas**), the raw Excel workbook used in the demonstration, and the **report_overview_config.xlsx** file. 
- The **report_overview_config.xlsx** workbook stores information about the report and is used to create on of the worksheets. 

### development
The **development** folder will contain programs for the entire process. 
- You will need to run each program incrementally. The programs will create one worksheet at a time for testing purposes. 
- Modify the path in the **config.sas** program to point to the project folder. 
- All Excel output will be placed in the **development** > **output** folder.

### production
- The **production** folder will have a single program named **create_excel_workbook.sas** that will create a single Excel workbook with all of the worksheets from development using those programs. 
- The Modify the **folder_path** macro variable in the **create_excel_workbook.sas** program to point to the project folder. 
- The programs from development are placed in the **production** > **programs** folder. Minor modifications were made to the programs to create a single Excel workbook.
- The final workbook will be placed in the **production** > **output** folder.