# Excel Proficiency with SAS: Unlocking Data Insights and Dynamic Reporting

## Requirements
- Access to SAS9 or SAS Viya

## Description
Join us for an immersive experience where data mastery meets reporting excellence! Discover the art of seamlessly reading and writing Excel data with SAS, unlocking a realm of possibilities. From Excel data ingestion to crafting dynamic reports in Excel using SAS, this session empowers you to navigate the convergence of SAS and Excel. Learn efficient SAS techniques to harness the power of both tools, ensuring a seamless workflow for comprehensive data analysis and reporting. Elevate your skills and gain the expertise needed to unleash the full potential of SAS in creating impactful Excel reports. 

## Introduction
You will learn how to
- Import data from Excel workbooks and generate SAS tables.
- Prepare the imported worksheets using PROC SQL.
- Generate reports using SAS procedures including PROC SQL, SGPLOT, and others.
- Construct Excel worksheets leveraging SAS ODS EXCEL and its diverse options.
- Dynamically assemble a final production-grade project for automated creation of Excel reports.


## Setup
1. You will also have to modify **path** macro variable in the **development** > **00_config.sas** program to reference the location of this main project folder.
2. You will also have to modify **folder_path** macro variable in the **production** > **create_excel_workbook.sas** program to reference the location of this main project folder.

## Folder descriptions

### data 
The **data** folder contains the following:
1. The program to create the raw data for the workshop (**cre8data.sas**). You do not need to run this program as the Excel file is already created.
2. The raw Excel workbook used in the demonstration (**2024M04_emp_info_raw.xlsx**). This is the input data required for the workshohp.
3. The **report_overview_config.xlsx** file. This workbook stores information about the report and is used to store report overview information for the workbook and worksheets.

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

### Final Note
- The last worksheet creates a schedule plot. The schedule plot uses a static date for the reference line. However, the data label is dynamic, and will always be the date you run the program. **So the reference line and the date will not match**. This is done to show that if the data was updated on the backend, you can dynamically update the visualization. In this scenario, the data is static.