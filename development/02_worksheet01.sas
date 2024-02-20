/*********************************************************************************
 WORKSHEET 1 - Workbook Description                                             
**********************************************************************************
 REQUIREMENT: The 00_config and 01_prepare_data programs need to be run.                  
*********************************************************************************/


/*****************************************/
/* SET TEXT FONT SIZES FOR THE WORKSHEET */
/*****************************************/
%let heading1_size = 22pt;
%let p_size = 14pt;  
%let list_size = 12pt; 
%let footnote_size = 10pt;



/******************************************************/
/* READ THE report_overview_config.xlsx FILE          */
/******************************************************/
/* Read the xlsx file with information about the workbook and create SAS tables */
libname xlconfig xlsx "&data_path/report_overview_config.xlsx";

/* Preview the worksheets */
proc print data = xlconfig.overview;
run;

proc print data = xlconfig.worksheets;
run;



/**********************/
/* OUTPUT TO EXCEL    */
/**********************/
/* Create Excel workbook */
ods excel file = "&dev_outpath/worksheet_01.xlsx" 
          options(sheet_name = "Workbook Overview");

 
/* Add workbook title and description using the overview worksheet */
proc odstext data = xlconfig.overview;
	/* Workbook title */
	p catx(' - ', Title, "&currMonthYear") / style = [color = &sasDarkBlue 
	                                                  fontsize = &heading1_size 
	                                                  tagattr = 'mergeacross:13'];
	/* Workbook description */
	p WorkbookDescription / style = [color = &sasDarkBlue 
	                                 fontsize = &p_size 
	                                 tagattr = 'mergeacross:13'];
	/* Insert blank line */                                 
	p ' ' / style = [color = white];
	p 'The workbook includes the following worksheets:' / style=[color=&sasDarkBlue 
	                                                             fontsize=&p_size 
																 tagattr = 'mergeacross:13'];
run;


/* Dynamic worksheet descriptions based off the XLSX file worksheet */
proc odstext data = xlconfig.worksheets;
	list / style = [color = &sasDarkBlue 
	                fontsize = &list_size 
	                tagattr = 'mergeacross:13'];
		item catx(' - ', WorksheetName, Description);
	end;
run;


/* Footer workbook creation information using the XLSX overview worksheet */
proc odstext data = xlconfig.overview;
	p ' ';
	p 'Maintained by ' || CreatedBy || ' in the ' || Department || ' department. Contact at ' || Phone || ' or ' || Email || '.' / 
	      style = [color = &sasDarkBlue 
	               fontsize = &footnote_size 
				   tagattr = 'mergeacross:13'];
	p "Report created on &currDate" / style = [color = &sasDarkBlue 
	                                           fontsize = &footnote_size 
											   tagattr = 'mergeacross:13'];
	end;
quit;


/* Clear the library reference */
libname xlconfig clear;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;