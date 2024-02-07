/**********************************************************************************/
/* WORKSHEET 1 - Workbook Description                                             */
/**********************************************************************************/
/* REQUIREMENT: The 01_excel_setup program needs to be run first                  */
/**********************************************************************************/


/*****************************************/
/* SET TEXT FONT SIZES FOR THE WORKSHEET */
/*****************************************/
%let heading1_size = 22pt;
%let p_size = 14pt;  
%let list_size = 12pt; 
%let footnote_size = 10pt;



/***************************************/
/* READ JSON WORKBOOK INFORMATION FILE */
/***************************************/
/* Read the JSON file with information about the workbook and create SAS tables */
fileName myFile "&path/development - SAS Results/workbook_overview.json";
libname myFile json fileref = myFile noalldata;

/* View the contents of the JSON file */
data _null_;
	rc = jsonpp('myFile', 'log');

/* Preview JSON tables */
proc print data=myFile.root;
run;

proc print data=myFile.worksheets;
run;



/**********************/
/* OUTPUT TO EXCEL    */
/**********************/
ods excel file="&outpath/worksheet_01.xlsx" 
          options(sheet_name ="Workbook Overview");

  
/* Workbook introduction title and text using the JSON root table */
proc odstext data = myFile.root;
	p Title / style = [color = &sasDarkBlue fontsize = &heading1_size];
	p WorkbookDescription / style = [color = &sasDarkBlue 
	                                 fontsize = &p_size 
	                                 tagattr = 'mergeacross:10'];
	/* Insert blank line */                                 
	p ' ' / style = [color = white];
	p 'The workbook includes the following worksheets:' / style=[color=&sasDarkBlue fontsize=&p_size];
run;


/* Dynamic worksheet descriptions based off the JSON file information */
proc odstext data = myfile.worksheets;
	list / style = [color = &sasDarkBlue fontsize = &list_size tagattr = 'mergeacross:10'];
		item catx(' - ', WorksheetName, Description);
	end;
run;


/* Footer workbook creation information */
proc odstext data = myFile.root;
	p ' ';
	p 'Maintained by ' || CreatedBy || ' in the ' || Department || ' department. Contact at ' || Phone || ' or ' || Email || '.' / style = [color = &sasDarkBlue fontsize = &footnote_size];
	p "Report created on &currDate" / style = [color = &sasDarkBlue fontsize = &footnote_size];
	end;
quit;


/* Clear the library reference */
libname myFile clear;



/**********************/
/* CLOSE EXCEL OUTPUT */
/**********************/
ods excel close;