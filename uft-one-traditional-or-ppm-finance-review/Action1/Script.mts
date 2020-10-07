﻿'===========================================================
'20201007 - DJ: Initial creation
'===========================================================

'===========================================================
'Function to search for the PPM proposal in the appropriate status
'===========================================================
Function PPMProposalSearch (CurrentStatus, NextAction)
	'===========================================================================================
	'BP:  Click the Search menu item
	'===========================================================================================
	Browser("Search Requests").Page("Dashboard - IT Financial").Link("SEARCH").Click
	
	'===========================================================================================
	'BP:  Click the Requests text
	'===========================================================================================
	Browser("Search Requests").Page("Dashboard - IT Financial").Link("Requests").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter PFM - Proposal into the Request Type field
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").WebEdit("Request Type Field").Set "PFM - Proposal"
	Browser("Search Requests").Page("Search Requests").WebElement("Status Label").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter a status of "New" into the Status field
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").WebEdit("Status Field").Set CurrentStatus
	
	'===========================================================================================
	'BP:  Click the Search button (OCR not seeing text, use traditional OR)
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").Link("Search").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the first record returned in the search results
	'===========================================================================================
	DataTable.Value("dtFirstReqID") = Browser("Search Requests").Page("Request Search Results").WebTable("Req #").GetCellData(2,2)
	Browser("Search Requests").Page("Request Search Results").Link("First Record Request Link").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
End Function

Dim BrowserExecutable, Counter, mySendKeys

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon
Set AppContext2=Browser("CreationTime:=1")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Strategic Portfolio link
'===========================================================================================
Browser("Search Requests").Page("PPM Launch Page").Image("Strategic Portfolio Image").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Andy Stein (IT Financial Manager) link to log in as Andy Stein
'===========================================================================================
Browser("Search Requests").Page("Portfolio Management").WebArea("Andy Stein Image").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Search for propsals currently in a status of "Finance Review"
'===========================================================================================
PPMProposalSearch "Finance Review", "Approved"

'===========================================================================================
'BP:  Click the link for the Financial Summary
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Proposal Name Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Maximize the popup window
'===========================================================================================
AppContext2.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Add Costs link, use traditional OR as it isn't visible on the screen, but is on the page
'===========================================================================================
Browser("Create a Blank Staffing").Page("Financial Summary").Link("Add Costs").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the copy costs button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs").WebElement("Copy Costs Button").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Copy from Another Request text 
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").WebElement("Copy from Another Request").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Include Project radio button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").Frame("copyCostLinesFSSearchDialogIF").WebRadioGroup("Import Type Radio Button Group").Select "Project"

'===========================================================================================
'BP:  Type Web for One World into the Include Project text bos
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").Frame("copyCostLinesFSSearchDialogIF").WebEdit("Project Name Text Box").Set "Web for One World"

'===========================================================================================
'BP:  Click the Copy Cost Lines text to get the application to run the value entry validation
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").Frame("copyCostLinesFSSearchDialogIF").WebElement("Copy Cost Lines from Another").Click

'===========================================================================================
'BP:  Click the Add button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").Frame("copyCostLinesFSSearchDialogIF").WebButton("addButton").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Copy Forecast Values check box
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").Frame("copyCostLinesFSSearchDialogIF_2").WebCheckBox("copyForecast").Set "ON"

'===========================================================================================
'BP:  Click the Copy Copy button, detection improvement submitted.
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_2").Frame("CopyCostsDialog").WebButton("CopyButton").Click

'===========================================================================================
'BP:  Click the first 0.00 field and type 100
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_3").WebElement("0.000").Click
Window("Edit Costs").Type "100" @@ hightlight id_;_1771790_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Create a Blank Staffing").Page("Edit Costs_3").WebElement("Contractor").Click

'===========================================================================================
'BP:  Click the Done button, detection improvement submitted.
'===========================================================================================
Browser("Create a Blank Staffing").Page("Edit Costs_2").WebButton("Done").Click

'===========================================================================================
'BP:  Close the popup window
'===========================================================================================
AppContext2.Close																			'Close the application at the end of your script

'===========================================================================================
'BP:  Click the Save text
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebElement("Save").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
Browser("Search Requests").Page("Req Details").Link("Approved").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
Browser("Search Requests").Page("Req Details").Link("Sign Out").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

