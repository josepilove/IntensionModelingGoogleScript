//dev v1.6
// the intent of this dev (1.6) build experiment is to figure out why the script produces erratic results when multiple users are viewing the sheet.
//
/**

1) SUMMARY SCRIPTS
2) INITIALIZATION SCRIPTS
3) UTILITIES
4) VOTER & CONTINUUM UPDATES
5) VOTE SUMMARY & ANALYSIS PROCESSING
6) BUILD VISUALIZATIONS
7) EXPORT TO OMNIGRAFFLE

**/


/*Sheet Variables*/
var _spreadsheet = SpreadsheetApp.getActive();
var _sheet = _spreadsheet.getActiveSheet();
var _variablesSheetName = "Settings";
var _analysisSheetName = "Alignment Analysis";
var _vizTemplateName = "Visualization Template";
var _summarySheetName = "Summaries";
var _variSheet = _spreadsheet.getSheetByName(_variablesSheetName);
var _analySheet = _spreadsheet.getSheetByName(_analysisSheetName);
var _vizSheet = _spreadsheet.getSheetByName(_vizTemplateName);
var _sumSheet = _spreadsheet.getSheetByName(_summarySheetName);
var _todaySheet = _spreadsheet.getSheetByName('Vote Collection: Today');
var _futureSheet = _spreadsheet.getSheetByName('Vote Collection: Future');
var _firstColumn = 1; 
var _rowsPerVoter = 5;
var _headerRows = 8;
var _headerColumns = 6;
var _numberOfParticipantsCell = "B26";
var _numberOfContinuumsCell = "B17";
var _todayAverageRow = "3";
var _futureAverageRow = "4";
var _changeAverageRow = "5";
var _headerSheets = 7;
var _numberOfInputs = 1;
var continuumsRow = 19;
var continuumsArray = getFullContinuumsArray();
var continuumNamesArray = new Array();
var resultsArray = new Array();
var resultsObject;
var builtViz = false;
var _continuumVariableHeaderColumns = 1;
var _continuumVariableHeaderRows = 1;
var _vizHeaderColumns = 2;
var _vizHeaderRows = 2;
var _vizFooterRows = 3;
var _vizMaxChicklets = 2;
var _vizContinuumRow = _vizHeaderRows + (_vizMaxChicklets + 1);
var _vizBaseRows = 5;


/*Continuum Worksheet Output variables*/
var offset = 150; //space between models

/*Continuum Output Variables*/
var valuesArray = new Array();
var maxChicklets;
var idsOfAllChicklets = new Array();
var chickletWidth = 29.52;
var chickletHeight = 9.36;
var chickletMargin = 1.811;
var baseHeight = 20; 










/**
***********************************************
*                                             *
*   1)    S U M M A R Y   S C R I P T S       *
*                                             *
***********************************************

beginVoting()
refresh()
loadSettings()
summarizeVotes()
prepareVisualizations()
uniqueName()

**/










/**
* BEGIN VOTING
* Load Settings and go to the voting tab
*
*
*/
function beginVoting()
{
  Logger.log("beginVoting()");
  loadSettings();
  goToVoteCollection();
}


/**
* REFRESH
* Refreshes everything
*
*
*/
function refresh()
{
  Logger.log("refresh()");
  loadSettings( buildAllVisualizations() );
}



/**
* LOAD SETTINGS
* Reloads everything except the visualizations. 
*
*
*/
function loadSettings(callback)
{
  Logger.log("loadSettings()");
  syncParticipants(syncContinuums());
}



/**
* SUMMARIZE VOTES
* Builds output summary table and populates it with the current votes.
*
*
*/
function summarizeVotes(callback)
{
  Logger.log("summarizeVotes()");
  verifyOutputSummaryTable(1);
}



/**
* PREPARE VISUALIZATIONS
* Builds & populates summary table before creating visualizations
*
*
*/
function prepareVisualizations()
{
  Logger.log("prepareVisualizations()");
  summarizeVotes();
  buildAllVisualizations();
  _spreadsheet.getSheetByName(getContinuumName(1)).activate();
}



/**
* UNIQUE NAME
* Gets email and document title
*
*
*/
function uniqueName()
{
  var name = getUser() + " - " + getDocumentName();
  Logger.log(name);
  return name;
}









/**
***********************************************************
*                                                         *
*   2)   I N I T I A L I Z A T I O N   S C R I P T S      *
*                                                         *
***********************************************************

onOpen()
buildTriggers()
onCustomEdit()

**/










/**
* ON OPEN
* Builds the custom menus
*
*
*/
function onOpen()
{
SpreadsheetApp.getUi().createMenu("InTension")
.addItem('Open Help Sidebar', 'guidebar')
.addItem('Refresh All (if something seems weird)', 'refresh')
.addSeparator()
.addItem('Export Worksheets', 'exportWorksheetSidebar')
.addItem('Export Visualizations', 'exportOutputSidebar')
.addSeparator()
.addItem('Set/Update Workshop Size Variables', 'loadSettings')
.addSeparator()
.addItem('Build Visualizations', 'prepareVisualizations')
.addItem('Refresh All Visualizations', 'prepareVisualizations')
.addItem('Refresh This Visualization', 'refreshThisVisualization')
.addSeparator()
.addItem('Delete All Visualizations', 'deleteAllVisualizations')
.addItem('Delete & Reset All', 'resetEverything')
.addSubMenu(SpreadsheetApp.getUi()
            .createMenu("For Experimenting")
            .addItem("Generate Dummy Continuums", "addDummyContinuums")
            .addItem("Generate Dummy Today Votes", "addDummyTodayVotes")
            .addItem("Generate Dummy Future Votes", "addDummyFutureVotes")
           )
//.addSubMenu(SpreadsheetApp.getUi()
//            .createMenu("Developer Tools")
//            .addItem('Log Triggers', 'showTriggers')
//)
.addSeparator()
.addItem('About This Thing', 'aboutModal')
.addToUi();
}



/**
* GUIDEBAR
* Creates the sidebar to help guide the user
*
*
*/
function guidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Guide')
      .setTitle('InTension Help Sidebar')
      .setWidth(400);
      SpreadsheetApp.getUi()
      .showSidebar(html);
}  



/**
* BUILD TRIGGERS
*
*
*/
function buildTriggers()
{
  Logger.log("buildTriggers()"); 
  ScriptApp.newTrigger("onCustomEdit")
     .forSpreadsheet(_spreadsheet)
     .onEdit()
     .create();
}



/**
* ON EDIT
* 
*
*/
function onCustomEdit()
{
  Logger.log("onCustomEdit()");
  var activeSheetName = _sheet.getSheetName();
  var activeCellRange = _sheet.getActiveCell().getA1Notation();
  
  if (activeSheetName == _variSheet.getSheetName()){
    if( activeCellRange == _numberOfParticipantsCell ){
      syncParticipants();
    } else {
      Logger.log(activeCellRange+" !== "+ _numberOfParticipantsCell);
    }
    if( activeCellRange == _numberOfContinuumsCell){
      _variSheet.getRange("D14").setValue("Loading...");
      syncContinuums();
    } else {
      Logger.log(activeCellRange+" !== "+ _numberOfContinuums3Cell);
    }
  }
}










/**
*******************************************
*                                         *
*     3)      U T I L I T I E S           *
*                                         *
*******************************************

googleAnalytics()
getUser()
getDocumentName()
aboutModal()
listToMatrix()
getLastColumnInRow()
getNumberOfContinuums_()
getNumberOfVoters_()
columnToLetter()
letterToColumn()
convertToLiteral()
convertToIndex()
clearText()
goToAnalysis()
goToVoteCollection()
continuumNamesToArray()
getFullContinuumsArray()
getContinuumName()
resetEverything()
clearVotes()
guidGenerator()
findInRange()
setContinuumsValue()
reapplyFormulas()

**/









/**
* GOOGLE ANALYTICS
* Track Spreadsheet views with Google Analytics
*
* @param {string} gaaccount Google Analytics Account like UA-1234-56.
* @param {string} spreadsheet Name of the Google Spreadsheet.
* @param {string} sheetname Name of individual Google Sheet.
* @return The 1x1 tracking GIF image
* @customfunction
*/
function googleAnalytics(gaaccount, spreadsheet, sheetname) {
  
  /** 
  * Written by Amit Agarwal 
  * Web: www.ctrlq.org 
  * Email: amit@labnol.org 
  */
  
  var imageURL = [
    "https://ssl.google-analytics.com/collect?v=1&t=event",
    "&tid=" + gaaccount,
    "&cid=" + Utilities.getUuid(),
    "&z="   + Math.round(Date.now() / 1000).toString(),
    "&ec="  + encodeURIComponent("Intension Modeling v1.5"),
    "&ea="  + encodeURIComponent(spreadsheet || "Spreadsheet"),
    "&el="  + encodeURIComponent(sheetname || "Sheet")
  ].join("");
  
  return imageURL;
}



/**
* GET USER
* Get email address of owner
*
*
*/
function getUser(){
  return _spreadsheet.getOwner().getEmail();
}



/**
* GET DOCUMENT NAME
* Gets the name of the document
*
*
*/
function getDocumentName(){
  return _spreadsheet.getName();
}



/**
* ABOUT MODAL
* Description of what this is and isn't
*
*
*/
function aboutModal() {
  var text = "<div style='font-family:helvetica; font-size: 13px;'><p><strong>This is an experiment.</strong> It has no warranty, not guarantee, and no support. </p><p>It is based on the Intention Modeling work of Dan Klyn and has been created as an efficiency for conducting Intension Modeling Workshops at TUG. </p><p>We start all of our projects with this activity and think that all projects everywhere should start with a similar process. This is our first attempt at making a tool that will make this more accessible.</p><p>It has been shared with you as a trusted friend of TUG and fellow information architect so that you can play around with it and see if it can be useful in your own work.</p><p>If you have thoughts or feedback, you can email them to <a href='mailto:joe@understandinggroup.com' target='_blank'>Joe Elmendorf</a>.</p> <p><em style='font-size: 75%;'>We reserve all rights to this sheet and the underlying InTension Modeling principles.<br/> You can't distribute this on your own. You can't claim this as your own. You can't sell this.</em></div> ";
  var htmlOutput = HtmlService
      .createHtmlOutput(text)
      .setWidth(500)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About This Thing');
}


/**
* LIST TO MATRIX
* Takes a list and sub-divides it. Can also be used to take 
* a single-dimensional array and turn it into a two-
* dimensional array (elementsPerSubArray need to equal the 
* number of list items).
*
*
* @param list {Array} Array of items to sub-divide
* @param elementsPerSubArray {int} The number of elements that each sub array should have.
* @return {array}
* @customfunction
*/
function listToMatrix(list, elementsPerSubArray) {
    Logger.log("listToMatrix("+list+","+elementsPerSubArray+")");
    var matrix = [], i, k;

    for (i = 0, k = -1; i < list.length; i++) {
        if (i % elementsPerSubArray === 0) {
            k++;
            matrix[k] = [];
        }

        matrix[k].push(list[i]);
    }

    return matrix;
}



/**
* GET LAST COLUMN IN ROW
* Looks in a single row and returns the number of the last 
* column number of the last cell that has content.
*
*
* @param row {int} The row you want to search.
* @param sheet {Sheet Object} The sheet that has the row.
* @return {int} The number of the last column in the specified row with content.
* @customfunction
*/
function getLastColumnInRow(row, sheet)
{
  var finalColumn = sheet.getMaxColumns();
  var targetRange = sheet.getRange("A"+row+":"+columnToLetter(finalColumn)+row);
  var values = targetRange.getValues();
  var output = "";
  //Logger.log("finalColumn:"+finalColumn);
  //Logger.log("targetRange:"+targetRange);

  for (var i=0; i < values[0].length; i++){
    
    if(values[0][i] !== 'undefined' && values[0][i] !== null && values[0][i] !== ""){
      
    } else {
      output = (i+1);
    }
  }

  if(output !== ""){
    return output;
  } else {
    return finalColumn;
  }
}  



/**
* INTERNAL - GET NUMBER OF CONTINUUMS
* Gets the number of continuums AS BUILT in the analysis sheet.
* 
*
* @return {int}
* @customfunction
*/
function getNumberOfContinuums_()
{
	var totalColumns = _analySheet.getMaxColumns();
  	var continuums = (totalColumns - _headerColumns);
  	return continuums;
}



/**
* INTERNAL - GET NUMBER OF VOTERS
* Gets the number of voters AS BUILT in the analysis sheet.
*
*
* @return {int}
* @customfunction
*/
function getNumberOfVoters_()
{
  	var totalRows = _analySheet.getMaxRows();
  	var voters = (totalRows - (_headerRows + 1))/_rowsPerVoter;
  	return voters;
}



/**
* COLUMN TO LETTER
* Convert the number of a column to a letter for use in A1 notation
*
*
* @param {1,2,3} column The number that correlates with the column you are targetting. 
* @return {String}      The letter for use in A1 notation. 
* @customfunction
*/
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}



/**
* LETTER TO COLUMN
* Convert the letter of a column to a number
*
*
* @param {A,B,C} letter The letter that correlates with the column you are targetting. 
* @return {String}      The number of that column from the left origin. 
* @customfunction
*/
function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}



/**
* CONVERT TO LITERAL
* Converts a passed integer to a continuum literal position: 5-0-5
* 0 => -5; ... 5 => 0 ... 10 => 5;
*
* @param {int} int Continuum index value between 0 and 10.
* @return {int} Continuum literal value between -5 and 5.
*/
function convertToLiteral(int)
{
  if( int <= 10 && int >= 0){
    return (int-5);
  } else {
    return "Error: number out of bounds; First variable needs to be less than or equal to 10.";
  }
}



/**
* CONVERT TO INDEX
* Converts a passed integer to a continuum index: 0-10
* -5 => 0; ... 0 => 5 ... 5 => 10;
*
* @param {int} int Continuum literal value between -5 and 5.
* @return {int} Continuum index value between 0 and 10.
*/
function convertToIndex(int)
{
  if( int <= 5 && int >= -5){
    return (int+5);
  } else {
    return "Error: number out of bounds; First variable needs to be: -5 <= n >= 5";
  }
}



/**
* CLEAR TEXT
* Clears the text for the passed range.
*
* @param {string} context A valid Range string.
*/
function clearText(context)
{
  _variSheet.getRange(context).clearContent();
}



/**
* GO TO ANALYSIS
* 
*
*/
function goToAnalysis()
{
  _AnalySheet.activate();
}


/**
* GO TO VOTE COLLECTION
* 
*
*/
function goToVoteCollection()
{
  _todaySheet.activate();
}



/**
* CONTINUUM NAMES TO ARRAY
* Puts the names of the continuums into an array.
*
*
* @customfunction
*/
function continuumNamesToArray(callback)
{
  Logger.log("continuumNamesToArray()");
  continuumNamesArray = _analySheet.getRange( columnToLetter( (_headerColumns + 1) ) + "2:" + columnToLetter((_analySheet.getMaxColumns())) + "2").getDisplayValues();
}



/**
* GET FULL CONTINUUMS ARRAY
*
*
*/
function getFullContinuumsArray()
{
  var ca = new Array();
  var continuums = getNumberOfContinuums_();
  
  for(var i=1; continuums >= i; i++){
    var setArray = new Array();
    var column = columnToLetter(i+_continuumVariableHeaderColumns);
    setArray = _variSheet.getRange( column+(continuumsRow+1) + ":" + column + (continuumsRow+4) ).getDisplayValues();
    ca[(i-1)] = setArray;
  }
  return ca;
}



/**
* GET CONTINUUM NAME
* Gets the name of the continuum from the AnalySheet.
* 
*
* @param {int} continuumID The ID of the target continuum.
* @return {String} The name of the continuum.
*/
function getContinuumName(continuumID)
{
  Logger.log("getContinuumName("+continuumID+")");
  var continuumA1 = columnToLetter( (_headerColumns + continuumID)) + 2;
  var continuumName = _analySheet.getRange( continuumA1 ).getValue(); 
  return continuumName;
}



/**
* RESET EVERYTHING
*
*
*/
function resetEverything()
{
  Logger.log("resetEverything()");
  _variSheet.getRange(_numberOfParticipantsCell).setValue("1");
  _variSheet.getRange(_numberOfContinuumsCell).setValue("1");

  loadSettings(clearVotes(buildOutputSummaryTable(populateOutputSummaryTable())));
  _variSheet.getRange("B20").setValue("Example: Service");
  _variSheet.getRange("B21").setValue("Example Description");
  _variSheet.getRange("B22").setValue("Example: Acquire");
  _variSheet.getRange("B23").setValue("Example Description");
  summarizeVotes(deleteAllVisualizations(function(){_variSheet.activate();}));
}



/**
* CLEAR VOTES
* Clears all votes on the two voting sheets.
*
*
*/
function clearVotes(callback)
{
  Logger.log("clearVotes()");
  _todaySheet.getRange("B2:"+columnToLetter(_todaySheet.getMaxColumns())+_todaySheet.getMaxRows()).clear({ formatOnly: false, contentsOnly: true });
  _futureSheet.getRange("C2:"+columnToLetter(_futureSheet.getMaxColumns())+_futureSheet.getMaxRows()).clear({ formatOnly: false, contentsOnly: true });
}



/**
* GUID GENERATOR
* Generates a random, and hopefully, unique string.
*
*
*/
function guidGenerator() {
    var S4 = function() {
       return (((1+Math.random())*0x10000)|0).toString(16).substring(1);
    };
    return "a"+(S4()+S4());
}



/**
* FIND IN RANGE
* Finds the row containing the provided search term
*
* @param {string} needle String to find in haystack
* @param {Range} haystack Sheet Range to search in
* @param {0/1} opt_returnMultiple 0: (default) returns first instance; 1: returns array of all instances
*
* @return {int} The either first or all rows where needle is found in haystack.
**/
function findInRange(needle, haystack, opt_returnMultiple)
{
  Logger.log("findInRange('"+needle+"', "+haystack.getA1Notation()+")");
  var output = new Array();
  if (opt_returnMultiple == null) {opt_returnMultiple = 0;}
  var data = haystack.getValues();
  for (var i=0; i<data.length; i++){
    if(data[i] == needle){
      Logger.log("row = "+(i+1));
      output.push((i+1));
    }
  }
  if(opt_returnMultiple == 1){
    return output;
  } else {
    return output[0];
  }
}




/**
* SET PARTICIPANTS VALUE
* Sets the value of the participant variable on VariSheet based 
* on the form in the sidebar (currently hidden).
*
* 
* @param {int} value The value to set the variable
* @param {1,2} opt_update 0 (default): Don't run sync; 1: Do run sync.
**/
function setParticipantsValue(value, opt_update)
{
  if(opt_update == null){opt_update = 0;}
  _variSheet.getRange(_numberOfParticipantsCell).setValue(value); 
  if( opt_update == 1){
    syncParticipants();
  }
}



/**
* SET CONTINUUMS VALUE
* Sets the value of the continuums variable on VariSheet based 
* on the form in the sidebar (currently hidden).
*
* 
* @param {int} value The value to set the variable
* @param {1,2} opt_update 0 (default): Don't run sync; 1: Do run sync.
**/
function setContinuumsValue(value, opt_update)
{
  if(opt_update == null){opt_update = 0;}
  _variSheet.getRange(_numberOfContinuumsCell).setValue(value); 
  if( opt_update == 1){
    syncContinuums();
  }
}



/**
* REAPPLY FORMULAS
* Function stores and resets all the cell formulas for the worksheet.
* This function can be called as a failsafe if something is acting weird.
*
*
**/
function reapplyFormulas(name)
{
  var formula;
  var range;
  var sheet;
  
  if(name == "vote-collection-continuum-names"){
    formula = '=concat(concat((column()-2), ": "), concat( concat(INDIRECT(concat("Settings!", concat(columnToLetter(column()-1), 20))), " <-> "), INDIRECT(concat("Settings!", concat(columnToLetter(column()-1), 22)))))';
    range = 'C1';
    _todaySheet.getRange(range).setFormula(formula);
    _futureSheet.getRange(range).setFormula(formula);
  }
  if(name == "vote-collection-voter-names"){
    formula = "='Vote Collection: Today'!B2";
    range = "B2";
    _futureSheet.getRange(range).setFormula(formula);
  }
  if(name == "analysis-continuum-names"){
    formula = '=CONCATENATE(Settings!B20," <-> ",Settings!B22)';
    range = "G2";
    _analySheet.getRange(range).setFormula(formula);
  }
  if(name == "analysis-vote-today"){
    formula = '=if(ISBLANK(INDIRECT(concat(concat(Settings!$C$11,"!"), concat(columnToLetter(column()-4),($B10+1))))),,(INDIRECT(concat(concat(Settings!$C$11,"!"),concat(columnToLetter(column()-4),($B10+1))))-5))';
    range = "G10";    
    _analySheet.getRange(range).setFormula(formula);
  }
  if(name == "analysis-vote-future"){
    formula = '=if(ISBLANK(INDIRECT(concat(concat(Settings!$C$12,"!"), concat(columnToLetter(column()-4),($B10+1))))),,(INDIRECT(concat(concat(Settings!$C$12,"!"),concat(columnToLetter(column()-4),($B10+1))))-5))';
    range = "G10";    
    _analySheet.getRange(range).setFormula(formula);
  }
  if(name == "analysis-voter-name"){
    formula = '=INDIRECT(concat(concat(Settings!$C$11,"!"),concat(columnToLetter(column()-1), ($B10+1))))';
    range = 'C10';
    _analySheet.getRange(range).setFormula(formula);
  }
  if(name == "analysis-voter-misalignments"){
    formula = '=if( isblank(G13),,countif(13:13, ">=6") )';
    range = 'A10';
    _analySheet.getRange(range).setFormula(formula);
  }

}












/**
****************************************************************
*                                                              *
*   4)   V O T E R   &   C O N T I N U U M   U P D A T E S     *
*                                                              *
****************************************************************

-------

   *
V  *  addVotersAtOnce()
O  *  addVoterTallieRows()
T  * 
E  *  removeVoter()
R  *  removeVoterTallieRows()
S  *  
   *  syncParticipants()
   *
   
--------

C  *  
O  *  addContinuumsAtOnce()
N  *  addContinuumTallieColumns()
T  *  
I  *  
N  *  removeContinuum()
U  *  removeContinuumTallieColumns()
U  *  
M  *  buildContinuumNameInputs()
S  *  syncContinuums()

-------


**/










/**
* ADD VOTERS AT ONCE
* Adds additional voter rows to the AnalySheet
*
*
* @param {int} numberOfVotersToAdd The number of voters that should be added to the Analysis sheet.
* @customfunction
*/
function addVotersAtOnce(numberOfVotersToAdd)
{
  Logger.log("addVotersAtOnce("+numberOfVotersToAdd+")");
  var lastVoterRow = _analySheet.getMaxRows();
  var voterTemplate = _analySheet.getRange(
    (lastVoterRow - _rowsPerVoter),
    _firstColumn,
    _rowsPerVoter,
    _analySheet.getMaxColumns()
  );

  _analySheet.insertRowsAfter(
	  lastVoterRow, 
	  (_rowsPerVoter * numberOfVotersToAdd)
  );    
  addVoterTallieRows(numberOfVotersToAdd);
  
  var sourceAndNew = _analySheet.getRange(
    (lastVoterRow - _rowsPerVoter),
    _firstColumn,
    ( (_rowsPerVoter * 2) + (_rowsPerVoter * numberOfVotersToAdd) ),
    _analySheet.getMaxColumns()
  );
  
  voterTemplate.autoFill(
    sourceAndNew,
    SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
  );  
}



/**
* ADD VOTER TALLIE ROWS
* Adds row in both voter tallie sheets ("Today", "Future").
*
*
*/
function addVoterTallieRows(numberOfRowsToAdd)
{
  Logger.log("addVoterTallieRows("+numberOfRowsToAdd+")");
  _todaySheet.insertRowsAfter(_todaySheet.getMaxRows(), numberOfRowsToAdd);
  var template = "B"+_futureSheet.getMaxRows();

  _futureSheet.insertRowsAfter(_futureSheet.getMaxRows(), numberOfRowsToAdd);
  
  _futureSheet.getRange(template).autoFill(
    _futureSheet.getRange(template+":"+"B"+_futureSheet.getMaxRows()),
    SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
  );
}



/**
* REMOVE VOTER
* Removes rows from the AnalySheet one-by-one. 
*
*
*/
function removeVoter()
{
	Logger.log("removeVoter()");
	if( (_analySheet.getMaxRows() - _rowsPerVoter) > (_headerRows + 1)){
		_analySheet.deleteRows(
			(_analySheet.getMaxRows() - _rowsPerVoter), 
			_rowsPerVoter
		);
	} else {
		Browser.msgBox("Can't delete the last voter. Need to keep at least one voter.");
	}
}



/**
* REMOVE VOTER TALLIE ROWS
* Removes un-needed rows from both voter tallie sheets ("Today", "Future").
*
*
*/
function removeVoterTallieRows(totalParticipants)
{
  Logger.log("removeVoterTallieRows("+totalParticipants+")");
  var existingVoters = (_todaySheet.getMaxRows()-1);
  var numberToDelete = existingVoters - totalParticipants;
  
  if(numberToDelete > 0){
    _todaySheet.deleteRows((totalParticipants+1), numberToDelete);
    _futureSheet.deleteRows((totalParticipants+1), numberToDelete);
  }
}


/**
* SYNC PARTICIPANTS
* Figures out whether to add or remove voter participants 
* from the AnalySheet based on the configuration on the VariSheet.
*
*
*/
function syncParticipants(callback)
{
  Logger.log("syncParticipants()");	
  var numberOfParticipants = _variSheet.getRange( _numberOfParticipantsCell ).getValue();
  	var i = getNumberOfVoters_();
	
  	if( i > numberOfParticipants)
  	{
      removeVoterTallieRows(numberOfParticipants);
      while (i > numberOfParticipants)
    	{
      	  removeVoter();
      	i--;
    	}
  	} else {
    	var numberOfVotersToAdd = numberOfParticipants - i;
    
    	if( numberOfVotersToAdd > 0 )
    	{
      	  addVotersAtOnce(numberOfVotersToAdd);
    	}
  	}
  	
	buildContinuumAverages();
}



/**
* ADD CONTINUUMS AT ONCE
* Adds continuum columns to the AnalySheet
*
*
* @param {int} totalContinuums The total of continuums that should be on the Analysis sheet.
* @customfunction
*/
function addContinuumsAtOnce(totalContinuums)
{
  var firstContRange = columnToLetter(_headerColumns+1)+"1:"+columnToLetter(_headerColumns+1)+_analySheet.getMaxRows();
  var continuumTemplate = _analySheet.getRange(firstContRange);
  var numberOfExistingContinuums = (_analySheet.getMaxColumns() - _headerColumns);
  var additionalContinuums = (totalContinuums - numberOfExistingContinuums)
  if(additionalContinuums > 0)
  {
    _analySheet.insertColumnsAfter((_headerColumns+1), additionalContinuums);
    addContinuumTallieColumns(additionalContinuums);
  }
  var sourceAndNew = _analySheet.getRange(columnToLetter(_headerColumns+1)+"1:"+columnToLetter((totalContinuums+_headerColumns))+_analySheet.getMaxRows());
  
  continuumTemplate.autoFill(sourceAndNew, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
}



/**
* ADD CONTINUUM TALLIE COLUMNS
* Adds columns in both voter tallie sheets ("Today", "Future").
*
*
*/
function addContinuumTallieColumns(numberOfColumnsToAdd)
{
  Logger.log("addContinuumTallieColumns("+numberOfColumnsToAdd+")");
  _todaySheet.insertColumnsAfter(_todaySheet.getMaxColumns(), numberOfColumnsToAdd);
  _futureSheet.insertColumnsAfter(_futureSheet.getMaxColumns(), numberOfColumnsToAdd);
  var templateRange = "C1";
  _todaySheet.getRange(templateRange).autoFill(_todaySheet.getRange(templateRange+":"+columnToLetter(_todaySheet.getMaxColumns())+"1"),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  _futureSheet.getRange(templateRange).autoFill(_futureSheet.getRange(templateRange+":"+columnToLetter(_futureSheet.getMaxColumns())+"1"),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}



/**
* REMOVE CONTINUUM
* Removes columns from the AnalySheet one-by-one.
*
*
*/
function removeContinuum()
{
  Logger.log("removeContinuum()");
  	if( _analySheet.getMaxColumns() > (_headerColumns + 1) ){
	  	_analySheet.deleteColumn(
	    	_analySheet.getMaxColumns()
	  	);
        removeContinuumTallieColumns();
  	} else {
  		Browser.msgBox("Can't delete the last continuum. Need to keep at least one continuum.");
  	}
}



/**
* REMOVE CONTINUUM TALLIE COLUMNS
* Removes un-needed columns from both voter tallie sheets ("Today", "Future").
*
*
*/
function removeContinuumTallieColumns(totalContinuums)
{
  Logger.log("removeContinuumTallieColumns("+totalContinuums+")");
  var existingContinuums = (_todaySheet.getMaxColumns()-2);
  var numberToDelete = existingContinuums - totalContinuums;
  
  if(numberToDelete > 0){
    _todaySheet.deleteColumns((totalContinuums+2), numberToDelete);
    _futureSheet.deleteColumns((totalContinuums+2), numberToDelete);
  }
}

/**
* BUILD CONTINUUM NAME INPUTS
* Creates a frame for which continuum values can be added on the VariSheet.
*
*
* @param {int} numberOfContinuums The total number of continuums in the model.
*/
function buildContinuumNameInputs(numberOfContinuums, callback)
{
  Logger.log("buildContinuumNameInputs("+numberOfContinuums+")");  
  _numberOfInputs = (getLastColumnInRow(continuumsRow, _variSheet) - _continuumVariableHeaderColumns);  
  var template = _variSheet.getRange(columnToLetter((_continuumVariableHeaderColumns + 1)) + continuumsRow);  //B19
  var lastContinuumLetter = columnToLetter(numberOfContinuums + _continuumVariableHeaderColumns); 
  var lastColumn = _variSheet.getMaxColumns();  //7
  var minColumns = 4; //4
  var minContinuums = minColumns - _continuumVariableHeaderColumns; //3
  var existingContinuums = _numberOfInputs;  
  var existingColumns = lastColumn;
  var neededColumns = numberOfContinuums + _continuumVariableHeaderColumns;
  
  
  
  if (numberOfContinuums !== existingContinuums){
    Logger.log("numberOfContinuums["+numberOfContinuums+"] !== existingContinuums["+existingContinuums+"]");
    if (neededColumns > minColumns){
      Logger.log('neededColumns['+neededColumns+'] > minColumns['+minColumns+']');
      if(numberOfContinuums > existingContinuums){
        Logger.log('numberOfContinuums['+numberOfContinuums+'] > existingContinuums['+existingContinuums+']');
        if(existingContinuums < minContinuums){
          Logger.log('existingContinuums['+existingContinuums+'] < minContinuums['+minContinuums+']');
          Logger.log("Add "+Math.abs((existingContinuums - numberOfContinuums))+" columns after:"+columnToLetter( ((_numberOfInputs + _continuumVariableHeaderColumns) +1)));
          _variSheet.insertColumnsAfter(
            ((_numberOfInputs + _continuumVariableHeaderColumns) +1), 
            Math.abs(existingContinuums - numberOfContinuums)
          );
        } else { 
          Logger.log("Add "+Math.abs((existingContinuums - numberOfContinuums))+" columns after:"+columnToLetter(lastColumn));
          _variSheet.insertColumnsAfter(
            lastColumn, 
            Math.abs((existingContinuums - numberOfContinuums))
          );
        }
      } else {
        Logger.log("delete "+(existingContinuums - ( numberOfContinuums + _continuumVariableHeaderColumns))+ " columns starting with "+columnToLetter(numberOfContinuums + _continuumVariableHeaderColumns));
        
        _variSheet.deleteColumns(
          ((numberOfContinuums + _continuumVariableHeaderColumns)+1), 
          ((existingContinuums - ( numberOfContinuums + _continuumVariableHeaderColumns))+1)
        );
      }
    } else if((lastColumn - minColumns) !== 0){

      Logger.log("remove all columns but the minimum");
      _variSheet.deleteColumns(minColumns, (lastColumn - minColumns));
      
    }
  } else {
    Logger.log('All set, no need to adjust the number of continuums');
  }

  lastContinuumLetter = columnToLetter(neededColumns);
  
  var sourceAndNew = _variSheet.getRange(columnToLetter((1 + _continuumVariableHeaderColumns)) + continuumsRow + ":" + lastContinuumLetter + continuumsRow);
  template.autoFill(sourceAndNew, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  _variSheet.getRange(
    (columnToLetter((_continuumVariableHeaderColumns + 1)))+(continuumsRow+1)+":"+(columnToLetter((_continuumVariableHeaderColumns + 1)))+(continuumsRow+5))
    .copyFormatToRange(
      _variSheet, 
      (1 + _continuumVariableHeaderColumns), 
      (numberOfContinuums + _continuumVariableHeaderColumns), 
      (continuumsRow + 1), 
      (continuumsRow + 5)
    );
  
  if (neededColumns < minColumns){
    _variSheet.getRange( columnToLetter((numberOfContinuums + (1 + _continuumVariableHeaderColumns))) + continuumsRow + ":" + columnToLetter(_variSheet.getMaxColumns()) + (continuumsRow + 5)).setBackground("#d9d9d9").clearContent();
  }
  
}



/**
* SYNC CONTINUUMS
* Figures out whether to add or remove continuums from the 
* AnalySheet based on the configuration on the VariSheet.
*
*
*/
function syncContinuums(callback)
{
  Logger.log("syncContinuums()");
  var numberOfContinuums = _variSheet.getRange( _numberOfContinuumsCell ).getValue();
  var numberOfParticipants = _variSheet.getRange( _numberOfParticipantsCell ).getValue();

  continuumNamesToArray();
  buildContinuumAverages();
  buildContinuumNameInputs(numberOfContinuums, clearText('D14'));


  var i = getNumberOfContinuums_();

  if ( i > numberOfContinuums)
  { 
    removeContinuumTallieColumns(numberOfContinuums);
    while (i > numberOfContinuums)
    {
      removeContinuum();
      i--;
    }
  } else {
    var numberOfContinuumsToAdd = numberOfContinuums - i;
    //if ( numberOfContinuumsToAdd > 0 )
    //{  
      addContinuumsAtOnce(numberOfContinuums);
    //}
  }
}










/**
***********************************************************************************
*                                                                                 * 
*   5)   V O T E   S U M M A R Y   &   A N A L Y S I S    P R O C E S S I N G     *
*                                                                                 *
***********************************************************************************

buildContinuumAverages()
buildContinuumAverage()
verifyOutputSummaryTable()
buildOutputSummaryTable()
addSummaryRows()
deleteSummaryRows()
populateOutputSummaryTable()
addAveragesToVoteArray()

**/










/**
* BUILD CONTINUUM AVERAGES
* Initiates the average-creating for each of the 4 types of averages for each continuum.
*
*
*/
function buildContinuumAverages()
{
  Logger.log("buildContinuumAverages()");
  var i = 1;
  	while ( i <= 4 )
	{
    	buildContinuumAverage(i);
    	i++;
  	}
}  



/**
* BUILD CONTINUUM AVERAGE
* Calculates the formula needed for the AnalySheet continuum voting averages 
* and applies that formula across all continuum's average cells.
*
*
* @param {1,2,3,4} type 1:Today Average; 2:Future Average; 3:Change Average; 4:Alignment Average;
*/
function buildContinuumAverage(type)
{
	Logger.log("buildContinuumAverage("+type+")");
    var numberOfParticipants = getNumberOfVoters_();
	var continuumColumn = (_headerColumns + 1);		
	var fixedRow = (parseInt(type) + (_todayAverageRow - 1));			
	var averageA1 = columnToLetter(continuumColumn) + fixedRow;						
	var averageCell = _analySheet.getRange( averageA1 );				
	var numberOfContinuums = getNumberOfContinuums_();
	var summary = "=average(";
	var voter = 1;
	var c = 1;
	while ( c <= numberOfContinuums )
	{
		continuumColumn = columnToLetter( _headerColumns + c );
		while ( voter <= numberOfParticipants )
		{
			var row = _headerRows + (((voter*_rowsPerVoter) - 4) + type);		
			summary += continuumColumn+row;
			voter++;
			if ( voter <= numberOfParticipants )
            {
				summary += ",";
			}
		}
		c++;
	}
	summary += ")";
	averageCell.setFormula( summary );              
	var sourceAndNew = _analySheet.getRange(averageA1+":"+columnToLetter( _analySheet.getMaxColumns() ) + fixedRow);	
	averageCell.autoFill(                               
		sourceAndNew,                                     
		SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES      
	); 
}



/**
* VERIFY OUTPUT SUMMARY TABLE
* Quick check to see if the summary table is built.
*
*
**/
function verifyOutputSummaryTable(populate){
  if(populate == null || populate == "" || populate == "undefined"){
    populate = 0;
  }
  var summaryHeaderRows = 2;
  var numberOfContinuums =  _variSheet.getRange( _numberOfContinuumsCell ).getValue();
  var maxRows = _sumSheet.getMaxRows();
  if ((numberOfContinuums + summaryHeaderRows) !== maxRows){
    if(populate == 1){
      Logger.log("Output Summary Table needs rebuilding. Building and Populating.");
      buildOutputSummaryTable(populateOutputSummaryTable());
    } else if (populate == 0){
      Logger.log("Output Summary Table needs rebuilding. Building.");
      buildOutputSummaryTable();
    }
  } else {
    if(populate == 1){
      Logger.log("Output Summary Table ready. Populating...");
      populateOutputSummaryTable();
    } else {
      Logger.log("Output Summary Table ready. No build needed. ");
    }
  }
}


/**
* BUILD OUTPUT SUMMARY TABLE
* Creates an empty summary table on SumSheet based on the number of 
* continuums from the VariSheet.
*
*
*/
function buildOutputSummaryTable(callback)
{
  Logger.log("buildOutputSummaryTable()");
  var summaryHeaderRows = 2;
  var summaryRowsPer = 1;
  var numberOfContinuums =  _variSheet.getRange( _numberOfContinuumsCell ).getValue();
  var maxColumns = _sumSheet.getMaxColumns();
  
  var summaryTemplate = _sumSheet.getRange(                //The range of the last-formatted summary
    (summaryHeaderRows+summaryRowsPer),                    //Row
    _firstColumn,                                          //Column
    summaryRowsPer,                                        //number of Rows
    maxColumns                                             //number of Columns 
  );

  var lastRowHere = _sumSheet.getMaxRows();
  var numberOfRowsForContinuums = (summaryRowsPer * numberOfContinuums);
  var numberOfNewRows = ((numberOfRowsForContinuums + summaryHeaderRows) - lastRowHere);

  if (numberOfNewRows > 0){
    
    addSummaryRows(numberOfNewRows, lastRowHere);
    summaryTemplate.copyFormatToRange(
      _sumSheet, 
      _firstColumn, 
      maxColumns, 
      (summaryHeaderRows+1), 
      (numberOfRowsForContinuums + summaryRowsPer)
    );
  } else if (numberOfNewRows < 0){
    deleteSummaryRows(Math.abs(numberOfNewRows), (numberOfRowsForContinuums + summaryHeaderRows));
  }
}



/**
* ADD SUMMARY ROWS
* Adds rows to the SumSheet.
*
*
* @param {int} numberOfRowsToAdd The number of rows to add to the SumSheet.
* @param {int} lastRowHere The last row on the current SumSheet.
*/
function addSummaryRows(numberOfRowsToAdd, lastRowHere)
{
  Logger.log("addSummaryRows("+numberOfRowsToAdd+", "+lastRowHere+")");
  _sumSheet.insertRowsAfter(
    lastRowHere, 
    numberOfRowsToAdd
  );
}



/**
* DELETE SUMMARY ROWS
* Removes rows from the SumSheet.
*
*
* @param {int} numberOfRowsToDelete The number of rows to delete from SumSheet.
* @param {int} lastNeededRow The number of the last row needed on the SumSheet. 
*/
function deleteSummaryRows(numberOfRowsToDelete, lastNeededRow)
{
  Logger.log("deleleSummaryRows("+numberOfRowsToDelete+","+lastNeededRow+")");
  _sumSheet.deleteRows(
    lastNeededRow+1, 
    numberOfRowsToDelete
  );
}



/**
* POPULATE OUTPUT SUMMARY TABLE
* Translates the vote data from the AnalySheet into a summarized form on the SumSheet.
*
*
*/
function populateOutputSummaryTable()
{
  Logger.log("populateOutputSummaryTable()");
  var numberOfContinuums = getNumberOfContinuums_();
  for (var c=0; c < numberOfContinuums; c++){
    populateContinuumSummary(c);
  }

}



/**
* POPULATE CONTINUUM SUMMARY
* Translates the vote data from the AnalySheet into a summarized form on the SumSheet for a single continuum.
*
* @param {int} c The ID of the continuum to build a summary for
*/
function populateContinuumSummary(c)
{  
  continuumNamesToArray();
  var continuums = continuumNamesArray[0];
  var summaryHeaderRows = 2;
  var continuumID = (c+1);
  var contColumnLetter = columnToLetter( ((_headerColumns + c) + 1) );   //"G"
  var participantTodayVotes = new Array();
  var participantFutureVotes = new Array();
  var summaryTodayVotesByChicklet = new Array();
  var summaryFutureVotesByChicklet = new Array();
  for (var v=0; v < getNumberOfVoters_(); v++){
    var voterID = (v+1);
    var voterTodayRow = ( (_headerRows-3) + ( voterID * _rowsPerVoter ));
    var voterFutureRow = ( (_headerRows-2) + ( voterID * _rowsPerVoter ));
    var todayVote = _analySheet.getRange(contColumnLetter+voterTodayRow).getValue();
    var futureVote = _analySheet.getRange(contColumnLetter+voterFutureRow).getValue();
    if ( todayVote !== "" )
    {
      participantTodayVotes.push(todayVote);
    }
    if (futureVote !== "" )
    {
      participantFutureVotes.push(futureVote);
    }
  }    
  for(var chickletNumber = 0; chickletNumber <= 10; chickletNumber++){
    var numberOfTodayVotes = 0;
    var numberOfFutureVotes = 0;
    if( participantTodayVotes.length > 0 ){
      for( var j=0; j<participantTodayVotes.length; j++ ){
        if(participantTodayVotes[j] == (chickletNumber - 5)){
          numberOfTodayVotes++; 
        }
        summaryTodayVotesByChicklet[chickletNumber] = numberOfTodayVotes;
      }
    } else {
      summaryTodayVotesByChicklet[chickletNumber] = 0;
    }
    if( participantFutureVotes.length > 0){
      for( var k=0; k<participantFutureVotes.length; k++ ){
        if(participantFutureVotes[k] == (chickletNumber - 5)){
          numberOfFutureVotes++; 
        }
        summaryFutureVotesByChicklet[chickletNumber] = numberOfFutureVotes;
      }
    } else {
      summaryFutureVotesByChicklet[chickletNumber] = 0;
    }
  }
  var continuumsNamesRange = "A"+(summaryHeaderRows+continuumID);    
  var todaySummaryRange = "B"+(summaryHeaderRows+continuumID)+":L"+(summaryHeaderRows+continuumID);    
  var futureSummaryRange = "N"+(summaryHeaderRows+continuumID)+":X"+(summaryHeaderRows+continuumID);
  /*Continuum Names*/  _sumSheet.getRange(continuumsNamesRange).setValue(continuums[c]);
  /*Today Summary*/    _sumSheet.getRange(todaySummaryRange).setValues(listToMatrix(summaryTodayVotesByChicklet,11));
  /*Future Summary*/   _sumSheet.getRange(futureSummaryRange).setValues(listToMatrix(summaryFutureVotesByChicklet,11));
  valuesArray[c] = new Array();
  Logger.log("valuesArray["+c+"] = "+valuesArray[c]);
  valuesArray[c][0] = summaryTodayVotesByChicklet;
  valuesArray[c][1] = summaryFutureVotesByChicklet;
  addAveragesToVoteArray();
  Logger.log("Populated Summary Table for Continuum "+ (c+1));
}



/**
* ADD AVERAGES TO VOTE ARRAY
* Gets averages from the AnalySheet and adds them to valuesArray.
*
**/
function addAveragesToVoteArray()
{
  continuumNamesToArray();
  var continuums = continuumNamesArray[0];
  for (var c=0; c < continuums.length; c++){
    var continuumID = (c+1);
    var averageTodayVote = _analySheet.getRange( columnToLetter(continuumID+_headerColumns)+3 ).getValue();
    var averageFutureVote = _analySheet.getRange( columnToLetter(continuumID+_headerColumns)+4 ).getValue();
    if(valuesArray[c] == undefined){
      valuesArray[c] = new Array();
    }
    valuesArray[c][2] = new Array();
    valuesArray[c][2][0] = averageTodayVote; 
    valuesArray[c][2][1] = averageFutureVote;
  }
}










/**
******************************************************
*                                                    *
*   6)   B U I L D   V I S U A L I Z A T I O N S     *
*                                                    *
******************************************************

buildAllVisualizations()
buildVisualizationSheet()
refreshThisVisualization()
setSheetRows()
addSheetRows()
populateVisualizationSheet()
deleteAllVisualizations()
deleteVisualizationSheet()
goToNextVisualizationSheet()
buildFirstVisualiztaion()

**/










/**
* BUILD ALL VISUALIZATIONS
* Initiates visualizations if the continuum names are all filled-in. 
*
*
*/
function buildAllVisualizations()
{
  Logger.log("buildAllVisualizations()");
  if(getLastColumnInRow(continuumsRow, _variSheet) == getLastColumnInRow((continuumsRow+1), _variSheet) ){
    var i = 1;
    while( i <= getNumberOfContinuums_())
    {
      buildVisualizationSheet(i);
      i++;
    }
    builtViz = true;
  } else {
    _spreadsheet.toast("Fill out continuums before building visualizations");
    _variSheet.activate();
  }
}



/**
* BUILD VISUALIZATION SHEET
* Builds a new sheet for each continuum that does not already
* have a sheet.
* 
*
* @param {int} continuum The id of the target continuum
*/
function buildVisualizationSheet(continuum)
{
  Logger.log("buildVisualizationSheet("+continuum+")");
  var orientationCell = "C2"; 
  var templateSheet = _spreadsheet.getSheetByName(_vizTemplateName);
  var lastSheet = _spreadsheet.getNumSheets();
  var continuumName = getContinuumName(continuum); 
  
  if(_spreadsheet.getSheetByName(continuumName) == null)    //if sheet doesn't already exist, add it. 
  {
    _spreadsheet.insertSheet(continuumName, (lastSheet + 1),{template: templateSheet});
    var thisSheet = _spreadsheet.getSheetByName(continuumName);
    if( thisSheet.isSheetHidden() )
    {
      thisSheet.showSheet();
    }
    var continuumLetter = columnToLetter((continuum + _headerColumns));
    thisSheet.getRange(orientationCell).setValue(continuumLetter);
    thisSheet.getRange("B1").setValue(continuumName);
    setSheetRows(thisSheet, continuum, 0);
    populateVisualizationSheet(continuum);

  } else {
    setSheetRows(_spreadsheet.getSheetByName(continuumName), continuum, 1);
    populateVisualizationSheet(continuum);

  }
}



/**
* REFRESH THIS VISUALIZATION
* Updates the currently selected visualization sheet
*
*
**/
function refreshThisVisualization()
{
  var currentSheet = _spreadsheet.getActiveSheet();

  if(
    currentSheet.getName() !== _analySheet.getName() && 
    currentSheet.getName() !== _variSheet.getName() && 
    currentSheet.getName() !== _todaySheet.getName() && 
    currentSheet.getName() !== _futureSheet.getName() && 
    currentSheet.getName() !== _sumSheet.getName() && 
    currentSheet.getName() !== _vizSheet.getName()
    )
  {
    verifyOutputSummaryTable(0);
    var continuum = _analySheet.getRange(_sheet.getRange("C2").getValue()+"1").getValue();
    populateContinuumSummary(continuum-1)
    buildVisualizationSheet(continuum);
  } else {
    Browser.msgBox("This sheet is not a visualization. Select the visualization tab you want to refresh."); 
  }
}



/**
* SET SHEET ROWS
* Updates the visualization rows based on the max number of chicklets
*
*
* @param {sheet} thisSheet The visualization sheet for the target continuum
* @param {int} continuum The number of the target continuum
* @param {0/1} opt_redo 0: just adds rows; 1: DEFAULT: first deletes all non-required rows before adding new ones;
**/
function setSheetRows(thisSheet, continuum, opt_redo)
{
  if(opt_redo == null || opt_redo == "" || opt_redo == "undefined"){
    opt_redo = 1;
  }
  
  Logger.log("setSheetRows("+thisSheet+","+continuum+","+opt_redo+")");

  var baseRow = findInRange('x', thisSheet.getRange("A:A"));
  var numberOfTopRowsToDelete = (baseRow - 3) - _vizHeaderRows;
  var numberOfBottomRowsToDelete = (((thisSheet.getMaxRows() - _vizHeaderRows) - baseRow) - _vizFooterRows) ;
  
  
  populateContinuumSummary(continuum-1);
  _vizMaxTodayChicklets = highestNumberOfChicklets(valuesArray[(continuum-1)], 'today');  
  _vizMaxFutureChicklets = highestNumberOfChicklets(valuesArray[(continuum-1)], 'future');
  
  var todayRows = (baseRow - (_vizHeaderRows + 1)); // (5 - (2 + 1)) = 2
  var futureRows = (thisSheet.getMaxRows() - (_vizFooterRows + baseRow)); //(10 - (3 + 5) = 2
  
  Logger.log("max today: "+_vizMaxTodayChicklets);
  Logger.log("current today: "+todayRows);
  Logger.log("max future: "+_vizMaxFutureChicklets);
  Logger.log("current future: "+futureRows);

  //instead of approaching the full set of rows as a "delete" or "add" process, approach top and bottom separate and calculate add/delete for each as separate processes
  
  
  if( ( ( (_vizMaxTodayChicklets + _vizMaxFutureChicklets) + (_vizFooterRows + _vizHeaderRows) ) + 1) > thisSheet.getMaxRows() ){
    Logger.log("Need to add more Rows");
    var topRows = ((baseRow + _vizHeaderRows + 3) - _vizMaxTodayChicklets);
    Logger.log("topRows: "+topRows);
    var bottomRows = ( thisSheet.getMaxRows() - ( ( (_vizHeaderRows + _vizFooterRows)  + baseRow ) + 3) );
    Logger.log("bottomRows: "+bottomRows);
    addSheetRows(thisSheet, baseRow, topRows, bottomRows);
  } else if( ( ( (_vizMaxTodayChicklets + _vizMaxFutureChicklets) + (_vizFooterRows + _vizHeaderRows) +1 )  ) < thisSheet.getMaxRows() ){
    Logger.log("Need to delete Rows");
  } else if( ( ( (_vizMaxTodayChicklets + _vizMaxFutureChicklets) + (_vizFooterRows + _vizHeaderRows)+1 )  ) == thisSheet.getMaxRows() ){
    Logger.log("No need to futz with rows");
  }
  
  //addSheetRows(thisSheet, continuum);

}


function deleteSheetRows(thisSheet, numberOfTopRowsToDelete, numberOfBottomRowsToDelete)
{
  if(numberOfTopRowsToDelete > 0) {
    thisSheet.deleteRows((_vizHeaderRows + 1), numberOfTopRowsToDelete);
  }
  
  if(numberOfBottomRowsToDelete > 0) {
    thisSheet.deleteRows(((_vizHeaderRows + 1) + _vizBaseRows), numberOfBottomRowsToDelete);
  }
}



/**
* ADD SHEET ROWS
* Adds rows to the target continuum visualization sheet
*
*
* @param {sheet} thisShett The visualization sheet for the target continuum
* @param {int} continuum The number of the target continuum
**/
function addSheetRows(thisSheet, baseRow, topRows, bottomRows)
{
  Logger.log("addSheetRows("+thisSheet.getName()+", "+baseRow+", "+topRows+", "+bottomRows+")");
  
  if(topRows > 1){
    thisSheet.insertRowsBefore(
      (_vizHeaderRows+1),
      (topRows - 1)
    );
    
    var topTemplateA1 = "A" + ((_vizHeaderRows + topRows)) + ":O" + ((_vizHeaderRows + topRows));
    var topTemplate = thisSheet.getRange(topTemplateA1);
    
    topTemplate.autoFill(
      thisSheet.getRange(
        (_vizHeaderRows + 1),
        1,
        (topRows),
        thisSheet.getMaxColumns()
      ),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  }
  
  if(bottomRows > 1){
    thisSheet.insertRowsAfter(
      (thisSheet.getMaxRows() - _vizFooterRows),
      (bottomRows)
    );
    
    var bottomTemplate = thisSheet.getRange("A" + (thisSheet.getMaxRows() - _vizFooterRows) + ":O" + (thisSheet.getMaxRows() - _vizFooterRows));
    
    bottomTemplate.autoFill(
      thisSheet.getRange(
        (thisSheet.getMaxRows() - _vizFooterRows),
        1,
        (baseRow + bottomRows),
        thisSheet.getMaxColumns()
      ),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  }
}



/**
* POPULATE VISUALIZATION SHEET
* Adds data to the visualization sheets from the SumSheet.
*
*
* @param {int} continuum The id of the target continuum
*/
function populateVisualizationSheet(continuum)
{
  Logger.log("populateVisualizationSheet("+continuum+")");
  var thisSheet = _spreadsheet.getSheetByName(getContinuumName(continuum));
  thisSheet.activate();
  var todayRange = thisSheet.getRange((_spreadsheet.getActiveSheet().getMaxRows() - 1), 4, 1, (_spreadsheet.getActiveSheet().getMaxColumns()-4));
  var futureRange = thisSheet.getRange((_spreadsheet.getActiveSheet().getMaxRows()), 4, 1, (_spreadsheet.getActiveSheet().getMaxColumns()-4));
  var todaySummary = _sumSheet.getRange((2+continuum),2,1,11).getValues();
  var futureSummary = _sumSheet.getRange((2+continuum),14,1,11).getValues();
  
  var todayAvg = _analySheet.getRange(columnToLetter((_headerColumns+continuum))+"3").getValue();
  var futureAvg = _analySheet.getRange(columnToLetter((_headerColumns+continuum))+"4").getValue();
  var todayAvgRange = thisSheet.getRange((_spreadsheet.getActiveSheet().getMaxRows() - 1), (_spreadsheet.getActiveSheet().getMaxColumns()));
  var futureAvgRange = thisSheet.getRange(_spreadsheet.getActiveSheet().getMaxRows(), _spreadsheet.getActiveSheet().getMaxColumns());
  
  todayRange.setValues(todaySummary);
  todayAvgRange.setValue(todayAvg);
  futureRange.setValues(futureSummary);
  futureAvgRange.setValue(futureAvg);
  _spreadsheet.toast(thisSheet.getName()+" visualization ready.");
}



/**
* DELETE ALL VISUALIZATIONS
* Deletes all visualizations sheets.
*
*
*/
function deleteAllVisualizations(callback)
{
  Logger.log("deleteAllVisualizations()");
  deleteVisualizationSheet("all");
}



/**
* DELETE VISUALIZATION SHEET
* Deletes visualization sheet by continuum name. 
* Pass "all" to delete all visualization sheets.
*
*
* @param {String} continuumName The name of the contiuum.
*/
function deleteVisualizationSheet(continuumName)
{
  Logger.log("deleteVisualizationSheet("+continuumName+")");
  if (continuumName !== "all"){
    _spreadsheet.deleteSheet(_spreadsheet.getSheetByName(continuumName));
  } else {
    var sheetArray = _spreadsheet.getSheets();
    for (var s=0; s<sheetArray.length; s++){
      var sheetName = sheetArray[s].getName();
      if ( sheetName.indexOf("<->") > -1 ){
        _spreadsheet.deleteSheet(_spreadsheet.getSheetByName(sheetArray[s].getName()));
      }
    }
  }
}



/**
* GO TO NEXT VISUALIZATION SHEET
* !!!!! N E E D S   W O R K !!!!!
*
*/
function goToNextVisualizationSheet()
{
  var sheetsArray = _spreadsheet.getSheets();  
  
  var continuums = continuumsArray;
  
  var numberOfContinuums = continuums.length;
    
  var thisContinuumLetter = _sheet.getRange("c2").getValue();
  
  var thisContinuumID = (letterToColumn(thisContinuumLetter)-6);
  
  if( thisContinuumID < numberOfContinuums ){
    buildVisualizationSheet(thisContinuumID + 1);
  } else {
    _spreadsheet.toast("There are no more continuums.");
  }
  
}



/**
* BUILD FIRST VISUALIZATION
* 
*
*/
function buildFirstVisualization()
{
  //verifyOutputSummaryTable(1);
  buildVisualizationSheet(1);
}






/**
***************************************
*                                     *
*      D U M M Y   T H I N G S        *
*                                     *
***************************************

addDummyContinuums()
addDummyTodayVotes()
addDummyFutureVotes()
getRandomName()
getRandomVote()
getRandomWordPair()

**/










/**
* ADD DUMMY CONTINUUMS
* 
*
*/
function addDummyContinuums()
{  
  var numberOfContinuums = getNumberOfContinuums_();
  
  for (var i=1; i <= numberOfContinuums; i++)
  {
    var today = getRandomWordPair();
    var future = getRandomWordPair();
    _variSheet.getRange(columnToLetter(i + _continuumVariableHeaderColumns)+20).setValue(today);
    _variSheet.getRange(columnToLetter(i + _continuumVariableHeaderColumns)+22).setValue(future);
  }
}



/**
* ADD DUMMY TODAY VOTES
* 
*
*/
function addDummyTodayVotes()
{
  var numberOfVoters = getNumberOfVoters_();
  for (var i=1; i<=numberOfVoters; i++){
    _todaySheet.getRange("B"+(i+1)).setValue(getRandomName());
    for(var c=0; c<getNumberOfContinuums_();c++)
    {
      _todaySheet.getRange(columnToLetter(3+c)+(i+1)).setValue(getRandomVote());
    }
  }
  _todaySheet.activate();
}



/**
* ADD DUMMY FUTURE VOTES
* 
*
*/
function addDummyFutureVotes()
{
  var numberOfVoters = getNumberOfVoters_();
  for (var i=1; i<=numberOfVoters; i++){
    for(var c=0; c<getNumberOfContinuums_();c++)
    {
      _futureSheet.getRange(columnToLetter(3+c)+(i+1)).setValue(getRandomVote());
    }
  }
  _futureSheet.activate();
}



/**
* GET RANDOM NAME
* 
*
*/
function getRandomName()
{
  names = ["Keala", "Cahaya", "Yolotli", "Nsia", "Chao", "Lindsay", "Emery", "Hunter", "Ekene", "Kamalani", "Amahle", "Chinweuba", "Yannic", "Andy", "Dana", "Onyekachi", "Chandra", "Makoto", "Kapua"];
  modifiers = ["-Sue", "-Jo", "-Ann", "", " Jr."," VI", " III", " IX", " IV", " VII", "", "-Jo"];
  
  name = names[Math.floor(Math.random()*names.length)];
  name += modifiers[Math.floor(Math.random()*modifiers.length)];
  
  return name;
}



/**
* GET RANDOM VOTE
* 
*
*/
function getRandomVote()
{
  votes = [0,1,2,3,4,5,6,7,8,9,10];
  return votes[Math.floor(Math.random()*votes.length)];
}



/**
* GET RANDOM WORD PAIR
* 
*
*/
function getRandomWordPair()
{
  nouns = ["river", "things", "touch", "birth", "servant", "shelf", "level", "card", "rake", "sugar", "board", "cough", "dolls", "oil", "box", "memory", "health", "shame", "woman", "owl", "cats", "party", "reward", "support", "volcano", "bridge", "scale", "adjustment", "bat", "mint", "fold", "lead", "jellyfish", "spy", "motion", "behavior", "night", "caption", "toothbrush", "advice", "cars", "bubble", "celery", "queen", "brain", "industry", "trains", "wren", "surprise", "stranger"];
  adjectives = ["selective", "minor", "capricious", "equal", "successful", "pink", "energetic", "noisy", "didactic", "unable", "paltry", "smart", "whole", "disgusting", "meek", "gainful", "broken", "thirsty", "plausible", "tasteless", "unhealthy", "filthy", "busy", "glossy", "scary", "puzzled", "wiry", "high", "quizzical", "complete", "endurable", "incandescent", "worthless", "grateful", "serious", "tacky", "giant", "fearful", "melted", "familiar", "naive", "slow", "measly", "rough", "holistic", "superficial", "white", "ratty", "infamous", "colossal"];
  verbs = ["shed", "waste", "creep", "fly", "fling", "tree", "tire", "scream", "cost", "tempt", "do", "implore", "saunter", "put", "step", "tap", "compel", "forbid", "roar", "pull", "converge", "co-operate", "sanctify", "sort", "sack", "lick", "protect", "install", "allow", "compel", "kiss", "scab", "bash", "inflame", "paint", "scare", "awake", "love", "remain", "catch", "withdraw", "vie", "believe", "sterilize", "touch", "knock", "learn", "sting", "classify", "consist"];
  types = ["verb", "noun"];
  
  var type = types[Math.floor(Math.random()*types.length)];
  
  if(type == "verb"){
    var verb = verbs[Math.floor(Math.random()*verbs.length)];
    var adjective = adjectives[Math.floor(Math.random()*adjectives.length)];
    return adjective+" "+verb;
  } else if (type == "noun"){
    var noun = nouns[Math.floor(Math.random()*nouns.length)];
    var adjective = adjectives[Math.floor(Math.random()*adjectives.length)];
    return adjective+" "+noun;
  }  
}










/**
*******************************************************
*                                                     *
*   7)   E X P O R T   T O   O M N I G R A F F L E    *
*                                                     *
*******************************************************

escapeAndRunInOG()
exportWorksheetSidebar()
exportOutputSidebar()

**/










/**
* ESCAPE AND RUN IN OG
* URL escapes the code that builds the omnigraffle exports and prepends the 
* header needed to open in OmniGraffle.
*
* @param {string} string The script to be opened in omnigraffle
* @return {string} targetURL Contains properly formatted string to be opened in OG. 
*/
function escapeAndRunInOG(script){
   var scriptCode = encodeURIComponent(script);           
   var openOG = "omnigraffle:///omnijs-run?script="
   var canvas = 'var canvas = document.windows[0].selection.canvas;';
   var targetURL = openOG + encodeURIComponent(canvas) + scriptCode;
   return targetURL;
}



/**
* EXPORT WORKSHEET SIDEBAR
* Creates the sidebar for the extra goodies
*
*
*/
function exportWorksheetSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Worksheet')
      .setTitle('Generate Worksheets')
      .setWidth(400);
      SpreadsheetApp.getUi()
      .showSidebar(html);
}



/**
* EXPORT OUTPUT SIDEBAR
* Creates the sidebar for the extra goodies
*
*
*/
function exportOutputSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Output')
      .setTitle('Generate Output')
      .setWidth(400);
      SpreadsheetApp.getUi()
      .showSidebar(html);
}  