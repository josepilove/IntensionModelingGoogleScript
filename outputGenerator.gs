//Template v1.5
/**
*******************************************************************
*                                                                 *
*  C R E A T E   V O T E   R E S U L T S   O U T P U T   C O D E  *
*                                                                 *
*******************************************************************

generateModelOutput()
highestNumberOfChicklets()
createChickletsFromSummary()
chicklet()
outputContinuum()


**/



/**
* MAKE ADEQUATE ROOM
* Make sure that the offset is sufficient based on the max number of chicklets. 
*
**/
function makeAdequateRoom()
{
  offset = 150;
  if ( (2*(maxChicklets * (chickletHeight + chickletMargin))) > offset){
    offset = ((2*(maxChicklets * (chickletHeight + chickletMargin))) + 30);
  } else {
    Logger.log(offset+" > "+(2*(maxChicklets * (chickletHeight + chickletMargin))));
  }
}



/**
* GENERATE MODEL OUTPUT
* Initiates the process of building the export code for OmniGraffle.
*
**/
function generateModelOutput()
{
  continuumsArray = getFullContinuumsArray();
  populateOutputSummaryTable();
  var fullModel = '';
  for(var c=0; continuumsArray.length > c; c++){
    var continuumID = (c+1);
    idsOfAllChicklets[c] = new Array();    
    maxChicklets = highestNumberOfChicklets(valuesArray[c]);
    makeAdequateRoom();
    var base = outputContinuum(continuumID, continuumsArray[c][0], continuumsArray[c][1], continuumsArray[c][2], continuumsArray[c][3]);
    var parsedText = valuesArray[c];
    var output = createChickletsFromSummary(parsedText, continuumID);
    fullModel += base;
    fullModel += output;
  }
  return escapeAndRunInOG(fullModel);
}



/**
* HIGHEST NUMBER OF CHICKLETS
* Figures out the tallest stack of chicklets across all continuums
*
*
* @param {Array} chickSummary Array of vote values
* @param {string} opt_which DEFAULT: "overall", "today", "future"
* @return {int} number of the tallest stack of chicklets
**/
function highestNumberOfChicklets(chickSummary, opt_which){
  if(opt_which == null){ opt_which = "overall";}
  Logger.log("highestNumberOfChicklets("+chickSummary+", "+opt_which+")");
  var current = chickSummary[0];
  var future = chickSummary[1];
  var largestCurrent = Math.max.apply(null, current);
  var largestFuture = Math.max.apply(null, future);
  var largestOverall = Math.max.apply(null, [largestCurrent, largestFuture]);
  
  if(opt_which == "overall"){
	return largestOverall;
  } else if(opt_which == "today"){
    return largestCurrent;
  } else if(opt_which == "future"){
    return largestFuture;
  } else {
    return largestOverall;
  }
}



/**
* CREATE CHICKLETS FROM SUMMARY
* Builds a chicklets for each vote
*
*
* @param {Array} chickSummary Array of vote values
* @param {int} continuumID ID of the continuum where these chicklets should be placed. 
* @return {string} Javascript formatted for generating a set of chicklets for a continuum in Omnigraffle 
**/
function createChickletsFromSummary(chickSummary, continuumID){
	var continuumIndex = (continuumID-1);
    var currentSummary = chickSummary[0];
	var futureSummary = chickSummary[1];

    var currentAverage = Math.round(chickSummary[2][0]);
    var futureAverage = Math.round(chickSummary[2][1]);
	
	var output = "";
	var max = maxChicklets;
    for (var cs=0; currentSummary.length > cs; cs++){
      for(var i=0; max>i; i++){
        if(currentSummary[cs] > i){
          output += chicklet(continuumID, 'current', cs, i);
        } else {
          output += chicklet(continuumID, 'currentBlank', cs, i);
        }
      }
    }
  
    for (var fs=0; futureSummary.length > fs; fs++){
      for(var i=0; max>i; i++){
        if(futureSummary[fs] > i){
          output += chicklet(continuumID, 'future', fs, i);
        } else {
          output += chicklet(continuumID, 'futureBlank', fs, i);
        }
      }
    }
  
    output += chicklet(continuumID, 'todayAverage', convertToIndex(currentAverage), 1);
    output += chicklet(continuumID, 'futureAverage', convertToIndex(futureAverage), 1);
  
	var chickletGroup = guidGenerator();
	chickletGroupId = chickletGroup;
	output += 'var '+chickletGroup+' = new Group([';
	
    for(var i=0; idsOfAllChicklets[continuumIndex].length > i; i++){
      
      if( (idsOfAllChicklets[continuumIndex].length-1) > i){
        output += idsOfAllChicklets[continuumIndex][i]+", ";
      } else {
        output += idsOfAllChicklets[continuumIndex][i];
      }
    }
	
	output += ']);';
	return output;	
}



/**
* CHICKLET
* Build a single chicklet.
*
*
* @param {int} continuumID Cid of the continuum where this chicklet should be positioned
* @param {string} type Visual type of a chicklet "current", "currentBlank", "future", "futureBlank", "todayAverage", "futureAverage"
* @param {int} xcolumn A number 0-10 indicating which column
* @param {int} yrow A number 0-highestNumberOfChicklets indicating the row( from the base-line )
* @return {string} Javascript formatted for generating a single chicklets for a continuum in Omnigraffle 
*/
function chicklet(continuumID ,type, xcolumn, yrow){
	var g1 = guidGenerator();
	var sideMargin = chickletMargin;
	var leftOffset = 244;
	var width = chickletWidth;
    var continuumIndex = (continuumID-1);
	
	var height = chickletHeight;
	
	var color;
	var stroke;
	var mod;
	var bh;
	var fillType;
	var strokeType;
	if( type == 'future' ){
		color = 'Color.RGB(0.278857, 0.63459, 0.94468, 0.297652)';
		stroke = 'Color.RGB(0.0, 0.5597, 1.0)';
		fillType = 'FillType.Solid';
		strokeType = 'StrokeType.Single';
		mod = 1;
		bh = baseHeight;
	} else if( type == 'current' ){
		color = 'Color.RGB(0.4717, 0.769585, 0.272121, 0.300138)'; 
		stroke = 'Color.RGB(0.355022, 0.82873, 0.0)';
		fillType = 'FillType.Solid';
		strokeType = 'StrokeType.Single';
		mod = -1; //will be used to flip the directions the chicklets build from. 
		bh = 0;
	} else if( type == 'currentBlank'){
		color = 'Color.RGB(0.4717, 0.769585, 0.272121, 0.300138)'; 
		stroke = 'Color.RGB(0.355022, 0.82873, 0.0)';
		fillType = 'null';
		strokeType = 'null';
		mod = -1;
		bh = 0;
	} else if( type == 'futureBlank'){
		color = 'Color.RGB(0.278857, 0.63459, 0.94468, 0.297652)';
		stroke = 'Color.RGB(0.0, 0.5597, 1.0)';
		fillType = 'null';
		strokeType = 'null';
		mod = 1;
		bh = baseHeight;
    } else if( type == 'todayAverage'){
        color = 'Color.RGB(0.4717, 0.769585, 0.272121, 0.300138)'; 
        stroke = 'Color.RGB(0.191224053797468, 0.522549900520651, 7.83883646463155e-05)';
        fillType = 'FillType.Solid';
		strokeType = 'StrokeType.Single';
        bh = (baseHeight/2);
        mod = -1;
    } else if( type == 'futureAverage'){
		color = 'Color.RGB(0.278857, 0.63459, 0.94468, 0.297652)';
		stroke = 'Color.RGB(0.0162022773736569, 0.282415463538999, 0.818936204663212)';
		fillType = 'FillType.Solid';
		strokeType = 'StrokeType.Single';
        bh = 0 - (baseHeight/2);
        mod = 1;
    }
	
	var topMargin = chickletMargin*mod;
	var baseAndChickletHeight = bh+height;
	var x = leftOffset + (parseInt( xcolumn ) * (width + sideMargin));
	var y = (offset*continuumID) - 2.5 + baseAndChickletHeight + ((parseInt( yrow )*mod) * (height + (topMargin*mod))); //base offset from line + (row number * (height + top/bottom margin) )
  
	var chicklet = 'var '+g1+' = canvas.newShape();';
	chicklet += g1+'.flippedVertically = true;';
   
    if(type == 'futureAverage' || type == 'todayAverage'){
      width = 17;
      height = 17;
      sideMargin = ((chickletWidth - width) + sideMargin);
      x = (leftOffset + 6) + (parseInt( xcolumn ) * (width + sideMargin)); //base offset from left + (column number * (width + right margin) )
      y = (offset*continuumID)+13;
      chicklet += g1+'.shape = "Circle";';
      chicklet += g1+'.geometry = new Rect(' +x+ ', ' +y+ ', ' +width+ ', ' +height+ ');';
      chicklet += g1+'.strokeThickness = 2.0;';
    } else {
      chicklet += g1+'.geometry = new Rect(' +x+ ', ' +y+ ', ' +width+ ', ' +height+ ');';
      chicklet += g1+'.strokeThickness = 0.5;';
    }
	chicklet += g1+'.shadowColor = null;';
	chicklet += g1+'.strokeColor = ' +stroke+ ';';
	chicklet += g1+'.fillColor = ' +color+ ';';
	chicklet += g1+'.flippedHorizontally = true;';
	chicklet += g1+'.fillType = '+fillType+';';
	chicklet += g1+'.strokeType ='+strokeType+';';
	
	idsOfAllChicklets[continuumIndex].push(g1);
	return chicklet;
}



/**
* OUTPUT CONTINUUM
* Builds an summary-styled continuum base and formats it for OmniGraffle
* v.2 - updating origination point & re-grouping the parts to make more sense.
*
*
* @param {int} Cid The desired number to show at the left of the continuum (usually the continuumID)
* @param {string} Aterm The bold text on the left of the continuum
* @param {string} Adesc The sub-text on the left of the continuum
* @param {string} Bterm The bold text on the right of the continuum
* @param {string} Bdesc The sub-text on the right of the continuum
* @return {string} Javascript formatted for generating continuum in Omnigraffle 
**/
function outputContinuum(Cid, Aterm, Adesc, Bterm, Bdesc){
	var g1 = guidGenerator();
	var g2 = guidGenerator();
	var g3 = guidGenerator();
	var g4 = guidGenerator();
	var g5 = guidGenerator();
	var g6 = guidGenerator();
	var g7 = guidGenerator();
	var g8 = guidGenerator();
	var g9 = guidGenerator();
	var g10 = guidGenerator();
	var g11 = guidGenerator();
	var g12 = guidGenerator();
	var g13 = guidGenerator();
	var g14 = guidGenerator();
	var g15 = guidGenerator();
	var g16 = guidGenerator();
	var g17 = guidGenerator();
	var g18 = guidGenerator();
	var g19 = guidGenerator();
	var g20 = guidGenerator();
	var g21 = guidGenerator();
	var g22 = guidGenerator();

	var code = 'var '+g1+' = canvas.newShape();';
	code += g1+'.fillType = null;';
	code += g1+'.text = "'+Bdesc+'";';
	code += g1+'.textVerticalPadding = 4.718570232391357;';
	code += g1+'.textSize = 9.381;';
	code += g1+'.textHorizontalAlignment = HorizontalTextAlignment.Left;';
	code += g1+'.strokeType = null;';
	code += g1+'.geometry = new Rect(594.79, 27.98, 222.25, 32.44);';
	code += g1+'.shadowColor = null;';
	code += g1+'.strokeColor = null;';
	code += g1+'.textVerticalPlacement = VerticalTextPlacement.Top;';
	code += g1+'.fillColor = null;';
	code += g1+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
	code += g1+'.strokeThickness = 0.9437140811153025;';
	code += g1+'.textHorizontalPadding = 4.718570232391357;';
	code += g1+'.flippedHorizontally = true;';

	code += 'var '+g2+' = canvas.newShape();';
	code += g2+'.fillType = null;';
	code += g2+'.text = "'+Bterm+'";';
	code += g2+'.textVerticalPadding = 0;';
	code += g2+'.textSize = 12.212;';
	code += g2+'.textHorizontalAlignment = HorizontalTextAlignment.Left;';
	code += g2+'.strokeType = null;';
	code += g2+'.geometry = new Rect(594.79, 0.00, 222.25, 16.00);';
	code += g2+'.shadowColor = null;';
	code += g2+'.strokeColor = null;';
	code += g2+'.textVerticalPlacement = VerticalTextPlacement.Bottom;';
	code += g2+'.fillColor = null;';
	code += g2+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
	code += g2+'.strokeThickness = 0.9437140811153025;';
	code += g2+'.textHorizontalPadding = 4.718570232391357;';
	code += g2+'.fontName = "HelveticaNeue-Bold";';
	code += g2+'.flippedHorizontally = true;';

	code += 'var '+g3+' = canvas.newShape();';
	code += g3+'.fillType = null;';
	code += g3+'.text = "'+Adesc+'";';
	code += g3+'.textVerticalPadding = 4.718570232391357;';
	code += g3+'.textSize = 9.381;';
	code += g3+'.textHorizontalAlignment = HorizontalTextAlignment.Right;';
	code += g3+'.strokeType = null;';
	code += g3+'.geometry = new Rect(21.29, 27.48, 222.25, 32.44);';
	code += g3+'.shadowColor = null;';
	code += g3+'.strokeColor = null;';
	code += g3+'.textVerticalPlacement = VerticalTextPlacement.Top;';
	code += g3+'.fillColor = null;';
	code += g3+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
	code += g3+'.strokeThickness = 0.9437140811153025;';
	code += g3+'.flippedHorizontally = true;';

	code += 'var '+g4+' = canvas.newShape();';
	code += g4+'.fillType = null;';
	code += g4+'.text = "'+Aterm+'";';
	code += g4+'.textVerticalPadding = 0;';
	code += g4+'.textSize = 12.212;';
	code += g4+'.textHorizontalAlignment = HorizontalTextAlignment.Right;';
	code += g4+'.strokeType = null;';
	code += g4+'.geometry = new Rect(21.29, 0.00, 222.25, 16.00);';
	code += g4+'.shadowColor = null;';
	code += g4+'.strokeColor = null;';
	code += g4+'.textVerticalPlacement = VerticalTextPlacement.Bottom;';
	code += g4+'.fillColor = null;';
	code += g4+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
	code += g4+'.strokeThickness = 0.9437140811153025;';
	code += g4+'.fontName = "HelveticaNeue-Bold";';
	code += g4+'.flippedHorizontally = true;';

	code += 'var '+g5+' = canvas.newShape();';
	code += g5+'.fillType = null;';
	code += g5+'.text = "'+Cid+'";';
	code += g5+'.textVerticalPadding = 0;';
	code += g5+'.autosizing = TextAutosizing.Full;';
	code += g5+'.textHorizontalAlignment = HorizontalTextAlignment.Left;';
	code += g5+'.strokeType = null;';
	code += g5+'.geometry = new Rect(0.00, 14.22, 91.00, 20.00);';
	code += g5+'.shadowColor = null;';
	code += g5+'.strokeColor = null;';
	code += g5+'.fillColor = null;';
	code += g5+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
	code += g5+'.fontName = "HelveticaNeue-Bold";';
	code += g5+'.textWraps = false;';

	code += 'var '+g6+' = canvas.newShape();';
	code += g6+'.strokeType = null;';
	code += g6+'.geometry = new Rect(582.90, 18.70, 7.54, 6.08);';
	code += g6+'.shadowColor = null;';
	code += g6+'.shapeVertices = [new Point(582.90, 24.78), new Point(590.44, 24.78), new Point(590.44, 18.70), new Point(582.90, 18.70)];';
	code += g6+'.strokeColor = null;';
	code += g6+'.fillColor = Color.RGB(0.70613, 0.706147, 0.706138);';
	code += g6+'.textColor = Color.RGB(0.475635, 0.475647, 0.47564);';
	code += g6+'.strokeThickness = 0.6166579729272318;';
	code += g6+'.name = "B-Square";';
	code += g6+'.flippedHorizontally = true;';

	code += 'var '+g7+' = canvas.newShape();';
	code += g7+'.strokeType = null;';
	code += g7+'.geometry = new Rect(240.44, 20.71, 350.00, 2.07);';
	code += g7+'.shadowColor = null;';
	code += g7+'.shapeVertices = [new Point(240.44, 22.78), new Point(590.44, 22.78), new Point(590.44, 20.71), new Point(240.44, 20.71)];';
	code += g7+'.strokeColor = null;';
	code += g7+'.fillColor = Color.RGB(0.70613, 0.706147, 0.706138);';
	code += g7+'.textColor = Color.RGB(0.475635, 0.475647, 0.47564);';
	code += g7+'.strokeThickness = 0.6166579729272318;';
	code += g7+'.name = "Line";';
	code += g7+'.flippedHorizontally = true;';

	code += 'var '+g8+' = canvas.newShape();';
	code += g8+'.strokeType = null;';
	code += g8+'.geometry = new Rect(240.45, 18.70, 7.54, 6.08);';
	code += g8+'.shadowColor = null;';
	code += g8+'.shapeVertices = [new Point(240.45, 24.78), new Point(247.99, 24.78), new Point(247.99, 18.70), new Point(240.45, 18.70)];';
	code += g8+'.strokeColor = null;';
	code += g8+'.fillColor = Color.RGB(0.70613, 0.706147, 0.706138);';
	code += g8+'.textColor = Color.RGB(0.475635, 0.475647, 0.47564);';
	code += g8+'.strokeThickness = 0.6166579729272318;';
	code += g8+'.name = "A-Square";';
	code += g8+'.flippedHorizontally = true;';

	code += 'var '+g9+' = new Group(['+g8+', '+g7+', '+g6+']);';
	code += g9+'.geometry = new Rect(240.44, 18.70, 350.00, 6.08);';

	code += 'var '+g10+' = canvas.newShape();';
	code += g10+'.text = "5";';
	code += g10+'.textVerticalPadding = 0;';
	code += g10+'.textSize = 7.721;';
	code += g10+'.strokeType = null;';
	code += g10+'.geometry = new Rect(567.08, 19.58, 8.39, 4.32);';
	code += g10+'.shadowColor = null;';
	code += g10+'.strokeColor = null;';
	code += g10+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g10+'.strokeThickness = 0.6166579729272318;';
	code += g10+'.textHorizontalPadding = 0;';
	code += g10+'.name = "5";';
	code += g10+'.fontName = "HelveticaNeue-Light";';
	code += g10+'.textWraps = false;';
	code += g10+'.flippedHorizontally = true;';

	code += 'var '+g11+' = canvas.newShape();';
	code += g11+'.text = "4";';
	code += g11+'.textVerticalPadding = 0;';
	code += g11+'.textSize = 7.721;';
	code += g11+'.strokeType = null;';
	code += g11+'.geometry = new Rect(535.64, 19.58, 8.39, 4.32);';
	code += g11+'.shadowColor = null;';
	code += g11+'.strokeColor = null;';
	code += g11+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g11+'.strokeThickness = 0.6166579729272318;';
	code += g11+'.textHorizontalPadding = 0;';
	code += g11+'.name = "4";';
	code += g11+'.fontName = "HelveticaNeue-Light";';
	code += g11+'.textWraps = false;';
	code += g11+'.flippedHorizontally = true;';

	code += 'var '+g12+' = canvas.newShape();';
	code += g12+'.text = "3";';
	code += g12+'.textVerticalPadding = 0;';
	code += g12+'.textSize = 7.721;';
	code += g12+'.strokeType = null;';
	code += g12+'.geometry = new Rect(504.75, 19.58, 8.39, 4.32);';
	code += g12+'.shadowColor = null;';
	code += g12+'.strokeColor = null;';
	code += g12+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g12+'.strokeThickness = 0.6166579729272318;';
	code += g12+'.textHorizontalPadding = 0;';
	code += g12+'.name = "3";';
	code += g12+'.fontName = "HelveticaNeue-Light";';
	code += g12+'.textWraps = false;';
	code += g12+'.flippedHorizontally = true;';

	code += 'var '+g13+' = canvas.newShape();';
	code += g13+'.text = "2";';
	code += g13+'.textVerticalPadding = 0;';
	code += g13+'.textSize = 7.721;';
	code += g13+'.strokeType = null;';
	code += g13+'.geometry = new Rect(473.58, 19.58, 8.39, 4.32);';
	code += g13+'.shadowColor = null;';
	code += g13+'.strokeColor = null;';
	code += g13+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g13+'.strokeThickness = 0.6166579729272318;';
	code += g13+'.textHorizontalPadding = 0;';
	code += g13+'.name = "2";';
	code += g13+'.fontName = "HelveticaNeue-Light";';
	code += g13+'.textWraps = false;';
	code += g13+'.flippedHorizontally = true;';

	code += 'var '+g14+' = canvas.newShape();';
	code += g14+'.text = "1";';
	code += g14+'.textVerticalPadding = 0;';
	code += g14+'.textSize = 7.721;';
	code += g14+'.strokeType = null;';
	code += g14+'.geometry = new Rect(442.41, 19.58, 8.39, 4.32);';
	code += g14+'.shadowColor = null;';
	code += g14+'.strokeColor = null;';
	code += g14+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g14+'.strokeThickness = 0.6166579729272318;';
	code += g14+'.textHorizontalPadding = 0;';
	code += g14+'.name = "1";';
	code += g14+'.fontName = "HelveticaNeue-Light";';
	code += g14+'.textWraps = false;';
	code += g14+'.flippedHorizontally = true;';

	code += 'var '+g15+' = canvas.newShape();';
	code += g15+'.text = "0";';
	code += g15+'.textVerticalPadding = 0;';
	code += g15+'.textSize = 7.721;';
	code += g15+'.strokeType = null;';
	code += g15+'.geometry = new Rect(411.24, 19.58, 8.39, 4.32);';
	code += g15+'.shadowColor = null;';
	code += g15+'.strokeColor = null;';
	code += g15+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g15+'.strokeThickness = 0.6166579729272318;';
	code += g15+'.textHorizontalPadding = 0;';
	code += g15+'.name = "0";';
	code += g15+'.fontName = "HelveticaNeue-Light";';
	code += g15+'.textWraps = false;';
	code += g15+'.flippedHorizontally = true;';

	code += 'var '+g16+' = canvas.newShape();';
	code += g16+'.text = "1";';
	code += g16+'.textVerticalPadding = 0;';
	code += g16+'.textSize = 7.721;';
	code += g16+'.strokeType = null;';
	code += g16+'.geometry = new Rect(380.07, 19.58, 8.39, 4.32);';
	code += g16+'.shadowColor = null;';
	code += g16+'.strokeColor = null;';
	code += g16+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g16+'.strokeThickness = 0.6166579729272318;';
	code += g16+'.textHorizontalPadding = 0;';
	code += g16+'.name = "-1";';
	code += g16+'.fontName = "HelveticaNeue-Light";';
	code += g16+'.textWraps = false;';
	code += g16+'.flippedHorizontally = true;';

	code += 'var '+g17+' = canvas.newShape();';
	code += g17+'.text = "2";';
	code += g17+'.textVerticalPadding = 0;';
	code += g17+'.textSize = 7.721;';
	code += g17+'.strokeType = null;';
	code += g17+'.geometry = new Rect(348.90, 19.58, 8.39, 4.32);';
	code += g17+'.shadowColor = null;';
	code += g17+'.strokeColor = null;';
	code += g17+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g17+'.strokeThickness = 0.6166579729272318;';
	code += g17+'.textHorizontalPadding = 0;';
	code += g17+'.name = "-2";';
	code += g17+'.fontName = "HelveticaNeue-Light";';
	code += g17+'.textWraps = false;';
	code += g17+'.flippedHorizontally = true;';

	code += 'var '+g18+' = canvas.newShape();';
	code += g18+'.text = "3";';
	code += g18+'.textVerticalPadding = 0;';
	code += g18+'.textSize = 7.721;';
	code += g18+'.strokeType = null;';
	code += g18+'.geometry = new Rect(317.73, 19.58, 8.39, 4.32);';
	code += g18+'.shadowColor = null;';
	code += g18+'.strokeColor = null;';
	code += g18+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g18+'.strokeThickness = 0.6166579729272318;';
	code += g18+'.textHorizontalPadding = 0;';
	code += g18+'.name = "-3";';
	code += g18+'.fontName = "HelveticaNeue-Light";';
	code += g18+'.textWraps = false;';
	code += g18+'.flippedHorizontally = true;';

	code += 'var '+g19+' = canvas.newShape();';
	code += g19+'.text = "4";';
	code += g19+'.textVerticalPadding = 0;';
	code += g19+'.textSize = 7.721;';
	code += g19+'.strokeType = null;';
	code += g19+'.geometry = new Rect(286.29, 19.58, 8.39, 4.32);';
	code += g19+'.shadowColor = null;';
	code += g19+'.strokeColor = null;';
	code += g19+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g19+'.strokeThickness = 0.6166579729272318;';
	code += g19+'.textHorizontalPadding = 0;';
	code += g19+'.name = "-4";';
	code += g19+'.fontName = "HelveticaNeue-Light";';
	code += g19+'.textWraps = false;';
	code += g19+'.flippedHorizontally = true;';

	code += 'var '+g20+' = canvas.newShape();';
	code += g20+'.text = "5";';
	code += g20+'.textVerticalPadding = 0;';
	code += g20+'.textSize = 7.721;';
	code += g20+'.strokeType = null;';
	code += g20+'.geometry = new Rect(255.40, 19.58, 8.39, 4.32);';
	code += g20+'.shadowColor = null;';
	code += g20+'.strokeColor = null;';
	code += g20+'.textColor = Color.RGB(0.475634932518005, 0.475646734237671, 0.475640416145325);';
	code += g20+'.strokeThickness = 0.6166579729272318;';
	code += g20+'.textHorizontalPadding = 0;';
	code += g20+'.name = "-5";';
	code += g20+'.fontName = "HelveticaNeue-Light";';
	code += g20+'.textWraps = false;';
	code += g20+'.flippedHorizontally = true;';

	code += 'var '+g21+' = new Group(['+g20+', '+g19+', '+g18+', '+g17+', '+g16+', '+g15+', '+g14+', '+g13+', '+g12+', '+g11+', '+g10+']);';
	code += g21+'.geometry = new Rect(255.40, 19.58, 320.07, 4.32);';
	
	//Full Group
	code += 'var '+g22+' = new Group(['+g21+', '+g9+', '+g5+', '+g4+', '+g3+', '+g2+', '+g1+']);';
	code += g22+'.geometry = new Rect(0.00, '+( parseInt(Cid)*parseInt(offset) )+', 817.04, 60.42);';
	baseGroupId = g22;

	return code;
}