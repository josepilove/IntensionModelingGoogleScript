//Template v1.5
/**
*************************************************************
*                                                           *
*  C R E A T E   W O R K S H E E T   O U T P U T   C O D E  *
*                                                           *
*************************************************************

generateIntensionModelingWorksheet()
buildWorksheetGenerationScript()
continuum()

**/











/**
* GENERATE INTENSION MODELING WORKSHEET
*
*
*/
function generateIntensionModelingWorksheet()
{
  var array = getFullContinuumsArray();
  var script = buildWorksheetGenerationScript(array);
  return escapeAndRunInOG(script);
}



/**
* BUILD WORKSHEET GENERATION SCRIPT
* Builds the script for omnigraffle from the continuum values.
*
* @param {array} values An array of continuum values {[0=>"TermA", 1=>"DescA", 2=>"TermB, 3=>"DescB"]}
* @return {string} cont A string that contains the non-escaped omnigraffle script, without headers.
*/
function buildWorksheetGenerationScript(values){
  var cont = "";
  Logger.log(values);
  for (var i=0; (values.length) > i; i++){
    var continuumID = (i+1);
    cont += continuum(continuumID, values[i][0], values[i][1], values[i][2], values[i][3]);
  }
  return cont;
}



/**
* CONTINUUM
* Builds a worksheet-styled continuum base and formats it for OmniGraffle
* v.4 - Added ghosted numbers 0-10 below the 5-0-5 numbers 
*       to aid the input into Google Forms, since Google 
*       forms doesn't allow for 5-5 numbering.
*
* @param {int} Cid The desired number to show at the left of the continuum (usually the continuumID)
* @param {string} Aterm The bold text on the left of the continuum
* @param {string} Adesc The sub-text on the left of the continuum
* @param {string} Bterm The bold text on the right of the continuum
* @param {string} Bdesc The sub-text on the right of the continuum
* @return {string} Javascript formatted for generating continuum in Omnigraffle
**/
//insert variables into continuum format
//return a string of the script
function continuum(Cid, Aterm, Adesc, Bterm, Bdesc){

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
	var g23 = guidGenerator();
	var g24 = guidGenerator();
	var g25 = guidGenerator();
	var g26 = guidGenerator();
	var g27 = guidGenerator();
	var g28 = guidGenerator();
	var g29 = guidGenerator();
	var g30 = guidGenerator();
	var g31 = guidGenerator();
	var g32 = guidGenerator();
	var g33 = guidGenerator();
	var g34 = guidGenerator();
	var g35 = guidGenerator();
	var g36 = guidGenerator();
	var g37 = guidGenerator();
	var g38 = guidGenerator();
	var g39 = guidGenerator();
	var g40 = guidGenerator();
	var g41 = guidGenerator();
	var g42 = guidGenerator();
	var g43 = guidGenerator();
	var g44 = guidGenerator();
	
	    //Right Triangle
	var code = 'var '+g1+' = canvas.newShape();';
		code += g1+'.textUnitRect = new Rect(0.00, 0.50, 0.60, 0.50);';
		code += g1+'.shape = "RightTriangle";';
		code += g1+'.textHorizontalAlignment = HorizontalTextAlignment.Right;';
		code += g1+'.strokeType = null;';
		code += g1+'.geometry = new Rect(250.69, 7.23, 261.71, 21.79);';
		code += g1+'.shadowColor = null;';
		code += g1+'.strokeColor = null;';
		code += g1+'.fillColor = Color.RGB(1.0, 0.781338, 0.683827, 0.5);';
		code += g1+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
		code += g1+'.strokeThickness = 0.5288657895318044;';
		code += g1+'.textHorizontalPadding = 20;';
		code += g1+'.fontName = "Helvetica-Bold";';
	    //Right Triangle
		code += 'var '+g2+' = canvas.newShape();';
		code += g2+'.textUnitRect = new Rect(0.00, 0.50, 0.60, 0.50);';
		code += g2+'.shape = "RightTriangle";';
		code += g2+'.textHorizontalAlignment = HorizontalTextAlignment.Right;';
		code += g2+'.strokeType = null;';
		code += g2+'.flippedVertically = true;';
		code += g2+'.geometry = new Rect(250.70, 7.20, 261.71, 21.79);';
		code += g2+'.shadowColor = null;';
		code += g2+'.strokeColor = null;';
		code += g2+'.fillColor = Color.RGB(0.704052, 0.911016, 1.0, 0.5);';
		code += g2+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
		code += g2+'.strokeThickness = 0.5288657895318044;';
		code += g2+'.textHorizontalPadding = 20;';
		code += g2+'.name = "Right Triangle";';
		code += g2+'.fontName = "Helvetica-Bold";';
		code += g2+'.flippedHorizontally = true;';
		//B Description
		code += 'var '+g3+' = canvas.newShape();';
		code += g3+'.fillType = null;';
		code += g3+'.text = ';
		code += '"'+Bdesc+'"';
		code += ';';
		code += g3+'.textVerticalPadding = 4.727363804573168;';
		code += g3+'.textSize = 9.398482520461659;';
		code += g3+'.textHorizontalAlignment = HorizontalTextAlignment.Left;';
		code += g3+'.strokeType = null;';
		code += g3+'.geometry = new Rect(515.77, 16.59, 222.84, 42.96);';
		code += g3+'.shadowColor = null;';
		code += g3+'.strokeColor = null;';
		code += g3+'.textVerticalPlacement = VerticalTextPlacement.Top;';
		code += g3+'.fillColor = null;';
		code += g3+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
		code += g3+'.strokeThickness = 0.9454727956162144;';
		code += g3+'.textHorizontalPadding = 4.727363804573168;';
		//B Term
		code += 'var '+g4+' = canvas.newShape();';
		code += g4+'.fillType = null;';
		code += g4+'.text = ';
		code += '"'+Bterm+'"';
		code += ';';
		code += g4+'.textVerticalPadding = 0;';
		code += g4+'.textSize = 12.234758398878348;';
		code += g4+'.textHorizontalAlignment = HorizontalTextAlignment.Left;';
		code += g4+'.strokeType = null;';
		code += g4+'.geometry = new Rect(515.77, 0.76, 222.84, 15.83);';
		code += g4+'.shadowColor = null;';
		code += g4+'.strokeColor = null;';
		code += g4+'.textVerticalPlacement = VerticalTextPlacement.Bottom;';
		code += g4+'.fillColor = null;';
		code += g4+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
		code += g4+'.strokeThickness = 0.9454727956162144;';
		code += g4+'.textHorizontalPadding = 4.727363804573168;';
		code += g4+'.fontName = "HelveticaNeue-Bold";';
		//Right-Side(B) Table
		code += 'var '+g5+' = Table.withRowsColumns(2, 1, ['+g4+', '+g3+']);';
		code += g5+'.rows = 2;';
		code += g5+'.columns = 1;';
		code += g5+'.name = "Right-Side Table";';
		code += g5+'.geometry = new Rect(515.77, 0.76, 222.84, 58.79);';
		//Right 5
		code += 'var '+g6+' = canvas.newShape();';
		code += g6+'.fillType = null;';
		code += g6+'.text = "5";';
		code += g6+'.textVerticalPadding = 2.6443288326263428;';
		code += g6+'.textSize = 10.462;';
		code += g6+'.geometry = new Rect(492.99, 7.20, 19.42, 21.71);';
		code += g6+'.shadowColor = null;';
		code += g6+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g6+'.fillColor = null;';
		code += g6+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g6+'.strokeThickness = 0.2644328947659022;';
		code += g6+'.textHorizontalPadding = 2.6443288326263428;';
		code += g6+'.name = "Right 5";';
		code += g6+'.fontName = "Helvetica-Bold";';
		//Spacer 10
		code += 'var '+g7+' = canvas.newShape();';
		code += g7+'.autosizing = TextAutosizing.Vertical;';
		code += g7+'.strokeType = null;';
		code += g7+'.geometry = new Rect(488.13, 7.23, 4.85, 21.76);';
		code += g7+'.shadowColor = null;';
		code += g7+'.strokeColor = null;';
		code += g7+'.strokeThickness = 0.8687541750383465;';
		code += g7+'.name = "Spacer10";';
		//Right 4
		code += 'var '+g8+' = canvas.newShape();';
		code += g8+'.fillType = null;';
		code += g8+'.text = "4";';
		code += g8+'.textVerticalPadding = 2.6443288326263428;';
		code += g8+'.textSize = 10.462;';
		code += g8+'.geometry = new Rect(468.72, 7.20, 19.42, 21.71);';
		code += g8+'.shadowColor = null;';
		code += g8+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g8+'.fillColor = null;';
		code += g8+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g8+'.strokeThickness = 0.2644328947659022;';
		code += g8+'.textHorizontalPadding = 2.6443288326263428;';
		code += g8+'.name = "Right 4";';
		code += g8+'.fontName = "Helvetica-Bold";';
		//Spacer 9
		code += 'var '+g9+' = canvas.newShape();';
		code += g9+'.autosizing = TextAutosizing.Vertical;';
		code += g9+'.strokeType = null;';
		code += g9+'.geometry = new Rect(463.87, 7.23, 4.85, 21.76);';
		code += g9+'.shadowColor = null;';
		code += g9+'.strokeColor = null;';
		code += g9+'.strokeThickness = 0.8687541750383465;';
		code += g9+'.name = "Spacer9";';
		//Right 3
		code += 'var '+g10+' = canvas.newShape();';
		code += g10+'.fillType = null;';
		code += g10+'.text = "3";';
		code += g10+'.textVerticalPadding = 2.6443288326263428;';
		code += g10+'.textSize = 10.462;';
		code += g10+'.geometry = new Rect(444.46, 7.20, 19.42, 21.71);';
		code += g10+'.shadowColor = null;';
		code += g10+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g10+'.fillColor = null;';
		code += g10+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g10+'.strokeThickness = 0.2644328947659022;';
		code += g10+'.textHorizontalPadding = 2.6443288326263428;';
		code += g10+'.name = "Right 3";';
		code += g10+'.fontName = "Helvetica-Bold";';
		//Spacer 8
		code += 'var '+g11+' = canvas.newShape();';
		code += g11+'.autosizing = TextAutosizing.Vertical;';
		code += g11+'.strokeType = null;';
		code += g11+'.geometry = new Rect(439.60, 7.23, 4.85, 21.76);';
		code += g11+'.shadowColor = null;';
		code += g11+'.strokeColor = null;';
		code += g11+'.strokeThickness = 0.8687541750383465;';
		code += g11+'.name = "Spacer8";';
		//Right 2
		code += 'var '+g12+' = canvas.newShape();';
		code += g12+'.fillType = null;';
		code += g12+'.text = "2";';
		code += g12+'.textVerticalPadding = 2.6443288326263428;';
		code += g12+'.textSize = 10.462;';
		code += g12+'.geometry = new Rect(420.19, 7.20, 19.42, 21.71);';
		code += g12+'.shadowColor = null;';
		code += g12+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g12+'.fillColor = null;';
		code += g12+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g12+'.strokeThickness = 0.2644328947659022;';
		code += g12+'.textHorizontalPadding = 2.6443288326263428;';
		code += g12+'.name = "Right 2";';
		code += g12+'.fontName = "Helvetica-Bold";';
		//Spacer 7
		code += 'var '+g13+' = canvas.newShape();';
		code += g13+'.autosizing = TextAutosizing.Vertical;';
		code += g13+'.strokeType = null;';
		code += g13+'.geometry = new Rect(415.33, 7.23, 4.85, 21.76);';
		code += g13+'.shadowColor = null;';
		code += g13+'.strokeColor = null;';
		code += g13+'.strokeThickness = 0.8687541750383465;';
		code += g13+'.name = "Spacer7";';
		//Right 1
		code += 'var '+g14+' = canvas.newShape();';
		code += g14+'.fillType = null;';
		code += g14+'.text = "1";';
		code += g14+'.textVerticalPadding = 2.6443288326263428;';
		code += g14+'.textSize = 10.462;';
		code += g14+'.geometry = new Rect(395.92, 7.20, 19.42, 21.71);';
		code += g14+'.shadowColor = null;';
		code += g14+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g14+'.fillColor = null;';
		code += g14+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g14+'.strokeThickness = 0.2644328947659022;';
		code += g14+'.textHorizontalPadding = 2.6443288326263428;';
		code += g14+'.name = "Right 1";';
		code += g14+'.fontName = "Helvetica-Bold";';
		//Spacer 6
		code += 'var '+g15+' = canvas.newShape();';
		code += g15+'.autosizing = TextAutosizing.Vertical;';
		code += g15+'.strokeType = null;';
		code += g15+'.geometry = new Rect(391.06, 7.20, 4.85, 21.76);';
		code += g15+'.shadowColor = null;';
		code += g15+'.strokeColor = null;';
		code += g15+'.strokeThickness = 0.8687541750383465;';
		code += g15+'.name = "Spacer6";';
		//Zero
		code += 'var '+g16+' = canvas.newShape();';
		code += g16+'.fillType = null;';
		code += g16+'.text = "0";';
		code += g16+'.textVerticalPadding = 2.6443288326263428;';
		code += g16+'.textSize = 10.462;';
		code += g16+'.geometry = new Rect(371.66, 7.20, 19.42, 21.71);';
		code += g16+'.shadowColor = null;';
		code += g16+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g16+'.fillColor = null;';
		code += g16+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g16+'.strokeThickness = 0.2644328947659022;';
		code += g16+'.textHorizontalPadding = 2.6443288326263428;';
		code += g16+'.name = "Zero";';
		code += g16+'.fontName = "Helvetica-Bold";';
		//Spacer 5
		code += 'var '+g17+' = canvas.newShape();';
		code += g17+'.autosizing = TextAutosizing.Vertical;';
		code += g17+'.strokeType = null;';
		code += g17+'.geometry = new Rect(366.80, 7.23, 4.85, 21.76);';
		code += g17+'.shadowColor = null;';
		code += g17+'.strokeColor = null;';
		code += g17+'.strokeThickness = 0.8687541750383465;';
		code += g17+'.name = "Spacer5";';
		//Left 1
		code += 'var '+g18+' = canvas.newShape();';
		code += g18+'.fillType = null;';
		code += g18+'.text = "1";';
		code += g18+'.textVerticalPadding = 2.6443288326263428;';
		code += g18+'.textSize = 10.462;';
		code += g18+'.geometry = new Rect(347.39, 7.20, 19.42, 21.71);';
		code += g18+'.shadowColor = null;';
		code += g18+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g18+'.fillColor = null;';
		code += g18+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g18+'.strokeThickness = 0.2644328947659022;';
		code += g18+'.textHorizontalPadding = 2.6443288326263428;';
		code += g18+'.name = "Left 1";';
		code += g18+'.fontName = "Helvetica-Bold";';
		//Spacer 4
		code += 'var '+g19+' = canvas.newShape();';
		code += g19+'.autosizing = TextAutosizing.Vertical;';
		code += g19+'.strokeType = null;';
		code += g19+'.geometry = new Rect(342.53, 7.20, 4.85, 21.76);';
		code += g19+'.shadowColor = null;';
		code += g19+'.strokeColor = null;';
		code += g19+'.strokeThickness = 0.8687541750383465;';
		code += g19+'.name = "Spacer4";';
		//Left 2
		code += 'var '+g20+' = canvas.newShape();';
		code += g20+'.fillType = null;';
		code += g20+'.text = "2";';
		code += g20+'.textVerticalPadding = 2.6443288326263428;';
		code += g20+'.textSize = 10.462;';
		code += g20+'.geometry = new Rect(323.12, 7.20, 19.42, 21.71);';
		code += g20+'.shadowColor = null;';
		code += g20+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g20+'.fillColor = null;';
		code += g20+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g20+'.strokeThickness = 0.2644328947659022;';
		code += g20+'.textHorizontalPadding = 2.6443288326263428;';
		code += g20+'.name = "Left 2";';
		code += g20+'.fontName = "Helvetica-Bold";';
		//Spacer 3
		code += 'var '+g21+' = canvas.newShape();';
		code += g21+'.autosizing = TextAutosizing.Vertical;';
		code += g21+'.strokeType = null;';
		code += g21+'.geometry = new Rect(318.26, 7.20, 4.85, 21.76);';
		code += g21+'.shadowColor = null;';
		code += g21+'.strokeColor = null;';
		code += g21+'.strokeThickness = 0.8687541750383465;';
		code += g21+'.name = "Spacer3";';
		//Left 3
		code += 'var '+g22+' = canvas.newShape();';
		code += g22+'.fillType = null;';
		code += g22+'.text = "3";';
		code += g22+'.textVerticalPadding = 2.6443288326263428;';
		code += g22+'.textSize = 10.462;';
		code += g22+'.geometry = new Rect(298.85, 7.20, 19.42, 21.71);';
		code += g22+'.shadowColor = null;';
		code += g22+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g22+'.fillColor = null;';
		code += g22+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g22+'.strokeThickness = 0.2644328947659022;';
		code += g22+'.textHorizontalPadding = 2.6443288326263428;';
		code += g22+'.name = "Left 3";';
		code += g22+'.fontName = "Helvetica-Bold";';
		//Spacer 2
		code += 'var '+g23+' = canvas.newShape();';
		code += g23+'.autosizing = TextAutosizing.Vertical;';
		code += g23+'.strokeType = null;';
		code += g23+'.geometry = new Rect(293.99, 7.20, 4.85, 21.76);';
		code += g23+'.shadowColor = null;';
		code += g23+'.strokeColor = null;';
		code += g23+'.strokeThickness = 0.8687541750383465;';
		code += g23+'.name = "Spacer2";';
		//Left 4
		code += 'var '+g24+' = canvas.newShape();';
		code += g24+'.fillType = null;';
		code += g24+'.text = "4";';
		code += g24+'.textVerticalPadding = 2.6443288326263428;';
		code += g24+'.textSize = 10.462;';
		code += g24+'.geometry = new Rect(274.59, 7.20, 19.42, 21.71);';
		code += g24+'.shadowColor = null;';
		code += g24+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g24+'.fillColor = null;';
		code += g24+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g24+'.strokeThickness = 0.2644328947659022;';
		code += g24+'.textHorizontalPadding = 2.6443288326263428;';
		code += g24+'.name = "Left 4";';
		code += g24+'.fontName = "Helvetica-Bold";';
		//Spacer 1
		code += 'var '+g25+' = canvas.newShape();';
		code += g25+'.autosizing = TextAutosizing.Vertical;';
		code += g25+'.strokeType = null;';
		code += g25+'.geometry = new Rect(269.73, 7.15, 4.85, 21.76);';
		code += g25+'.shadowColor = null;';
		code += g25+'.strokeColor = null;';
		code += g25+'.strokeThickness = 0.8687541750383465;';
		code += g25+'.name = "Spacer1";';
		//Left 5
		code += 'var '+g26+' = canvas.newShape();';
		code += g26+'.fillType = null;';
		code += g26+'.text = "5";';
		code += g26+'.textVerticalPadding = 2.6443288326263428;';
		code += g26+'.textSize = 10.462;';
		code += g26+'.geometry = new Rect(250.32, 7.20, 19.42, 21.71);';
		code += g26+'.shadowColor = null;';
		code += g26+'.strokeColor = Color.RGB(0.582505, 0.582519, 0.582512);';
		code += g26+'.fillColor = null;';
		code += g26+'.textColor = Color.RGB(0.260517418384552, 0.260524392127991, 0.26052063703537);';
		code += g26+'.strokeThickness = 0.2644328947659022;';
		code += g26+'.textHorizontalPadding = 2.6443288326263428;';
		code += g26+'.name = "Left 5";';
		code += g26+'.fontName = "Helvetica-Bold";';
		//A Description
		code += 'var '+g27+' = canvas.newShape();';
		code += g27+'.fillType = null;';
		code += g27+'.text = ';
		code += '"'+Adesc+'"';
		code += ';';
		code += g27+'.textVerticalPadding = 4.727363804573168;';
		code += g27+'.textSize = 9.398482520461659;';
		code += g27+'.textHorizontalAlignment = HorizontalTextAlignment.Right;';
		code += g27+'.strokeType = null;';
		code += g27+'.geometry = new Rect(23.94, 16.59, 222.84, 32.09);';
		code += g27+'.shadowColor = null;';
		code += g27+'.strokeColor = null;';
		code += g27+'.textVerticalPlacement = VerticalTextPlacement.Top;';
		code += g27+'.fillColor = null;';
		code += g27+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
		code += g27+'.strokeThickness = 0.9454727956162144;';
		code += g27+'.textHorizontalPadding = 4.727363804573168;';
		//A Term
		code += 'var '+g28+' = canvas.newShape();';
		code += g28+'.fillType = null;';
		code += g28+'.text = ';
		code += '"'+Aterm+'"';
		code += ';';
		code += g28+'.textVerticalPadding = 0;';
		code += g28+'.textSize = 12.234758398878348;';
		code += g28+'.textHorizontalAlignment = HorizontalTextAlignment.Right;';
		code += g28+'.strokeType = null;';
		code += g28+'.geometry = new Rect(423.94, 0.76, 222.84, 15.83);';
		code += g28+'.shadowColor = null;';
		code += g28+'.strokeColor = null;';
		code += g28+'.textVerticalPlacement = VerticalTextPlacement.Bottom;';
		code += g28+'.fillColor = null;';
		code += g28+'.textColor = Color.RGB(0.0, 0.0, 0.0);';
		code += g28+'.strokeThickness = 0.9454727956162144;';
		code += g28+'.textHorizontalPadding = 4.727363804573168;';
		code += g28+'.fontName = "HelveticaNeue-Bold";';
		//Left Side (A) Table
		code += 'var '+g29+' = Table.withRowsColumns(2, 1, ['+g28+', '+g27+']);';
		code += g29+'.rows = 2;';
		code += g29+'.columns = 1;';
		code += g29+'.name = "Left Side Table";';
		code += g29+'.geometry = new Rect(23.94, 0.76, 222.84, 47.91);';
		//Base Line
		code += 'var '+g30+' = canvas.newLine();';
		code += g30+'.strokeJoin = LineJoin.Miter;';
		code += g30+'.strokeThickness = 1.676970677222307;';
		code += g30+'.name = "Base Line";';
		code += g30+'.shadowColor = null;';
		code += g30+'.strokeCap = LineCap.Butt;';
		code += g30+'.points = [new Point(250.67, 56.55), new Point(512.91, 56.55)];';
		//ID Marker
		code += 'var '+g31+' = canvas.newShape();';
		code += g31+'.fillType = null;';
		code += g31+'.text = ';
		code += '"'+Cid+'"';
		code += ';';
		code += g31+'.textSize = 18;';
		code += g31+'.strokeType = null;';
		code += g31+'.geometry = new Rect(0.00, 0.00, 21.02, 32.58);';
		code += g31+'.shadowColor = null;';
		code += g31+'.strokeColor = null;';
		code += g31+'.fillColor = null;';
		code += g31+'.textColor = Color.RGB(0.647058844566345, 0.647058844566345, 0.647058844566345);';
		code += g31+'.fontName = "HelveticaNeue-Bold";';
		code += g31+'.textWraps = false;';
		//Ghost 0
		code += 'var '+g32+' = canvas.newShape();';
		code += g32+'.fillType = null;';
		code += g32+'.text = ';
		code += '"0"';
		code += ';';
		code += g32+'.textVerticalPadding = 0;';
		code += g32+'.autosizing = TextAutosizing.Full;';
		code += g32+'.textSize = 8;';
		code += g32+'.strokeType = null;';
		code += g32+'.geometry = new Rect(258.40, 32.00, 5.00, 11.00);';
		code += g32+'.shadowColor = null;';
		code += g32+'.strokeColor = null;';
		code += g32+'.fillColor = null;';
		code += g32+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g32+'.textHorizontalPadding = 0;';
		code += g32+'.name = "ghost 0";';
		code += g32+'.fontName = "HelveticaNeue-Italic";';
		code += g32+'.textWraps = false;';
		//Ghost 1
		code += 'var '+g33+' = canvas.newShape();';
		code += g33+'.fillType = null;';
		code += g33+'.text = ';
		code += '"1"';
		code += ';';
		code += g33+'.textVerticalPadding = 0;';
		code += g33+'.autosizing = TextAutosizing.Full;';
		code += g33+'.textSize = 8;';
		code += g33+'.strokeType = null;';
		code += g33+'.geometry = new Rect(282.58, 32.00, 5.00, 11.00);';
		code += g33+'.shadowColor = null;';
		code += g33+'.strokeColor = null;';
		code += g33+'.fillColor = null;';
		code += g33+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g33+'.textHorizontalPadding = 0;';
		code += g33+'.name = "ghost 1";';
		code += g33+'.fontName = "HelveticaNeue-Italic";';
		code += g33+'.textWraps = false;';
		//Ghost 2
		code += 'var '+g34+' = canvas.newShape();';
		code += g34+'.fillType = null;';
		code += g34+'.text = ';
		code += '"2"';
		code += ';';
		code += g34+'.textVerticalPadding = 0;';
		code += g34+'.autosizing = TextAutosizing.Full;';
		code += g34+'.textSize = 8;';
		code += g34+'.strokeType = null;';
		code += g34+'.geometry = new Rect(306.76, 32.00, 5.00, 11.00);';
		code += g34+'.shadowColor = null;';
		code += g34+'.strokeColor = null;';
		code += g34+'.fillColor = null;';
		code += g34+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g34+'.textHorizontalPadding = 0;';
		code += g34+'.name = "ghost 2";';
		code += g34+'.fontName = "HelveticaNeue-Italic";';
		code += g34+'.textWraps = false;';
		//Ghost 3
		code += 'var '+g35+' = canvas.newShape();';
		code += g35+'.fillType = null;';
		code += g35+'.text = ';
		code += '"3"';
		code += ';';
		code += g35+'.textVerticalPadding = 0;';
		code += g35+'.autosizing = TextAutosizing.Full;';
		code += g35+'.textSize = 8;';
		code += g35+'.strokeType = null;';
		code += g35+'.geometry = new Rect(330.94, 32.00, 5.00, 11.00);';
		code += g35+'.shadowColor = null;';
		code += g35+'.strokeColor = null;';
		code += g35+'.fillColor = null;';
		code += g35+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g35+'.textHorizontalPadding = 0;';
		code += g35+'.name = "ghost 3";';
		code += g35+'.fontName = "HelveticaNeue-Italic";';
		code += g35+'.textWraps = false;';
		//Ghost 4
		code += 'var '+g36+' = canvas.newShape();';
		code += g36+'.fillType = null;';
		code += g36+'.text = ';
		code += '"4"';
		code += ';';
		code += g36+'.textVerticalPadding = 0;';
		code += g36+'.autosizing = TextAutosizing.Full;';
		code += g36+'.textSize = 8;';
		code += g36+'.strokeType = null;';
		code += g36+'.geometry = new Rect(355.12, 32.00, 5.00, 11.00);';
		code += g36+'.shadowColor = null;';
		code += g36+'.strokeColor = null;';
		code += g36+'.fillColor = null;';
		code += g36+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g36+'.textHorizontalPadding = 0;';
		code += g36+'.name = "ghost 4";';
		code += g36+'.fontName = "HelveticaNeue-Italic";';
		code += g36+'.textWraps = false;';
		//Ghost 5
		code += 'var '+g37+' = canvas.newShape();';
		code += g37+'.fillType = null;';
		code += g37+'.text = ';
		code += '"5"';
		code += ';';
		code += g37+'.textVerticalPadding = 0;';
		code += g37+'.autosizing = TextAutosizing.Full;';
		code += g37+'.textSize = 8;';
		code += g37+'.strokeType = null;';
		code += g37+'.geometry = new Rect(379.30, 32.00, 5.00, 11.00);';
		code += g37+'.shadowColor = null;';
		code += g37+'.strokeColor = null;';
		code += g37+'.fillColor = null;';
		code += g37+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g37+'.textHorizontalPadding = 0;';
		code += g37+'.name = "ghost 5";';
		code += g37+'.fontName = "HelveticaNeue-Italic";';
		code += g37+'.textWraps = false;';
		//Ghost 6
		code += 'var '+g38+' = canvas.newShape();';
		code += g38+'.fillType = null;';
		code += g38+'.text = ';
		code += '"6"';
		code += ';';
		code += g38+'.textVerticalPadding = 0;';
		code += g38+'.autosizing = TextAutosizing.Full;';
		code += g38+'.textSize = 8;';
		code += g38+'.strokeType = null;';
		code += g38+'.geometry = new Rect(403.48, 32.00, 5.00, 11.00);';
		code += g38+'.shadowColor = null;';
		code += g38+'.strokeColor = null;';
		code += g38+'.fillColor = null;';
		code += g38+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g38+'.textHorizontalPadding = 0;';
		code += g38+'.name = "ghost 6";';
		code += g38+'.fontName = "HelveticaNeue-Italic";';
		code += g38+'.textWraps = false;';
		//Ghost 7
		code += 'var '+g39+' = canvas.newShape();';
		code += g39+'.fillType = null;';
		code += g39+'.text = ';
		code += '"7"';
		code += ';';
		code += g39+'.textVerticalPadding = 0;';
		code += g39+'.autosizing = TextAutosizing.Full;';
		code += g39+'.textSize = 8;';
		code += g39+'.strokeType = null;';
		code += g39+'.geometry = new Rect(427.66, 32.00, 5.00, 11.00);';
		code += g39+'.shadowColor = null;';
		code += g39+'.strokeColor = null;';
		code += g39+'.fillColor = null;';
		code += g39+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g39+'.textHorizontalPadding = 0;';
		code += g39+'.name = "ghost 7";';
		code += g39+'.fontName = "HelveticaNeue-Italic";';
		code += g39+'.textWraps = false;';
		//Ghost 8
		code += 'var '+g40+' = canvas.newShape();';
		code += g40+'.fillType = null;';
		code += g40+'.text = ';
		code += '"8"';
		code += ';';
		code += g40+'.textVerticalPadding = 0;';
		code += g40+'.autosizing = TextAutosizing.Full;';
		code += g40+'.textSize = 8;';
		code += g40+'.strokeType = null;';
		code += g40+'.geometry = new Rect(451.84, 32.00, 5.00, 11.00);';
		code += g40+'.shadowColor = null;';
		code += g40+'.strokeColor = null;';
		code += g40+'.fillColor = null;';
		code += g40+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g40+'.textHorizontalPadding = 0;';
		code += g40+'.name = "ghost 8";';
		code += g40+'.fontName = "HelveticaNeue-Italic";';
		code += g40+'.textWraps = false;';
		//Ghost 9
		code += 'var '+g41+' = canvas.newShape();';
		code += g41+'.fillType = null;';
		code += g41+'.text = ';
		code += '"9"';
		code += ';';
		code += g41+'.textVerticalPadding = 0;';
		code += g41+'.autosizing = TextAutosizing.Full;';
		code += g41+'.textSize = 8;';
		code += g41+'.strokeType = null;';
		code += g41+'.geometry = new Rect(476.02, 32.00, 5.00, 11.00);';
		code += g41+'.shadowColor = null;';
		code += g41+'.strokeColor = null;';
		code += g41+'.fillColor = null;';
		code += g41+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g41+'.textHorizontalPadding = 0;';
		code += g41+'.name = "ghost 9";';
		code += g41+'.fontName = "HelveticaNeue-Italic";';
		code += g41+'.textWraps = false;';
		//Ghost 10
		code += 'var '+g42+' = canvas.newShape();';
		code += g42+'.fillType = null;';
		code += g42+'.text = ';
		code += '"10"';
		code += ';';
		code += g42+'.textVerticalPadding = 0;';
		code += g42+'.autosizing = TextAutosizing.Full;';
		code += g42+'.textSize = 8;';
		code += g42+'.strokeType = null;';
		code += g42+'.geometry = new Rect(497.70, 32.00, 10.00, 11.00);';
		code += g42+'.shadowColor = null;';
		code += g42+'.strokeColor = null;';
		code += g42+'.fillColor = null;';
		code += g42+'.textColor = Color.RGB(0.64706, 0.64706, 0.64706);';
		code += g42+'.textHorizontalPadding = 0;';
		code += g42+'.name = "ghost 10";';
		code += g42+'.fontName = "HelveticaNeue-Italic";';
		code += g42+'.textWraps = false;';
		
		code += 'var '+g43+' = new Group(['+g42+', '+g41+', '+g40+', '+g39+', '+g38+', '+g37+', '+g36+', '+g35+', '+g34+', '+g33+', '+g32+']);';
		code += g43+'.name = "Ghosts";';
		code += g43+'.geometry = new Rect(258.40, 32.00, 249.30, 11.00);';
		//Full Group
		code += 'var '+g44+' = new Group(['+g43+', '+g31+', '+g30+', '+g29+', '+g26+', '+g25+', '+g24+', '+g23+', '+g22+', '+g21+', '+g20+', '+g19+', '+g18+', '+g17+', '+g16+', '+g15+', '+g14+', '+g13+', '+g12+', '+g11+', '+g10+', '+g9+', '+g8+', '+g7+', '+g6+', '+g5+', '+g2+', '+g1+']);';
		code += g44+'.geometry = new Rect(0, '+ ( ( parseInt(Cid)*parseInt(offset)  ) )+', 738.61, 59.55);';
	return code;
}