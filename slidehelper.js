/**
 * @OnlyCurrentDoc Limits the script to only accessing the current presentation.
 */

function onOpen(event) {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem("Slide Helper", "showSidebar")
    .addToUi();
}
var scriptProperties = PropertiesService.getScriptProperties();
function storeProperties(endSlide) {
  var start = new Date().getTime();
  var presentation = SlidesApp.getActivePresentation();
  var returnProp;
  var elementNum;
  Logger.log("storeProperties() is called.")
  slides=presentation.getSlides();
  //scriptProperties.setProperties("currSlideNumber", currSlideNumber);
  var start, end;
  if(!endSlide) {
    var currSlideNumber = presentation.getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '');
    start = 0;
    end = currSlideNumber-1;
    scriptProperties.setProperty("lastStoredIndex", end);
  }
  else {
    start = parseInt(scriptProperties.getProperty("lastStoredIndex"));
    end = endSlide;
    scriptProperties.setProperty("lastStoredIndex", end);
  }
  var comma = ',';
  for(var i=start;i<end;i++) //stores the previous slide element text properties
  {
    returnProp='';
    elementNum=1;
    slides[i].getPageElements().forEach(function(pageElement) {
      if(pageElement.getPageElementType().toString()==='SHAPE')
      {
        var textRange = pageElement.asShape().getText();
        textRange.getRuns().forEach(function(run) {
          var textStyle= run.getTextStyle();
          if(run.getLength()>1)
          {
            returnProp=returnProp.concat(textStyle.isBold(),comma);
            returnProp=returnProp.concat(textStyle.isItalic(),comma);
            returnProp=returnProp.concat(textStyle.isUnderline(),comma);
            returnProp=returnProp.concat(textStyle.getFontSize(),comma);
            returnProp=returnProp.concat(textStyle.getFontFamily(),comma);
            returnProp=returnProp.concat(getColorString(textStyle,'Foreground'),comma);
            returnProp=returnProp.concat(getColorString(textStyle,'Background'),comma);
            returnProp=returnProp.concat(pageElement.getPageElementType(),comma);
            returnProp=returnProp.concat(pageElement.getHeight(),comma);
            returnProp=returnProp.concat(pageElement.getWidth(),comma);
            returnProp=returnProp.concat(pageElement.getTop(),comma);
            returnProp=returnProp.concat(pageElement.getLeft());
            returnProp=returnProp.concat("\n");
            
            elementNum++; 
          }
        });

      }
        
      });
      var temp = "slide"+i;
      scriptProperties.setProperty(temp, returnProp);
      Logger.log("returnProp:", returnProp);
      //Logger.log("Learned "+key);
  }
   
  /**
   * stores the selected element text properties
   */
  var end = new Date().getTime();
  Logger.log("storeProperties() takes "+(end-start)+" miliseconds to complete.");
  return returnProp;
}

function splitString(key){
  var start = new Date().getTime();
  if (typeof key == 'number'){
    key = "slide" + key;
  }
  var slideElementsString = scriptProperties.getProperty(key);
  var slideElementsArray = slideElementsString.split("\n");
  slideElementsArray.pop();
  var newSlideElementsArray = [];
  slideElementsArray.forEach(function(slideElement){
    var newTextProperties = [];
    textProperties = slideElement.split(",");
    textProperties.forEach(function(property){
      if (property=="true"){
        newTextProperties.push(true);
      }
      else if (property=="false"){
        newTextProperties.push(false);
      }
      else if (isNaN(parseFloat(property))){
        newTextProperties.push(property);
      }
      else {
        newTextProperties.push(parseFloat(property));
      }
    });
    newSlideElementsArray.push(newTextProperties);
  });
  var end = new Date().getTime();
  Logger.log("splitString() takes: "+(end-start)+" miliseconds to complete.");
  return newSlideElementsArray;
}

var scriptProperties = PropertiesService.getScriptProperties();

/**
 * Open the Add-on upon install.
 * @param {Event} event The install event.
 */
function onInstall(event) {
  onOpen(event);
  
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile("sidebar").setTitle(
    "Slide Helper"
  );
  SlidesApp.getUi().showSidebar(ui);

}

/**
 * Recursively gets child text elements a list of elements.
 * @param {PageElement[]} elements The elements to get text from.
 * @return {Text[]} An array of text elements.
 */
function getElementTexts(elements) {
  var texts = [];
  elements.forEach(function(element) {
    switch (element.getPageElementType()) {
      case SlidesApp.PageElementType.GROUP:
        element
          .asGroup()
          .getChildren()
          .forEach(function(child) {
            texts = texts.concat(getElementTexts(child));
          });
        break;
      case SlidesApp.PageElementType.TABLE:
        var table = element.asTable();
        for (var y = 0; y < table.getNumColumns(); ++y) {
          for (var x = 0; x < table.getNumRows(); ++x) {
            texts.push(table.getCell(x, y).getText());
          }
        }
        break;
      case SlidesApp.PageElementType.SHAPE:
        texts.push(element.asShape().getText());
        break;
    }
  });
  return texts;
}

function selectedStuff() {
  var selection = SlidesApp.getActivePresentation().getSelection();
  var selectionType = selection.getSelectionType();
  var texts = [];
  switch (selectionType) {
    case SlidesApp.SelectionType.PAGE:
      var pages = selection
        .getPageRange()
        .getPages()
        .forEach(function(page) {
          texts = texts.concat(getElementTexts(page.getPageElements()));
        });
      break;
    case SlidesApp.SelectionType.PAGE_ELEMENT:
      var pageElements = selection.getPageElementRange().getPageElements();
      texts = texts.concat(getElementTexts(pageElements));
      break;
    case SlidesApp.SelectionType.TABLE_CELL:
      var cells = selection
        .getTableCellRange()
        .getTableCells()
        .forEach(function(cell) {
          texts.push(cell.getText());
        });
      break;
    case SlidesApp.SelectionType.TEXT:
      var elements = selection
        .getPageElementRange()
        .getPageElements()
        .forEach(function(element) {
          texts.push(element.asShape().getText());
        });
      break;
  }
  return texts;
}

function italicizeSelectedElements(toggle) {
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setItalic(toggle);
  });
}

function boldSelectedElements(toggle) {
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setBold(toggle);
  });
}

function resizeAndPosition() {
  
  var highestWidth=0;
  var tempWidth;
  var left=9999999999999;
  var currPage=SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var pageElements=currPage.getPageElements();
  pageElements.forEach(function(pageElement) {
    tempWidth = pageElement.getWidth();
    if(highestWidth<tempWidth) {
      highestWidth = tempWidth;
      left = pageElement.getLeft();
    }
    if(pageElement.getLeft()<left) 
      left = pageElement.getLeft();
  });

  pageElements.forEach(function(pageElement) {
    pageElement.setWidth(highestWidth);
    pageElement.setLeft(left);
  });
}

function underlineSelectedElements(toggle) {
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setUnderline(toggle);
  });
}

function changeFontSize(size){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setFontSize(size);
  });
}
function changeFontType(font){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setFontFamily(font);
  });
}

function changeTextBackgroundColor(color){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setBackgroundColor(color);
  });
}
function changeTextForegroundColor(color){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setForegroundColor(color);
  });
}

function getTextProp() {
  var docProperties = [false,false,false,0,'','','',0,0,0,0,0];
  var selection = SlidesApp.getActivePresentation().getSelection();
  var textStyle= selection.getTextRange().getTextStyle();
  var element = selection.getPageElementRange().getPageElements()[0];
  var bold = textStyle.isBold();
  var italic = textStyle.isItalic();
  var underline = textStyle.isUnderline();
  docProperties[0] = bold;
  docProperties[1] = italic;
  docProperties[2] = underline;
  docProperties[3] = textStyle.getFontSize();
  docProperties[4] = textStyle.getFontFamily();
  docProperties[5] = getColorString(textStyle,'Foreground');
  docProperties[6] = getColorString(textStyle,'Background');
  docProperties[7] = textStyle.getBaselineOffset();
  docProperties[8] = element.getHeight();
  docProperties[9] = element.getWidth();
  docProperties[10] = element.getTop();
  docProperties[11] = element.getLeft();
  docProperties[12] = selection.getTextRange().asString();
  
  return docProperties;

}
function setTextProp() {
  var getProp = getTextProp();
  //Logger.log(getProp[0]+' '+getProp[1]+' '+getProp[2]);
  boldSelectedElements(getProp[0]);
  italicizeSelectedElements(getProp[1]);
  underlineSelectedElements(getProp[2]);
  changeFontSize(getProp[3]);
  changeFontType(getProp[4]);
  changeTextBackgroundColor(getProp[5]);
  changeTextForegroundColor(getProp[6]);

}

function alignParagraphText() {
  
  SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide().getShapes().forEach(function(shape) {
    shape.getText().getParagraphs().forEach(function(paragraph) {
      paragraph.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    })
  });
}

/////////////////////////////////////////////Alignment like Master \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function getMasterAlignment() {
  var alignments = [];
  var i=0;
  var master = SlidesApp.getActivePresentation().getMasters()[0];
  master.getPageElements().forEach(function(pageElement) {
    alignments[i]=pageElement.asShape().getText().getParagraphStyle().getParagraphAlignment();
    i++;
  })
  return alignments;
}
function setAlignmentsAsMaster() {
  var alignments = getMasterAlignment();
  alignTextLikeMaster(alignments);
}
function alignTextLikeMaster(masterAlignments) {
  var i=0;
  SlidesApp.getActivePresentation().getSelection().getCurrentPage().getPageElements().forEach(function(pageElement) {
    pageElement.asShape().getText().getParagraphStyle().setParagraphAlignment(masterAlignments[i]);
      i++;
    
  });
}
/////////////////////////////////////////////////////Positions and Dimensions like Master\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

function getMasterPositionDimension() {

  var master = SlidesApp.getActivePresentation().getMasters()[0];
  var temp=[];
  var i=0;
  Logger.log('Number of placeholders in the master: ' + master.getPlaceholders().length);
  master.getPlaceholders().forEach(function(pageElement) {

    temp.push([pageElement.getLeft(),pageElement.getHeight(),pageElement.getTop(),pageElement.getWidth()]);
    i++;
  });
  return temp;
}
function setPositionsDimensionsAsMaster() {
  
  var positionsAndDimensions =  getMasterPositionDimension();
  resizeAndPositionLikeMaster(positionsAndDimensions);

}

function resizeAndPositionLikeMaster(masterIndents) {
  
  
  var currPage=SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var pageElements=currPage.getPageElements();
  
  var i=0;
  pageElements.forEach(function(pageElement) {
    
    if(masterIndents[i]) {
      pageElement.setLeft(masterIndents[i][0]);
      pageElement.setHeight(masterIndents[i][1]);
      pageElement.setTop(masterIndents[i][2]);
      pageElement.setWidth(masterIndents[i][3]);
    }
      
    i++;
  });
}

//////////////////////////////////////////////////Predictions\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function predictBold(currentProperties) {
  var categoryAttr = 'bold';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  //Logger.log("getTextProp() says: current element is bold: "+selectedProp[0]+", italic: "+selectedProp[1]+", underline: "+selectedProp[2]);
  var currElementInput = {italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3],fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  //var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  var prediction = parseProperties(categoryAttr,ignoredAttr, currElementInput);
  prediction = prediction.substring(prediction.indexOf('should be')+10,prediction.length-4);
  if(prediction === 'true'){
    prediction = 'The selected text should be bold.';
  }
  else{
    prediction = 'The selected text should not be bold.'
  }
  return prediction;
}
function predictItalic() {
  var categoryAttr = 'italic';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0],underline: selectedProp[2], fontsize: selectedProp[3],fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = parseProperties(categoryAttr,ignoredAttr, currElementInput);
  prediction = prediction.substring(prediction.indexOf('should be')+10,prediction.length-4);
  if(prediction === 'true'){
    prediction = 'The selected text should be italicized.';
  }
  else{
    prediction = 'The selected text should not be italicized.'
  }
  return prediction;
}
function predictUnderline() {
  var categoryAttr = 'underline';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1],fontsize: selectedProp[3],fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = parseProperties(categoryAttr,ignoredAttr, currElementInput);
  prediction = prediction.substring(prediction.indexOf('should be')+10,prediction.length-4);
  if(prediction === 'true'){
    prediction = 'The selected text should be underlined.';
  }
  else{
    prediction = 'The selected text should not be underlined.'
  }
  return prediction;
}
function predictFontSize() {
  var start = new Date().getTime();
  var categoryAttr = 'fontsize';
  var ignoredAttr = null;
  var selectedText = SlidesApp.getActivePresentation().getSelection().getTextRange().asString();
  Logger.log('Selected Text: '+selectedText);
  var selectedProp = getTextProp();
  
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2],fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = parseProperties(categoryAttr,ignoredAttr, currElementInput);
  prediction = prediction.substring(prediction.indexOf('should be')+10,prediction.length-4);
  Logger.log('Predicted font size: '+prediction);
  var end = new Date().getTime();
  Logger.log("Time taken for predictFontSize(): "+(end - start) + ' miliseconds.');
  return "The font size of the selected text should be "+prediction;
}
function predictFontSizeBeforeRuns() {
  var start = new Date().getTime();
  var categoryAttr = 'fontsize';
  var ignoredAttr = 'fgcolor';
  var selectedText = SlidesApp.getActivePresentation().getSelection().getTextRange().asString();
  Logger.log('Selected Text: '+selectedText);
  var selectedProp = getTextProp();
  
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  Logger.log('Predicted font size: '+prediction);
  var end = new Date().getTime();
  Logger.log("Time taken for predictFontSize(): "+(end - start) + ' miliseconds.');
  return prediction;
}
function predictTop() {
  var categoryAttr = 'top';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2],fontsize: selectedProp[3], fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], width: selectedProp[9], left: selectedProp[11]};
  var prediction = parseProperties(categoryAttr,ignoredAttr, currElementInput);
  
  
  return prediction;
}
function predictLeft() {
  var categoryAttr = 'left';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2],fontsize: selectedProp[3], fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10]};
  var prediction = parseProperties(categoryAttr,ignoredAttr, currElementInput);
  
  return prediction;
}

function predictWidth(){
  var categoryAttr = 'width';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2],fontsize: selectedProp[3], fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], height: selectedProp[8], top: selectedProp[10],left: selectedProp[11]};
  var prediction = parseProperties(categoryAttr, ignoredAttr, currElementInput);
  return prediction;
}

function predictHeight(){
  var categoryAttr = 'height';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2],fontsize: selectedProp[3], fontfamily: selectedProp[4], fgcolor: selectedProp[5], bgcolor: selectedProp[6], width: selectedProp[9], top: selectedProp[10],left: selectedProp[11]};
  var prediction = parseProperties(categoryAttr, ignoredAttr, currElementInput);
  return prediction;
}

function fixPosition() {
  
  var changedPosition='';
  var top = predictTop();
  top = top.substring(top.indexOf("should be")+10,top.length-4);
  var left = predictLeft();
  left = left.substring(left.indexOf("should be")+10,left.length-4);
  var width = predictWidth();
  width = width.substring(width.indexOf("should be")+10, width.length-4);
  var height = predictHeight();
  height = height.substring(height.indexOf("should be")+10, height.length-4);
  var currProp = getTextProp();
  var currTop = currProp[10];
  var currLeft = currProp[11];
  var currWidth = currProp[9];
  var currHeight = currProp[8];
  Logger.log("Left is: "+left+", Top is: "+top+", Width is: "+width+", Height is: "+height);
  if(currTop===top && currLeft===left) {
    changedPosition ='Positioned perfectly.'
  }
  else {
    var element = SlidesApp.getActivePresentation().getSelection().getPageElementRange().getPageElements()[0] ;
    element.setLeft(left);
    var leftTilt = (currProp[11]-left).toFixed(2);
    element.setTop(top);
    var topTilt = (currTop-top).toFixed(2);
    if(leftTilt < 0){
      changedPosition = 'Shifted element '+Math.abs(leftTilt)+' points to the right. ';
    }
    else{
      changedPosition = 'Shifted element '+leftTilt+' points to the left. ';
    }
    if(topTilt < 0){
      changedPosition += 'Shifted element '+Math.abs(topTilt)+' points downward. ';
    }
    else{
      changedPosition += 'Shifted element '+topTilt+' points upward. ';
    }
  }
  if(currWidth===width && currHeight===height){
    changedPosition += 'Sized perfectly.'
  }
  else{
    var element = SlidesApp.getActivePresentation().getSelection().getPageElementRange().getPageElements()[0] ;
    element.setHeight(height);
    element.setWidth(width);
    var wChange = (currWidth-width).toFixed(2);
    var hChange = (currHeight-height).toFixed(2);
    if(wChange<0){
      changedPosition += 'Expanded element width by '+Math.abs(wChange)+' points. ';
    }
    else{
      changedPosition += 'Shortened element width by '+wChange+' points. ';
    }
    if(hChange<0){
      changedPosition += 'Expanded element height by '+Math.abs(hChange)+' points. ';
    }
    else{
      changedPosition += 'Shortened element height by '+hChange+' points. ';
    }
  }
  return changedPosition;
}

function parseProperties(category, ignore, currElementInput){
  var start = new Date().getTime();
  var presentation = SlidesApp.getActivePresentation();
  var currSlideNumber = parseInt(presentation.getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '')-1);
  var lastStoredIndex = scriptProperties.getProperty("lastStoredIndex");
  if(currSlideNumber>lastStoredIndex) {
    storeProperties(currSlideNumber);
  }
  var data = [];
 
  for(var i=0;i<(currSlideNumber-1);i++) {    
    slideNo="Slide: "+(i+1);
    slideNumber = '' + i;
    var jsonSlide = splitString(i);
    jsonSlide.forEach(function(p){
      var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fontfamily: p[4],fgcolor: p[5],bgcolor: p[6],height: p[8], width: p[9], top: p[10], left: p[11]};
      data.push(newObj);
    });

  }
  var startTrain = new Date().getTime();
  var result = trainSet(data, category, ignore, currElementInput);
  var endTrain = new Date().getTime();
  Logger.log('trainSet() takes '+(endTrain-startTrain)+' milliseconds to complete.');
  var end = new Date().getTime();
  Logger.log('parseProperties() takes '+(end-start)+' milliseconds to complete.');
  return result;
}

function extractTextProperties(category, ignore, currElementInput) {
  var start = new Date().getTime();
  var presentation = SlidesApp.getActivePresentation();
  var slides = [];
  var tempOutput= [[]];
  var output= [[]];
  var p = [];
  tempOutput.pop();
  output.pop();
  output.push(["Bold", "Italic", "Underline", "Font Size", "Font Family", "Foreground Color", "Background Color", "Element Type", "Element Height", "Element Width", "Element position from top", "Element position from left"]);
  var slideNo="";
  var i=0,k=0;
  var slidek = 0;
  var data = [];
  var texts = [];
  slides=presentation.getSlides();
  
 
  var currSlideNumber = SlidesApp.getActivePresentation().getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '');
  Logger.log("Currently at slide: "+currSlideNumber);
  for(i=0;i<currSlideNumber;i++) {
    var lStart = new Date().getTime();
    slideNo="Slide: "+(i+1);
    //Logger.log(slideNo+", Slide ID: "+slides[i].getObjectId().toString());
    
    
    slides[i].getPageElements().forEach(function(pageElement) {
      
      
      if(pageElement.getPageElementType().toString()==='SHAPE')
      {
        //pageElement.select();
        //Logger.log(pageElement.getObjectId()+" -> "+pageElement.getPageElementType());
        
        //var texts = pageElement.asShape().getText().asString().split(" ");
      
        var textStyle= pageElement.asShape().getText().getTextStyle();
      
        p[0]=textStyle.isBold();
        p[1]= textStyle.isItalic();
        p[2]=textStyle.isUnderline();
        p[3]=textStyle.getFontSize();
        p[4]=textStyle.getFontFamily();
        p[5]=textStyle.getForegroundColor();
        p[6]=textStyle.getBackgroundColor();
        p[7]=pageElement.getPageElementType();
        p[8]=pageElement.getHeight();
        p[9]=pageElement.getWidth();
        p[10]=pageElement.getTop();
        p[11]=pageElement.getLeft();
        
        if(p[3]!==null )//&& p[10]<400 && p[11]<400
        {
          
          output.push([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]);
          //tempOutput.push(([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]));
          //Logger.log(tempOutput.pop());
          var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5], height: p[8], width: p[9], top: p[10], left: p[11]};
          data.push(newObj);
          texts.push(pageElement.asShape().getText().asString());
          k++;
          slidek++;
        }
        
      } 
      
    });
    var lEnd = new Date().getTime();
    // Logger.log('Time taken for slide '+(i+1)+': '+(lEnd-lStart)+' miliseconds.');
    // Logger.log('Page Elements per Slide: '+slidek);
    slidek = 0;
  }
  Logger.log("Number of page elements: "+k);
  

   for(var j=0;j<=k;j++)
   {
    
      // Logger.log(output[j]);   
  
   }
   var trainSetStart = new Date().getTime();
   var result = trainSet(data, category, ignore, currElementInput);
   var trainSetEnd = new Date().getTime();
   Logger.log('trainSet() takes '+(trainSetEnd-trainSetStart)+' miliseconds to compute.');
   var end = new Date().getTime();
   return result;
}

function extractTextPropertiesAfterRuns(category, ignore, currElementInput) {
  var start = new Date().getTime();
  var presentation = SlidesApp.getActivePresentation();//.getSlideById()[0];
  var slides = [];
  var tempOutput= [[]];
  var output= [[]];
  var p = [];
  tempOutput.pop();
  output.pop();
  output.push(["Bold", "Italic", "Underline", "Font Size", "Font Family", "Foreground Color", "Background Color", "Element Type", "Element Height", "Element Width", "Element position from top", "Element position from left", "Run"]);
  var slideNo="";
  var i=0,k=0;
  var data = [];
  slides=presentation.getSlides();
  
  //var noOfSlidesBeforePresent = slides.length-50;
  var currSlideNumber = SlidesApp.getActivePresentation().getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '');
  Logger.log("Currently at slide: "+currSlideNumber);
  for(i=0;i<currSlideNumber;i++) {
    var lStart = new Date().getTime();
    slideNo="Slide: "+(i+1);
    //Logger.log(slideNo+", Slide ID: "+slides[i].getObjectId().toString());
    
    
    slides[i].getPageElements().forEach(function(pageElement) {
      
      
      if(pageElement.getPageElementType().toString()==='SHAPE')
      {
          var textRange = pageElement.asShape().getText(); // Get text belonging to current page element
          var elementType =pageElement.getPageElementType();
          var elementHeight=pageElement.getHeight();
          var elementWidth=pageElement.getWidth();
          var elementTop=pageElement.getTop();
          var elementLeft=pageElement.getLeft();
          textRange.getRuns().forEach(function(run) { // Loop through all runs in text
          var textStyle = run.getTextStyle(); // Get current row text style
          if(run.getLength()>1)
          {
            //Logger.log('Run: '+run.asString()+'--> Bold: '+textStyle.isBold()+', Italic: '+textStyle.isItalic()+', Underline: '+textStyle.isUnderline());
            p[0]=textStyle.isBold();
            p[1]= textStyle.isItalic();
            p[2]=textStyle.isUnderline();
            p[3]=textStyle.getFontSize();
            p[4]=textStyle.getFontFamily();
            p[5]=textStyle.getForegroundColor();
            p[6]=textStyle.getBackgroundColor();
            p[7]=elementType;
            p[8]=elementHeight;
            p[9]=elementWidth;
            p[10]=elementTop;
            p[11]=elementLeft;
            p[12]=run.asString();
            output.push([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11],p[12]]);
          //tempOutput.push(([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]));
          //Logger.log(tempOutput.pop());
          var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5], height: p[8], width: p[9], top: p[10], left: p[11], run: p[12]};
          data.push(newObj);
          k++;
          }
        });
        
        
      } 
      
    });
    var lEnd = new Date().getTime();
    // Logger.log('Time taken for slide '+(i+1)+': '+(lEnd-lStart)+' miliseconds.');
  }
  Logger.log("Number of page elements: "+k);
  

   for(var j=0;j<=k;j++)
   {
    
      // Logger.log(output[j]);   
  
   }
  var result = trainSet(data, category, ignore, currElementInput);
  Logger.log('extractProperties returns: '+result);
  var end = new Date().getTime();
  Logger.log('Time taken for extractProperties(): '+ (end-start)+' miliseconds.');
  return result;

}

function comboExtract(ignore){
  var start = new Date().getTime();
  var data = [];
  var currSlideNumber = scriptProperties.getProperty("currSlideNumber");
  for(var i=0;i<(currSlideNumber-1);i++) {    
    slideNo="Slide: "+(i+1);
    slideNumber = '' + i;
    var jsonSlide = splitString(i);
    jsonSlide.forEach(function(p){
      var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5], height: p[8], width: p[9], top: p[10], left: p[11]};
      data.push(newObj);
    });

  }
  var boldTree = TreeExtraction(data,'bold',ignore);
  var italicTree = TreeExtraction(data,'italic',ignore);
  var underlineTree = TreeExtraction(data,'underline',ignore);
  var fontSizeTree = TreeExtraction(data,'fontsize',ignore);
  var rets = [data,boldTree,italicTree,underlineTree,fontSizeTree];
  return rets;
}

function TreeExtraction(data,category,ignore){
  var start = new Date().getTime();
  var config = {
    trainingSet: data,
    categoryAttr: category,
    ignoredAttributes: [ignore]
  
  };
  var decisionTree = new dt.DecisionTree(config);
  return decisionTree;
}

function trainSet(data, category, ignore, currElementInput) {
  var start = new Date().getTime();
  var config = {
    trainingSet: data,
    categoryAttr: category,
    ignoredAttributes: [ignore]
  
  };
  var decisionTree = new dt.DecisionTree(config);
  var testPredict = {italic: true, underline: false, fontsize: 14, fgcolor: 'Color'};
  var selectedProp = getTextProp();
  //var currElementInput = {italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5]};
  var decisionTreePrediction = decisionTree.predict(currElementInput); 
  decisionTreePrediction = decisionTreePrediction.replace("undefined",category);
  // var sim = findSimilarElements(currElementInput, decisionTree, decisionTreePrediction.substring(decisionTreePrediction.indexOf("should be")+10,decisionTreePrediction.length-4),category);
  var end = new Date().getTime();
  return decisionTreePrediction;

  
}

//Helper for random number generator
function RandomInt(max){
  return parseInt(Math.floor(Math.random() * max)+"");
}

//This method is done by Somasekhar, modified by Safwat
function assessFontSizePrediction() {
  var count = 0;
  var numOfSamples = 10;
  var correctPredictions = 0;
  var totalDeviation = 0;
  var percDeviation;
  var returnString = "Accuracy level: ";
  var accuracy;
  var page;
  var element;
  while(count<numOfSamples){
    page = Math.floor(Math.random() * 56)+10; //Last slide index = 66, starting from slide 10 so trainSet can get enough data
    element = Math.floor(Math.random() * 3); //Shuffling between 0-2 in element numbers
    var shapeElement = selectTextRun(page, element);
    if(shapeElement==false) //checking if element is of type SHAPE, if not then continue loop
    {
        continue;
    }
    var props = getTextProp();
    var inp = {bold: props[0], italic: props[1], underline: props[2], height: props[8], width: props[9], top: props[10], left: props[11]};
    var observed = extractTextProperties('fontsize','fgcolor', inp);
    Logger.log('Predicted Value: '+observed);
    var expected = props[3];
    Logger.log('Expected Value: '+expected);
    if(observed==expected)
    {
        correctPredictions++;
    }
    var deviation = Math.abs(100*(observed-expected)/expected);
    
    totalDeviation += deviation;
    count++;
  }
  accuracy = correctPredictions/numOfSamples*100;
  percDeviation = totalDeviation/numOfSamples;
  returnString+=accuracy;
  returnString+="%, Percentage of Deviation: ";
  returnString+=percDeviation;
  returnString+="%";
  return returnString;
}

//This method is done by Stuti, modified by Safwat

function assessBoldPrediction() {
  
  var numOfSamples = 10;
  var count = 0;
  var correctPredictions=0;
  var returnString="Accuracy level: ";
  var accuracy;
  while(count<numOfSamples)
  {
    page = Math.floor(Math.random() * 56)+10; //Last slide index = 66, starting from slide 10 so trainSet can get enough data
    element = Math.floor(Math.random() * 3); //Shuffling between 0-2 in element numbers
    var shapeElement = selectTextRun(page, element);
    if(shapeElement==false) //checking if element is of type SHAPE, if not then continue loop
    {
        continue;
    }
    var actual_value = getTextProp()[0];
    var predicted_value = predictBold();
    if (String(actual_value) == String(predicted_value)) {
      correctPredictions++;
    }
    count++;
  }
  accuracy=count/10*100;
  returnString+=accuracy;
  returnString+="%";
  return returnString;
}

function selectTextRun(page,element) {
  var slide = SlidesApp.getActivePresentation().getSlides()[page];
  slide.selectAsCurrentPage();
  var pageElement = slide.getPageElements()[element];
  if(pageElement.getPageElementType()=="SHAPE")
  {
    var textRange = pageElement.asShape().getText();
    var run = textRange.getRuns()[0];
    run.select();
  }
  else return false;
  
}

function simSelect(page,element) {
  var slide = SlidesApp.getActivePresentation().getSlides()[page];
  // slide.selectAsCurrentPage();
  var pageElement = slide.getPageElements()[element];
  if(pageElement === undefined){
    return false;
  }
  if(pageElement.getPageElementType()=="SHAPE")
  {
    var textRange = pageElement.asShape().getText();
    var run = textRange.getRuns()[0];
    // run.select();
    var textStyle = run.getTextStyle();
    var candidate = {bold: (textStyle.isBold()==null?false:textStyle.isBold()), italic: (textStyle.isItalic()==null?false:textStyle.isItalic()), underline: (textStyle.isUnderline()==null?false:textStyle.isUnderline()),fontsize: textStyle.getFontSize(),fgcolor: textStyle.getForegroundColor(), height: pageElement.getHeight(), width: pageElement.getWidth(), top: pageElement.getTop(), left: pageElement.getLeft(), run: run.asString()};
    return candidate;
  }
  else return false;
  
}

function sScore(fElement, sElement){
  var score = 0;
  var props = Object.keys(fElement);
  for(var i = 0;i<props.length;i++){
    var attr = props[i];
    if(attr === "fontsize" || attr === "height" || attr === "top" || attr === "left" || attr === "width"){
      var error = (sElement[attr]-fElement[attr])/(fElement[attr])
      // Logger.log(error);
      score += error;
    }
    else if ((attr === "bold" || attr === "italic" || attr === "underline") && (fElement[attr] !== sElement[attr])){
      score += (1/7);
    }
  }
  return score;
}
function findSimilarElements(currElementInput,trees,comps){
  Logger.log("In the Sim");
  var count = 0
  var retu = "Similar Element: ";
  var currSlideNumber = SlidesApp.getActivePresentation().getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '');
  Logger.log(currSlideNumber)
  while(count<1){
    var pa = Math.floor(Math.random() * currSlideNumber); 
    var e = Math.floor(Math.random() * SlidesApp.getActivePresentation().getSlides()[pa].getPageElements().length);
    var res = simSelect(pa,e);
    while(res === false){
      pa = Math.floor(Math.random() * currSlideNumber); 
      e = Math.floor(Math.random() * 3);
      res = simSelect(pa,e);
    }
    var simScore = sScore(currElementInput,res);
    if(Math.abs(simScore)<0.2){
      var btemp = res['bold'];
      var itemp = res['italic']
      var utemp = res['underline']
      var ftemp = res['fontsize']
      delete res['bold'];
      var compb = trees[0].predict(res);
      compb = compb.substring(compb.indexOf("should be")+10,compb.length-4);
      res['bold'] = btemp;
      delete res['italic'];
      var compi = trees[1].predict(res);
      compi = compi.substring(compi.indexOf("should be")+10,compi.length-4);
      res['italic'] = itemp;
      delete res['underline'];
      var compu = trees[2].predict(res);
      compu = compu.substring(compu.indexOf("should be")+10,compu.length-4);
      res['underline'] = utemp;
      delete res['fontsize'];
      var compf = trees[3].predict(res);
      compf = compf.substring(compf.indexOf("should be")+10,compf.length-4);
      res['fontsize'] = ftemp;
      var samecount = 0
      if(compb===comps[0]){
        samecount++;
      }
      if(compi===comps[1]){
        samecount++;
      }
      if(compu===comps[2]){
        samecount++;
      }
      if(compf===comps[3]){
        samecount++;
      }
      if(samecount>=3){
        retu += res['run'] +" @ "+"Slide "+(pa+1);
        count += 1;
      }
    }
  }
  return retu;
}

function accuracyCheck(category, tree, currElementInput){
  var checking = currElementInput[category]+"";
  delete currElementInput[category];
  var res = tree.predict(currElementInput);
  res = res.substring(res.indexOf("should be")+10,res.length-4);
  var error = 0;
  if(category === "fontsize"){
    error = 1 - Math.abs((res-checking)/(checking));
  }
  else if(checking === res){
    error = 1;
  }
  else{
    error = 0;
  }
  currElementInput[category] = checking;
  return error;
}

function comboPredict(){
  Logger.log("Welcome");
  var predictions = "";
  var selectedProp = getTextProp();
  var boldInput = {italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var italicInput = {bold: selectedProp[0], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var underlineInput = {bold: selectedProp[0], italic: selectedProp[1], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var fontSizeInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var ret = comboExtract('fgcolor');
  Logger.log("Finished Extraction");
  var data = ret[0];
  Logger.log("Size of Dataset: "+data.length);
  var trees = ret.slice(1,5);
  var b = trees[0].predict(boldInput);
  b = b.substring(b.indexOf("should be")+10,b.length-4);
  var it = trees[1].predict(italicInput);
  it = it.substring(it.indexOf("should be")+10,it.length-4);
  var u = trees[2].predict(underlineInput);
  u = u.substring(u.indexOf("should be")+10,u.length-4);
  var f = trees[3].predict(fontSizeInput);
  f = f.substring(f.indexOf("should be")+10,f.length-4);
  Logger.log(b+" "+selectedProp[0])
  Logger.log(it+" "+selectedProp[1])
  Logger.log(u+" "+selectedProp[2])
  Logger.log(f+" "+selectedProp[3])
  if(b !== selectedProp[0]+""){
    if(b==='false'){
      predictions += "The selected element should not be bold, "
    }
    else{
      predictions += "The selected element should be bold, "
    }
  }
  if(it !== selectedProp[1]+""){
    if(it==='false'){
      predictions += "should not be italicized, "
    }
    else{
      predictions += "should be italicized, "
    }
  }
  if(u !== selectedProp[2]+""){
    if(u==='false'){
      predictions += "should not be underlined, "
    }
    else{
      predictions += "should be underlined, "
    }
  }
  if(f !== selectedProp[3]+""){
    predictions += "font size should be "+f+"."
  }
  if(predictions === ""){
    predictions = "No changes needed";
  }
  Logger.log("Finished Compilation");
  return predictions;
}

function logicExplain(){
  var exp = "";
  var info = comboExtract('fgcolor');
  var data = info[0];
  var trees = info.slice(1,5);
  var boldacc = 0;
  var italicacc = 0;
  var underlineacc = 0;
  var fontsizeacc = 0;
  var count = 0;
  var selectedProp = getTextProp();
  var boldInput = {italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var italicInput = {bold: selectedProp[0], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var underlineInput = {bold: selectedProp[0], italic: selectedProp[1], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var fontSizeInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var bigInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var b = trees[0].predict(boldInput);
  b = b.substring(b.indexOf("should be")+10,b.length-4);
  var it = trees[1].predict(italicInput);
  it = it.substring(it.indexOf("should be")+10,it.length-4);
  var u = trees[2].predict(underlineInput);
  u = u.substring(u.indexOf("should be")+10,u.length-4);
  var f = trees[3].predict(fontSizeInput);
  f = f.substring(f.indexOf("should be")+10,f.length-4);
  var combles = [b,it,u,f];
  for(var i = 0;i<data.length;i+=3){
    count += 1
    boldacc += accuracyCheck('bold',trees[0],data[i]);
    italicacc += accuracyCheck('italic',trees[1],data[i]);
    underlineacc += accuracyCheck('underline',trees[2],data[i]);
    fontsizeacc += accuracyCheck('fontsize',trees[3],data[i]);
  }
  boldacc *= (100/count);
  italicacc *= (100/count);
  underlineacc *= (100/count);
  fontsizeacc *= (100/count);
  exp += "bold: (acc: " + boldacc+"%) <br>";
  exp += "italic: (acc: " + italicacc+"%) <br>";
  exp += "underline: (acc: " + underlineacc+"%) <br>";
  exp += "fontsize: (acc: " + fontsizeacc+"%) <br><br>";
  var sim = findSimilarElements(bigInput,trees,combles);
  exp += sim;
  return exp;
}

function getColorString(textStyle,ground) {
  var groundColor;
  if(ground==='Foreground') groundColor = textStyle.getForegroundColor();
  else if(ground==='Background') groundColor = textStyle.getBackgroundColor()
  var hexString;
  if(groundColor.getColorType().toString()==='RGB') 
    hexString = groundColor.asRgbColor().asHexString();
  else if(groundColor.getColorType().toString()==='THEME') 
    hexString = groundColor.asThemeColor().getThemeColorType().toString();
  return hexString;
}