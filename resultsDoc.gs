function resultsDoc(){
 var ssa = SpreadsheetApp;
 var ss = ssa.getActiveSpreadsheet();
 var sheets = ss.getSheets();


  var ui = ssa.getUi();
  

  var doca= DocumentApp;
  var drive = DriveApp;
  var template = "17Yhu-39iHAL7-VK--ZukXUuV6oJj6qLV_OB0n3ECLKo";
  var source = doca.openById(template);
  var resDocCopy =drive.getFileById(template).makeCopy('Tournament Results');
  var resDocId = resDocCopy.getId();
  var resDoc = doca.openById(resDocId);
  var docBody = resDoc.getBody();
  var resDocUrl = resDoc.getUrl();
  var asset = drive.getFileById(resDocId);
  sheets.shift();
  for(var dis in sheets){
  var thisSheet = sheets[dis];
  var letters = ["a","b","c","d","e","f","g","h","i","j","k","f"];
  var disLetter = letters[dis];
 var eventa = thisSheet.getName(); 
  var aOne = thisSheet.getRange('B3').getValue()+"  "+thisSheet.getRange('A3').getValue();
  var aTwo = thisSheet.getRange('B4').getValue()+"  "+thisSheet.getRange('A4').getValue();
  var aThree= thisSheet.getRange('B5').getValue()+"  "+thisSheet.getRange('A5').getValue();
  var aFour= thisSheet.getRange('B6').getValue()+"  "+thisSheet.getRange('A6').getValue();
  var aFive= thisSheet.getRange('B7').getValue()+"  "+thisSheet.getRange('A7').getValue();
  var aSix= thisSheet.getRange('B8').getValue()+"  "+thisSheet.getRange('A8').getValue();
  
  docBody.replaceText('<<'+disLetter+'>>',eventa);
  docBody.replaceText('<<'+disLetter+'1>>',aOne);
  docBody.replaceText('<<'+disLetter+'2>>',aTwo);
  docBody.replaceText('<<'+disLetter+'3>>',aThree);
  docBody.replaceText('<<'+disLetter+'4>>',aFour);
  docBody.replaceText('<<'+disLetter+'5>>',aFive);
  docBody.replaceText('<<'+disLetter+'6>>',aSix);
  
  }
  ss.insertSheet("Links");
  var linksSheet = ss.getSheetByName("Links");
  linksSheet.getRange(1,1).setValue("Link to results page:");
  linksSheet.getRange(2,1).setValue(resDocUrl);
  linksSheet.getRange(2,2).setValue(resDocId);
  linksSheet.hideColumn(linksSheet.getRange("B1"));
  asset.setSharing(drive.Access.ANYONE,drive.Permission.EDIT);
  resDoc.saveAndClose();

}
function sweepsDoc(){
  uiSweeps();
   var doca= DocumentApp;
  var linksSheet = ss.getSheetByName("Links");
  var resDocId = linksSheet.getRange(2,2).getValue();
  Logger.log(resDocId);
  var resDoc = doca.openById(resDocId);
  var docBody = resDoc.getBody();
  
  var sweepsSheet = ss.getSheetByName('Sweeps');
  
  docBody.replaceText("<<sweeps1>>",sweepsSheet.getRange('B2').getValue());
  docBody.replaceText("<<sweeps2>>",sweepsSheet.getRange('B3').getValue());
  docBody.replaceText("<<sweeps3>>",sweepsSheet.getRange('B4').getValue());

 
}