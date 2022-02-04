var ssa = SpreadsheetApp;
var ss = ssa.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var ui = ssa.getUi();
var sNum = ss.getSheets().length -1
 var tRange = "A3:Z";

var sweepsMenu = HtmlService
.createHtmlOutputFromFile('sweepsMenu')
.setWidth(350)
.setHeight(150);


function onOpen(){
//Create the menu  
  var menu =  [{name:"Round 1", functionName:"oneEntry"},{name:"Round 2", functionName:"twoEntry"},{name:"Round 3",functionName:"threeEntry"},{name: "Break", functionName: "fBreak"},{name: "Finals", functionName: "finals"},
               {name: "Sweeps", functionName: "uiSweeps"},{name: "Results Doc" , functionName: "printResults"},];
  ss.addMenu("Tabs",menu);
  
}
               
function oneEntry(){

var col1 = 4;

  for(i=0;i<sNum;i++){
    var thisSheet = ss.getSheets()[i];
    var thisRange =thisSheet.getRange(tRange);

    if(thisSheet.getLastColumn() > col1){
    thisRange.sort(col1);
    }
  }
  sweeps();
}
               
function twoEntry(){
  
  var col1 = 7;

  for(i=0;i<sNum;i++){
    var thisSheet = ss.getSheets()[i];
    var thisRange =thisSheet.getRange(tRange)
     if(thisSheet.getLastColumn() > col1){
    thisRange.sort(col1);
    }
  
  }
  sweeps();
}
               
function threeEntry(){
var col1 = 10;

  for(i=0;i<sNum;i++){
    var thisSheet = ss.getSheets()[i];
    var thisRange =thisSheet.getRange(tRange)
    if(thisSheet.getLastColumn() > col1){
    thisRange.sort(col1);
    }
  
  }
  sweeps();
}

function finals(){

var col1 = 24;
var col2= 23;
  for(i=0;i<sNum;i++){
    var thisSheet = ss.getSheets()[i];
    var thisRange =thisSheet.getRange(tRange)
    if(thisSheet.getLastColumn()>col1){
    thisRange.sort([{column: col1, ascending: true},{column: col2, ascending:false}]);
    }
  }
  sweeps();
}

function sweeps(){
  var formulae = [];
  
ssa.flush();
  var sweepSheet = ss.getSheetByName("Sweeps");
  var teams = sweepSheet.getDataRange().getValues();
  teams = transpose(teams);
  teams = teams[0];
  teams = clearBlanks(teams);
  teams.shift();
  var formulae = fixSweeps(teams);
  formulae = oneToTwo(formulae);
  sweepSheet.getRange(2, 3, teams.length).setFormulas(formulae);
 
  
  
}
function fBreak(){
var col1 = 14;
var col2= 15;
  for(i=0;i<sNum;i++){
    var thisSheet = ss.getSheets()[i];
    var thisRange =thisSheet.getRange(tRange)
    if(thisSheet.getLastColumn() > col1){
    thisRange.sort([{column: col1, ascending: true},{column: col2, ascending:false}]);
    }
  
  }
  sweeps();
}
function printResults(){
  var sweepSheet = ss.getSheetByName("Sweeps");
  var linksSheet = ss.getSheetByName("Links");
  if (linksSheet != null){
   var resDocId = linksSheet.getRange(2,2).getValue();
   
    DriveApp.getFileById(resDocId).setTrashed(true);
    
    ss.deleteSheet(linksSheet);}
finals();
  ssa.flush();
  sweepSheet.activate();
 resultsDoc();
  
  ssa.flush();
  sweepsDoc();
  //ss.getSheetByName('Sheet1').hideSheet();
  
  
}

function fixSweeps(teams){
ssa.flush();
var formulae=[];
  for(q=0;q<teams.length;q++){
   formulae.push("=sum(") 
  }
var sheets = ss.getSheets();
sheets.pop();
sheets.shift();
  

  for (i=0;i<sheets.length;i++){

   var sheet = sheets[i];
  var data = sheets[i].getDataRange().getValues();
   data = transpose(data);
  
  var codes = clearBlanks(data[1]);
   
    for (t=0;t<teams.length;t++){
  var team = teams[t];
      var formula = formulae[t];  
    for (j=1;j<codes.length;j++){
     var row = j*1+2;
     var tCode = codes[j];
   
    var alpha = tCode.slice(-1);
    var teamCode = tCode.split(alpha)[0];
      if(teamCode == team){
       formula = formula +sheet.getName()+ "!Y"+ row + "+"; 
       
      }
    
    }

  
      formulae[t]=formula;
    }
  
  }
  for(w=0;w<formulae.length;w++){
   formulae[w] = formulae[w]+"0)";
  }
  Logger.log(formulae);
 return formulae; 

}  


function transpose(a){
 
  
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function clearBlanks(array){
     var newArray = [];
     for(n=0;n<array.length;n++){
            if(array[n]!=""){newArray.push(array[n]);}
                                     }
  return newArray;}

function oneToTwo(array){
    var newArray = [];
  
  for(i=0;i<array.length;i++){
var arrayItem = [];
  arrayItem.push(array[i]);
  newArray.push(arrayItem);
  
  } 
  return newArray;
}

function selectSweeps(num){
  
  var sweepCounts = {};
  var sheets = ss.getSheets();
  var sweepSheet = ss.getSheetByName("Sweeps");
  var teams = sweepSheet.getDataRange().getValues();
  teams = transpose(teams);
  teams = teams[0]; 
  teams = clearBlanks(teams);
  
  for (var s in sheets){
    var sheet = sheets[s];
    if (sheet.getName() == "Sweeps"){
      continue
    }
    var sRange = sheet.getRange(3, 3, 150, 1);
    sRange.setValue('x');
    
  }
  
  for (i in teams){
    var team = teams[i];
    
    sweepCounts[team] = [];
    
  }
  Logger.log(sweepCounts)
  for(var s in sheets){
    var sheet = sheets[s];
    var data = sheet.getDataRange().getValues();
    data = transpose(data);
    var perCodes = data[1];
    if (perCodes == undefined){
     continue 
    }
    perCodes = clearBlanks(perCodes);
    
    var scores = data[24];
    if (scores == undefined){
     continue 
    }
    scores.shift();
    
    for (var p in perCodes){
      var tCode = perCodes[p];
      var alpha = tCode.slice(-1);
      var teamCode = tCode.split(alpha)[0];
      var score = scores[p];
      
      sweepCounts[teamCode].push(score)
    
    
    }
    
    
    
   
  }
  var teamScores = {};
  for (var key in sweepCounts){
    var tops = topNum(sweepCounts[key],num);
    sweepCounts[key] = tops
    var tScore = addArray(sweepCounts[key]);
    if (isNaN(tScore)){continue}
    teamScores[key] = tScore      
  }
  
  var scoreData = sweepSheet.getDataRange().getValues();
  scoreData = clearBlanks(scoreData);
  for (var u in scoreData){
   var thisTeam = scoreData[u][0];
    if (thisTeam == 'Code'){
    continue
    }
   scoreData[u][2] = teamScores[thisTeam];
   
  }
  sweepSheet.getDataRange().setValues(scoreData);
Logger.log(teamScores)
}




function addArray(array){
  var total = 0;
  for (var n in array){
   var num = array[n];
    total +=1* num;
  }
 return total;
  
}

function topNum(array, num){
  var newArray = [];
  array.sort(function(a, b){return b - a});
  newArray = array.slice(0,num)
  return newArray;
}



function uiSweeps(){
ui.showModalDialog(sweepsMenu, 'Sweeps Menu')
}

function rSweepsMenu(form){
 tSweeps = form.tSweep
 num = form.num
 
 if(tSweeps == 'tops'){
   if(num == 0){
   ui.alert('Please select a number of contestants to count.');
   uiSweeps();
   }
   selectSweeps(num);
  
 }
  if(tSweeps == 'everything'){
    allSweep();
                             }
  if(tSweeps == 'aB'){
    abSweep();
  }
  
 Logger.log(tSweeps)
 Logger.log(num)
 ssa.flush();
 var sweepSheet = ss.getSheetByName('Sweeps')
  sweepSheet.sort(3,false);
  var len = clearBlanks(transpose(sweepSheet.getDataRange().getValues()))[0].length - 1
  sweepSheet.getRange(2,4,len).setFormula('=Rank((R[0]C[-1]),$C$2:C)')
  
 ui.alert('Success')
}

function  abSweep(){
  var sheets = ss.getSheets();
  for (var s in sheets){
    var sheet = sheets[s];
    if (sheet.getName() == "Sweeps"){
      continue
    }
    var sRange = sheet.getRange(3, 3, 150, 1);
    sRange.setFormulaR1C1('=if(RIGHT(R[0]C[-1],1)="A","x",if(RIGHT(R[0]C[-1],1)="B","x",))')
    
  }
  sweeps()
}


function  allSweep(){
  var sheets = ss.getSheets();
  for (var s in sheets){
    var sheet = sheets[s];
    if (sheet.getName() == "Sweeps"){
      continue
    }
    if (sheet.getName() == "Links"){
      continue
    }
    var sRange = sheet.getRange(3, 3, 150, 1);
    sRange.setValue('x');
    
  }
  sweeps()
}


function test(){
  var sheets = ss.getSheets();
  var sweepSheet = ss.getSheetByName("Sweeps");
  var teams = sweepSheet.getDataRange().getValues();
  teams = transpose(teams);
  teams = teams[0];
  
  Logger.log(teams);
  
  
}