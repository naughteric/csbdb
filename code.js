var url = "https://docs.google.com/spreadsheets/d/166HRPIz616o5TVXc-e1SYqtP7UHMD5I47t7Q2m0xpnw/edit?folder=0AEgXVPdo6S40Uk9PVA#gid=0";
//This is the URL for the main Spreadsheet

function facultyName(){
  
  /*
  This function first retrieves the Last and First Name from the Faculty Sheet
  Once the list of the names are retireved they are mapped to a 1d array declared as LNameList and FNameList respectively
  The for loop merges the Last and First Name together and is inserted into CombinedFL
  The list of names in combinedFL is returned whenever the function is called.
  */
  
    var ss = SpreadsheetApp.openByUrl(url);
    var getFacWs = ss.getSheetByName("Faculty");
    var getLastName = getFacWs.getRange(2,2,getFacWs.getLastRow(),1).getValues(); 
    var getFirstName = getFacWs.getRange(2,3,getFacWs.getLastRow(),1).getValues();

    var LNameList = getLastName.map(function(r){return r[0]; }); 
    var FNameList = getFirstName.map(function(r){return r[0]; });

    var combinedFL = [];
    var tempName = "";

     for (i=0; i< LNameList.length; i++){
        tempName = LNameList[i] + " " + FNameList[i];
        combinedFL[i] = tempName;
     }
  
    return combinedFL;
}

//----------------------------------------------------------------------------------------------------------------------------------------------
//START HTMLSERVICE
function doGet(pgload) {
  
  if (pgload.parameters.v == "facform") {
     return loadFacultyForm();
  } else if (pgload.parameters.v == "ohform") {
     return loadOfficeHoursForm();
  } else if (pgload.parameters.v == "upform") {
     return loadUploadForm();
  } else if (pgload.parameters.v == "rptform") {
     return loadReportsForm();
  } else {
     return HtmlService.createTemplateFromFile("home").evaluate();
  }
  
  /*
  When the WebApp loads the default page is home.html
  The If Else statement checks the URL parameter and calls the required function to load the page
  */
  
}
//END HTMLSERVICE
//----------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------
//START FACULTY
//Start of functions related to faculty form

function loadFacultyForm(){

  /*
  The Spreadsheet is called by the URL
  Each Worksheet (Ws) is called by their name
  The values are retrieved from each Ws and stored into their own array
  The arrays are assigned to another array (tlist, slist, dlist, etc..) that will be declared when the page (faculty) loads
  */
  var ss = SpreadsheetApp.openByUrl(url);
  var getClassWs = ss.getSheetByName("Classification");
  var listCl = getClassWs.getRange(1,1,getClassWs.getRange("A1").getDataRegion().getLastRow()).getValues();
  
  var getTbWs = ss.getSheetByName("Timebase");
  var listTb = getTbWs.getRange(1,1,getTbWs.getRange("A1").getDataRegion().getLastRow()).getValues();
  
  var getStWs = ss.getSheetByName("Status");
  var listSt = getStWs.getRange(1,1,getStWs.getRange("A1").getDataRegion().getLastRow()).getValues();
  
  var getDptWs = ss.getSheetByName("Department");
  var listDpt = getDptWs.getRange(1,1,getDptWs.getRange("A1").getDataRegion().getLastRow()).getValues();
  
  var getTypeWs = ss.getSheetByName("Type");
  var listType = getTypeWs.getRange(1,1,getTypeWs.getRange("A1").getDataRegion().getLastRow()).getValues();
  
  var editList = facultyName();
  
  var tmp = HtmlService.createTemplateFromFile("faculty")
  tmp.list = listCl; //List of all values in Classification WorkSheet; for populating Classification Dropdown Box
  tmp.tlist = listTb; //List of all values in TimeBase Worksheet; for populating TimeBase Dropdown Box
  tmp.slist = listSt; //List of all values in Status Worksheet; for populating Status Dropdown Box
  tmp.dlist = listDpt; //List of all values in Department Worksheet; for populating Department Dropdown Box
  tmp.tylist = listType; //List of values in Type Worksheet; for populating Type Dropdown box
  tmp.edlist = editList; //List of all Faculty members; for populating Faculty Member Edit Dropdown Box
  return tmp.evaluate();
  
}

function userClicked(psid,lname,fname,classif,timebase,status,emtype,department,officenum,officephone,email,ferp,homeaddress,homephone){
  var testDB = SpreadsheetApp.openByUrl(url);
  var getTB = testDB.getSheetByName("Faculty");
  
  
  getTB.appendRow([psid,lname,fname,classif,timebase,status,emtype,department,officenum,officephone,email,ferp,homeaddress,homephone]);
  
  var sortRng = getTB.getRange(2, 1, getTB.getLastRow(),14);
  sortRng.sort({column : 2});
  sortRng.setNumberFormat("@")
}


function checkDupe(idToCheck){


    var getOHDB = SpreadsheetApp.openByUrl(url);
    var getFacPSID = getOHDB.getSheetByName("Faculty");
    var getRefPsid = getFacPSID.getRange(2,1,getFacPSID.getLastRow(),1).getValues();

    var refPsid = getRefPsid.map(function(r){return r[0]; });

    var isDupe = 0;

    for(i=0; i<refPsid.length; i++) {
      var testId = refPsid[i].toString();
        if (idToCheck === testId){
            isDupe = 1;
        }
    }
  
    return isDupe;
}

//End of functions related to Faculty Form
//END FACULTY
//----------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------
//START OFFICEHOURS FORM
//Start of functions related to Office Hours Form

function loadOfficeHoursForm(){
 
    var fullFacName = facultyName();
    
    var oftmp = HtmlService.createTemplateFromFile("officehours")
    oftmp.idlist = fullFacName
    return oftmp.evaluate();
}

function getPosOfName(fullName){

    var getOHDB = SpreadsheetApp.openByUrl(url);
    var getFacPSID = getOHDB.getSheetByName("Faculty");
    var getRefPsid = getFacPSID.getRange(2,1,getFacPSID.getLastRow(),1).getValues();
    var combinedFL = facultyName();
  
    var refPsid = getRefPsid.map(function(r){return r[0]; });
  
    var namePos = combinedFL.indexOf(fullName);
      
     if (namePos > -1) {
       var passingId = refPsid[namePos];  
       return passingId;
     } else {
       return 'Not Listed';
     }

}
 
 function ofHrsSaved(ofDays,startTime,endTime,comments,psid){
   var testDB = SpreadsheetApp.openByUrl(url);
   var getTB = testDB.getSheetByName("Officehours");
   
   
   getTB.appendRow([ofDays,startTime,endTime,comments,psid]);
   
   var setFormat = getTB.getRange(2, 1, getTB.getLastRow(), 5);
   setFormat.setNumberFormat("@");
 }

//End of functions related to Office Hours Form
//END OFFICE HOURS
//----------------------------------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------------------------------
//START UPLOAD
//Start of functions related to Upload File Form

function loadUploadForm(){

  var uptmp = HtmlService.createTemplateFromFile("upload")
  return uptmp.evaluate();

}

function uploadFiles(data)
{
 var file = data.myFile;
 var folder = DriveApp.getFolderById('1prEoUQXN3hhkgn_eNUwLGqdJmd9fSMyT');
 var createFile = folder.createFile(file);
 importData();
 formatReport();
 var returnMsg = "PeopleSoft Report Uploaded";
 return returnMsg;
}

function formatReport(){
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("Classes");
  ws.deleteColumns(7, 8);
  ws.deleteColumns(11, 5);
  ws.deleteColumns(13,4);
  ws.deleteColumns(14,12);
  
  var sub = ws.getRange(2, 2, ws.getLastRow(), 1).getValues();
  var num = ws.getRange(2, 3, ws.getLastRow(), 1).getValues();
  var mapsub = sub.map(function(r){return r[0]; });
  var mapnum = num.map(function(r){return r[0]; });
  
  var temp = "";
  var combined = [];
  for (var i = 0; i<=mapsub.length; i++){
    temp = mapsub[i]+""+mapnum[i];
    combined[i] = temp;
  }
  
  ws.deleteColumns(2, 2);
  ws.insertColumnAfter(1);
  
  var newRange = "";
  var actCell = 2;
  for (var n = 0; n<=combined.length; n++){
    var newCell = n+actCell;
    newRange = "B"+newCell;
    ws.getRange(newRange).setValue(combined[n]);
  }
  
  var numFormat = ws.getRange(1, 1, ws.getLastRow(), ws.getLastColumn());
  numFormat.setNumberFormat("@");
  
}

//START CSV CONVERT
function importData() {
  var fSource = DriveApp.getFolderById('1prEoUQXN3hhkgn_eNUwLGqdJmd9fSMyT'); // reports_folder_id = id of folder where csv reports are saved *EDITED*
  var fi = fSource.getFilesByName('report.csv'); // latest report file
  var ss = SpreadsheetApp.openByUrl(url); // data_sheet_id = id of spreadsheet that holds the data to be updated with new report data

  if ( fi.hasNext() ) { // proceed if "report.csv" file exists in the reports folder
    var file = fi.next();
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv); // see below for CSVToArray function
    var newsheet = ss.getSheetByName("Classes"); // create a 'NEWDATA' sheet to store imported data *EDITED*
    // loop through csv data array and insert (append) as rows into 'NEWDATA' sheet
    for ( var i=0, lenCsv=csvData.length; i<lenCsv; i++ ) {
      newsheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    }
    /*
    ** report data is now in 'NEWDATA' sheet in the spreadsheet - process it as needed,
    ** then delete 'NEWDATA' sheet using ss.deleteSheet(newsheet)
    */
    // rename the report.csv file so it is not processed on next scheduled run
    file.setName("report-"+(new Date().toString())+".csv");
  }
};


// http://www.bennadel.com/blog/1504-Ask-Ben-Parsing-CSV-Strings-With-Javascript-Exec-Regular-Expression-Command.htm
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.

function CSVToArray( strData, strDelimiter ) {
  // Check to see if the delimiter is defined. If not,
  // then default to COMMA.
  strDelimiter = (strDelimiter || ",");

  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){

      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );

    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } else {

      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];

    }

    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }

  // Return the parsed data.
  return( arrData );
};

//END CSV CONVERT
//----------------------------------------------------------------------------------------------------------------------------------------------

//End of functions related to Upload File Form
//END UPLOAD
//----------------------------------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------------------------------
//START REPORTING
//Start of functions related to Creating reports

function loadReportsForm(){
  var getDB = SpreadsheetApp.openByUrl(url);
  var getdptTB = getDB.getSheetByName("Department");
  var listdpt = getdptTB.getRange(1,1,getdptTB.getRange("A1").getDataRegion().getLastRow()).getValues();

  var rpttmp = HtmlService.createTemplateFromFile("report")
  rpttmp.dlist = listdpt;
  return rpttmp.evaluate();

}

//----------------------------------------------------------------------------------------------------------------------------------------------
//Functions for Faculty List Report

function queryRegions(){
  
  var ss = SpreadsheetApp.openByUrl(url);
  var facws = ss.getSheetByName("Faculty");
  var classws = ss.getSheetByName("Classes");
  var officews = ss.getSheetByName("Officehours");
  var facRegion = facws.getDataRange().getA1Notation();
  var classRegion = classws.getDataRange().getA1Notation();
  var offRegion = officews.getDataRange().getA1Notation();
  
  var arrRegions = [];
  arrRegions[0] = facRegion;
  arrRegions[1] = classRegion;
  arrRegions[2] = offRegion;
  

  
  return arrRegions;

}
function sendRptMail(link, wsName){

  var sendTo = "csbdatabase@mail.fresnostate.edu"; //Change to whoever needs copy of reports
  MailApp.sendEmail(sendTo, wsName, link)

}

function sortDepartment(dpt){
  
  var regions = queryRegions();
  var facRegion = regions[0];
  
  var ss = SpreadsheetApp.openByUrl(url);
  var sortws = ss.getSheetByName("Sorting");
  var facws = ss.getSheetByName("Faculty");

  
  
  if(dpt === "Select a Department"){ 
    var idsToQry = facws.getRange(2,1, facws.getRange("A2").getDataRegion().getLastRow(),1).getValues();
    var listIds = idsToQry.map(function(r){return r[0]; });  
  } else {
    var qryIdByDpt = "=QUERY(Faculty!"+facRegion+",\"select A where H contains '"+dpt+"' \",0)";
    sortws.getRange("A1").setFormula(qryIdByDpt);
    var idsToQry = sortws.getRange(1,1,sortws.getRange("A1").getDataRegion().getLastRow(),1).getValues();
    var listIds = idsToQry.map(function(r){return r[0]; }); 
  }

  return listIds;

}

function queryToCell(dpt){

    var getSS = SpreadsheetApp.openByUrl(url);
    var toSheet = getSS.getSheetByName("Faclist"); 
  
  // Start - Determine Name of New WS
  var getDate = new Date();
  var YY = getDate.getFullYear();
  var MM = getDate.getMonth()+1;
  
  if (MM <= 5){
    var fasp = "SP";
  } else {
    var fasp = "FA";
  }
  
  if (dpt === "Select a Department"){
    var rptDpt = "ALL";
  }
  else{
    var rptDpt = dpt;
  }
  
  var newWsName = fasp+YY+"-"+rptDpt;
  
  var rptss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Oa2Wpi9neUdqeGi20evNIuzbfv64JFKrHyMRDKsdhYI/edit#gid=0"); 
  var newws = rptss.insertSheet(newWsName, 1);
  //End - Determine Name of New WS
 
  
  
    var listIds = sortDepartment(dpt);
  
  

    var regions = queryRegions();
    var facRegion = regions[0];
  
    var headToCell1 = toSheet.getRange("A1");
    headToCell1.setValue("Name");
    var headToCell3 = toSheet.getRange("C1");
    headToCell3.setValue("Office");
    var headToCell4 = toSheet.getRange("D1");
    headToCell4.setValue("Phone");
    var headToCell5 = toSheet.getRange("E1");
    headToCell5.setValue("Department");
    var headToCell6 = toSheet.getRange("F1");
    headToCell6.setValue("Email");
  
    if (dpt === "Select a Department"){ //If Department if
      var selectDpt = dpt;
      var listLength = listIds.length-1;
      
      for(var i = 0; i<=listLength; i++){
      
        var currId = listIds[i];
        var qryFac = "=QUERY(Faculty!"+facRegion+",\"Select B, C, I, J, H, K where A matches '"+currId+"' \",0)";
        
        var newRec = i+2;
        var newCell = "A"+newRec;
        
        toSheet.getRange(newCell).setFormula(qryFac)
      
      }
      
    } else {
      var selectDpt = dpt;
      var listLength = listIds.length-1;
      
      for(var i = 0; i<=listLength; i++){
      
        var currId = listIds[i];
        var qryFac = "=QUERY(Faculty!"+facRegion+",\"Select B, C, I, J, H, K where A matches '"+currId+"' \",0)";
        
        var newRec = i+2;
        var newCell = "A"+newRec;
        
        toSheet.getRange(newCell).setFormula(qryFac)
      
      }
      
    }
    
    readFromSpreadsheetAndWriteOnDoc(newWsName);
  
  var start, end;
  start = 1
  end = toSheet.getLastRow();
  toSheet.deleteRows(start, end);
  
  //Test Send Email
  //var linkRpt = rptss.getUrl()
  //sendRptMail(linkRpt, newWsName); EMAIL_TO_DAYNA_IF_NECESSARY
  
  var rptDone = newWsName+" Faculty Name Sort "+"Report Complete";
  
  return rptDone;

}

//Source - https://stackoverflow.com/questions/34687772/copy-a-range-of-spreadsheet-to-a-doc
function readFromSpreadsheetAndWriteOnDoc(newWsName) {
  
  var source = SpreadsheetApp.openByUrl(url);
  var sourcesheet = source.getSheetByName('Faclist'); //Change

  var rptSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Oa2Wpi9neUdqeGi20evNIuzbfv64JFKrHyMRDKsdhYI/edit#gid=0"); 
  var writeRpt = rptSheet.getSheetByName(newWsName);
  
  
  var numRecords = sourcesheet.getRange("A1").getDataRegion().getLastRow();
  
  //ADDING A HEADER - Faculty Name Report
  var headLine = writeRpt.getLastRow();

  if (headLine === 0){
  
    var header = [" "," ","Faculty Name Sort"+" - "+newWsName]; 
    writeRpt.appendRow(header);
    writeRpt.getRange("C1").setFontSize(14);
    writeRpt.getRange("C1").setFontWeight("bold");
    
    var endRpt = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn()).getLastRow();
    var skipLine = endRpt+2;
    var lastCell = "A"+skipLine;
    writeRpt.getRange(lastCell).setValue("     ")
  }
  //END ADDING HEADER
  
  for (var n=1; n<=numRecords; n++){
  
    var nextLine = sourcesheet.getRange(n,1,1,6).getValues();

    var arrToPrint = [];
    var valHold = "";
    var name = ""
    for (var h = 0; h<6; h++){
    
      if (h === 0 && n != 1){
      
        valHold = nextLine[0][0] + ", " + nextLine[0][1]
        arrToPrint[h] = valHold;
      
      } else if (h === 1){
      
        valHold = " "
        arrToPrint[h] = valHold
      
      } else {
        
        valHold = nextLine[0][h];
        arrToPrint[h] = valHold;
        
      } 

    }
    
    if (arrToPrint[0] === "Name"){
    
      var formNext = writeRpt.getLastRow()+1
      var cellA = "A"+formNext;
      var cellC = "C"+formNext;
      var cellD = "D"+formNext;
      var cellE = "E"+formNext;
      var cellF = "F"+formNext;
      writeRpt.getRange(cellA).setFontSize(11);
      writeRpt.getRange(cellA).setFontWeight("bold");
      writeRpt.getRange(cellA).setFontColor("blue")
      writeRpt.getRange(cellC).setFontSize(11);
      writeRpt.getRange(cellC).setFontWeight("bold");
      writeRpt.getRange(cellC).setFontColor("blue")
      writeRpt.getRange(cellD).setFontSize(11);
      writeRpt.getRange(cellD).setFontWeight("bold");
      writeRpt.getRange(cellD).setFontColor("blue")
      writeRpt.getRange(cellE).setFontSize(11);
      writeRpt.getRange(cellE).setFontWeight("bold");
      writeRpt.getRange(cellE).setFontColor("blue")
      writeRpt.getRange(cellF).setFontSize(11);
      writeRpt.getRange(cellF).setFontWeight("bold");
      writeRpt.getRange(cellF).setFontColor("blue")
    
    }
    
    writeRpt.appendRow(arrToPrint);
    
  }
  
  writeRpt.deleteColumns(2,1);
  writeRpt.autoResizeColumn(1)

}
//End of functions for Faculty List Report
//----------------------------------------------------------------------------------------------------------------------------------------------
//Start of functions for Door Post Query (ALL Staff/Faculty)

function doorByDpt(dpt){
  var ss = SpreadsheetApp.openByUrl(url);
  var postws = ss.getSheetByName("Doorpost"); 
  
  // Start - Determine Name of New WS
  var getDate = new Date();
  var YY = getDate.getFullYear();
  var MM = getDate.getMonth()+1;
  
  if (MM <= 5){
    var fasp = "SP";
  } else {
    var fasp = "FA";
  }
  
  if (dpt === "Select a Department"){
    var rptDpt = "ALL";
  }
  else{
    var rptDpt = dpt;
  }
  
  var newWsName = fasp+YY+"-"+rptDpt;
  
  var rptss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1PZNO3PnJ-CtGeU1MnP8nZT9axfssaTkK-z7sUHqyqn8/edit#gid=0"); 
  var newws = rptss.insertSheet(newWsName, 1);
  //End - Determine Name of New WS
  
  
  
  var listIds = sortDepartment(dpt);
  
  

  var regions = queryRegions();
  var facRegion = regions[0];
  var classRegion = regions[1];
  var offRegion = regions[2];

  
  if (dpt === "Select a Department"){
    var selectDpt = " ";
  } else {
    var selectDpt = dpt;
  }

  var listLength = listIds.length-1;
  
  for(var i=0; i<=listLength;i++){
    
    var currId = listIds[i];
    
    var qryName = "=QUERY(Faculty!"+facRegion+",\"Select B, C where A matches '"+currId+"' \",0)";
    var qryOffNumPh = "=QUERY(Faculty!"+facRegion+",\"Select I, J where A matches '"+currId+"' \",0)";
    var qryEmail = "=QUERY(Faculty!"+facRegion+",\"Select K where A matches '"+currId+"' \",0)";
    var qryOffice = "=QUERY(Officehours!"+offRegion+",\"select A, B, C where E matches '"+currId+"' order by A asc \",0)";
    var qryComm = "=QUERY(Officehours!"+offRegion+",\"select D where E matches '"+currId+"' order by A asc \",0)";
    var qryClasses = "=QUERY(Classes!"+classRegion+",\"select B, E, F, G, H, I where J matches '"+currId+"' order by G asc \",0)";
    
    var qryDpt = "=QUERY(Department!A1:B21, \"Select B where A matches '"+selectDpt+"' \",0)";
    
    postws.getRange("A1").setFormula(qryName);
    postws.getRange("F1").setFormula(qryDpt);
    postws.getRange("A2").setValue("Office");
    postws.getRange("B2").setValue("Phone");
    postws.getRange("D2").setValue("Email");
    postws.getRange("A3").setFormula(qryOffNumPh);
    postws.getRange("D3").setFormula(qryEmail);
    postws.getRange("A4").setValue("Office Hours");
    postws.getRange("A5").setValue("Day(s)");
    postws.getRange("B5").setValue("Start Time");
    postws.getRange("C5").setValue("End Time");
    postws.getRange("E5").setValue("Comments");
    postws.getRange("A6").setFormula(qryOffice);
    postws.getRange("E6").setFormula(qryComm);
    
    
    var nextLine = postws.getRange("A6").getDataRegion().getLastRow();
    var headingRow = nextLine+1;
    var courseHead = "A"+headingRow;
    var lblsRow = nextLine+2;
    var courseLbl = "A"+lblsRow;
    var schedLbl = "B"+lblsRow;
    var roomLbl = "C"+lblsRow;
    var dayLbl = "D"+lblsRow;
    var startLbl = "E"+lblsRow;
    var endLbl = "F"+lblsRow;
    var qryRow = nextLine+3;
    var setQry = "A"+qryRow;
  
    postws.getRange(courseHead).setValue("Courses");
    postws.getRange(courseLbl).setValue("Course");
    postws.getRange(schedLbl).setValue("Schedule");
    postws.getRange(roomLbl).setValue("Room");
    postws.getRange(dayLbl).setValue("Day(s)");
    postws.getRange(startLbl).setValue("Start Time");
    postws.getRange(endLbl).setValue("End Time");
    postws.getRange(setQry).setFormula(qryClasses);
    

    deptOfficePost(newWsName);

    var start, end;
    start = 1
    end = postws.getLastRow();
    postws.deleteRows(start, end);
  
  }
  var rptDone = newWsName+" Office Hours Door Posting "+"Report Complete";
  
  
  //var linkRpt = rptss.getUrl()
  //sendRptMail(linkRpt, newWsName); EMAIL_TO_DAYNA_IF_NECESSARY
  
  return rptDone;
}

function deptOfficePost(newWsName) {
  //This function reads each row one by one into an array and written to a new line
  // get the Spreadsheet by sheet URL
  var source = SpreadsheetApp.openByUrl(url);
  var sourcesheet = source.getSheetByName('Doorpost'); 
  
  //Finding and determining which file to write to
  var rptSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1PZNO3PnJ-CtGeU1MnP8nZT9axfssaTkK-z7sUHqyqn8/edit#gid=0"); 
  var writeRpt = rptSheet.getSheetByName(newWsName);  
  
  
  
  var numRows = sourcesheet.getRange(1, 1, sourcesheet.getLastRow(), sourcesheet.getLastColumn()).getLastRow();
  
  //ADDING A HEADER 
  var headLine = writeRpt.getLastRow();

  if (headLine === 0){
  
    var header = [" ","Office Hours Door Post"+" - "+newWsName]; 
    writeRpt.appendRow(header);
    writeRpt.getRange("B1").setFontSize(14);
    writeRpt.getRange("B1").setFontWeight("bold");
    
    var endRpt = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn()).getLastRow();
    var skipLine = endRpt+2;
    var lastCell = "A"+skipLine;
    writeRpt.getRange(lastCell).setValue("     ")
  }
  //END ADDING HEADER
  
  for(var n = 1; n<=numRows; n++){
    
     var srcData = sourcesheet.getRange(n,1,1,7).getValues(); 
 
  
     var arrToPrint = [];
     var valHold = "";
     //Convert a 2d array to 1d array
     for(var h=0; h<7; h++){
        valHold = srcData[0][h];
        arrToPrint[h] = valHold;
     }
    
    writeRpt.appendRow(arrToPrint);
    
    if (n === 1){
      var formNext = writeRpt.getLastRow();
      var cellB = "B"+formNext;
      var cellA = "A"+formNext;
      writeRpt.getRange(cellA).setFontSize(14);
      writeRpt.getRange(cellA).setFontWeight("bold");
      writeRpt.getRange(cellB).setFontSize(14);
      writeRpt.getRange(cellB).setFontWeight("bold");
    }
    if (n === 3){
    
      var formNext = writeRpt.getLastRow();
      var cellA = "A"+formNext;
      var cellB = "B"+formNext;
      var cellC = "C"+formNext;
      var cellD = "D"+formNext;
      var cellE = "E"+formNext;
      var cellF = "F"+formNext;
      writeRpt.getRange(cellA).setFontSize(12);
      writeRpt.getRange(cellB).setFontSize(12);
      writeRpt.getRange(cellD).setFontSize(12);
    
      writeRpt.getRange(cellA).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellB).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellC).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellD).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellE).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellF).setBorder(false, false, true, false, false, false);
    
    }
    if (n === 4){
    
      var formNext = writeRpt.getLastRow();
      var cellA = "A"+formNext;
      writeRpt.getRange(cellA).setFontColor("red");
      writeRpt.getRange(cellA).setFontSize(11);
      writeRpt.getRange(cellA).setFontWeight("bold");
    }
    if (n === 5){
    
      var formNext = writeRpt.getLastRow();
      var cellA = "A"+formNext;
      var cellB = "B"+formNext;
      var cellC = "C"+formNext;
      var cellD = "D"+formNext;
      var cellE = "E"+formNext;
      var cellF = "F"+formNext;
      writeRpt.getRange(cellA).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellB).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellC).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellD).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellE).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellF).setBorder(false, false, true, false, false, false);
    }
    if(arrToPrint[0] === "Courses"){
    
      var formNext = writeRpt.getLastRow();
      var cellA = "A"+formNext;
      writeRpt.getRange(cellA).setFontColor("red");
      writeRpt.getRange(cellA).setFontSize(11);
      writeRpt.getRange(cellA).setFontWeight("bold");
    }
    if(arrToPrint[0] === "Course"){
      
      var formNext = writeRpt.getLastRow();
      var cellA = "A"+formNext;
      var cellB = "B"+formNext;
      var cellC = "C"+formNext;
      var cellD = "D"+formNext;
      var cellE = "E"+formNext;
      var cellF = "F"+formNext;
      writeRpt.getRange(cellA).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellB).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellC).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellD).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellE).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellF).setBorder(false, false, true, false, false, false);
    }
    
  }
  
  
  if (n > numRows){
  
    var endRpt = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn()).getLastRow();
    var skipLine = endRpt+10;
    var lastCell = "A"+skipLine;
    writeRpt.getRange(lastCell).setValue("     ")
  }
  
   var setFormat = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn());
   setFormat.setNumberFormat("@");
  writeRpt.autoResizeColumn(1);
  
}

function doorQry(dpt){
  
  var ss = SpreadsheetApp.openByUrl(url);
  var toSheet = ss.getSheetByName("Faccourse"); 
  
  // Start - Determine Name of New WS
  var getDate = new Date();
  var YY = getDate.getFullYear();
  var MM = getDate.getMonth()+1;
  
  if (MM <= 5){
    var fasp = "SP";
  } else {
    var fasp = "FA";
  }
  
  if (dpt === "Select a Department"){
    var rptDpt = "ALL";
  }
  else{
    var rptDpt = dpt;
  }
  
  var newWsName = fasp+YY+"-"+rptDpt;
  
  var rptss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1NY7Sz24iEd56TAVj_2H3x08P9JNUX1r4ZmiTCV7KyYs/edit#gid=0");
  var newws = rptss.insertSheet(newWsName, 1);
  //End - Determine Name of New WS
  
  
  
  var listIds = sortDepartment(dpt);
  
  

  var regions = queryRegions();
  var facRegion = regions[0];
  var classRegion = regions[1];
  var offRegion = regions[2];

  
  if (dpt === "Select a Department"){
    var selectDpt = " ";
  } else {
    var selectDpt = dpt;
  }

    var idlength = listIds.length-1;
    for (var i=0; i<=idlength; i++){
      
        
      
        var currId = listIds[i]; //Grabs the current PSID in the array to plug into the query
      
      Logger.log(currId);

        var qryLname = "=QUERY(Faculty!"+facRegion+",\"Select B where A matches '"+currId+"' \",0)";
        var qryFname = "=QUERY(Faculty!"+facRegion+",\"Select C where A matches '"+currId+"' \",0)";
        var qryOffNum = "=QUERY(Faculty!"+facRegion+",\"Select I where A matches '"+currId+"' \",0)";
        var qryOffPho = "=QUERY(Faculty!"+facRegion+",\"Select J where A matches '"+currId+"' \",0)";
        var qryEmail = "=QUERY(Faculty!"+facRegion+",\"Select K where A matches '"+currId+"' \",0)";
        var qryCourse = "=QUERY(Classes!"+classRegion+",\"select B where J matches '"+currId+"' order by G asc \",0)";
        var qryClassNum = "=QUERY(Classes!"+classRegion+",\"Select E where J matches '"+currId+"' order by G asc \",0)";
        var qryClassDay = "=QUERY(Classes!"+classRegion+",\"select G where J matches '"+currId+"' order by G asc \",0)";
        var qryClassStart = "=QUERY(Classes!"+classRegion+",\"select H where J matches '"+currId+"' order by G asc \",0)";
        var qryClassEnd = "=QUERY(Classes!"+classRegion+",\"select I where J matches '"+currId+"' order by G asc \",0)";
        var qryClassRoom = "=QUERY(Classes!"+classRegion+",\"select F where J matches '"+currId+"' order by G asc \",0)";
        var qryOfficeHours = "=QUERY(Officehours!"+offRegion+",\"select A, B, C where E matches '"+currId+"' order by A asc \",0)"
        var qryOffComm = "=QUERY(Officehours!"+offRegion+",\"select D where E matches '"+currId+"' order by A asc \",0)"
        
        var qryDpt = "=QUERY(Department!A1:B21, \"Select B where A matches '"+selectDpt+"' \",0)";

        toSheet.getRange('B1').setFormula(qryFname);
        toSheet.getRange('A1').setFormula(qryLname);
        toSheet.getRange('C1').setFormula(qryOffNum);
        toSheet.getRange('D1').setFormula(qryOffPho);
        toSheet.getRange('E1').setFormula(qryEmail);
        toSheet.getRange("G1").setFormula(qryDpt);
        toSheet.getRange('A2').setValue("Courses");
        toSheet.getRange('B3').setFormula(qryCourse);
        toSheet.getRange('C2').setValue("Class #");
        toSheet.getRange('C3').setFormula(qryClassNum);
        toSheet.getRange('D3').setFormula(qryClassDay);
        toSheet.getRange('E3').setFormula(qryClassStart);
        toSheet.getRange('F3').setFormula(qryClassEnd);
        toSheet.getRange('G3').setFormula(qryClassRoom);
      
      var nextLine = toSheet.getRange("B3").getDataRegion().getLastRow();
      var lblRow = nextLine+1;
      var HrsRow = nextLine+2;
      var setlbl = "A"+lblRow;
      var setHrs = "D"+HrsRow;
      var setComm = "G"+HrsRow;
      
        toSheet.getRange(setlbl).setValue("Office Hours");
        toSheet.getRange(setHrs).setFormula(qryOfficeHours);
        toSheet.getRange(setComm).setFormula(qryOffComm);

        officeHrPostToDoc(newWsName);

        var start, end;
        start = 1
        end = toSheet.getLastRow();
        toSheet.deleteRows(start, end);

    }
  var rptDone = newWsName+" Faculty Courses/Office Hours "+"Report Complete";
  
  //var linkRpt = rptss.getUrl()
  //sendRptMail(linkRpt, newWsName); EMAIL_TO_DAYNA_IF_NECESSARY
  
  return rptDone;
}

function officeHrPostToDoc(newWsName) 
{
  //This function reads each row one by one into an array and written to a new line
  // get the Spreadsheet by sheet URL
  var source = SpreadsheetApp.openByUrl(url);
  var sourcesheet = source.getSheetByName('Faccourse');
  
  
  var rptSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1NY7Sz24iEd56TAVj_2H3x08P9JNUX1r4ZmiTCV7KyYs/edit#gid=0"); 
  var writeRpt = rptSheet.getSheetByName(newWsName);
  
  var numRows = sourcesheet.getRange("A1").getDataRegion().getLastRow();
  
  //ADDING A HEADER
  var headLine = writeRpt.getLastRow();

  if (headLine === 0){
  
    var header = [" ","Faculty Courses/Office Hours"+" - "+newWsName]; 
    writeRpt.appendRow(header);
    writeRpt.getRange("B1").setFontSize(14);
    writeRpt.getRange("B1").setFontWeight("bold");
    
    var endRpt = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn()).getLastRow();
    var skipLine = endRpt+2;
    var lastCell = "A"+skipLine;
    writeRpt.getRange(lastCell).setValue("     ")
  }
  //END ADDING HEADER
  
  for(var n = 1; n<=numRows; n++){
    
     var srcData = sourcesheet.getRange(n,1,1,7).getValues(); 
  
     var arrToPrint = [];
     var valHold = "";
     //Convert a 2d array to 1d array
     for(var h=0; h<7; h++){
        valHold = srcData[0][h];
        arrToPrint[h] = valHold;
     }
    
     writeRpt.appendRow(arrToPrint);
    
    if (n === 1){
    
      var nextRow = writeRpt.getLastRow();
      var cellA = "A"+nextRow;
      var cellB = "B"+nextRow;
      var cellC = "C"+nextRow;
      var cellD = "D"+nextRow;
      var cellE = "E"+nextRow;
      var cellG = "G"+nextRow;
      writeRpt.getRange(cellA).setFontSize(11);
      writeRpt.getRange(cellA).setFontWeight("bold");
      writeRpt.getRange(cellB).setFontSize(11);
      writeRpt.getRange(cellB).setFontWeight("bold");
      writeRpt.getRange(cellC).setFontWeight("bold");
      writeRpt.getRange(cellD).setFontWeight("bold");
      writeRpt.getRange(cellE).setFontWeight("bold");
      writeRpt.getRange(cellG).setFontWeight("bold");
     
    }
    if (n === 2){    
      var nextRow = writeRpt.getLastRow();
      var cellA = "A"+nextRow;
      var cellC = "C"+nextRow;
      writeRpt.getRange(cellA).setFontColor("blue");
      writeRpt.getRange(cellA).setFontWeight("bold");
      writeRpt.getRange(cellC).setFontColor("blue");
      writeRpt.getRange(cellC).setFontWeight("bold");
    }
    if (arrToPrint[0] === "Office Hours"){
    
      var nextRow = writeRpt.getLastRow();
      var cellA = "A"+nextRow;
      writeRpt.getRange(cellA).setFontColor("red");
      writeRpt.getRange(cellA).setFontWeight("bold");
    }
    if (n === numRows){
    
      var formNext = writeRpt.getLastRow();
      var cellA = "A"+formNext;
      var cellB = "B"+formNext;
      var cellC = "C"+formNext;
      var cellD = "D"+formNext;
      var cellE = "E"+formNext;
      var cellF = "F"+formNext;
      var cellG = "G"+formNext;
      writeRpt.getRange(cellA).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellB).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellC).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellD).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellE).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellF).setBorder(false, false, true, false, false, false);
      writeRpt.getRange(cellG).setBorder(false, false, true, false, false, false);
     
    }
    
  }
  
  
  if (n > numRows){
    var endRpt = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn()).getLastRow();
    var skipLine = endRpt+1;
    var lastCell = "A"+skipLine;
    writeRpt.getRange(lastCell).setValue("     ")
  }
  
  var setFormat = writeRpt.getRange(1, 1, writeRpt.getLastRow(), writeRpt.getLastColumn());
  setFormat.setNumberFormat("@");
  writeRpt.autoResizeColumn(1);
}

//End of functions for Door Post Query (ALL Staff/Faculty)
//END REPORTING
//----------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------
//START EDITING/MAINTANENCE
//Start of Faculty editing functions
function retrieveInfo(posi){

      var getDB = SpreadsheetApp.openByUrl(url);
      var getFacPSID = getDB.getSheetByName("Faculty");
      var getRefFac = getFacPSID.getRange(2,1,getFacPSID.getLastRow(),14).getValues();
  
      var refPsid = getRefFac.map(function(r){return r[0]; });
      var refLn = getRefFac.map(function(r){return r[1]; });
      var refFn = getRefFac.map(function(r){return r[2]; });
      var refClass = getRefFac.map(function(r){return r[3]; });
      var refTime = getRefFac.map(function(r){return r[4]; });
      var refStat = getRefFac.map(function(r){return r[5]; });
      var refType = getRefFac.map(function(r){return r[6]; });
      var refDept = getRefFac.map(function(r){return r[7]; });
      var refOffNum = getRefFac.map(function(r){return r[8]; });
      var refOffPh = getRefFac.map(function(r){return r[9]; });
      var refEmail = getRefFac.map(function(r){return r[10]; });
      var refFerp = getRefFac.map(function(r){return r[11]; });
      var refHomeAdd = getRefFac.map(function(r){return r[12]; });
      var refHomePh = getRefFac.map(function(r){return r[13]; });
  
      var endOfSheet = getFacPSID.getLastRow();
      var facMember = [];
  
      if (posi <= endOfSheet){
          facMember[0] = refPsid[posi];
          facMember[1] = refLn[posi];
          facMember[2] = refFn[posi];
          facMember[3] = refClass[posi];
          facMember[4] = refTime[posi];
          facMember[5] = refStat[posi];
          facMember[6] = refType[posi];
          facMember[7] = refDept[posi];
          facMember[8] = refOffNum[posi];
          facMember[9] = refOffPh[posi];
          facMember[10] = refEmail[posi];
          facMember[11] = refFerp[posi];
          facMember[12] = refHomeAdd[posi];
          facMember[13] = refHomePh[posi];
      }
  
      return facMember;

}


function scrThr(facIndex){
  
    var combinedFL = facultyName();
    var namePos = combinedFL.indexOf(facIndex);
  
    var returnedInfo = retrieveInfo(namePos);
  
    return returnedInfo;
  
}
function searchID(inputID){

  var getSS = SpreadsheetApp.openByUrl(url);
  var getWS = getSS.getSheetByName("Faculty");
  var getID = getWS.getRange(2,1,getWS.getLastRow(),1).getValues();
  var mapID = getID.map(function(r){return r[0]; });
  
  var idposi = mapID.indexOf(inputID);
  
  var returnedInfo = retrieveInfo(idposi);
  
  return returnedInfo;

}
function deleteFaculty(psid){
  
  Logger.log(psid);
  
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("Faculty");
  var psids = ws.getRange(2,1,ws.getLastRow(),1).getValues();
  
  var mapIds = psids.map(function(r){return r[0]; });

  
  var position = mapIds.indexOf(psid);

  var start, end;
  start = position + 2;
  end = position + 2;
  ws.deleteRows(start, 1);
  
  return "Faculty Member Deleted"

}


function saveUpdt(facIndex,psid,lname,fname,classif,timeb,stat,emtype,dpt,offnum,offph,email,ferp,homeadd,homeph){
    var getSS = SpreadsheetApp.openByUrl(url);
    var getWS = getSS.getSheetByName("Faculty");
    var getID = getWS.getRange(2,1,getWS.getLastRow(),1).getValues();
    var mapID = getID.map(function(r){return r[0]; });
  
    var namePos = mapID.indexOf(facIndex)+2;
  
    var psidRng = "A"+namePos;
    var lameRng = "B"+namePos;
    var fnameRng = "C"+namePos;
    var clssRng = "D"+namePos;
    var timeRng = "E"+namePos;
    var statRng = "F"+namePos;
    var typeRng = "G"+namePos;
    var dptRng = "H"+namePos;
    var offRng = "I"+namePos;
    var offPhRng = "J"+namePos;
    var emailRng = "K"+namePos;
    var ferpRng = "L"+namePos;
    var homeRng = "M"+namePos;
    var phRng = "N"+namePos;
  
  

    var updtPsid = getWS.getRange(psidRng)
    updtPsid.setValue(psid)
    var updtLName = getWS.getRange(lameRng)
    updtLName.setValue(lname)
    var updtFName = getWS.getRange(fnameRng);
    updtFName.setValue(fname)
    var updtClss = getWS.getRange(clssRng);
    updtClss.setValue(classif)
    var updtTime = getWS.getRange(timeRng);
    updtTime.setValue(timeb)
    var updtStat = getWS.getRange(statRng);
    updtStat.setValue(stat)
    var updtType = getWS.getRange(typeRng);
    updtType.setValue(emtype)
    var updtDpt = getWS.getRange(dptRng);
    updtDpt.setValue(dpt)
    var updtOffNum = getWS.getRange(offRng);
    updtOffNum.setValue(offnum)
    var updtOffPh = getWS.getRange(offPhRng);
    updtOffPh.setValue(offph)
    var updtEmail = getWS.getRange(emailRng);
    updtEmail.setValue(email)
    var updtFerp = getWS.getRange(ferpRng);
    updtFerp.setValue(ferp)
    var updtHomeAd = getWS.getRange(homeRng);
    updtHomeAd.setValue(homeadd)
    var updtHmPh = getWS.getRange(phRng);
    updtHmPh.setValue(homeph)
    
}

function deleteClasses(resp){ 

  var getSS = SpreadsheetApp.openByUrl(url);
  var getWS = getSS.getSheetByName("Classes");
  
     var start, end;
     start = 1;
     end = getWS.getLastRow();
     getWS.deleteRows(start, end);
  return "Delete Successful";

}

function deleteHours(){

  var getSS = SpreadsheetApp.openByUrl(url);
  var getWS = getSS.getSheetByName("Officehours");
  
     var start, end;
     start = 2
     end = getWS.getLastRow();
     getWS.deleteRows(start, end);
  
  return "Delete Successful";

}


//End Faculty Editing Functions
//END EDITING/MAINTENANCE
//----------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------