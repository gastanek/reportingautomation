
//global variables
var configSheet
var processStep
var lastProcDate
var masterReportLocation
var masterReport
var masterReportLastSheet
var masterCurrentSheetName
var newReportSheet
var reportingErrorMessage = "None";


//control variables
var newWeekSheetReady = false;
var reportingError = false;

function initializeConfigs() {
  //grab the config sheet
  configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master');
  
  //get the local configs grab the config range
  processStep = configSheet.getRange("B3").getValue();
  lastProcDate = configSheet.getRange("B4").getValue();
  masterReportLocation = configSheet.getRange("B5").getValue();
  masterReportLastSheet = configSheet.getRange("B6").getValue();
  

}

function checkMasterReport() {
  masterReport = DriveApp.getFileById(masterReportLocation);
  //open as a spreadsheet and grab the latest sheet
  var reportSheet = SpreadsheetApp.open(masterReport);
  masterCurrentSheetName = reportSheet.getActiveSheet().getSheetName();
  
  //now compare the last sheet we wrote too and this current one
  //this may be a bit of a bug as it makes it a one time shot
  //perhaps I want to add a flag for "new" or "update"
  //but I mostly want to be sure in this case that I'm only
  //writing to the new sheet when the Master has created it and not writing over the old one
  
  if(masterCurrentSheetName != masterReportLastSheet) {
    ///new sheet is present, set the flag to true so processing can continue
    //assuming the active sheet is the new sheet and not just an older one (Yikes!)
    configSheet.getRange("C6").setValue(masterCurrentSheetName);
    newWeekSheetReady = true;
  }    
  configSheet.getRange("B7").setValue(newWeekSheetReady);
}


//START SLACK INTERACTION SECTION - CODE IN THIS AREA IS SPECIFICALLY RELATED TO ACCEPTING AND PROCESSING SLACK MESSAGES

//slack response comments
var cheekyResponseComments = new Array("That comment was added.",
                                       "Pushed it onto the stack.", 
                                       "And another comment is in there.", 
                                       "You got it.  Done.", 
                                       "You aren't just making this up, are you?", 
                                       "Done.  You get what you measure.",
                                       "Impressive. :-/",
                                       "I can't believe it's not butter",
                                       "Threw that shrimp on the barbie",
                                       "Do you think there is a special void in space where these comments go?",
                                       "That update right there will stop the presses.",
                                       "Welp, Allstar Comment!!!",
                                       "Good job.  You can take a vacation now.",
                                       "I've seen better comments from Twitter trolls.",
                                       "Look the other way...I just wrote that to dev -> null" );

function doPost(e){
  try{
    postMessage = JSON.parse(e.postData.contents)
    if(postMessage.type == "url_verification"){
      //this is the challenge request
      return ContentService.createTextOutput(postMessage.challenge);
    }
  }
    catch(error){
      console.log(error)
      }
  
    numComments = acceptSlackMessage(e);
    
    if(numComments == 0) {
      return ContentService.createTextOutput("Added that for you.  That is the first comment this week.");
    }
    else {
      totalcomments = numComments + 1;
      cheekyResponse = cheekyResponseComments[Math.floor(Math.random() * (cheekyResponseComments.length - 1))] 
      return ContentService.createTextOutput(cheekyResponse);
  
    }

}

function acceptSlackMessage(messagedata){
  
  //do some work on the messagedata to determine if this is comment or stats on onboardings
  var msgbits = messagedata.parameter.text.split(/::/);
  
  var sheet = msgbits[0];
  
  if (sheet.indexOf("Onboarding") >=0) {
   //this is an onboarding  message
    var incrementnumber = parseInt(msgbits[1]);
    var currentValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(5,3).getValue();
    currentValue += incrementnumber;
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet).getRange(5,3).setValue(currentValue);
  }
  
  //if it isn't an onboarding, it's comments
  if (sheet.indexOf("Comments") >=0) {
  var testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);
  lastRow = testSheet.getLastRow();

  //post the specific text directly to the sheet  
  testSheet.getRange(lastRow+1, 1).setValue(msgbits[1]);    
  }
    if (sheet.indexOf("Action") >=0) {
   //this is an instruction to perform an action
   //for now the only action is publish report so we'll just call main.
    //if I come up with more actions to perform I'll need to parse msgbits[1] to know what to do
      
    mainLoop();
  }
  
  //else this is something unknown, just ignore for now
  
}
  
  
//END OF SLACK INTERACTION SECTION

function getMyReports() {
  //a fragile, hacky way to get my report data from my email
  //the really good way would be to access the reports directly from SFDC
  
  var stagingSheets = new Array("ThisQtr", "NextQtr", "TotalQtr", "OppCount", "POCData", "Comments");
  
  var label = GmailApp.getUserLabelByName("SfdcReport");
  //get the threads in the report inbox - there should never be more than what we need -  which is to say this is a likely spot for a bug becuase it assumes the content with the label is pristine
  var threads = label.getThreads();
  
  for (var tcount=0; tcount< threads.length; tcount++){
    var thread = threads[tcount];
    var mail = thread.getMessages()[0];
    var reportSubject = mail.getSubject();
    var reportMessage = mail.getBody();
    var stripHTML = reportMessage.replace(/<[^>]*>/g, '');
    var stripSpacing = stripHTML.replace(/&nbsp;*/g, '');
    var lineArray = stripSpacing.split(/\r?\n/);
    var reportSheet = "Garbage";
    //set the sheet we are working with
    if(reportSubject.indexOf("Opps close this qtr") >=0) { reportSheet = stagingSheets[2] };
    if(reportSubject.indexOf("this week, close this qtr") >=0) { reportSheet = stagingSheets[0] };
    if(reportSubject.indexOf("ths week, future qtrs") >=0) { reportSheet = stagingSheets[1] };
    if(reportSubject.indexOf("All Opps with SA Actions") >=0) { reportSheet = stagingSheets[3] };
    if(reportSubject.indexOf("POC actions") >=0) {reportSheet = stagingSheets[4] };
    if(reportSubject.indexOf("Onboarding") >=0) {
      var currentValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(3,3).getValue();
      currentValue += 1;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(3,3).setValue(currentValue);
    } //and then just write out whatever to Garbage
      
    var theSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reportSheet);  
    var a=2;
    for (i=0; i<lineArray.length; i++){
      
      if(lineArray[i].toString()) { //linArray val is not empty
        theSheet.getRange(a,1).setValue(lineArray[i].replace(/Record Count/, ''));
        a++;
        
      }
    }    
  }
}


  
  
  function cleanItAllUp() {
   //post processing will txfr nbrs from the individual tabs to the "anchor" tab.  This tab is exactly like the section in the main reporting page
   //we can make some safe assumptions about location of records in this spreadsheet - but the assumptions means this is also a likely spot for a bug at some point.
   //column C is our description section
   
    var anchorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anchor");
    
    var sheetList = new Array( SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ThisQtr"),
                               SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NextQtr"),
                               SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TotalQtr"),
                               SpreadsheetApp.getActiveSpreadsheet().getSheetByName("POCData"));
    
    //process actions in this quarter
    for(var i=10; i<=sheetList[0].getLastRow(); i++){
      //skip ahead to 10 because there is a lot of noise from the SFDC report at the top
      var americas, international;
      
      //iterate over each row in the sheet, find "Grand Total" and process the next two
      if(sheetList[0].getRange(i, 1).getValue().toString().indexOf("Grand Total") >=0 ) {
        //get the next two values yo
        americas = sheetList[0].getRange(i+1, 1).getValue();
        international = sheetList[0].getRange(i+2, 1).getValue();
      }
    //push those numbers into the anchor sheet
      anchorSheet.getRange(12,4).setValue(americas);
      anchorSheet.getRange(12,5).setValue(international);
    
    
    }
    
    //process actions for next quarter
    for(var i=10; i<=sheetList[1].getLastRow(); i++){
      //skip ahead to 10 because there is a lot of noise at the top
      var americas, international;
      
      //iterate over each row in the sheet, find "Grand Total" and process the next two
      if(sheetList[1].getRange(i, 1).getValue().toString().indexOf("Grand Total") >=0 ) {
        //get the next two values yo
        americas = sheetList[0].getRange(i+1, 1).getValue();
        international = sheetList[0].getRange(i+2, 1).getValue();
      }
    //push those numbers into the anchor sheet
      anchorSheet.getRange(13,4).setValue(americas);
      anchorSheet.getRange(13,5).setValue(international);
    
    
    }
    
    //process all actions for opps this quarter
    for(var i=8; i<=sheetList[2].getLastRow(); i++){
      //from "International - Field" to "Grand Total"
      var start, end;
      if(sheetList[2].getRange(i, 1).getValue().toString().indexOf("International - Field") >=0 ) {
        start = i;
      }
      if(sheetList[2].getRange(i, 1).getValue().toString().indexOf("Grand Total") >=0 ) {
        end = i;
      }
    }

      //process content between start and end
      
      for(start; start < end; start++){
        //anchor rows: 1. 17; 2. 18; 3. 19; 4. 20; 5. 21; 6. 22; 7. 23; 8. 24; 
        //anchor - cols 4,5; 
        
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("01. Discovery Conversation") >=0 ) {
        
        anchorSheet.getRange(17,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(17,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("02. Solution/Use Case") >=0 ) {
        
        anchorSheet.getRange(18,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(18,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("03. Capabilities Presentation") >=0 ) {
        
        anchorSheet.getRange(19,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(19,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("04. Solution or Architecture Design Conversation") >=0 ) {
        
        anchorSheet.getRange(20,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(20,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("05. Opportunity Demo") >=0 ) {
        
        anchorSheet.getRange(21,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(21,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("06. Solution Support") >=0 ) {
        
        anchorSheet.getRange(22,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(22,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("07. POC") >=0 ) {
        
        anchorSheet.getRange(23,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(23,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[2].getRange(start, 1).getValue().toString().indexOf("08. Developer/Customer Workshop") >=0 ) {
        
        anchorSheet.getRange(24,4).setValue(sheetList[2].getRange(start+1, 1).getValue());
        anchorSheet.getRange(24,5).setValue(sheetList[2].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
     }    

    //poc data
    for(var i=10; i<=sheetList[3].getLastRow(); i++){
      //from "International - Field" to "Grand Total"
      var start, end;
      if(sheetList[3].getRange(i, 1).getValue().toString().indexOf("International - Field") >=0 ) {
        start = i;
      }
      if(sheetList[3].getRange(i, 1).getValue().toString().indexOf("Grand Total") >=0 ) {
        end = i;
      }
    }

      //process content between start and end
      
      for(start; start < end; start++){
        //anchor rows: 1. 5; 2. 6; 3. 7; 4. 8; 5. 9;
        //anchor - cols 4,5; 
        
      if(sheetList[3].getRange(start, 1).getValue().toString().indexOf("Planned") >=0 ) {
        
        anchorSheet.getRange(5,4).setValue(sheetList[3].getRange(start+1, 1).getValue());
        anchorSheet.getRange(5,5).setValue(sheetList[3].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[3].getRange(start, 1).getValue().toString().indexOf("In-Progress") >=0 && sheetList[3].getRange(start, 1).getValue().toString().indexOf("Stalled") < 0) {
        
        anchorSheet.getRange(6,4).setValue(sheetList[3].getRange(start+1, 1).getValue());
        anchorSheet.getRange(6,5).setValue(sheetList[3].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[3].getRange(start, 1).getValue().toString().indexOf("Stalled") >=0 ) {
        
        anchorSheet.getRange(7,4).setValue(sheetList[3].getRange(start+1, 1).getValue());
        anchorSheet.getRange(7,5).setValue(sheetList[3].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[3].getRange(start, 1).getValue().toString().indexOf("Completed - Successful") >=0 ) {
        
        anchorSheet.getRange(8,4).setValue(sheetList[3].getRange(start+1, 1).getValue());
        anchorSheet.getRange(8,5).setValue(sheetList[3].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }
      if(sheetList[3].getRange(start, 1).getValue().toString().indexOf("Completed - Problematic") >=0 ) {
        
        anchorSheet.getRange(9,4).setValue(sheetList[3].getRange(start+1, 1).getValue());
        anchorSheet.getRange(9,5).setValue(sheetList[3].getRange(start+2, 1).getValue());
        start+=2;
        continue;
      }

    
      }
    
    //now take care of Onboardings
    //current active = currentactivelast week + new this week - completed this week
      var lastActive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(11,2).getValue();
      var currentNew = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(3,3).getValue();
      var currentClosed = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(5,3).getValue();
      var newActive = lastActive + currentNew - currentClosed;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(4,3).setValue(newActive);
      var lastQtrClosed = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(12,2).getValue();
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(6,3).setValue(lastQtrClosed + currentClosed);
    
      //copy this over to the anchorsheet
      anchorSheet.getRange(32,3,4).setValues(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(3,3,4).getValues());
 
    //finish with the comments
    var newComments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comments").getRange(1,1,8).getValues();
    var anchorComments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anchor").getRange(40,2,8);
    anchorComments.setValues(newComments);
    
    
  }
  

function sanityCheck() {
 //do a basic sanity check on the content before I go through posting
  //make sure numbers are actually numbers, etc
  
  var anchorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anchor");
  
  //poc data check
  var values = anchorSheet.getRange(5,4,5,2).getValues();
  for (var i=0; i<5;i++){
    for (var a=0; a<2; a++) {
    var thevalue = values[i][a];
      if (isNaN(thevalue)) {
      //we've got a problem in our sheet
      reportingError = true;
      reportingErrorMessage = "Non-number set in the reporting fields";
      }
    }
  }
  //actions data check
  var values = anchorSheet.getRange(12,4,2,2).getValues();
  for (var i=0; i<2;i++){
    for (var a=0; a<2; a++) {
    var thevalue = values[i][a];
      if (isNaN(thevalue)) {
      //we've got a problem in our sheet
      reportingError = true;
      reportingErrorMessage = "Non-number set in the reporting fields";
      }
    }
  }
  //total actions data check
  var values = anchorSheet.getRange(17,4,8,2).getValues();
  for (var i=0; i<8;i++){
    for (var a=0; a<2; a++) {
    var thevalue = values[i][a];
      if (isNaN(thevalue)) {
      //we've got a problem in our sheet
      reportingError = true;
      reportingErrorMessage = "Non-number set in the reporting fields";
      }
    }
    
    
  }  
  
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(8,2).setValue(reportingError);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(9,2).setValue(reportingErrorMessage);
  
}

  function postToMaster() {
   //this is where we are taking the well prepared values and putting them into the master report
   initializeConfigs();
   checkMasterReport();
   var x,y;
   var masterReport = DriveApp.getFileById(masterReportLocation);
   //open as a spreadsheet and grab the latest report sheet
   var reportSheet = SpreadsheetApp.open(masterReport);
  
   currentSheet = reportSheet.getSheetByName(masterCurrentSheetName);
   //find my anchor point in the master report
   var bigRange = currentSheet.getRange(1,3,100,2).getValues();
    for (var i=0; i<100; i++) {
      for(var a=0; a<2; a++) {
        if(bigRange[i][a].indexOf("POC Activity for active Opportunities") >= 0){
          x=i+1;
          y=a+3;
        }
      }
    }
    if (x<1 && y<1) {
     //failed to find the anchor point. 
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(8,2).setValue(true);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(9,2).setValue("Failed to find the anchor point in the master report sheet.");
      return;
    }
    
    //x and y represent my range anchor points.  now just copy from my current sheet to the master report sheet
    //note this is a likely spot for a bug: if someone changes my expected content, I'll fail.  If someone changes the specific row/col orientation, my position assumptions will fail
    //anchorsheet
    var anchorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anchor");

    //poc data
    var masterPocRange = currentSheet.getRange(x,y+2, 5, 2);
    var anchorPocRange = anchorSheet.getRange(5,4,5,2).getValues();
    masterPocRange.setValues(anchorPocRange);
    
    //weekly actions
    var masterWeekActions = currentSheet.getRange(x+7, y+2, 2,2);
    var anchorWeekActions = anchorSheet.getRange(12,4,2,2).getValues();
    masterWeekActions.setValues(anchorWeekActions);
    
    //quarter actions
    var masterQuarterActions = currentSheet.getRange(x+12, y+2, 8,2);
    var anchorQuarterActions = anchorSheet.getRange(17,4,8,2).getValues();
    masterQuarterActions.setValues(anchorQuarterActions);
    
    //onboarding actions
    var masterOnboarding = currentSheet.getRange(x+28, y+1, 4);
    var anchorOnboarding = anchorSheet.getRange(32,3,4).getValues();
    masterOnboarding.setValues(anchorOnboarding);
    
    //comments - max of 8
    var masterComments = currentSheet.getRange(x+36, y, 8);
    var anchorComments = anchorSheet.getRange(40,2,8).getValues();
    masterComments.setValues(anchorComments);
    
    //we have completed posting our stuff to the master report    
  }
         
         
  function prepForNextTime() {
    //this is run after all the processing is complete and cleans things up so the next run is current
    //archive the tab
    var anchorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anchor");
    var copySheet = anchorSheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
    copySheet.setName(new Date());
    
    //set anchor sheet values back to 0
    anchorSheet.getRange(5,4,5).setValue(0);
    anchorSheet.getRange(5,5,5).setValue(0);
    anchorSheet.getRange(12,4,2).setValue(0);
    anchorSheet.getRange(12,5,2).setValue(0);
    anchorSheet.getRange(17,4,8).setValue(0);
    anchorSheet.getRange(17,5,8).setValue(0);
    anchorSheet.getRange(32,3,4).setValue(0);

    //delete content from the staging pages, do this so that I avoid duplicating numbers
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ThisQtr").deleteRows(1, 50);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NextQtr").deleteRows(1, 50);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TotalQtr").deleteRows(1, 50);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("POCData").deleteRows(1, 50);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comments").deleteRows(1, 50);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(3,3,4).setValue(0);
    
    //update the Master sheet so we know that we processed this week (and don't dup work if Master report hasn't updated his sheet)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master').getRange("B6").setValue(masterCurrentSheetName);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(11,2).setValue(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(4,3).getValue());
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(12,2).setValue(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding").getRange(6,3).getValue());

    //remove the sfdcreport label from the emails from my mail box
    var label = GmailApp.getUserLabelByName("SfdcReport");
    var threads = label.getThreads();
  
    for (var tcount=0; tcount< threads.length; tcount++){
      threads[tcount].removeLabel(label);
    }

    //update the master records on last run time
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(4,2).setValue(new Date());

    
    //send me a message letting me know how things went
    copyForMail();
    
  }

function copyForMail() {
  //we only get here because we've already checked that everything is super awesome
  var subject = "Reporting completed successfully";
  var destination = "emailaddress.com";
  var message = "Weekly reporting has completed without errors.";

  var anchorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anchor");
  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("TempReportingSheet");
  anchorSheet.copyTo(newSpreadsheet);
  
  // Delete the sheet we don't want
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  var tempReportFile = DriveApp.getFileById(newSpreadsheet.getId());
  var pdfVersion = tempReportFile.getBlob().getAs('application/pdf').setName("TempReportPDF");

  var tempPdfReport = DriveApp.createFile(pdfVersion);
  
  // Send me a copy via email
  MailApp.sendEmail(destination, subject, message, {attachments:[tempPdfReport]});
  
  // Delete the temp files
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
  DriveApp.getFileById(tempPdfReport.getId()).setTrashed(true);
}
  
  
  function mainLoop() {
    //initialize the configs
    initializeConfigs();
    
    //check first if we should be doing this because the Master report is ready
    checkMasterReport();  
    
    //if we are not looking at a new sheet and this isn't an update, then break
    if(newWeekSheetReady != true){
      //email me and let me know this broke
      var subject = "Reporting failed - no new sheet ready";
      var destination = "emailaddress.com";
      var message = "Weekly reporting failed because we did not find a new sheet ready in the Master report.";
      MailApp.sendEmail(destination, subject, message);
      return;
    }
    
    //grab the mail
    getMyReports();
    
    //make it pretty
    cleanItAllUp();
    
    //make sure it didn't go crazy
    sanityCheck();
    
    //if something is wrong, back out now and let me know
    if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(8,2).getValue() === true){
      var subject = "Reporting failed - error in processing";
      var destination = "emailaddress.com";
      var message = "Weekly reporting failed due to a processing error.";
      MailApp.sendEmail(destination, subject, message);
      
     return; 
    }

    //all good, keep going    
    //post it to the master report
    postToMaster();
    
    if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").getRange(8,2).getValue() === true){
      var subject = "Reporting failed - error while posting to master";
      var destination = "emailaddress.com";
      var message = "Weekly reporting failed while posting to the master report";
      MailApp.sendEmail(destination, subject, message);
      
     return; 
    }

    
    
    //make it ready for next time
    prepForNextTime();
    
    
}

