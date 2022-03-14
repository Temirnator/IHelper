var CONF;
///Password: IHelper2022#@!$
function onOpen() {
  var ui = SpreadsheetApp.getUi();  //For convenience
  var userEmail = Session.getActiveUser().getEmail();
  if(userEmail=='ihelper@nu.edu.kz'){
    //Creates a button 'Send All' in the top panel of the Spreadsheet
    ui.createMenu('Send All')
      .addItem('Email 01','sender_email_01')
      .addItem('Email 02','sender_email_02')
      .addItem('Email 03','sender_email_03')
      .addItem('Email 04','sender_email_04')
      .addItem('Email 05','sender_email_05')
      .addItem('Email 06','sender_email_06')
    .addToUi();

    //Creates a button 'Send Selected' in the top panel of the Spreadsheet
    ui.createMenu('Send Selected')
      .addItem('Email 01','slctsender_email_01')
      .addItem('Email 02','slctsender_email_02')
      .addItem('Email 03','slctsender_email_03')
      .addItem('Email 04','slctsender_email_04')
      .addItem('Email 05','slctsender_email_05')
      .addItem('Email 06','slctsender_email_06')
    .addToUi();

    //Creates a button 'Generate Certificates' in the top panel of the Spreadsheet
    ui.createMenu('Generate Certificates')
      .addItem('Certificate Standard','certificateGen_standard')
      .addItem('Certificate Tutors Award','certificateGen_tutorsAward')
    .addToUi();
  }
}

////////////
//Email 01//
////////////
function sender_email_01(){
  var subject = 'IHelper Email 01 - Welcome letter (and important information)';//'Subject' variable
  var template_name = 'ihelper_email_01';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources folder
  sender(subject,template_name, attachment);
}

function slctsender_email_01(){
  var subject = 'IHelper Email 01 - Welcome letter (and important information)';//'Subject' variable
  var template_name = 'ihelper_email_01';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources
  slctsender(subject,template_name, attachment);
}

////////////
//Email 02//
////////////
function sender_email_02(){
  var subject = 'IHelper Email 02 - Kid Assignment';//'Subject' variable
  var template_name = 'ihelper_email_02';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources folder
  sender(subject,template_name, attachment);
}

function slctsender_email_02(){
  var subject = 'IHelper Email 02 - Kid Assignment';//'Subject' variable
  var template_name = 'ihelper_email_02';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources
  slctsender(subject,template_name, attachment);
}

////////////
//Email 03//
////////////
function sender_email_03(){
  var subject = 'IHelper Email 03 - Assessment Report';//'Subject' variable
  var template_name = 'ihelper_email_03';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources folder
  sender(subject,template_name, attachment);
}

function slctsender_email_03(){
  var subject = 'IHelper Email 03 - Assessment Report';//'Subject' variable
  var template_name = 'ihelper_email_03';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources
  slctsender(subject,template_name, attachment);
}

////////////
//Email 04//
////////////
function sender_email_04(){
  var subject = 'IHelper Email 04 - Weekly Report';//'Subject' variable
  var template_name = 'ihelper_email_04';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources folder
  sender(subject,template_name, attachment);
}

function slctsender_email_04(){
  var subject = 'IHelper Email 04 - Weekly Report';//'Subject' variable
  var template_name = 'ihelper_email_04';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources
  slctsender(subject,template_name, attachment);
}

////////////
//Email 05//
////////////
function sender_email_05(){
  var subject = 'IHelper Email 05 - End of Period, Certificate Award, Continuation survey';//'Subject' variable
  var template_name = 'ihelper_email_05';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources folder
  sender(subject,template_name, attachment);
}

function slctsender_email_05(){
  var subject = 'IHelper Email 05 - End of Period, Certificate Award, Continuation survey';//'Subject' variable
  var template_name = 'ihelper_email_05';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources
  slctsender(subject,template_name, attachment);
}

////////////
//Email 06//
////////////
function sender_email_06(){
  var subject = 'IHelper Email 06 - Ban from the project, Blocklisting';//'Subject' variable
  var template_name = 'ihelper_email_06';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources folder
  sender(subject,template_name, attachment);
}

function slctsender_email_06(){
  var subject = 'IHelper Email 06 - Ban from the project, Blocklisting';//'Subject' variable
  var template_name = 'ihelper_email_06';//template name
  var attachment = DriveApp.getFolderById('1jg7Q4-QueE8YdZHopjoTqSN0o2ljoAyU');//tutor resources
  slctsender(subject,template_name, attachment);
}

//////////
//Sender//
//////////
function sender(subject, template_name, attachment){
  //Spreadsheet navigation
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  //Sheets by names
  var tutorSheet = ss.getSheetByName('Tutors Active');
  var accountsSheet = ss.getSheetByName('Accounts');
  var mainPageSheet = ss.getSheetByName('Main Page');
  var emailHistorySheet = ss.getSheetByName('Email History');

  //Email contents
  var temp = HtmlService.createTemplateFromFile(template_name);//Generate an HTML template named temp from the 'template_name' template
  var date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');
  var senderData = senderClass();

  //User Data. Gets email manually
  var userEmailManual = ui.prompt('Enter your email:');
  CONF = ui.alert('Confirm to Send all', '',ui.ButtonSet.YES_NO);//Add Alert that asks Confirm Yes or No
  if (CONF !== ui.Button.YES) {//stop the function if the button YES is not pressed
    return;
  }

  //Gets name, phone, position from the 'Accounts' sheet 
  var userNameManual = null;
  var userPhoneManual = null;
  var userPositionManual = null;

  //Searches for the userEmailManual in the 'Accounts' sheet. 
  for(var i=1;i<accountsSheet.getLastRow();i++){
    var data = accountsSheet.getRange(i+1,1,1,accountsSheet.getLastColumn()).getValues()[0];
    if(userEmailManual==data[0]){
      userNameManual      = data[1];
      userPhoneManual     = data[2];
      userPositionManual  = data[3];
      break;
    }
  }

  //Validates that the userEmailManual is in the 'Accounts' list and its data can be allowed to send emails
  if(userNameManual==null){
    ui.alert('You are not authorized to send emails','',ui.ButtonSet.OK);
    return;
  } 
  //???Curious if it works
  senderData = {
    sendername: userNameManual, 
    senderposition: userPositionManual, 
    senderemail: userEmailManual, 
    senderphone: userPhoneManual
  };

  //Main for loop that sends emails for every row in the Tutors Active Sheet
  for(var i=1; i<lastrow;i++){
    var data = tutorSheet.getRange(i+1, 1, 1, tutorSheet.getLastColumn()).getValues()[0];//Transfers data from row to an array named 'data'
    senderData = {
      email:          data[2],
      fullname:       data[3],
      phone:          data[4],
      
      langPref:       data[8],
      genderPref:     data[9],
      gradePref:      data[10],
      specChildPref:  data[11],
      comments:       data[12],

      kidname:        data[13],
      kidphone:       data[14],

      kidlang:        data[15],
      kidgrade:       data[16],
      kidSVcategory:  data[17],
      
      parentname:     data[18],
      parentphone:    data[19],

      kidComments:    data[21],
      certificateLink: data[38]
      //2  B Email Address	
      //3  C Full name	
      //4  D Phone number	
      //5  E Certificates Link	
      //6  F Experience in tutoring	
      //7  G Why do you want to volunteer? 	
      //8  H Language Preference	
      //9  I Gender Preference	
      //10 J Grade Preference	
      //11 K SpecChild Preference	
      //12 L Comments
      //13 M Kid name	
      //14 N Kid phone (F)	
      //15 O Kid language (F)	
      //16 P Kid grade (F)	
      //17 Q Kid SVC (F)	
      //18 R Parent name (F)	
      //19 S Parent phone (F)	
      //20 T Docs link (F)	
      //21 U Comments (F)
      
    }
    
    temp.senderData=senderData;//get values for a template using fields from the 'tutor' class
    var message = temp.evaluate().getContent();//generate a message from the template and current variables
    MailApp.sendEmail({//Send email 
    to: senderData.email//to 'tutor.email'
    , subject: subject//with subject of the email as 'subject' variable contents,
    , htmlBody: message//with a message generated from an HTML template 'message'
    , attachments: attachment
    
    });
    //Emails History
    emailHistorySheet.getRange(emailHistorySheet.getLastRow()+1,1).setValue(date);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),2).setValue(senderData.senderemail);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),3).setValue(senderData.sendername);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),4).setValue(senderData.email);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),5).setValue(senderData.fullname);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),6).setValue(subject);
    //
    //22 V  Email 01 (S)	
    //23 W  Email 02 (S)	
    //24 X  Email 03 (S)	
    //25 Y  Email 04 (S)	
    //26 Z  Email 05 (S)	
    //27 AA Email 06 (S)
    if(template_name == "ihelper_email_01"){
      tutorSheet.getRange(i+1,22).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_02"){
      tutorSheet.getRange(i+1,23).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_03"){
      tutorSheet.getRange(i+1,24).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_04"){
      tutorSheet.getRange(i+1,25).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_05"){
      tutorSheet.getRange(i+1,26).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_06"){
      tutorSheet.getRange(i+1,27).setValue("Was sent on "+date);
    }
  }
  return;
  
}

function slctsender(subject, template_name, attachment){
  //Spreadsheet navigation
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  //Sheets by names
  var tutorSheet = ss.getSheetByName('Tutors Active');
  var accountsSheet = ss.getSheetByName('Accounts');
  var mainPageSheet = ss.getSheetByName('Main Page');
  var emailHistorySheet = ss.getSheetByName('Email History');

  //Email contents
  var temp = HtmlService.createTemplateFromFile(template_name);//Generate an HTML template named temp from the 'template_name' template
  var date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');
  var senderData = senderClass();

  //User Data. Gets email manually
  var userEmailManual = ui.prompt('Enter your email:');
  CONF = ui.alert('Confirm to Send all', '',ui.ButtonSet.YES_NO);//Add Alert that asks Confirm Yes or No
  if (CONF !== ui.Button.YES) {//stop the function if the button YES is not pressed
    return;
  }

  //Gets name, phone, position from the 'Accounts' sheet 
  var userNameManual = null;
  var userPhoneManual = null;
  var userPositionManual = null;

  //Searches for the userEmailManual in the 'Accounts' sheet. 
  for(var i=1;i<accountsSheet.getLastRow();i++){
    var data = accountsSheet.getRange(i+1,1,1,accountsSheet.getLastColumn()).getValues()[0];
    if(userEmailManual==data[0]){
      userNameManual      = data[1];
      userPhoneManual     = data[2];
      userPositionManual  = data[3];
      break;
    }
  }

  //Validates that the userEmailManual is in the 'Accounts' list and its data can be allowed to send emails
  if(userNameManual==null){
    ui.alert('You are not authorized to send emails','',ui.ButtonSet.OK);
    return;
  } 
  //???Curious if it works
  senderData = {
    sendername: userNameManual, 
    senderposition: userPositionManual, 
    senderemail: userEmailManual, 
    senderphone: userPhoneManual
  };

  //difference #1 between sender and slctsender functions: the range variable is only in slctsender
  var range =  tutorSheet.getActiveRange();//Gets active range in the Tutors Active Sheet

  //Main for loop that sends emails for every row in the Tutors Active Sheet
  //difference #2 between sender and slctsender functions: for loop ends at range.getNumRows()
  for(var i=1; i<range.getNumRows();i++){
    //Transfers data from row to an array named 'data'
    //difference #3 between sender and slctsender functions: 'range.getRow()+i' instead of 'i+1'
    var data = tutorSheet.getRange(range.getRow()+i, 1, 1, tutorSheet.getLastColumn()).getValues()[0];
    senderData = {
      email:          data[2],
      fullname:       data[3],
      phone:          data[4],
      
      langPref:       data[8],
      genderPref:     data[9],
      gradePref:      data[10],
      specChildPref:  data[11],
      comments:       data[12],

      kidname:        data[13],
      kidphone:       data[14],

      kidlang:        data[15],
      kidgrade:       data[16],
      kidSVcategory:  data[17],
      
      parentname:     data[18],
      parentphone:    data[19],

      kidComments:    data[21],
      certificateLink: data[38]
      //2  B Email Address	
      //3  C Full name	
      //4  D Phone number	
      //5  E Certificates Link	
      //6  F Experience in tutoring	
      //7  G Why do you want to volunteer? 	
      //8  H Language Preference	
      //9  I Gender Preference	
      //10 J Grade Preference	
      //11 K SpecChild Preference	
      //12 L Comments
      //13 M Kid name	
      //14 N Kid phone (F)	
      //15 O Kid language (F)	
      //16 P Kid grade (F)	
      //17 Q Kid SVC (F)	
      //18 R Parent name (F)	
      //19 S Parent phone (F)	
      //20 T Docs link (F)	
      //21 U Comments (F)
      
    }
    
    temp.senderData=senderData;//get values for a template using fields from the 'tutor' class
    var message = temp.evaluate().getContent();//generate a message from the template and current variables
    MailApp.sendEmail({//Send email 
    to: senderData.email//to 'tutor.email'
    , subject: subject//with subject of the email as 'subject' variable contents,
    , htmlBody: message//with a message generated from an HTML template 'message'
    , attachments: attachment
    
    });
    //Emails History
    emailHistorySheet.getRange(emailHistorySheet.getLastRow()+1,1).setValue(date);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),2).setValue(senderData.senderemail);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),3).setValue(senderData.sendername);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),4).setValue(senderData.email);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),5).setValue(senderData.fullname);
    emailHistorySheet.getRange(emailHistorySheet.getLastRow(),6).setValue(subject);
    //
    //22 V  Email 01 (S)	
    //23 W  Email 02 (S)	
    //24 X  Email 03 (S)	
    //25 Y  Email 04 (S)	
    //26 Z  Email 05 (S)	
    //27 AA Email 06 (S)
    
    //difference #4 between sender and slctsender functions: 'range.getRow()+i' instead of 'i+1'
    if(template_name == "ihelper_email_01"){
      tutorSheet.getRange(range.getRow()+i,22).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_02"){
      tutorSheet.getRange(range.getRow()+i,23).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_03"){
      tutorSheet.getRange(range.getRow()+i,24).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_04"){
      tutorSheet.getRange(range.getRow()+i,25).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_05"){
      tutorSheet.getRange(range.getRow()+i,26).setValue("Was sent on "+date);
    }else if(template_name == "ihelper_email_06"){
      tutorSheet.getRange(range.getRow()+i,27).setValue("Was sent on "+date);
    }
  }


  return;
}

//replace for globaltutor
function senderClass(){
  //Spreadsheet navigation
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainPageSheet = ss.getSheetByName('Main Page');

  var senderClassData = {//senderClassData class
    //some fields change as the row changes (fullname, phone, email,etc)
      fullname : null, phone: null, email: null
    , langPref: null, genderPref: null, gradePref: null, specChildPref: null, comments: null
    
    //Kid data
    , kidname: null, kidphone: null
    , kidlang: null, kidgrade: null, kidgender: null, kidSVcategory: null, kidComments: null
    
    //Parent data
    , parentname: null, parentphone: null
    
    //Sender data
    , sendername: null, senderposition: null, senderemail: null, senderphone: null
    
    //Main Coordinator data
    , maincoord:      mainPageSheet.getRange(2,2).getValue()//others remain the same for all rows (maincoord, maincoordnum, maincoordemail,etc)
    , maincoordnum:   mainPageSheet.getRange(3,2).getValue()
    , maincoordemail: mainPageSheet.getRange(4,2).getValue()
    
    //NU Red Crescent Society Head data
    , nurchead:       mainPageSheet.getRange(7,2).getValue()//others remain the same for all rows (maincoord, maincoordnum, maincoordemail,etc)
    , nurcheadnum:    mainPageSheet.getRange(8,2).getValue()
    , nurcheademail:  mainPageSheet.getRange(9,2).getValue()
    
    //Links to resources
    , tutorResources:       mainPageSheet.getRange(12,2).getValue()
    , telegramchat:       mainPageSheet.getRange(13,2).getValue()
    , academicMaterials:  mainPageSheet.getRange(14,2).getValue()
  

    //Links to surveys
    , endOfTermSurvey:                  mainPageSheet.getRange(17,2).getValue() // email 5
    , feedbackForm:                     mainPageSheet.getRange(18,2).getValue() // email 1,2,3,4,5,6
    , kidAssignmentConfirmationForm:    mainPageSheet.getRange(19,2).getValue() // email 2
    , lessonReportsForm:                mainPageSheet.getRange(20,2).getValue() // email 3,4,6
    , participationAgreementForm:       mainPageSheet.getRange(21,2).getValue() // email 1
    , tutorsRecruitmentForm:            mainPageSheet.getRange(22,2).getValue() // recruitment email
    , kidEnrollment:                    mainPageSheet.getRange(23,2).getValue() // enrollment
    , withdrawalForm:                   mainPageSheet.getRange(24,2).getValue() // email 2,3,4
    , explanatoryForm:                  mainPageSheet.getRange(25,2).getValue() // email 6
    //17 End of Term - Continuation Survey 
    //18 Feedback Form
    //19 Kid Assignment Confirmation Form
    //20 Lesson Reports Form
    //21 Participation Agreement Form
    //22 Tutors Recruitment Form
    //23 Kid Enrollment / IHelper жобасына оқушыны қосу / Добавление ученика в проект IHelper
    //24 Withdrawal Form
    //25 Explanatory Form

    //Today's date
    , today:Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy")

    //Link to Certificate
    , certificateLink: null
  };

  return senderClassData;
}

////////////////////////
//Standard Certificate//
////////////////////////
function certificateGen_standard(){
  var template_id = '1MTdQSdX6MopbUXSAqy-gE_nRLfmRScflv6d0g-sd02E';//template id
  certificateGen(template_id);
}

////////////////////////////
//Tutors Award Certificate//
////////////////////////////
function certificateGen_tutorsAward(){
  var template_id = '';//template id
  certificateGen(template_id);
}

//////////////////////////
//Certificates Generator//
//////////////////////////
function certificateGen(template_id){
  
  //Spreadsheet navigation
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //Sheets by names
  var tutorSheet = ss.getSheetByName('Tutors Active');
  var range =  tutorSheet.getActiveRange();
  
  const PDF_folder = DriveApp.getFolderById("1J8o_MpJOFr--0RgwPI6gdq7VJIbXXbcA");
  const CERTIFICATES_FOLDER = DriveApp.getFolderById("10xKx3DCI-7dinSIEndBDspbnRk9fS7zi");
  const PDF_Template = DriveApp.getFileById(template_id);
  
  for(var i=0; i<range.getNumRows();i++){
    var newTempFile = PDF_Template.makeCopy(tutorSheet.getRange(range.getRow()+i,3).getValue(), CERTIFICATES_FOLDER);
    var OpenDoc = DocumentApp.openById(newTempFile.getId());
    var body = OpenDoc.getBody();
    console.log(body);
    body.replaceText("{Name Surname}", tutorSheet.getRange(range.getRow()+i,3).getValue());
    OpenDoc.saveAndClose();
    var BLOBPDF = newTempFile.getAs(MimeType.PDF);
    var pdfFile =  PDF_folder.createFile(BLOBPDF).setName(tutorSheet.getRange(range.getRow()+i,3).getValue());
    tutorSheet.getRange(range.getRow()+i,38).setValue(pdfFile.getUrl());
  }
  return;
}
