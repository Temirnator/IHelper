/*function onOpen() { 
  var ui = SpreadsheetApp.getUi(); 
  ui.createMenu('Show Sheet ID').addItem('Show Active Sheet ID','getSheetbyId' ).addToUi(); 
} 
*/

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Get Kid Information").addItem("Update Kid Info","getKidInfo").addToUi();
}

function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for(var row = 0; row < range.length; row++){

    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;

    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
};

function getKidInfo(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Kid Information");
  var source_ss = SpreadsheetApp.openById("1AaMNmhopqblI0o87lCO1DEbHFg_6bH81vtmJH4Boj8Y");
  var source_sheet = source_ss.getSheetByName("Kids");
  var kidFolderName = DriveApp.getFileById(ss.getId()).getParents().next().getName();
  var lastrow = getLastRowSpecial(source_sheet.getRange("A2:A").getValues());
  for(var i=1;i<=lastrow;i++){
    if(kidFolderName==source_sheet.getRange(i+1,1).getValue()){
      for(var j=2; j<=12;j++){
        sheet.getRange(j+1,2).setValue(source_sheet.getRange(i+1,j-1).getValue());
      }
    }
  }

}

function getSheetbyId(){ 
  var sheetID = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId(); 
  SpreadsheetApp.getUi().alert(sheetID); 
} 

function onEdit(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user_email = Session.getActiveUser().getEmail();

  var kidInfoSheet = getSheetById(0);
  var lessonReportsSheet = getSheetById(150657855);
  var kidPerformanceSheet = getSheetById(1746314528);
  var lessonPlanSheet = getSheetById(715668535);
  var teacherDraftSheet = getSheetById(1420972940);
  var bookOfComSugSheet = getSheetById(1959866622);

  //0 - Kid Information
  //150657855 - Lesson Reports
  //1746314528 - Kid Performance Reports
  //715668535 - Lesson Plan
  //1420972940 - Teacher's Draft Diary
  //1959866622 - Book of Complaints and Suggestions

  var lessonReportfirstEntryRow = 4;
  var kidPerformancefirstEntryRow = 4;
  var lessonPlanfirstEntryRow = 3;
  var teacherDraftfirstEntryRow = 3;
  var bookOfComSugfirstEntryRow = 3;


  if ((lessonReportsSheet.getActiveCell().getRow() >= lessonReportfirstEntryRow) && (lessonReportsSheet.getActiveCell().getRow() <= lessonReportsSheet.getLastRow())&&(lessonReportsSheet.getActiveCell().getColumn() <= lessonReportsSheet.getLastColumn()-2)&&((lessonReportsSheet.getActiveCell().getRow())%3!=(lessonReportfirstEntryRow%3-1))) {
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(), 10).setValue(user_email);
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(), 11).setValue(new Date);
  }
  if((lessonReportsSheet.getActiveCell().getRow() >= lessonReportfirstEntryRow)&&(lessonReportsSheet.getActiveCell().getColumn() == 1)&&(lessonReportsSheet.getActiveCell().getValue()== false)&&(lessonReportsSheet.getActiveCell().getRow()!=4)&&(lessonReportsSheet.getActiveCell().getRow()!=5)&&(lessonReportsSheet.getActiveCell().getRow()!=19)&&(lessonReportsSheet.getActiveCell().getRow()!=20)&&(lessonReportsSheet.getActiveCell().getRow()!=34)&&(lessonReportsSheet.getActiveCell().getRow()!=35)){
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),4).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),5).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),6).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),7).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),8).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),9).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),10).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),11).clearContent();
  }else if((lessonReportsSheet.getActiveCell().getRow() >= lessonReportfirstEntryRow)&&(lessonReportsSheet.getActiveCell().getColumn() == 1)&&(lessonReportsSheet.getActiveCell().getValue()== false)){
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),4).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),5).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),7).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),8).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),9).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),10).clearContent();
    lessonReportsSheet.getRange(lessonReportsSheet.getActiveCell().getRow(),11).clearContent();
}

  if ((kidPerformanceSheet.getActiveCell().getRow() >= kidPerformancefirstEntryRow) && (kidPerformanceSheet.getActiveCell().getRow() <= kidPerformanceSheet.getLastRow())&&(kidPerformanceSheet.getActiveCell().getColumn() <= kidPerformanceSheet.getLastColumn()-2)&&((kidPerformanceSheet.getActiveCell().getRow())%3!=(kidPerformancefirstEntryRow%3-1))) {
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(), 9 ).setValue(user_email);
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(), 10).setValue(new Date);
  }

  if((kidPerformanceSheet.getActiveCell().getRow() >= kidPerformancefirstEntryRow)&&(kidPerformanceSheet.getActiveCell().getColumn() == 1)&&(kidPerformanceSheet.getActiveCell().getValue()== false)&&((kidPerformanceSheet.getActiveCell().getRow())%3!=(kidPerformancefirstEntryRow%3-1))){
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),4).clearContent();
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),5).clearContent();
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),6).clearContent();
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),7).clearContent();
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),8).clearContent();
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),9).clearContent();
    kidPerformanceSheet.getRange(kidPerformanceSheet.getActiveCell().getRow(),10).clearContent();
  } 
    
    if ((lessonPlanSheet.getActiveCell().getRow() >= lessonPlanfirstEntryRow)&&(lessonPlanSheet.getActiveCell().getColumn() == 1)&&(lessonPlanSheet.getActiveCell().getValue()== true)){
      lessonPlanSheet.getRange(lessonPlanSheet.getActiveCell().getRow(),11).setValue(user_email);
      lessonPlanSheet.getRange(lessonPlanSheet.getActiveCell().getRow(),12).setValue(new Date);
    }else if((lessonPlanSheet.getActiveCell().getRow() >= lessonPlanfirstEntryRow)&&(lessonPlanSheet.getActiveCell().getColumn() == 1)&&(lessonPlanSheet.getActiveCell().getValue()== false)){
      lessonPlanSheet.getRange(lessonPlanSheet.getActiveCell().getRow(),11).clearContent();
      lessonPlanSheet.getRange(lessonPlanSheet.getActiveCell().getRow(),12).clearContent();
    }

    if ((teacherDraftSheet.getActiveCell().getRow() >= teacherDraftfirstEntryRow)&&(teacherDraftSheet.getActiveCell().getColumn() == 1)&&(!teacherDraftSheet.getActiveCell().isBlank())){
      teacherDraftSheet.getRange(teacherDraftSheet.getActiveCell().getRow(),2).setValue(user_email);
      teacherDraftSheet.getRange(teacherDraftSheet.getActiveCell().getRow(),3).setValue(new Date);
    }else if((teacherDraftSheet.getActiveCell().getRow() >= 2)&&(teacherDraftSheet.getActiveCell().getColumn() == 1)&&(teacherDraftSheet.getActiveCell().isBlank())){
      teacherDraftSheet.getRange(teacherDraftSheet.getActiveCell().getRow(),2).clearContent();
      teacherDraftSheet.getRange(teacherDraftSheet.getActiveCell().getRow(),3).clearContent();
    }

    
    if ((bookOfComSugSheet.getActiveCell().getRow() >= bookOfComSugfirstEntryRow)&&(bookOfComSugSheet.getActiveCell().getColumn() == 1)&&(!bookOfComSugSheet.getActiveCell().isBlank())){
      bookOfComSugSheet.getRange(bookOfComSugSheet.getActiveCell().getRow(),2).setValue(user_email);
      bookOfComSugSheet.getRange(bookOfComSugSheet.getActiveCell().getRow(),3).setValue(new Date);
    }else if((bookOfComSugSheet.getActiveCell().getRow() >= bookOfComSugfirstEntryRow)&&(bookOfComSugSheet.getActiveCell().getColumn() == 1)&&(bookOfComSugSheet.getActiveCell().isBlank())){
      bookOfComSugSheet.getRange(bookOfComSugSheet.getActiveCell().getRow(),2).clearContent();
      bookOfComSugSheet.getRange(bookOfComSugSheet.getActiveCell().getRow(),3).clearContent();
    }
    
}

/**/

/* functino notifyCoordinators(){

}
*/

/* function seeContactInformation(){
  
}
 */

function getSheetById(id) { 
  return SpreadsheetApp.getActive().getSheets().filter( 
    function(s) {return s.getSheetId() === id;} 
  )[0]; 
}
