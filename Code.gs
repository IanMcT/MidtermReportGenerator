/*Midterm generator
By Ian McTavish
Goal: create folders for every teacher and a midterm doc for every student in their class.
Generate from Google Sheet set up as follows:
Column A Teacher
Column B Section code (i.e. ICS3U)
Column C Student
Column D Blank
Column E Period

Note - you need a Google doc as a template.  The id (it is in the URL then needs to be set (look in the variables section))
Create this script in a sheets file (put in a folder called progress reports)
Issues - with large school and large template execution time can take too long.  
There are strategies to deal with this (i.e. delete students that have been done, run again). Due to time constraints an automated approach wasn't attempted.
Improvement ideas: include teacher email in the spreadsheet so you can auotomate permissions
Include student email in the spreadsheet so you can create a program to automatically email the file to the students at a certain date.
*/
function getStudents() {
  //Variables
  var templateID = "ReplaceWithTheTemplateID";//You need to replace this!!!
  var sheet = SpreadsheetApp.getActiveSheet();
  var currentDirectoryID = DriveApp.getFileById(sheet.getParent().getId()).getParents().next().getId();
  var currentDirectory = DriveApp.getFolderById(currentDirectoryID);
  var data = sheet.getDataRange().getValues();
  var currentTeacher = "";
  var previousTeacher = "-1";
  var currentStudent;
  var folder;
  Logger.log(data.length);
  //start at 1 to eliminate header row
  for (var i = 1; i < data.length; i++) {
    //set teacher and student
    currentTeacher = data[i][0];
    currentStudent = data[i][2];// + " " + data[i][3];
    if(currentTeacher == previousTeacher)
    {
      //same teacher
    }else{
      //Create a folder for the teacher
      Logger.log("New teacher: " + currentTeacher);
      folder = currentDirectory.createFolder(currentTeacher);
    }
    
    //Create the progress report by copying the template
    var newDocId = DriveApp.getFileById(templateID).makeCopy(folder).getId();
    DriveApp.getFileById(newDocId).setName(currentStudent);

  //replace text
  var doc = DocumentApp.openById(newDocId);
  var body = doc.getBody();
  //Your template needs text as follows to replace the information from the spreadsheet.
  body.replaceText('«Last_Name», «First_Name»',data[i][2]);
  body.replaceText('«Course_Name»',data[i][1]);
  body.replaceText('«Period»',data[i][4]);
  body.replaceText('«Teacher»',currentTeacher);

    previousTeacher = currentTeacher;

  }//end loop
}//end function
