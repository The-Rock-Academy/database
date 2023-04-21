class Database {
  constructor(ss) {
    this.ss = ss
    this.currentTerm = this.getDatabaseTerm()
    this.databaseData = newDatabaseData(ss);
    this.archiveFolder = DriveApp.getFolderById(this.databaseData.getVariable("Archive Folder"));
  }

  /**
   * Very simple function simply makes a copy of the database and moves it to the archive folder.
   * It will ask a user to overwrite the previous archive if there is already an archive.
   */
  archive() {
    let archiveName = "Archive " + this.currentTerm
    
    //Check for previous archive
    let previousArchive = this.archiveFolder.getFilesByName(archiveName);

    if (previousArchive.hasNext()) {
      let currentArchive = previousArchive.next();
      let ui = SpreadsheetApp.getUi()
      let questionResponse = ui.alert("There is already an archive for this database that was made on " + currentArchive.getLastUpdated().toLocaleString() + "\nWould you like to overwrite it?", ui.ButtonSet.YES_NO)
      if (questionResponse == ui.Button.NO) {
        return
      } else {
        currentArchive.setTrashed(true);
      }
    }
    DriveApp.getFileById(this.ss.getId()).makeCopy(archiveName, this.archiveFolder)
  }

  /**
   * This is for grbbing the spreadsheets current term it is set too
   * Used by the various attendance sheets
   */
  getDatabaseTerm() {
    return this.ss.getSheetByName("Data").getRange(1,2).getValue();
  }

  setDatabaseTerm(newValue){
    this.ss.getSheetByName("Data").getRange(1,2).setValue(newValue);
  }

  /**
   * Will go from the current term and find out the next term.
   */
  getNextDatabaseTerm() {
    let currentTermSplit = this.currentTerm.split(" ")

    let nextTerm = "Term " + ((parseInt(currentTermSplit[1]) + 1) < 5 ? (parseInt(currentTermSplit[1]) + 1) + " " + currentTermSplit[2] : 1 + " " + (parseInt(currentTermSplit[2]) + 1));

    return nextTerm;
  }

  /**
   * Ask the user for the term dates.
   * However it will prepopoulate the form with what we already know and guess.
   * It will get the start date, number of weeks and the term number and year.
   * 
   * After it is submitted it will carry on the reseting procedure
   */
  getTermDates() {
    let htmlForm = ' <div> <div id="form"> <h1>Please enter details about new term</h1> <form id="feedbackForm"> <label for="termNum">Term number</label> <input type="number" id="termNum" name="termNum" value="termNumDefault"><br><br> <label for="termYear">Term year</label> <input type="number" id="termYear" name="termYear", value="termYearDefault"><br><br> <label for="termDate">Term date start (Should be the first monday)</label> <input type="date" id="termDate" name="termDate"><br><br> <label for="termWeeks">Term week number</label> <input type="number" id="termWeeks" name="termWeeks", value=10><br><br> <div> <input type="button" value="Submit" onclick="functionToBeAdded"> </form> </div> </div> <div id="thanks" style="display: none;"> <p>Thank you, Your new term dates are being set.</p> </div>'

    let nextTerm = this.getNextDatabaseTerm().split(" ");
    //Now I will go through and replace the default values in the form
    htmlForm = htmlForm.replace("termNumDefault", nextTerm[1]).replace("termYearDefault", nextTerm[2]);
      //Here I add in the function. I have to add it like this as there is problems with escaping and "" ' problems. However this seems to work.
    htmlForm = htmlForm.replace("functionToBeAdded", "function submitForm(){google.script.run.handleTermDatesFormSubmit(document.getElementById('feedbackForm')); document.getElementById('form').style.display = 'none'; document.getElementById('thanks').style.display = 'block';};submitForm();")


    var htmlOutput = HtmlService.createHtmlOutput(htmlForm).setHeight(400).setWidth(400)
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "New term dates");
  }

  /**
   * This should recieved the form that was submitted iwth the next term dates.
   */
  reset(form) {
    this.archive();

    //Set the term number on the Home sheet
    let previousTerm = this.getDatabaseTerm();
    let nextTerm = "Term " + form.termNum + " " + form.termYear
    this.setDatabaseTerm(nextTerm);

    let termStartDate = new Date(form.termDate)


    let nextTermDates = [...Array(parseInt(form.termWeeks)).keys()].map(weekNumber => {
      let result = new Date(termStartDate);
      result.setDate(result.getDate() + (7 * (weekNumber)))
      return result
      }
    )

    //Call the attendance sheet resetter
    let attendanceSheet = this.ss.getSheetByName(AttendanceManager.sheetName());
    let attendanceManager = new AttendanceManager(attendanceSheet, previousTerm)
    attendanceManager.reset(nextTermDates, nextTerm);

    //Call SHP sheet resetter
    SHPManager.newFromSS(this.ss).reset(nextTermDates);


    this.ss.toast("The resetting to term " + nextTerm + " is completed.")
  }
}
function handleTermDatesFormSubmit(form) {
  let database = new Database(SpreadsheetApp.getActiveSpreadsheet());
  database.reset(form)
}

function newDatabase(ss) {
  return new Database(ss)
}

function getSheetManagerForType(ss, type) {
  switch(type) {
    case "term":
      return newAttendanceSheet(ss.getSheetByName("Master Sheet"));
    case "shp":
      return SHPManager.newFromSS(ss, undefined);
    case "band":
      return BandSchoolManager.newFromSS(ss);
    default:
      throw new Error("You are trying to get a sheet for type: '" + type + "' which does not exist.");
  }
}