class Database {
  constructor(ss) {
    this.ss = ss
    this.currentTerm = this.getDatabaseTerm()
    this.databaseData = newDatabaseData(ss);
    this.archiveFolder = DriveApp.getFolderById(this.databaseData.getVariable("Archive Folder"));
  }

  static createSpreadSheetCopy(ss, folder, archiveName) {
    let previousArchive = folder.getFilesByName(archiveName);

    if (previousArchive.hasNext()) {
      let currentArchive = previousArchive.next();
      let ui = SpreadsheetApp.getUi()
      let questionResponse = ui.alert("There is already an archive for " + archiveName + " that was made on " + currentArchive.getLastUpdated().toLocaleString() + "\nWould you like to overwrite it?", ui.ButtonSet.YES_NO)
      if (questionResponse == ui.Button.NO) {
        return
      } else {
        currentArchive.setTrashed(true);
      }
    }

    let databaseFile = DriveApp.getFileById(ss.getId())

    let oldDatabaseCopy = databaseFile.makeCopy(ss.getName(), folder);
    oldDatabaseCopy.setName(archiveName);

    return oldDatabaseCopy

  }

  /**
   * Very simple function simply makes a copy of the database and moves it to the archive folder.
   * It will ask a user to overwrite the previous archive if there is already an archive.
   */
  archive() {
    let copied = Database.createSpreadSheetCopy(this.ss, this.archiveFolder, this.currentTerm + " Database");
    let copied_ss = SpreadsheetApp.openById(copied.getId());
    // Remove duplicated form
    SHPBookingsNewFromSS(copied_ss).deleteAttachedForm();
    newDatabaseData(copied_ss).setVariable("Invoice Sender", copied.getId())

    return copied_ss;
  }

  /**
   * This is for grabbing the spreadsheets current term it is set too
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
   * This should received the form that was submitted iwth the next term dates.
   */
  reset(form, old_database) {
    

    //Set the term number on the Home sheet
    let previousTerm = this.getDatabaseTerm();
    let nextTerm = "Term " + form.termNum + " " + form.termYear
    this.setDatabaseTerm(nextTerm);

    let termStartDate = new Date(form.termDate)

    this.ss.rename(nextTerm + " Database")


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
    SHPManager.newFromSS(this.ss, null, 2).reset(nextTermDates);

    //Call Band School sheet resetter
    let bandSchoolManager = newBandSchoolSheet(this.databaseData.getBandSchoolSheet());
    let old_band_school = bandSchoolManager.archive(previousTerm, old_database)
    bandSchoolManager.reset(nextTermDates, nextTerm);

    let bandSchoolInvoicingManager = newBandSchoolInvoicingSheet(this.databaseData.getBandSchoolInvoicingSheet());
    let old_band_school_invoice = bandSchoolInvoicingManager.archive(previousTerm, old_band_school);
    bandSchoolInvoicingManager.reset(nextTermDates, nextTerm);

    //Update the data sheet
    let databaseData = newDatabaseData(old_database);
    databaseData.setVariable("Main Database SS", old_database.getUrl());
    databaseData.setVariable("Band School ID", old_band_school.getUrl());
    databaseData.setVariable("Band School Invoicing URL", old_band_school_invoice.getUrl());


    this.ss.toast("The resetting to term " + nextTerm + " is completed.")
  }
}

function newDatabase(ss) {
  return new Database(ss)
}

function getSheetManagerForTypeForInvoicing(ss, type, term) {
  switch(type) {
    case "term":
      return newAttendanceSheet(ss.getSheetByName("Master Sheet"));
    case "shp":
      // If term has jan or Jan in it then it is the second week
      if (term.toLowerCase().includes("jan")) {
        return SHPManager.newFromSS(ss, undefined, 2);
      } else {
        return SHPManager.newFromSS(ss, undefined);
      }
    case "band":
      return BandSchoolInvoicingManager.newFromSS(ss);
    default:
      throw new Error("You are trying to get a sheet for type: '" + type + "' which does not exist.");
  }
}