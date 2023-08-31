/**
 * This class is designed to manage the data sheet which is found in the Data tab in the database.
 * 
 * This is where all sorts of config setting can be made.
 */

class DatabaseData {
  constructor(database) {
    this.sheet = database.getSheetByName("Data");
  }

  getVariable(variableName) {
    let textFinderResults = this.sheet.getRange(1,1, this.sheet.getMaxRows(), 2).createTextFinder(variableName).matchEntireCell(true).findAll();

    if (textFinderResults.length == 0) {
      throw new Error("Tried to find '" + variableName + "' in the data sheet in the Database. It could not be found. Make sure that it exists or that there isnt a typo somehwere.");
    } else if (textFinderResults.length > 1) {
      this.sheet.getParent().toast("More than one occurance of '" + variableName + "' was found in the data sheet. Make sure there is only one as having multiple could cause errors")
    } else {
      return this.sheet.getRange(textFinderResults[0].getRow(), 2).getValue();
    }
  }

  //Various things that are designed to help do common tasks in the database
  getTemplateSS() {
    return SpreadsheetApp.openByUrl(this.getVariable("Email templates SS"));
  }

  getBandSchoolSheet() {
    return SpreadsheetApp.openByUrl(this.getVariable("Band School ID")).getSheetByName(getBandSchoolSheetName());
  }
}

function newDatabaseData(database) {
  return new DatabaseData(database);
}