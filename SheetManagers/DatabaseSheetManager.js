/**
 * This is the class that each database sheet will extend.
 */
class DatabaseSheetManager {
    constructor(sheet, currentTerm) {
      if (currentTerm == undefined) {
        currentTerm = newDatabase(sheet.getParent()).getDatabaseTerm();
      }
      this.sheet = sheet;
      this.currentTerm = currentTerm;
      this.databaseData = newDatabaseData(this.sheet.getParent());
    }

    // ------------------------- General sheet -------------------------
  
    clean() {
      throw new Error("clean has not been implemented yet");
    }
    /**
     * Reset the sheet from term to term
     * @param {FormData} nextTermDetails This is the form which is returned from the reset database popup.
     * @param {string} nextTerm Next term in the string format
     */
    reset(nextTermDetails, nextTerm) {
      throw new Error("reset has not been implemented yet");
    }

    // ------------------------- Invoices -------------------------
    /**
     * Update sheet after invoice sent.
     * @param {integer} invoiceNumber 
     * @param {float} totalCost 
     * @param {Date} sentDate 
     */
    updateSheetAfterInvoiceSent(invoiceNumber, totalCost, sentDate) {
      throw new Error("updateSheetAfterInvoiceSent has not bee implemented yet");
    }

    /**
     * Clear the invoice number from the sheet after a user has cleared the invoice sender.
     * @param {number} invoiceNumber 
     */
    clearInvoiceNumber(invoiceNumber) {
      throw new Error("clearInvoiceNumber has not been implemented yet");
    }

  /**
   * This will take a column name and find the column number. This is used so that I dont have to worry about having correct column numbers.
   * If column names are on different rows than I can easily change it in one place here.
   */
    /**
     * Searches for the column number of the column name given
     * @param {string} columnName What is the column name you are looking for
     * @returns Return the column int or throws and error.
     */
    getColumn(columnName) {
      let searchResult = this.sheet.getRange(1,1, 1, this.sheet.getMaxColumns()).createTextFinder(columnName).findNext();
      if (searchResult == null) {
        throw "Cant find the column named: " + columnName + " in " + this.sheet.getName() + " sheet";
      } else {
        return searchResult.getColumn()
      }
    }

    resetInvoiceColumns(nextTerm) {
      let currentInvoiceColumn = this.getColumn("Current Invoice " + this.currentTerm);
      let nextInvoiceColumn = currentInvoiceColumn + 4;
  
      //Rename current to previous
      this.sheet.getRange(1, currentInvoiceColumn).setValue("Previous Invoice " + this.currentTerm);
      //Add in new column
      this.sheet.insertColumns(nextInvoiceColumn, 4);
      //Copy everything across
      this.sheet.getRange(1, currentInvoiceColumn, this.sheet.getMaxRows(), 4).copyTo(this.sheet.getRange(1, nextInvoiceColumn, this.sheet.getMaxRows(), 4))
      //Set next term name
      this.sheet.getRange(1,nextInvoiceColumn).setValue("Current Invoice " + nextTerm)
  
      //Clear old contents and comments
      this.sheet.getRange(3, nextInvoiceColumn, this.sheet.getMaxRows(), 4).clear({contentsOnly: true});
      this.sheet.getRange(1, nextInvoiceColumn, this.sheet.getMaxRows(), 4).clear({commentsOnly: true});
  
      //Delete old column
      this.sheet.deleteColumns(currentInvoiceColumn-4, 4);
    }

    resetCommentColumns() {
      let currentCommentsColumn = this.getColumn("Current term comments");
      let currentCommentsTitle = this.sheet.getRange(1, currentCommentsColumn);

      //Add new coments column
      this.sheet.insertColumnsAfter(currentCommentsColumn,1);
      let nextCommentsTitle = this.sheet.getRange(1, currentCommentsColumn+1);

      //Rename current and next
      currentCommentsTitle.setValue("Previous term comments")
      nextCommentsTitle.setValue("Current term comments")

      //Copy formatting over
      currentCommentsTitle.copyTo(nextCommentsTitle, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      currentCommentsTitle.copyTo(nextCommentsTitle, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false)

      //Delete previous comments
      this.sheet.deleteColumns(currentCommentsColumn-1);
    }
  }