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

  /**
   * Will clean the sheet.
   * It will do things like move things to the inactive sections etc.
   */
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

  /**
   * Get invoice information from the sheet and load it into the invoice sender
   * @param {string} currentTerm The string representation of the current term
   * @param {number} row The row of which the invoice is from in the database sheet
   * @param {bool} send Whether to send the invoice after preparing
   */
  prepareInvoice(currentTerm, row, send = false) {
    throw new Error("prepareInvoice has not been implemented yet");
  }

  /**
   * This will loop through each row of the users selected range and send the invoice.
   * @param {Range} range User selected range
   */
  prepareAndSendInvoice(range) {
    for (let row = range.getRowIndex(); row < range.getHeight() + range.getRowIndex(); row++) {
      this.prepareInvoice(row, true);
    }
  }

  // -------------------------//------------------------ Invoices -------------------------//------------------------
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

  clearSheetAfterClearingInvoiceSender(invoiceNumber) {
    this.clearInvoiceNumber(invoiceNumber);
  }

  /**
   * Searches for the column number of the column name given
   * @param {string} columnName What is the column name you are looking for
   * @returns Return the column int or throws and error.
   */
  getColumn(columnName, contains = false) {
    let searchResult = this.sheet.getRange(1,1, 1, this.sheet.getMaxColumns()).createTextFinder(columnName).matchEntireCell(contains).findNext();
    if (searchResult == null) {
      throw "Cant find the column named: " + columnName + " in " + this.sheet.getName() + " sheet";
    } else {
      return searchResult.getColumn()
    }
  }

  /**
   * Get the invoice range for the current term
   * @param {number} invoiceNumber 
   * @returns 
   */
  getInvoiceRanges(invoiceNumber, previousTerm = false) {
    console.log("Current invoice column: " + this.currentInvoiceColumn);
    let invoiceRow = this.getInvoiceRow(invoiceNumber,  this.currentInvoiceColumn);
    console.log("Row of invoice number " + invoiceNumber + " is " + invoiceRow);
    if (invoiceRow == -1) {
      return null
    }
    if (previousTerm) {
      return {
        number: this.sheet.getRange(invoiceRow,  this.previousInvoiceColumn),
        amount: this.sheet.getRange(invoiceRow,  this.previousInvoiceColumn + 1),
        date: this.sheet.getRange(invoiceRow,  this.previousInvoiceColumn + 2),
        paidDate: this.sheet.getRange(invoiceRow,  this.previousInvoiceColumn + 3)
      }
    } else {
      return {
        number: this.sheet.getRange(invoiceRow,  this.currentInvoiceColumn),
        amount: this.sheet.getRange(invoiceRow,  this.currentInvoiceColumn + 1),
        date: this.sheet.getRange(invoiceRow,  this.currentInvoiceColumn + 2),
        paidDate: this.sheet.getRange(invoiceRow,  this.currentInvoiceColumn + 3)
      }
    }

  }

    /**
   * Get the pupil row for a particualr invoice number
   * @param {*} invoiceNumber Invoice number in question
   * @param {*} currentInvoiceColumn The current term invoice column in sheet
   * @returns The number of the row or -1 if it hasnt found a row
   */
  getInvoiceRow(invoiceNumber, currentInvoiceColumn = this.currentInvoiceColumn) {
    let foundRowCurrent = this.sheet.getRange(3, currentInvoiceColumn, this.sheet.getMaxRows(), 1).createTextFinder(invoiceNumber).matchEntireCell(true).findNext();
    if (foundRow != null) return foundRow.getRowIndex()
    else  {
      console.warn("Could not find a row for invoice number: " + invoiceNumber);
      return -1
    }
  }

  /**
   * Retrieves the invoice number of a row
   * @param {number} rowNumber Row number you are checking
   * @returns 
   */
  getInvoiceNumberOfRow(rowNumber, previousInvoiceColumn=undefined) {
    console.log("searching for row invoice: " + rowNumber )
    console.log("Current invoice column: " + this.currentInvoiceColumn)
    console.log(this.sheet.getRange(rowNumber,previousInvoiceColumn == undefined?this.currentInvoiceColumn:previousInvoiceColumn).getValue())
    return this.sheet.getRange(rowNumber, previousInvoiceColumn == undefined?this.currentInvoiceColumn:previousInvoiceColumn).getValue()
  }
  // -------------------------//------------------------ Resetting -------------------------//-------------------------//

  resetInvoiceColumns(nextTerm,nextTermInvoiceColumn=undefined) {
    this.currentInvoiceColumn = this.getColumn("Current Invoice " + this.currentTerm);

    let nextInvoiceColumn = nextTermInvoiceColumn == undefined ? this.currentInvoiceColumn + 4: nextTermInvoiceColumn;

    //Rename current to previous
    this.sheet.getRange(1, this.currentInvoiceColumn).setValue("Previous Invoice " + this.currentTerm);
    //Add in new column
    this.sheet.insertColumns(nextInvoiceColumn, 4);
    this.sheet.setColumnWidths(nextInvoiceColumn, 4, 80);
    //Copy everything across
    this.sheet.getRange(1, this.currentInvoiceColumn, this.sheet.getMaxRows(), 4).copyTo(this.sheet.getRange(1, nextInvoiceColumn, this.sheet.getMaxRows(), 4))
    //Set next term name
    this.sheet.getRange(1,nextInvoiceColumn).setValue("Current Invoice " + nextTerm)

    //Clear old contents and comments
    this.sheet.getRange(3, nextInvoiceColumn, this.sheet.getMaxRows(), 4).clear({contentsOnly: true});
    this.sheet.getRange(1, nextInvoiceColumn, this.sheet.getMaxRows(), 4).clear({commentsOnly: true});

    //Delete old column
    if (nextTermInvoiceColumn == undefined) {
      this.sheet.deleteColumns(this.currentInvoiceColumn, 4);
    }
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