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

    // ------ General sheet ----------
  
    clean() {
      throw new Error("clean has not been implemented yet");
    }
  
    reset() {
      throw new Error("reset has not been implemented yet");
    }

    // ------ Invoices ---------
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
  }