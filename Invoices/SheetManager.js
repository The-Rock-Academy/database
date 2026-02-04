/**
 * This class is designed to actually manage invoice sender sheet.
 * It will communicate with the databaseSheetManager. This sheet manager could be for any of the sheets.
 * 
 */
class SheetManager {
  constructor(invoiceSheet, databaseSheetManager) {
    this.sheet = invoiceSheet;
    this.ss = invoiceSheet.getParent();
    this.invoiceArchive = new InvoiceFolder(databaseSheetManager.databaseData.getVariable("Invoice Folder"))
    this.invoiceRanges = new InvoiceSheetRangeManager(this.sheet);
    this.invoiceCleaner = new InvoiceCleaner(this.sheet);
    this.databaseSheetManager = databaseSheetManager;
  }


  sendInvoice(updating = false) {
    console.log("Starting sending process");
    if (!updating && this.invoiceRanges.getInvoiceInfoRange().getValues()[0][2] == "update") {
      updating = true;
    }
    
    console.log("Checking if invoice exists")
    if (!updating && this.invoiceArchive.invoiceExists(this.invoiceRanges.getInvoiceNumberRange().getValue())) {
      let ui = SpreadsheetApp.getUi()
      let clearAnswer = ui.alert("This invoice already exists in the archive and has probably already been sent to parent. Would you really like to send this invoice?", ui.ButtonSet.YES_NO);
      if (clearAnswer == ui.Button.NO){
        console.log("Not sending invoice as user has decided against it")
        return
      } else {
        updating = true;
      }
    }
    // Remove the old invoice
    if (updating) {
      this.invoiceArchive.removeInvoice(this.invoiceRanges.getInvoiceNumberRange().getValue());
    }

    let invoicePDF = this.archiveInvoice()
    console.log("  Archived")
    this.emailInvoice(invoicePDF, updating, this.invoiceRanges.getInvoiceInfoRange().getValues()[0][1]);
    console.log("  Invoice sent")
    try{
    this.updateDatabaseSheet();
    console.log("  Updated attendance sheet");
    this.clearInvoice(true);
    console.log("  Sent, archived and cleared invoice")
    } catch (error) {
      throw new Error("Error occurred while sending invoice: " + error.message + "\nThe invoice was either not updated in database or not cleared from invoice sender. The invoice PDF has been archived. And parent was emailed.");
    }
  }

  emailInvoice(invoicePDF, updating=false, type) {

    let invoiceInfo = {
      invoiceNumber: this.invoiceRanges.getInvoiceNumberRange().getValue(),
      term: this.invoiceRanges.getTermRange().getValue(),
      parentName: this.invoiceRanges.getParentName(),
      invoicePrice: this.invoiceRanges.getTotalCostRange().getValue()
    };

    let recipient = this.invoiceRanges.getEmail();

    let templateSS = this.databaseSheetManager.databaseData.getTemplateSS();
    let template_type = updating ? "updating" : "default";

    try {
      let emailer = Emails.newEmailer(templateSS, (type+" invoice"), template_type);

      let replyTo = this.databaseSheetManager.databaseData.getReplyToEmail();
      emailer.sendEmail([recipient], invoiceInfo, [invoicePDF], replyTo);
    } catch (error) {
      // Delete the saved PDF if emailing fails
      invoicePDF.setTrashed(true);
      throw new Error("Failed to send email: " + error.message);
    }
  }

  archiveInvoice() {

    let invoiceBuilderSpreadSheet = SpreadsheetApp.openById(this.databaseSheetManager.databaseData.getVariable("Invoice Builder"));

    let invoiceBuilt = this.sheet.copyTo(invoiceBuilderSpreadSheet);
    // Add some unique timestamp to the name to avoid name conflicts
    let timestamp = new Date().toISOString().replace(/[-:.]/g, "");
    let uniqueName = "Invoice " + this.invoiceRanges.getInvoiceNumberRange().getValue() + " " + timestamp;
    invoiceBuilt.setName(uniqueName);
    invoiceBuilderSpreadSheet.getSheets().
      filter(sheet => sheet.getName() != uniqueName).
      forEach(sheet => invoiceBuilderSpreadSheet.deleteSheet(sheet));

    let invoicePDF;
    try {
      invoicePDF = DriveApp.createFile(invoiceBuilderSpreadSheet.getAs("application/PDF"));
      invoicePDF.setName(this.invoiceRanges.getInvoiceNumberRange().getValue() + ".pdf");
      this.invoiceArchive.addInvoice(invoicePDF);
      invoiceBuilt.setName(uniqueName + " (sent)");
    } catch (error) {
      // Clean up the sheet in the invoice builder if DriveApp.createFile fails
      throw new Error("Failed to create invoice PDF:\n" + error.message);
    }

    return invoicePDF;
  }

  /**
   * This will update the attendance sheet with the sent date and the total amount.
   */
  updateDatabaseSheet() {
    let totalCost = this.invoiceRanges.getTotalCostRange().getValue();

    let sentDate = new Date();

    let invoiceNumber = this.invoiceRanges.getInvoiceNumberRange().getValue();
    let numberOfLessons = this.invoiceRanges.getNumberOfLessonsRange().getValue();

    this.databaseSheetManager.updateSheetAfterInvoiceSent(invoiceNumber, totalCost,sentDate, numberOfLessons);
  }

  clearInvoice(afterSending = false) {
    let invoiceNumber = this.invoiceRanges.getInvoiceNumberRange().getValue();
    if (!afterSending) {
      this.databaseSheetManager.clearSheetAfterClearingInvoiceSender(invoiceNumber);
    }
    this.invoiceCleaner.clean();
  }

  loadInvoice(invoice) {
    let invoiceLoader = new InvoiceLoader(this.sheet, invoice);

    invoiceLoader.loadInvoice(this.databaseSheetManager.sheet.getParent());
  }

  /**
   * This will return true or false depending on if the invoice is loaded.
   */
  isInvoiceLoaded() {
    return !(this.invoiceRanges.getInvoiceNumberRange().getValue() == "invoice number")
  }
}

function newSheetManager(databaseSheetManager, invoiceSenderSheet) {
  return new SheetManager(invoiceSenderSheet, databaseSheetManager)
}

function newSheetManagerFromSender(invoiceSender) {
  let invoiceRanges = new InvoiceSheetRangeManager(invoiceSender);
  let invoiceInfo = invoiceRanges.getInvoiceInfoRange().getValues();
  let originManager = getSheetManagerForTypeForInvoicing(SpreadsheetApp.openById(invoiceInfo[0][0]), invoiceInfo[0][1], invoiceRanges.getTermRange().getValue());
  return newSheetManager(originManager, invoiceSender)
}