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
    debug("Starting sending process");
    if (!updating && this.invoiceRanges.getInvoiceInfoRange().getValues()[0][2] == "update") {
      updating = true;
    }
    
    debug("Checking if invoice exists")
    if (!updating && this.invoiceArchive.invoiceExists(this.invoiceRanges.getInvoiceNumberRange().getValue())) {
      let ui = SpreadsheetApp.getUi()
      let clearAnswer = ui.alert("This invoice already exists in the archive and has probably already been sent to parent. Would you really like to send this invoice?", ui.ButtonSet.YES_NO);
      if (clearAnswer == ui.Button.NO){
        debug("Not sending invoice as user has decided against it")
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
    debug("Archived")
    this.emailInvoice(invoicePDF, updating, this.invoiceRanges.getInvoiceInfoRange().getValues()[0][1]);
    debug("Invoice sent")
    this.updateDatabaseSheet();
    debug("Updated attendance sheet");
    this.clearInvoice(true);
    debug("Sent, archived and cleared invoice")
  }

  emailInvoice(invoicePDF, updating=false, type) {

    let invoiceInfo = {
      invoiceNumber: this.invoiceRanges.getInvoiceNumberRange().getValue(),
      term: this.invoiceRanges.getTermRange().getValue(),
      parentName: this.invoiceRanges.getParentName(),
      invoicePrice: this.invoiceRanges.getTotalCostRange().getValue()
    }

    let recipient = this.invoiceRanges.getEmail();

    let templateSS = this.databaseSheetManager.databaseData.getTemplateSS();
    let template_type = updating ? "updating" : "default";
    let emailer = Emails.newEmailer(templateSS, (type+" invoice"), template_type);

    emailer.sendEmail([recipient], invoiceInfo, [invoicePDF])

  }

  archiveInvoice() {
    this.ss.getSheets().filter(sheet => sheet.getName() != "Invoice").forEach(sheet => sheet.hideSheet());

    let invoicePDF = DriveApp.createFile(this.ss.getAs("application/PDF"))
    invoicePDF.setName(this.invoiceRanges.getInvoiceNumberRange().getValue() + ".pdf")
    this.invoiceArchive.addInvoice(invoicePDF)
    return invoicePDF
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
  let invoiceInfo = new InvoiceSheetRangeManager(invoiceSender).getInvoiceInfoRange().getValues();
  let originManager = getSheetManagerForType(SpreadsheetApp.openById(invoiceInfo[0][0]), invoiceInfo[0][1]);
  return newSheetManager(originManager, invoiceSender)
}