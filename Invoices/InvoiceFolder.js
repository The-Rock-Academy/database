class InvoiceFolder {
  constructor(invoiceFolderID) {
    this.folder = DriveApp.getFolderById(invoiceFolderID);
  }

  getNumberOfInvoices() {
    let files = this.folder.getFilesByType("application/PDF");
    let numberOfFiles = 0;

    while(files.hasNext()) {
      numberOfFiles++;
      files.next();
    }

    return numberOfFiles
  }

  invoiceExists(invoiceNumber) {
    let invoice;
    if (this.folder.getFilesByName(invoiceNumber + ".pdf").hasNext()) {
      invoice = this.folder.getFilesByName(invoiceNumber + ".pdf");
    } else {
      invoice =  this.folder.getFilesByName(invoiceNumber);
    }

    if(invoice.hasNext()) {
      return true
    } else {
      return false
    }
  }

  getInvoice(invoiceNumber) {
    if (this.folder.getFilesByName(invoiceNumber + ".pdf").hasNext()) {
      return this.folder.getFilesByName(invoiceNumber + ".pdf").next();
    } else {
      return this.folder.getFilesByName(invoiceNumber).next();
    }
  }

  addInvoice(file) {
    file.moveTo(this.folder)
  }

  /**
   * This is only to be used when you are updating an invoice
   */
  removeInvoice(invoiceNumber) {
    let invoice;
    if (this.folder.getFilesByName(invoiceNumber + ".pdf").hasNext()) {
      invoice = this.folder.getFilesByName(invoiceNumber + ".pdf").next();
    } else {
      invoice = this.folder.getFilesByName(invoiceNumber).next();
    }

    //Check owner
    let owner = invoice.getOwner().getEmail();
    let user = Session.getActiveUser().getEmail();
    if (owner == user) {
      invoice.setTrashed(true);
    } else {
      throw new Error("The invoice you are trying to update is owned by " + owner + " and you are " + user + ". This needs to be done by the owner.");
    }
  }

}

function newInvoiceFolder(invoiceFolderID) {
  return new InvoiceFolder(invoiceFolderID);
}
