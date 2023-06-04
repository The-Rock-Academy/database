/**
 * These classes here are designed to be the connections from the Invoice to the sheet.
 * This means that the Invoice doesnt need to know anything about the Sheet and vice versa.
 * 
 * There is the super classs that is effectivly just a dictonary of ranges.
 * Then the two sub classes are a cleaner and a loader.
 * One of these takes all of the ranges and clears them filling some with place holders.
 * The other one takes an invoice and loads them into all the needed ranges.
 */

class InvoiceSheetRangeManager {
  constructor(sheet) {
    this.sheet = sheet;
  }

  getParentRange() {
    return this.sheet.getParent().getRangeByName("parents");
  }

  getEmail() {
    return this.getParentRange().getValues()[1][0]
  }

  getParentName() {
    return this.getParentRange().getValues()[0][0]
  }

  getLogoRange() {
    return this.sheet.getParent().getRangeByName("logo")
  }

  getAddressRange() {
    return this.sheet.getParent().getRangeByName("address")
  }

  getPaymentRange() {
    return this.sheet.getParent().getRangeByName("payment")
  }

  getCostRange(itemNum, numberOfItems) {
    let row = this.sheet.getParent().getRangeByName("firstCost").getRow()
    return this.sheet.getRange(row - 1 + itemNum, 2,numberOfItems,5)
  }

  getCostsRange() {
    return this.getCostRange(1,3)
  }

  getNumberOfLessonsRange() {
    let row = this.sheet.getParent().getRangeByName("firstCost").getRow();
    return this.sheet.getRange(row, 5);
  }

  getInvoiceNumberRange() {
    return this.sheet.getParent().getRangeByName("invoiceNumber")
  }

  getTermRange() {
    return this.sheet.getParent().getRangeByName("term")
  }

  getNotesRange() {
    return this.sheet.getParent().getRangeByName("notes")
  }

  getDivisionMessageLocation() {
    return this.sheet.getParent().getRangeByName("divisionMessage")
  }

  getTotalCostRange() {
    return this.sheet.getParent().getRangeByName("totalCost")
  }

  getInvoiceInfoRange() {
    return this.sheet.getParent().getRangeByName("invoiceInfo")
  }

}

class InvoiceCleaner extends InvoiceSheetRangeManager {
  constructor(sheet) {
    super(sheet);
  }

  clean() {
    //Clear all old things
    super.getParentRange().clear({contentsOnly: true})
    // super.getLogoRange().clear({contentsOnly: true})
    super.getAddressRange().clear({contentsOnly: true})
    super.getPaymentRange().clear({contentsOnly: true})
    super.getCostsRange().clear({contentsOnly: true})
    super.getInvoiceNumberRange().clear({contentsOnly: true})
    super.getTermRange().clear({contentsOnly: true})
    super.getNotesRange().clear({contentsOnly: true})
    super.getDivisionMessageLocation().clear({contentsOnly: true})
    super.getInvoiceInfoRange().clear({contentsOnly: true})

    //Add in placeholders
    // super.getLogoRange().setValue("Logo image here")
    super.getParentRange().setValues([["parent name"], ["parents email"]])
    super.getAddressRange().setValues([["Address 1"], ["Address 2"], ["Phone number"]])
    super.getPaymentRange().setValues([["Company name"], ["Bank account"]])
    super.getInvoiceNumberRange().setValue("invoice number")
    super.getTermRange().setValue("term number and year")

    // Delete all unused rows by calculating the total cost row.
    let totalCostRow = super.getTotalCostRange().getRow();
    let firstCostRow = super.getCostsRange().getRow();
    let numberOfRows = totalCostRow - firstCostRow;
    let extraRows = numberOfRows - 3;
    if (extraRows > 0) {
      this.sheet.deleteRows(firstCostRow + 3, extraRows);
    }
  }
}

class InvoiceLoader extends InvoiceSheetRangeManager{
  constructor(sheet, invoice) {
    super(sheet);
    this.invoice = invoice;
  }

  loadInvoice(ssOfOrigin) {
    debug("Loading invoice information into sheet")
    this.loadParentInfo()
    this.loadCosts()
    this.loadInvoiceDetails()
    this.loadCompanyInfo()
    super.getNotesRange().setValue(this.invoice.note)
    super.getInvoiceInfoRange().setValues([[ssOfOrigin.getId(), this.invoice.type, this.invoice.updated ? "update" : "new"]])
    SpreadsheetApp.flush();
    debug("Finished loading")
  }

  loadParentInfo() {
    super.getParentRange().setValues([[this.invoice.parentName], [this.invoice.email]]);
  }

  /**
   * This will load the compay information into the sheet
   * There are 3 areas:
   * logo
   * address
   * payment
   */
  loadCompanyInfo() {
    let companyInfo = this.invoice.getCompanyInfo();

    // Logo
    //Disabled at the moment as all have the same image.
    // let logo = SpreadsheetApp.newCellImage().setSourceUrl(companyInfo.image);
    // super.getLogoRange().setValue(logo)

    //Address
    super.getAddressRange().setValues(companyInfo.address);

    //payment
    super.getPaymentRange().setValues([[companyInfo.name],[companyInfo.bankAccount]]);

    if (companyInfo.name.localeCompare("The Rock Academy") != 0) {
      super.getDivisionMessageLocation().setValue(companyInfo.name.concat(" - is a division of The Rock Academy"))
    }
  }

  /**
   * Currently only supports three cost items.
   */
  loadCosts() {
    //Loop through all of the costs and add them in.
    this.invoice.getCosts().forEach((cost, i) => {
      super.getCostRange(1+i, 1).setValues([[cost.desc, , , cost.quantity, cost.price]])
    })
  }

  loadInvoiceDetails() {
    super.getInvoiceNumberRange().setValue(this.invoice.number);
    super.getInvoiceNumberRange().setHorizontalAlignment("left");
    super.getTermRange().setValue(this.invoice.term);
  }
}


