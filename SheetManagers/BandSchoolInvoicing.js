class BandSchoolInvoicingManager extends DatabaseSheetManager {
    static sheetName() {
        return "Band School";
    }

    static newFromSS(ss) {
        let sheet = ss.getSheetByName(BandSchoolInvoicingManager.sheetName());
        return new BandSchoolInvoicingManager(sheet);
    }

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn("Current Invoice", false);
    }

    reset(nextTermDates, nextTerm) {
        this.resetInvoiceColumns(nextTerm, this.getColumn("Current Invoice", false)+4);
        this.sheet.getParent().setName(nextTerm + " Band School Invoicing - TRA");
    }

    archive(term) {
        let bandSchoolFolder = DriveApp.getFolderById(this.databaseData.getVariable("Band School Folder"));

        Database.createSpreadSheetCopy(this.sheet.getParent(), bandSchoolFolder, term + " Band School Invoicing - TRA Archive");
    }
    
    prepareAndSendInvoice(range, previousTerm = false) {
        for (let row = range.getRowIndex(); row < range.getHeight() + range.getRowIndex(); row++) {
          this.prepareInvoice(row, true, previousTerm);
        }
    }

    invoiceCurrent(invoiceNumber) {
        let foundRowCurrent = this.sheet.getRange(3, this.currentInvoiceColumn, this.sheet.getMaxRows(), 1).createTextFinder(invoiceNumber).matchEntireCell(true).findNext();
        if (foundRowCurrent != null) return true
        else return false
      }

    prepareInvoice(row, send = false, previousTerm = false) {
        let invoiceSheet = newSheetManager(this, SpreadsheetApp.openById(this.databaseData.getVariable("Invoice Sender")).getSheetByName(this.databaseData.getVariable("Invoice Sender sheet name")));
        let ui = SpreadsheetApp.getUi();
        let activeRow = row;
        
        // --------------------------------
        // Checking for previous invoices
        // --------------------------------
        // ---- pupilName -----
        // Getting earlier as need to have this information for error message.
        let pupilName = this.sheet.getRange(activeRow, this.getColumn("Student Name")).getValue();

        // Check if invoice has already been sent.
        let invoiceNumberOfRow = this.getInvoiceNumberOfRow(activeRow, undefined)
        let updating = false;
        if (invoiceNumberOfRow != "" && !this.getInvoiceRanges(invoiceNumberOfRow, previousTerm).date.isBlank()) {
            let answer = ui.alert("It appears you have already made and sent an invoice for " + pupilName + " for the term.\nThe new invoice you create for this pupil will override the previous one you had.\nWould you like to continue with making a new one?", ui.ButtonSet.YES_NO)
            if (answer == ui.Button.NO) {
            return
            } else {
            updating = true
            }
        }

        // Check if the invoice sender is already occupied.
        if (invoiceSheet.isInvoiceLoaded()) {
        let answer = ui.alert("It appears the invoice sender already has an invoice loaded. Would you like to overide that invoice?", ui.ButtonSet.YES_NO)
        if (answer == ui.Button.NO) {
            return
        } else {
            invoiceSheet.clearInvoice();
        }
        }

        // --------------------------------
        // Collecting information for invoice
        // --------------------------------

        let numberOfWeeks = this.sheet.getRange(1,this.getColumn(previousTerm ? "Previous Term" : "Current Term")).getMergedRanges()[0].getNumColumns()

        let parentName = this.sheet.getRange(activeRow, this.getColumn("Guardian")).getValue();

        let email = this.sheet.getRange(activeRow, this.getColumn("Email")).getValue();

        let billingCompany = this.sheet.getRange(activeRow, this.getColumn("Pupils Billing Company")).getValue();

        let invoiceTerm = this.currentTerm;

        let totalPrice = this.sheet.getRange(activeRow, this.getColumn("Pupil cost")).getValue();

        if (!(parentName && email && billingCompany && pupilName &&this.currentTerm  && totalPrice)) {
            SpreadsheetApp.getUi().alert("Sorry the invoice for row " + row +" cannot be made as it is missing values. Please check all values are present for the pupil.")
        return;
        }

        // -----------------------------
        // Create and load invoice into the invoice sheet
        // -----------------------------
        let invoice = newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, 1, 0, totalPrice, "", billingCompany,invoiceTerm, "band");
        if (updating) {
            invoice.number = this.getInvoiceNumberOfRow(row, previousTerm);
            let previousInvoiceInformation = this.getInvoiceRanges(invoice.number, previousTerm)
            invoice.note = "This invoice is an updated version of an invoice sent on " + previousInvoiceInformation.date.getValue().toLocaleString('en-NZ') + ", for $" + previousInvoiceInformation.amount.getValue() + ".";
            invoice.updated = true;
        }
        invoiceSheet.loadInvoice(invoice);

        // ----------------------------
        // Load the invoice number into the attendance sheet
        // ----------------------------
        this.sheet.getRange(activeRow, this.currentInvoiceColumn).setValue(invoice.number);


        SpreadsheetApp.flush(); //This ensures that the invoice is actaully loaded at this point

        // -----------------------------
        // Potentially send the invoice if needed
        // -----------------------------
        if(send) {
        invoiceSheet.sendInvoice(updating);
        }
    }

    updateSheetAfterInvoiceSent(invoiceNumber, totalCost, sentDate, numberOfLessons) {
        let invoiceRange = this.getInvoiceRanges(invoiceNumber, !this.invoiceCurrent(invoiceNumber));
        invoiceRange.amount.setValue(totalCost);
        invoiceRange.date.setValue(sentDate)
        invoiceRange.paidDate.clearContent();
    }

    clearInvoiceNumber(invoiceNumber) {
        console.log("Trying to clear " + invoiceNumber + " from the shp sheet")
        try {
          let invoiceInfo = this.getInvoiceRanges(invoiceNumber, !this.invoiceCurrent(invoiceNumber));
    
          //Dont clear the number if the user was trying to update the a invoice but decided against it.
          if (invoiceInfo.date.isBlank() && invoiceInfo.amount.isBlank()) {
            invoiceInfo.number.clearContent();
          }
        }
        catch(err) {
          console.warn("You have tried to clear a invoice number that couldnt be found.\n" + err);
        }
    }

    sendReminder(range) {
        let invoiceCollector = new InvoiceCollector(this.sheet, this.currentInvoiceColumn+2, this.currentInvoiceColumn+3, this.getColumn("Guardian"), this.getColumn("Email"), this.getColumn("Student Name"), this.currentInvoiceColumn, this.currentInvoiceColumn+1, this.getColumn("Invoice reminder"));
    
        invoiceCollector.sendReminders(range);
    }
      
}

function newBandSchoolInvoicingSheet(sheet) {
    return new BandSchoolInvoicingManager(sheet);
}

function getBandSchoolInvoicingSheetName() {
    return BandSchoolInvoicingManager.sheetName();
}