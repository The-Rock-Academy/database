class BandSchoolManager extends DatabaseSheetManager {
    static sheetName() {
        return "Band School";
    }

    static newFromSS(ss) {
        let sheet = ss.getSheetByName(BandSchoolManager.sheetName());
        return new BandSchoolManager(sheet);
    }

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn("Current Invoice", false);
        this.previousInvoiceColumn = this.getColumn("Previous Invoice ");
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
        let invoiceNumberOfRow = this.getInvoiceNumberOfRow(activeRow, previousTerm?this.previousInvoiceColumn: undefined)
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

        let invoiceTerm = previousTerm? this.sheet.getRange(1,this.previousInvoiceColumn).getValue().slice(17) : this.currentTerm;

        if (!(parentName && email && billingCompany && pupilName &&this.currentTerm)) {
            SpreadsheetApp.getUi().alert("Sorry the invoice for row " + row +" cannot be made as it is missing values. Please check all values and are present for the pupil.")
        return;
        }

        // -----------------------------
        // Create and load invoice into the invoice sheet
        // -----------------------------
        let invoice = newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, numberOfWeeks, 0, 30, "", billingCompany,invoiceTerm, "band");
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
        this.sheet.getRange(activeRow, previousTerm? this.previousInvoiceColumn: this.currentInvoiceColumn).setValue(invoice.number);


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

    newStudent(dayTime, Student_name, Parent_name, email, phone, instrument, billingCompany) {
        // Find row to add the new pupil too.

        let dayTimeSearcher = this.sheet.getRange(3, this.getColumn("Phone"), this.sheet.getMaxRows(), 1).createTextFinder(dayTime);

        let dayTimeRow = dayTimeSearcher.findNext();
        if (dayTimeRow == null) {
            throw new Error("Could not find the day time " + dayTime + " in the sheet " + this.sheet.getName());
        }
        dayTimeRow = dayTimeRow.getRow();

        let instrumentCol = this.getColumn("Instrument");

        let nextFreeRow = this.sheet.getRange(dayTimeRow+1, instrumentCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()+1;

        // Add the new pupil to the sheet.
        this.sheet.insertRowBefore(nextFreeRow);

        this.sheet.getRange(nextFreeRow, this.getColumn("Student name")).setValue(Student_name);
        this.sheet.getRange(nextFreeRow, this.getColumn("Guardian")).setValue(Parent_name);
        this.sheet.getRange(nextFreeRow, this.getColumn("Email")).setValue(email);
        this.sheet.getRange(nextFreeRow, this.getColumn("Phone")).setValue(phone);
        this.sheet.getRange(nextFreeRow, this.getColumn("Pupils Billing Company")).setValue(billingCompany);
        this.sheet.getRange(nextFreeRow, this.getColumn("Instrument")).setValue(instrument);
    
    }
}



function newBandSchoolSheet(sheet) {
    return new BandSchoolManager(sheet);
}