class BandSchoolManager extends DatabaseSheetManager {
    static sheetName() {
        return "Band School";
    }

    static newFromSS(ss) {
        return new BandSchoolManager(ss.getSheetByName(BandSchoolManager.sheetName()));
    }

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn("Current Invoice", false);
    }

    prepareInvoice(row, send = false) {
        let invoiceSheet = Invoices.newSheetManager(this, SpreadsheetApp.openById(this.databaseData.getVariable("Invoice Sender")).getSheetByName(this.databaseData.getVariable("Invoice Sender sheet name")));
        let ui = SpreadsheetApp.getUi();
        let activeRow = row;
        
        // --------------------------------
        // Checking for previous invoices
        // --------------------------------
        // ---- pupilName -----
        // Getting earlier as need to have this information for error message.
        let pupilName = this.sheet.getRange(activeRow, this.getColumn("Student Name")).getValue();

        // Check if invoice has already been sent.
        let invoiceNumberOfRow = this.getInvoiceNumberOfRow(activeRow)
        let updating = false;
        if (invoiceNumberOfRow != "" && !this.getInvoiceRanges(invoiceNumberOfRow).date.isBlank()) {
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

        //Getting lesson cost. Current it is very much hard coded but this can be fixed once I have gotten guidance from Geoff.
        let numberOfWeeks = this.sheet.getRange(1,this.getColumn("Current Term")).getMergedRanges()[0].getNumColumns()
        let totalCost = this.sheet.getColumn(activeRow, this.getColumn("Pupil cost")).getValue();
        let chargedLessons = numberOfWeeks;
        let costOfLesson = totalCost / chargedLessons;

        // ---- parentName -----
        let parentName = this.sheet.getRange(activeRow, this.getColumn("Guardian")).getValue();

        let email = this.sheet.getRange(activeRow, this.getColumn("Email")).getValue();

        let billingCompany = this.sheet.getRange(activeRow, this.getColumn("Pupils Billing Company")).getValue();

        if (!(parentName && email && billingCompany && pupilName && costOfLesson && (chargedLessons || chargedLessons  == 0) &&this.currentTerm)) {
        SpreadsheetApp.getUi().alert("Sorry the invoice for row " + row +" cannot be made as it is missing values. Please check all values and are present for the pupil.")
        return;
        }

        // -----------------------------
        // Create and load invoice into the invoice sheet
        // -----------------------------
        let invoice = Invoices.newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, chargedLessons, 0, costOfLesson, "", billingCompany,this.currentTerm, "band");
        if (updating) {
            invoice.number = this.getInvoiceNumberOfRow(row);
            let previousInvoiceInformation = this.getInvoiceRanges(invoice.number)
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

    updateSheetAfterInvoiceSent(invoiceNumber, totalCost, sentDate) {
        let invoiceRange = this.getInvoiceRanges(invoiceNumber);
        invoiceRange.amount.setValue(totalCost);
        invoiceRange.date.setValue(sentDate)
        invoiceRange.paidDate.clearContent();
    }

    clearInvoiceNumber(invoiceNumber) {
        console.log("Trying to clear " + invoiceNumber + " from the shp sheet")
        try {
          let invoiceInfo = this.getInvoiceRanges(invoiceNumber);
    
          //Dont clear the number if the user was trying to update the a invoice but decided against it.
          if (invoiceInfo.date.isBlank() && invoiceInfo.amount.isBlank()) {
            invoiceInfo.number.clearContent();
          }
        }
        catch(err) {
          console.warn("You have tried to clear a invoice number that couldnt be found.\n" + err);
        }
    }
}

function newBandSchoolSheet(sheet) {
    return new BandSchoolManager(sheet);
}