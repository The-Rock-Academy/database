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
        this.currentTermAttendanceColumnNum = this.getColumn("Current Term");
        this.currentTermWeeks = this.getCurrentTermWeeks();
    }

    reset(nextTermDates, nextTerm) {
        this.resetAttendanceColumns(nextTermDates, nextTerm);
        this.sheet.deleteColumns(this.currentTermAttendanceColumnNum, this.currentTermWeeks);
        this.sheet.insertRowAfter(this.sheet.getMaxColumns())
        this.sheet.getParent().setName(nextTerm + " Band School - TRA");
    }

    archive(term) {
        let bandSchoolFolder = DriveApp.getFolderById(this.databaseData.getVariable("Band School Folder"));

        Database.createSpreadSheetCopy(this.sheet.getParent(), bandSchoolFolder, term + " Band School - TRA Archive");
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

    // Added in so that this calss doesnt have to be derived from attendancesheetmanager which requires a invoice column
    resetAttendanceColumns(nextTermDetails, nextTerm) {
        let currentTermNameRange = this.sheet.getRange(1,this.currentTermAttendanceColumnNum);
    
        //Rename current term
        currentTermNameRange.setValue("Previous " +this.currentTerm)
    
        //Add in new Attendance
        let columnOfNextTerm = 20
        let numberOfWeeksOfNextTerm = nextTermDetails.length
    
        //Add in column
        this.sheet.insertColumns(columnOfNextTerm, numberOfWeeksOfNextTerm)
        let nextTermNameRange = this.sheet.getRange(1,columnOfNextTerm, 1, numberOfWeeksOfNextTerm);
        let nextTermDateRange = this.sheet.getRange(2, columnOfNextTerm, 1, numberOfWeeksOfNextTerm)
        //Name the term area
        nextTermNameRange.setValue("Current " + nextTerm).merge()
        //Set date values
        nextTermDateRange.setValues([nextTermDetails])
    
        //Copy across formattting
          //Doing a complex format copying becuase there may be different numbers of weeks and the previous function did not copy column widths
        nextTermDetails.map((date, index) => {
          this.sheet.getRange(2,this.currentTermAttendanceColumnNum).copyTo(this.sheet.getRange(2,columnOfNextTerm + index), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
          this.sheet.getRange(2,this.currentTermAttendanceColumnNum).copyTo(this.sheet.getRange(2,columnOfNextTerm + index), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        })
    
        //Add in border on right
        this.sheet.getRange(1,columnOfNextTerm + numberOfWeeksOfNextTerm-1, this.sheet.getMaxRows()).setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.DOUBLE)
    
        //Copy term range format
        currentTermNameRange.copyTo(nextTermNameRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false)
    }
      
    getCurrentTermWeeks() {
    return this.sheet.getRange(1,this.currentTermAttendanceColumnNum).getMergedRanges()[0].getNumColumns()
    }
}

function getBandSchoolSheetName() {
    return BandSchoolManager.sheetName();
}

function newBandSchoolSheet(sheet) {
    return new BandSchoolManager(sheet);
}