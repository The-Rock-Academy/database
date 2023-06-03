class SHPManager extends DatabaseSheetManager {
    static sheetName() {
        return "SHP";
    }

    static newFromSS(ss, currentTerm) {
        return (new SHPManager(ss.getSheetByName(SHPManager.sheetName()), currentTerm));
    }

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn("Invoice", true)
    }

    clean() {
        console.log("Nothing to be cleaned yet on the SHP");
    }

    reset(nextTermDetails, nextTerm) {

        // Insert new header and rows
        this.sheet.insertRowsBefore(3, 31);
        let headerRow = this.sheet.getRange(3,1,1,this.sheet.getMaxColumns());

        let finalTermWCDate = nextTermDetails[nextTermDetails.length - 1];
        let tempDate = new Date(finalTermWCDate);
        tempDate.setDate(tempDate.getDate() + 7);
        let holidayStartDate = new Date(tempDate);
        tempDate.setDate(tempDate.getDate() + 5);
        let holidayEndDate = new Date(tempDate);
        let sameMonth = holidayStartDate.getMonth() == holidayEndDate.getMonth();
        
        let headerText = "End of " + this.currentTerm + " : " +
        holidayStartDate.getDate() + (sameMonth ? "" : " " + holidayStartDate.toLocaleDateString(undefined, { month: 'long'})) + " - "
        + holidayEndDate.getDate() + " " + holidayEndDate.toLocaleDateString(undefined, { month: 'long'})
        console.log(headerText);

        headerRow.setValue(headerText).merge();
        headerRow.setFontColor('black');
        headerRow.setFontSize(13);
        headerRow.setHorizontalAlignment("center");
        headerRow.setFontWeight("bold")
    }

    updateSheetAfterInvoiceSent(invoiceNumber, totalCost, sentDate, numberOfLessons) {
        let invoiceRange = this.getInvoiceRanges(invoiceNumber);
        invoiceRange.amount.setValue(totalCost);
        invoiceRange.date.setValue(sentDate)
        invoiceRange.paidDate.clearContent();
    }

    findRowTerm(startRow) {
        let row = startRow;
        while (true) {
          let range = this.sheet.getRange(row, 1);
          let merged = range.getMergedRanges();
          if (merged.length > 0 && merged[0].getNumColumns() == this.sheet.getLastColumn()) {
            return range.getValue();
          }
          row--;
          if (row < 3) {
            throw new Error("Couldnt find a term value for the row: " + startRow);
          }
        }
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

    clearSheetAfterClearingInvoiceSender(invoiceNumber) {
        this.clearInvoiceNumber(invoiceNumber);
    }

    prepareInvoice(row, send = false, automated = false) {

        let invoiceSheet = Invoices.newSheetManager(this, SpreadsheetApp.openById(this.databaseData.getVariable("Invoice Sender")).getSheetByName(this.databaseData.getVariable("Invoice Sender sheet name")));
        
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
        if (!automated) { //Dont check if things are dupes or the invoice sender is full if this is automated.
            let ui = SpreadsheetApp.getUi();
            if (invoiceNumberOfRow != "" && !this.getInvoiceRanges(invoiceNumberOfRow).date.isBlank()) {
            let answer = ui.alert("It appears you have already made and sent an invoice for " + pupilName + " for the SHP.\nThe new invoice you create for this pupil will override the previous one you had.\nWould you like to continue with making a new one?", ui.ButtonSet.YES_NO)
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
        }
    
        // --------------------------------
        // Collecting information for invoice
        // --------------------------------
    
        // ----- Get number of lessons -------
    
        //Find out what week to just assume all lessons will be attended
        let price = this.sheet.getRange(activeRow, this.getColumn("Price")).getValue();
    
        let numberOfLessons = price == 350 ? 5 : price / 75;
    
        // ---- parentName -----
        let parentName = this.sheet.getRange(activeRow, this.getColumn("Guardian")).getValue();
    
        let email = this.sheet.getRange(activeRow, this.getColumn("Email")).getValue();
    
        let instrumentHire = this.sheet.getRange(activeRow, this.getColumn("Instrument Hire cost")).getValue();
    
        let billingCompany = this.sheet.getRange(activeRow, this.getColumn("Pupils Billing Company")).getValue();
    
        if (!(parentName && email && billingCompany && pupilName && price && (numberOfLessons || numberOfLessons  == 0) &&this.currentTerm)) {
            throw new Error("Sorry the invoice for row " + row +" cannot be made as it is missing values. Please check all values and are present for the pupil.");
        }
        // -----------------------------
        // Create and load invoice into the invoice sheet
        // -----------------------------
        let invoice = Invoices.newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, numberOfLessons, 0, price / numberOfLessons, instrumentHire, billingCompany, "School holidays " + this.findRowTerm(row), "shp");
        if (updating) {
            invoice.number = this.getInvoiceNumberOfRow(row);
            invoice.note = "This invoice is an updated version of an invoice sent on " + this.getInvoiceRanges(invoice.number).date.getValue().toLocaleString();
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

    /**
     * Retrieves the invoice number of a row
     * @param {number} rowNumber Row number you are checking
     * @returns 
    */
    getInvoiceNumberOfRow(rowNumber) {
        return this.sheet.getRange(rowNumber, this.currentInvoiceColumn).getValue()
    }

    addBooking(Student_name, Parent_name, email, phone, days, notes, Emergency_contact) {
        this.sheet.insertRowBefore(4);

        //Add in price formula
    

        // Using rowRange to help with the problem of this being done at the same time. I.e if we have people submiting a form at the same time.
        let rowRange = this.sheet.getRange(4, 1, 1, this.sheet.getLastColumn());

        this.sheet.getRange(rowRange.getRow(), this.getColumn("Price")).setFormula("=IF(COUNTA(Z4:AD4)=5, 350, counta(Z4:AD4) * 75)");

        this.sheet.getRange(rowRange.getRow(), this.getColumn("Student Name")).setValue(Student_name);
        this.sheet.getRange(rowRange.getRow(), this.getColumn("Guardian Name")).setValue(Parent_name);
        this.sheet.getRange(rowRange.getRow(), this.getColumn("Email")).setValue(email);
        this.sheet.getRange(rowRange.getRow(), this.getColumn("Phone")).setValue(phone);
        this.sheet.getRange(rowRange.getRow(), this.getColumn("Emergency Contact")).setValue(Emergency_contact);
        this.sheet.getRange(rowRange.getRow(), this.getColumn("Notes and Allergies")).setValue(notes);

        this.sheet.getRange(rowRange.getRow(), this.getColumn("Mon Day 1"), 1, 5).setValues([days.map(day => {
            if (day) {
                return "x"
            } else {
                return ""
        }}
        )])

        // Fill in the billing company.

        let previousBillingCompany = this.sheet.getRange(rowRange.getRow()+1, this.getColumn("Pupils Billing Company")).getValue();
        switch (previousBillingCompany) {
            case "TRA":
                this.sheet.getRange(rowRange.getRow(), this.getColumn("Pupils Billing Company")).setValue("GML");
                break;
            case "GML":
                this.sheet.getRange(rowRange.getRow(), this.getColumn("Pupils Billing Company")).setValue("TSA");
                break;
            case "TSA":
                this.sheet.getRange(rowRange.getRow(), this.getColumn("Pupils Billing Company")).setValue("TRA");
                break;
            default:
                this.sheet.getRange(rowRange.getRow(), this.getColumn("Pupils Billing Company")).setValue("TRA");
        }
        this.prepareInvoice(rowRange.getRow(), true, true);
    }
}

function newSHPSheet(shpSheet) {
    return new SHPManager(shpSheet)
}