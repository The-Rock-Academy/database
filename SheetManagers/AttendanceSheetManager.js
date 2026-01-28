class AttendanceSheetManager extends DatabaseSheetManager {

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn("Current Invoice", false);
        this.currentTermAttendanceColumnNum = this.getColumn("Current Term");
        this.currentTermWeeks = this.getCurrentTermWeeks();
    }

    reset(nextTermDates, nextTerm) {
        this.resetAttendanceColumns(nextTermDates, nextTerm);
        this.resetInvoiceColumns(nextTerm, this.currentInvoiceColumn+4+nextTermDates.length);

        this.sheet.deleteColumns(this.currentTermAttendanceColumnNum, this.currentTermWeeks);
    }

    resetAttendanceColumns(nextTermDetails, nextTerm) {
        let currentTermNameRange = this.sheet.getRange(1,this.currentTermAttendanceColumnNum);
    
        //Rename current term
        currentTermNameRange.setValue("Previous " +this.currentTerm)
    
        //Add in new Attendance
        let columnOfNextTerm = this.currentInvoiceColumn+4
        let numberOfWeeksOfNextTerm = nextTermDetails.length
    
        //Add in column
        this.sheet.insertColumns(columnOfNextTerm, numberOfWeeksOfNextTerm)
        let nextTermNameRange = this.sheet.getRange(1,columnOfNextTerm, 1, numberOfWeeksOfNextTerm);
        let nextTermDateRange = this.sheet.getRange(2, columnOfNextTerm, 1, numberOfWeeksOfNextTerm)
        //Name the term area
        nextTermNameRange.setValue("Current " + nextTerm).merge()
        //Set date values
        nextTermDateRange.setValues([nextTermDetails])
    
        //Copy across formatting
          //Doing a complex format copying because there may be different numbers of weeks and the previous function did not copy column widths
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
        let merged_ranges = this.sheet.getRange(1,this.currentTermAttendanceColumnNum).getMergedRanges()
        if (merged_ranges.length == 0) {
            throw new Error("Current Term header is not merged. Merge the header cells for the current term attendance column.");
        }

    return merged_ranges[0].getNumColumns()
    }

    sendReminder(range) {
        let invoiceCollector = new InvoiceCollector(this.sheet, this.currentInvoiceColumn+2, this.currentInvoiceColumn+3, this.getColumn("Guardian"), this.getColumn("Email"), this.getColumn("Student Name"), this.currentInvoiceColumn, this.currentInvoiceColumn+1, this.getColumn("Invoice reminder"));
    
        invoiceCollector.sendReminders(range);
    }
}