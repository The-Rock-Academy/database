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