class SHPManager extends DatabaseSheetManager {
    static sheetName() {
        return "SHP";
    }

    static newFromSS(ss, currentTerm) {
        return new SHPManager(ss.getSheetByName(SHPManager.sheetName()), currentTerm);
    }

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn("Invoice")
    }

    clean() {
        console.log("Nothing to be cleaned yet on the SHP");
    }

    reset(nextTermDetails, nextTerm) {

        // Insert new header and rows
        this.sheet.insertRowsBefore(3, 31);
        let headerRow = this.sheet.getRange(3,1,1,this.currentInvoiceColumn);

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
    }

    
}