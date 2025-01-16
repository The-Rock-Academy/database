class BandSchoolInvoicingManager extends DatabaseSheetManager {
    static sheetName() {
        return "Band School";
    }

    static newFromSS(ss) {
        let sheet = ss.getSheetByName(BandSchoolManager.sheetName());
        return new BandSchoolManager(sheet);
    }

    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
    }

    reset(nextTermDates, nextTerm) {
        this.resetInvoiceColumns(nextTerm, this.getColumn("Current Invoice", false)+4);
        this.sheet.getParent().setName(nextTerm + " Band School Invoicing - TRA");
    }

    archive(term) {
        let bandSchoolFolder = DriveApp.getFolderById(this.databaseData.getVariable("Band School Folder"));

        Database.createSpreadSheetCopy(this.sheet.getParent(), bandSchoolFolder, term + " Band School Invoicing - TRA Archive");
    }
}

function newBandSchoolInvoicingSheet(sheet) {
    return new BandSchoolInvoicingManager(sheet);
}

function getBandSchoolInvoicingSheetName() {
    return BandSchoolInvoicingManager.sheetName();
}