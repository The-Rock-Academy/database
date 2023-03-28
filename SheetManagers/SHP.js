class SHPManager extends DatabaseSheetManager {
    constructor(sheet, currentTerm) {
        super(sheet, currentTerm);
        this.currentInvoiceColumn = this.getColumn()
    }
}