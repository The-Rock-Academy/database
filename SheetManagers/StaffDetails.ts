class StaffDetails {
    static sheetName: string = "Staff";
    sheet: GoogleAppsScript.Spreadsheet.Sheet;

    constructor(ss) {
        this.sheet = ss.getSheetByName(StaffDetails.sheetName);
    }

    getColumn(name:string) {
        let column = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues()[0].indexOf(name) + 1;
        if (column == 0) {
            throw new Error("The column " + name + " does not exist");
        } else {
            return column;
        }
    }

    getEmail(name: string): string {
        let column = this.getColumn("Short name");
        
        //Get the row of the staff member
        let row = this.sheet.getRange(2, column, this.sheet.getLastRow() - 1, 1).getValues().map((row) => {
            return row[0];
        }).indexOf(name) + 2;

        return this.sheet.getRange(row, this.getColumn("Email")).getValue();
    }
}