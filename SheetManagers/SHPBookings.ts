class SHPBookings {
    static sheetName: string = "SHP Bookings";
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet;

    constructor(ss) {
        this.ss = ss;
        this.sheet = ss.getSheetByName(SHPBookings.sheetName);
        if (this.sheet == null) {
            throw new Error("The sheet '" + SHPBookings.sheetName + "' does not exist");
        }
    }

    getColumn(name:string) {
        let column = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues()[0].indexOf(name) + 1;
        if (column == 0) {
            throw new Error("The column " + name + " does not exist");
        } else {
            return column;
        }
    }

    // Take the form submission and add booking to the SHP sheet.
    formSubmission() {
        let latestRow = this.sheet.getLastRow();
        if (latestRow == 1) {
            console.log("No bookings to process");
            return;
        }
        let pupilInformation = this.sheet.getRange(latestRow, 1, 1, this.sheet.getLastColumn()).getValues()[0];


        let days = pupilInformation[7].split(", ");
        let days_array = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].map((day) => {
            return days.includes(day) ? true : false;
        });

        console.log("Recieved booking with information: " + pupilInformation);

        let shpManager = SHPManager.newFromSS(this.ss);

        shpManager.addBooking(pupilInformation[1], pupilInformation[2], pupilInformation[5], pupilInformation[3], days_array, pupilInformation[6], pupilInformation[4]);

        this.sheet.deleteRow(latestRow);
    }


    deleteAttachedForm() {
        let attachedFormURL = this.sheet.getFormUrl();

        if (attachedFormURL == null) {
            throw new Error("The sheet does not have an attached form");
        }

        let form = FormApp.openByUrl(attachedFormURL);
        form.removeDestination();

        DriveApp.getFileById(form.getId()).setTrashed(true);
    }

}
function SHPBookingsSheetName() {
    return SHPBookings.sheetName;
}

function SHPBookingsNewFromSS(ss) {
    return new SHPBookings(ss);
}