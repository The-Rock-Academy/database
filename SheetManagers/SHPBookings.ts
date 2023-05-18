class SHPBookings {
    static sheetName: string = "School Holiday Program bookings";
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet;

    constructor(ss) {
        this.ss = ss;
        this.sheet = ss.getSheetByName(SHPBookings.sheetName);
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
        let pupilInformation = this.sheet.getRange(latestRow, 1, 1, this.sheet.getLastColumn()).getValues()[0];


        let days = pupilInformation[7].split(", ");
        let days_array = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].map((day) => {
            console.log("Looking at day: " + day + " and seeing if it is in " + days + ". Turns out it is " + days.includes(day) + ".");
            return days.includes(day) ? true : false;
        });

        console.log("Recieved booking with information: " + pupilInformation);

        let shpManager = SHPManager.newFromSS(this.ss);

        shpManager.addBooking(pupilInformation[1], pupilInformation[2], pupilInformation[5], pupilInformation[3], days_array, pupilInformation[6], pupilInformation[4]);

        this.sheet.deleteRow(latestRow);
    }
}
function SHPBookingsSheetName() {
    return SHPBookings.sheetName;
}

function SHPBookingsNewFromSS(ss) {
    return new SHPBookings(ss);
}