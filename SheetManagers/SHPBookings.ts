class SHPBookings {
    static sheetName: string = "SHP Bookings";
    static sheetName2: string = "SHP Bookings 2";
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet;
    week: number;

    constructor(ss, week =1) {
        this.ss = ss;
        this.week = week;

        if (week == 1) {
            this.sheet = ss.getSheetByName(SHPBookings.sheetName);
        }
        else if (week == 2) {
            this.sheet = ss.getSheetByName(SHPBookings.sheetName2);
        } else {
            throw new Error("Week to SHPBookings constructor must be 1 or 2");
        }

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


        // Make everything lower case
        let days = pupilInformation[7].split(", ");
        days = days.map((day) => {
            return day.toLowerCase();
        });

        let days_array = ["monday", "tuesday", "wednesday", "thursday", "friday"].map((day) => {
            return days.includes(day) ? true : false;
        });

        console.log("Recieved booking with information: " + pupilInformation);

        let shpManager = SHPManager.newFromSS(this.ss, null, this.week);

        shpManager.addBooking(pupilInformation[1], pupilInformation[2], pupilInformation[5], pupilInformation[3], days_array, pupilInformation[6], pupilInformation[4]);

        this.sheet.deleteRow(latestRow);
    }


    deleteAttachedForm() {
        // Delete main form

        let attachedFormURL = this.sheet.getFormUrl();

        if (attachedFormURL == null) {
            throw new Error("The sheet does not have an attached form");
        }

        let form = FormApp.openByUrl(attachedFormURL);
        form.removeDestination();

        DriveApp.getFileById(form.getId()).setTrashed(true);

        // Delete copy form

        let shp2Sheet = this.ss.getSheetByName(SHPBookings.sheetName2);
        let attachedFormURL2 = shp2Sheet.getFormUrl();
        if (attachedFormURL2 == null) {
            return;
        } else {
            let form2 = FormApp.openByUrl(attachedFormURL2);
            form2.removeDestination();
            DriveApp.getFileById(form2.getId()).setTrashed(true);
        }
    }

}
function SHPBookingsSheetName(week = 1) {
    if (week == 1) {
        return SHPBookings.sheetName;
    } else if (week == 2) {
        return SHPBookings.sheetName2;
    }
}

function SHPBookingsNewFromSS(ss, week = 1) {
    return new SHPBookings(ss, week);
}