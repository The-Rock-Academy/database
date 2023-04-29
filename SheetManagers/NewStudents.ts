class NewStudentManager {
    static sheetName: string = "New Students";
    sheet: GoogleAppsScript.Spreadsheet.Sheet;

    constructor(sheet) {
        this.sheet = sheet;
    }

    getColumn(name:string) {
        let column = this.sheet.getRange(2, 1, 1, this.sheet.getLastColumn()).getValues()[0].indexOf(name) + 1;
        if (column == 0) {
            throw new Error("The column " + name + " does not exist");
        } else {
            return column;
        }
    }


        
}

class StudentProcessor extends NewStudentManager {
    activeRow: number;


    constructor(sheet, activeRow:number) {
        super(sheet);
        this.activeRow = activeRow;
    }

    clearLine() {
        this.sheet.getRange(this.activeRow, 1, 1, this.sheet.getLastColumn()).clear();
    }

    filterBlankColumns(columnNames: string[]) {
        return columnNames.map((name) => {
            let range = this.sheet.getRange(this.activeRow, this.getColumn(name));
            if (range.isBlank()) {
                throw new Error("The " + name + " column is blank, please fill it in");
            } else{
                return range.getValue();
            }
        });
    }

    getGenericInfo(): string[] {
        let genericInfoColumnNames: string[] = ["Name", "Email", "Number", "Suburb", "Student name", "Billing Company"];

        return this.filterBlankColumns(genericInfoColumnNames);
    }

    processNewStudent() {
        //Look at what the student is interested in
        let weeklyLessons = this.sheet.getRange(this.activeRow, this.getColumn("Are you interested in Weekly lessons?")).getValue() == "Yes";
        let bandSchool = this.sheet.getRange(this.activeRow, this.getColumn("Are you interested in Band School?")).getValue() == "Yes";
        let shp = this.sheet.getRange(this.activeRow, this.getColumn("Are you interested in the School holiday program?")).getValue() == "Yes";

        console.log("This student is interested in weekly lessons: " + weeklyLessons + ", band school: " + bandSchool + ", and the school holiday program: " + shp);

        if (weeklyLessons) {
            this.prcoessNewWeeklyStudent();
        }

        if (bandSchool) {
            this.processNewBandSchoolStudent();
        }

        if (shp) {
            this.processNewSHPSudent();
        }

        this.clearLine();
    }
    processNewSHPSudent() {
        throw new Error("Method not implemented.");
    }
    processNewBandSchoolStudent() {
        throw new Error("Method not implemented.");
    }

    prcoessNewWeeklyStudent() {
        // Get student information
        let genericInformation = this.getGenericInfo();

        let weeklyLessonInformationColumnNames: string[] = ["Preferred days of week", "Lesson length", "Lesson cost", "Instrument hire", "Tutor"];
        let weeklyLessonInfo = this.filterBlankColumns(weeklyLessonInformationColumnNames);

        let newStudentInfo = genericInformation.concat(weeklyLessonInfo);

        console.log("The new student information for weekly lessons is: " + newStudentInfo);

        // Add the student to the weekly lessons sheet
        let attendanceManager = AttendanceManager.getObjFromSS(this.sheet.getParent());
        attendanceManager.addStudent.apply(attendanceManager, newStudentInfo);

        // -----Email the parent and the tutor with confirmation-----

        // Get the templates
        let templateSheet = this.sheet.getParent().getSheetByName("MobilePupilConfirmationTemplate");
        if (templateSheet == null) {
            throw new Error("The template sheet MobilePupilConfirmationTemplate does not exist");
        }
        let subject_template = templateSheet.getRange(1, 2).getValue();
        let body_template = templateSheet.getRange(2, 2).getValue();

        // Get the data
        let data = {
            "Name": genericInformation[0],
            "Tutor": weeklyLessonInfo[weeklyLessonInformationColumnNames.indexOf("Tutor")],
            "Lesson_length": weeklyLessonInfo[weeklyLessonInformationColumnNames.indexOf("Lesson length")],
            "Lesson_cost": weeklyLessonInfo[weeklyLessonInformationColumnNames.indexOf("Lesson cost")],
            "Instrument_hire": weeklyLessonInfo[weeklyLessonInformationColumnNames.indexOf("Instrument hire")],
            "Preferred_days_of_week": weeklyLessonInfo[weeklyLessonInformationColumnNames.indexOf("Preferred days of week")],
            "Number": genericInformation[2],
            "Suburb": genericInformation[3],
            "Student_name": genericInformation[4],
            "Billing_company": genericInformation[5],
        };  
        };
        
        let emailer = Emails.newEmailer("New Student", "New Student");

    }
    
}


function NewStudentManagerSheetName(): string {
    return NewStudentManager.sheetName;
}

function newStudentProcesser(activeSheet:  GoogleAppsScript.Spreadsheet.Sheet, activeRow: number): StudentProcessor {
    return new StudentProcessor(activeSheet, activeRow);
}