class NewStudentManager {
    static sheetName: string = "New Students";
    sheet: GoogleAppsScript.Spreadsheet.Sheet;

    static columnHeaderRow: number = 1;
    static columnCategoryRow: number = 2;

    constructor(sheet) {
        this.sheet = sheet;
    }

    getColumn(name:string, category:string="", row:number=NewStudentManager.columnHeaderRow) {
        let column;
        if (category == "") {
            column = this.sheet.getRange(row, 1, 1, this.sheet.getLastColumn()).getValues()[0].indexOf(name) + 1;
        } else {
            let categoryHeader = this.sheet.getRange(NewStudentManager.columnCategoryRow,this.getColumn(category, "", NewStudentManager.columnCategoryRow));
            column = this.sheet.getRange(row, categoryHeader.getColumn(), 1, this.sheet.getLastColumn()).getValues()[0].indexOf(name)  + categoryHeader.getColumn();
        }
        if (column == 0) {
            throw new Error("The column " + name + " does not exist");
        } else {
            return column;
        }
    }

    clean() {
        console.log("Cleaning " + this.sheet.getName());
        let tutorNameRange = this.sheet.getParent().getSheetByName("Staff").getRange("B3:B55");

        this.sheet.getRange(3,this.getColumn("Billing Company"),this.sheet.getLastRow(),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["TRA", "GML", "TSA"]).build());

        this.sheet.getRange(3,this.getColumn("Tutor", "Weekly Lessons"),this.sheet.getLastRow(),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(tutorNameRange).build());
        this.sheet.getRange(3,this.getColumn("Tutor", "Band School"),this.sheet.getLastRow(),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(tutorNameRange).build());

        this.sheet.getRange(3,this.getColumn("Day", "Band School"),this.sheet.getLastRow(),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["MON", "TUE", "WED", "THU", "FRI"]).build());
        this.sheet.getRange(3,this.getColumn("Time", "Band School"),this.sheet.getLastRow(),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["4pm", "5pm"]).build());
    
        ["Mon", "Tue", "Wed", "Thu", "Fri"].forEach((day) => {
            this.sheet.getRange(3,this.getColumn(day, "School Holiday Programme"),this.sheet.getLastRow(),1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
        };
        console.log("Cleaned " + this.sheet.getName());
    }
}

function NewStudentManagerFromSS(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    return new NewStudentManager(ss.getSheetByName(NewStudentManager.sheetName));
}

class StudentProcessor extends NewStudentManager {
    activeRow: number;
    mainSS: GoogleAppsScript.Spreadsheet.Spreadsheet;


    constructor(sheet, activeRow:number) {
        super(sheet);
        this.activeRow = activeRow;
        this.mainSS = SpreadsheetApp.openByUrl((new DatabaseData(this.sheet.getParent())).getVariable("Main Database SS"));
    }

    clearLine() {
        this.sheet.getRange(this.activeRow, 1, 1, this.sheet.getLastColumn()).clear();
    }

    filterBlankColumns(columnNames: string[], category: string): {} {
        const result = {};
        columnNames.forEach((name) => {
            let range = this.sheet.getRange(this.activeRow, this.getColumn(name, category));

            if (range.isBlank()) {
                throw new Error("The " + name + " column is blank, please fill it in");
            } else {
                result[name.replaceAll(" ", "_")] = range.getValue();
            }
        });
        return result;
    }
    
    notifyOfNewStudent() {
        let templateSS = (new DatabaseData(this.mainSS)).getTemplateSS();


        let enquiryInformation = this.filterBlankColumns(["Name", "Email", "Phone", "Suburb", "Instruments interested in", "Services interested in", "Message"], "General");

        let emailer = Emails.newEmailer(templateSS, "PupilEnquiryFormSubmissionNotification");

        let emailToSendTo = (new DatabaseData(this.mainSS)).getVariable("New Enquiry notification email");

        emailer.sendEmail([emailToSendTo], enquiryInformation, [], enquiryInformation["Email"]);
    }

    getGenericInfo(): {} {
        let genericInfoColumnNames: string[] = ["Name", "Email", "Phone", "Suburb", "Student name", "Billing Company", "Level", "Age", "Instruments interested in"];

        return this.filterBlankColumns(genericInfoColumnNames, "General");
    }

    processNewStudent() {
        //Find out what services they are interested in
        let formResponse = this.sheet.getRange(this.activeRow, this.getColumn("Services interested in")).getValue();

        let weeklyLessons = formResponse.includes("Mobile Lessons");
        let bandSchool = formResponse.includes("Band School");
        let shp = formResponse.includes("School Holiday Programme");


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

        let genericInformation = this.getGenericInfo();

        let shpInformationColumnNames: string[] = ["Message","Emergency contact", "Mon", "Tue", "Wed", "Thu", "Fri"];
        let shpInfo = this.filterBlankColumns(shpInformationColumnNames, "School Holiday Programme");

        let shpManager = new SHPManager(this.sheet.getParent().getSheetByName(SHPManager.sheetName()));
        shpManager.addBooking(genericInformation.Student_name,
            genericInformation.Name,
            genericInformation.Email,
            genericInformation.Phone, 
            [shpInfo.Mon, shpInfo.Tue, shpInfo.Wed, shpInfo.Thu, shpInfo.Fri],
            shpInfo.Message,
            shpInfo.Emergency_contact);
    }

    processNewBandSchoolStudent() {
        let genericInformation = this.getGenericInfo();

        let bandSchoolInformationColumnNames: string[] = ["Day", "Time", "Tutor"];
        let bandSchoolInfo = this.filterBlankColumns(bandSchoolInformationColumnNames, "Band School");

        const newStudentInfo = {...genericInformation, ...bandSchoolInfo};

        console.log("The new student information for band school is: " + JSON.stringify(newStudentInfo));

        // Add pupil to the band school sheet
        let bandSchool = (new BandSchoolManager(SpreadsheetApp.openByUrl((new DatabaseData(this.mainSS)).getVariable("Band School ID")).getSheetByName(BandSchoolManager.sheetName()), ""))

        let dayTime = newStudentInfo.Day + " " + newStudentInfo.Time;
        bandSchool.newStudent(dayTime, newStudentInfo.Student_name, newStudentInfo.Name, newStudentInfo.Email, newStudentInfo.Phone, newStudentInfo.Instruments_interested_in, newStudentInfo.Billing_Company);

        // Send confirmation email

        // Get recipient information
        let tutor_email =  (new StaffDetails(this.mainSS)).getEmail(newStudentInfo.Tutor);

        let emailer = Emails.newEmailer(
            (new DatabaseData(this.mainSS)).getTemplateSS(), "Band School Enrolment Confirmation");

        emailer.sendEmail([newStudentInfo.Email, tutor_email], newStudentInfo);

    }

    prcoessNewWeeklyStudent() {
        // Get student information
        let genericInformation = this.getGenericInfo();

        let weeklyLessonInformationColumnNames: string[] = ["Preferred days of week", "Lesson length", "Lesson cost", "Instrument hire", "Tutor"];
        let weeklyLessonInfo = this.filterBlankColumns(weeklyLessonInformationColumnNames, "Weekly Lessons");

        const newStudentInfo = {...genericInformation, ...weeklyLessonInfo};


        let staffDetails =  new StaffDetails(this.mainSS);
        newStudentInfo.Tutor_Email = staffDetails.getEmail(newStudentInfo.Tutor);
        newStudentInfo.Tutor_Phone = staffDetails.getPhoneNumber(newStudentInfo.Tutor)

        console.log("The new student information for weekly lessons is: " + JSON.stringify(newStudentInfo));

        // Add the student to the weekly lessons sheet
        let attendanceManager = AttendanceManager.getObjFromSS(this.mainSS);
        attendanceManager.addStudent( newStudentInfo.Name, newStudentInfo.Email, newStudentInfo.Phone, newStudentInfo.Suburb, newStudentInfo.Student_name, newStudentInfo.Billing_Company, newStudentInfo.Preferred_days_of_week, newStudentInfo.Lesson_length, newStudentInfo.Lesson_cost, newStudentInfo.Instrument_hire, newStudentInfo.Tutor, newStudentInfo.Instruments_interested_in);

        // Get the PDF booklets
        let instruments:string[] = newStudentInfo.Instruments_interested_in.split(",").map(s => s.trim());

        let bookletFolder = DriveApp.getFolderById((new DatabaseData(this.sheet.getParent())).getVariable("Instrument booklets"));
        let bookletFolderFiles = bookletFolder.getFiles();

        let instrumentBooklets = [];

        while (bookletFolderFiles.hasNext()) {
            let file = bookletFolderFiles.next();
            let is_match = instruments.find(file_name => file.getName().includes(file_name));
            if (is_match != undefined) {
                instrumentBooklets.push(file.getAs("application/pdf"));
            }
        }

        
        // -----Email the parent and the tutor with confirmation-----

        // Get recipient information
        

        let emailer = Emails.newEmailer((new DatabaseData(this.mainSS)).getTemplateSS(), "Mobile Pupil Confirmation");

        emailer.sendEmail([newStudentInfo.Email, newStudentInfo.Tutor_Email], newStudentInfo, instrumentBooklets);
    }
    
}


function NewStudentManagerSheetName(): string {
    return NewStudentManager.sheetName;
}

function newStudentProcesser(activeSheet:  GoogleAppsScript.Spreadsheet.Sheet, activeRow: number): StudentProcessor {
    return new StudentProcessor(activeSheet, activeRow);
}