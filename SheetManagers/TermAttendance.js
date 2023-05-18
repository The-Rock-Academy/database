class AttendanceManager extends DatabaseSheetManager {
  static sheetName(){return "Master sheet"}
  static numberOfColumnsBeforeAttendanceStart() {return 2}

  static getObjFromSS(SS) {
    return new AttendanceManager(SS.getSheetByName(AttendanceManager.sheetName()))
  }

  constructor(sheet, currentTerm) {
    super(sheet, currentTerm);

    this.currentTermAttendanceColumnNum = this.getColumn("Current " + this.currentTerm);
    this.currentTermWeeks = this.getCurrentTermWeeks();
    this.currentInvoiceColumn = this.getColumn("Current Invoice " + this.currentTerm, true)
    this.previousInvoiceColumn = this.getColumn("Previous Invoice ");
    this.previousTermWeeks = this.previousInvoiceColumn- AttendanceManager.numberOfColumnsBeforeAttendanceStart();
  }

  clean() {
    // Get things that will be needed throughout the clean.
    // Some of these might need to be refactoed to be part of the object.
    let statusColumnNumber = this.getColumn("Status");
    let statusColumn = this.sheet.getRange(3, statusColumnNumber, this.getInactiveRowNumber()-2, 1);
    
    // Remove all the inactive pupils to the inactive section
    let inactiveSearch = statusColumn.createTextFinder("Inactive");
    for (const i in Array.from(Array(inactiveSearch.findAll().length).keys())) {
      this.sheet.moveRows(inactiveSearch.findNext(), this.getInactiveRowNumber()+3)
    }

    //Sort the active section by the status type
    // To do this I need to get the active range and then sort it using the sort function
    let activeRange = this.sheet.getRange(3, 1, this.getInactiveRowNumber()-3, this.sheet.getMaxColumns());

    activeRange.sort({column: this.getColumn("Teacher"), ascending: true})
  }

  /**
   * Reset the attendance sheet.
   * This will move the current term to the previous term slots both for attendance and for invoice
   */
  reset(nextTermDetails, nextTerm) {
    console.log("Resetting the Attendance sheet")

    //----Refreshing the attendance section----------

    let currentTermNameRange = this.sheet.getRange(1,this.currentTermAttendanceColumnNum);
    let currentTermDateRange = this.sheet.getRange(2,this.currentTermAttendanceColumnNum, 1,this.currentTermWeeks)
    //Rename current term
    currentTermNameRange.setValue("Previous " +this.currentTerm)

    //Add in new Attendance
    let columnOfNextTerm = this.currentInvoiceColumn+4
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

    //Delete old term information
    this.sheet.deleteColumns(AttendanceManager.numberOfColumnsBeforeAttendanceStart(), this.currentTermAttendanceColumnNum - AttendanceManager.numberOfColumnsBeforeAttendanceStart())

    //-----Refreshing the invoice section-------------

    this.resetInvoiceColumns(nextTerm, columnOfNextTerm+numberOfWeeksOfNextTerm-(this.currentTermAttendanceColumnNum - AttendanceManager.numberOfColumnsBeforeAttendanceStart()));

  }

  getInactiveRowNumber() {
    return this.sheet.createTextFinder("Inactive Pupils").findNext().getRowIndex();

  }

  getCurrentTermWeeks() {
    return this.sheet.getRange(1,this.currentTermAttendanceColumnNum).getMergedRanges()[0].getNumColumns()
  }

  // -------------------------------------------------------------------------------------------------------
  // --------- Invoices -----------------------------------------------------------------------------------
  // -------------------------------------------------------------------------------------------------------
  // These methods here are all about dealing with the invoices in the attendance sheet.

  // ------------------
  // Non term length invoices
  // ------------------
  /**
   * Gets the previous invoices that are stored in the comments section
   */
  getPreviousInvoices(row) {
    let comments = this.sheet.getRange(row, this.getColumn("Comments")).getValue();
    let regexOutcome = /{Previousinvoices:((\d+[a-z]*,?)*)}/g.exec(comments.replace(/\s+/g, ''))
    if (regexOutcome != null) {
      let previousInvoices = new Array();
      let previousInvoicesStringSplit = regexOutcome[1].split(',');
      previousInvoicesStringSplit.forEach((element,index) => {
        previousInvoices.push({
          number: /\d+/g.exec(element)[0],
          paid: !element.includes("unpaid"),
          temp: element.includes("temp")
        })
      })
      return previousInvoices;
    } {
      return null
    }
  }

  insertPreviousInvoicesIntoComments(row, invoices) {
    let comments = this.sheet.getRange(row, this.getColumn("Comments"));
    
    // Take all of the invoices and turn it into a array that can be compressed into a string
    let previousInvoiceStringArr = ["{Previous invoices:"];
    invoices.forEach(invoice => {
      previousInvoiceStringArr.push(invoice.number + (invoice.paid ? "" : " unpaid") + (invoice.temp ? " temp" : ""))
    })

    let previousInvoiceString = previousInvoiceStringArr.toString().replace(",", "") + "}";
    let newComment;

    // Either this is the first time it is being added the previousInvoice section was alread there.
    if (this.getPreviousInvoices(row) == null) {
      newComment =  previousInvoiceString + " " + comments.getValue();
    } else {
      newComment = comments.getValue().replace(/{Previous invoices:.*}/g, previousInvoiceString)
    }

    comments.setValue(newComment)
  }

  /**
   * Take a invoice object and add it into the comments section
   */
  addPreviousInvoice(row, invoice) {
    //Start the intialisation of prievious invoices procedure
    let currentInvoices = this.getPreviousInvoices(row);
    if (currentInvoices == null) {
      this.insertPreviousInvoicesIntoComments(row, [invoice])
    } else {
      currentInvoices.push(invoice)
      this.insertPreviousInvoicesIntoComments(row, currentInvoices)
    }
  }

  /**
   * Add a temp invoice in using the addPreviousInvoicefunciton
   */
  addTempInvoice(row, number, paid) {
    this.addPreviousInvoice(row, {number: number, paid: paid, temp: true})
  }

  /**
   * This will remove the tmep invoice you just added
   */
  removeTempInvoice(row){
    let previousInvoices = this.getPreviousInvoices(row);
    if (previousInvoices != null) {
      let tempInvoice = previousInvoices.pop();
      if (tempInvoice.temp) {
        this.insertPreviousInvoicesIntoComments(row, previousInvoices);
        return tempInvoice.number;
      } else {
        return null
      }
    }
  }


  /**
   * Will make the most recently added invoice non temp if it was previously temp
   */
  makeTempInvoicePermanent(row) {
    let previousInvoices = this.getPreviousInvoices(row);
    if (previousInvoices == null) {
      return
    }
    let tempInvoice = previousInvoices[previousInvoices.length - 1];

    if (tempInvoice.temp) {
      previousInvoices[previousInvoices.length - 1] = {number: tempInvoice.number, paid: tempInvoice.paid, temp: false}
      this.insertPreviousInvoicesIntoComments(row, previousInvoices);
    }
  }

  /**
 * This will update the attendance section to make the P and I now be coloured
 */
  updateAttendanceToInvoiced(invoiceNumber) {
    console.log(this.getAttendanceRange(this.getInvoiceRow(invoiceNumber)).getValues())
    this.getAttendanceRange(this.getInvoiceRow(invoiceNumber), true, !this.invoiceCurrent(invoiceNumber)).forEach(range => {
      console.log("Checking: " + range.getValue());
      if (range.getBackground() != "#c8c8c8" && ["T", "P", "I"].includes(range.getValue())) {
        console.log("Updating colour")
        range.setBackground("#c8c8c8");
        if (range.getValue() == "I") range.clearContent()
      }
    })
  }

  /**
  * Remove all the I values from the attendance sheet
  */
  clearAttendanceNotInvoiced(invoiceNumber) {
    let attendanceRow = this.getInvoiceRow(invoiceNumber)
    if (attendanceRow != -1) {
      this.getAttendanceRange(attendanceRow, true, !this.invoiceCurrent(invoiceNumber)).forEach(range => {
        if (range.getValue() == "I") range.clearContent();
      })
    }
  }

  // ------------------
  // Term length invoices
  // ------------------

  /**
   * This will give you the current week number.
   * It will take the week number of the following monday.
   * If it cant be found it returns a negative one.
   * 
   * @returns week number
   */
  getWeekNumber() {
    let currentDate = new Date();
    currentDate.setHours(0,0,0,0)
    if (currentDate.getDay() == 0) {
      currentDate.setDate(currentDate.getDate() + 1);
    } else if (currentDate.getDay() != 1) {
      currentDate.setDate(currentDate.getDate() + 8 - currentDate.getDay())
    }
    let nextMondayDate = currentDate;
    let termDates = this.sheet.getRange(2, this.currentTermAttendanceColumnNum, 1, this.currentTermWeeks).getValues();
    return termDates[0].findIndex(date => date.getDate() == nextMondayDate.getDate() && date.getMonth() == date.getMonth())
  }

  /**
   * Takes the row number and returns the attendance range.
   */
  getAttendanceRange(row, asArray = false, previousTerm=false) {
    if (asArray) {
      let attendanceRanges = new Array();
      let weeks = Array(previousTerm?this.previousTermWeeks:this.currentTermWeeks).fill().map((element, index) => index)
      weeks.forEach(week => {
        attendanceRanges.push(this.sheet.getRange(row, (previousTerm?AttendanceManager.numberOfColumnsBeforeAttendanceStart():this.currentTermAttendanceColumnNum) + week))
      })
      return attendanceRanges
    } else {
      return this.sheet.getRange(row, previousTerm?AttendanceManager.numberOfColumnsBeforeAttendanceStart():this.currentTermAttendanceColumnNum, 1, previousTerm?this.previousTermWeeks:this.currentTermWeeks)

    }
  }


  prepareInvoice(row, send = false, forTerm = true, previousTerm = false) {
    let invoiceSheet = Invoices.newSheetManager(this, SpreadsheetApp.openById(this.databaseData.getVariable("Invoice Sender")).getSheetByName(this.databaseData.getVariable("Invoice Sender sheet name")));
    let ui = SpreadsheetApp.getUi();
    let activeRow = row;
    
    // --------------------------------
    // Checking for previous invoices
    // --------------------------------
    // ---- pupilName -----
    // Getting earlier as need to have this information for error message.
    let pupilName = this.sheet.getRange(activeRow, this.getColumn("Student Name")).getValue();

    // Check if invoice has already been sent.
    let invoiceNumberOfRow = this.getInvoiceNumberOfRow(activeRow, previousTerm?this.previousInvoiceColumn:undefined)

    let updating = false;
    if (invoiceNumberOfRow != "" && !this.getInvoiceRanges(invoiceNumberOfRow, previousTerm).date.isBlank()) {
      if (forTerm) {
        let answer = ui.alert("It appears you have already made and sent an invoice for " + pupilName + " for the term.\nThe new invoice you create for this pupil will override the previous one you had.\nWould you like to continue with making a new one?", ui.ButtonSet.YES_NO)
        if (answer == ui.Button.NO) {
          return
        } else {
          updating = true
        }
      } else {
        let invoiceInformation = this.getInvoiceRanges(this.getInvoiceNumberOfRow(row, previousTerm?this.previousInvoiceColumn:undefined))
        this.addTempInvoice(row, invoiceInformation.number.getValue(), !invoiceInformation.paidDate.isBlank())
        invoiceInformation.number.clear()
      }

    }

    // Check if the invoice sender is already occupied.
    if (invoiceSheet.isInvoiceLoaded()) {
      let answer = ui.alert("It appears the invoice sender already has an invoice loaded. Would you like to overide that invoice?", ui.ButtonSet.YES_NO)
      if (answer == ui.Button.NO) {
        return
      } else {
        invoiceSheet.clearInvoice();
      }
    }

    // --------------------------------
    // Collecting information for invoice
    // --------------------------------

    // ----- Get number of lessons -------

    //Find out what week to just assume all lessons will be attended
    let weekNumber = this.getWeekNumber();

    //Count lessons up until end of term simple
    console.log("Week number is: " + weekNumber);
    let lessonsToInvoice = this.getAttendanceRange(activeRow, true, previousTerm).filter((range, index) => {
      console.log("Looking at range: " + range.getValue() + " which is index: " + index);
      return (range.isBlank() || range.getValue() == "P") && (range.getBackground() != "#c8c8c8") && (index < (forTerm ? 100 : weekNumber))
    });

    let trialLessons = this.getAttendanceRange(activeRow, true, previousTerm).filter((range, index) => {
      console.log("Looking at range: " + range.getValue() + " which is index: " + index);
      return range.getValue() == "T" && (range.getBackground() != "#c8c8c8") && (index < (forTerm ? 100 : weekNumber))
    }).length
    console.log("this is how many trialLessons" + trialLessons)

    let chargedLessons = lessonsToInvoice.length

    // ---- cost of lesson -----
    let costOfLesson = this.sheet.getRange(activeRow, this.getColumn("Lesson Cost")).getValue();

    // ---- parentName -----
    let parentName = this.sheet.getRange(activeRow, this.getColumn("Guardian")).getValue();


    let email = this.sheet.getRange(activeRow, this.getColumn("Email")).getValue();

    let instrumentHire = this.sheet.getRange(activeRow, this.getColumn("Hire ")).getValue();

    let billingCompany = this.sheet.getRange(activeRow, this.getColumn("Pupils Billing Company")).getValue();

    let invoiceTerm = previousTerm? this.sheet.getRange(1,this.previousInvoiceColumn).getValue().slice(17) : this.currentTerm;

    if (!(parentName && email && billingCompany && pupilName && costOfLesson && (chargedLessons || chargedLessons  == 0) &&invoiceTerm)) {
      SpreadsheetApp.getUi().alert("Sorry the invoice for row " + row +" cannot be made as it is missing values. Please check all values and are present for the pupil.")
      this.clearAttendanceNotInvoiced();
      return;
    }

    //Mark the attendance cells that have been invoiced for
    lessonsToInvoice.forEach(range => {
      if (range.isBlank()) range.setValue("I")
    })

    // -----------------------------
    // Create and load invoice into the invoice sheet
    // -----------------------------
    let invoice = Invoices.newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, chargedLessons, trialLessons, costOfLesson, instrumentHire, billingCompany,invoiceTerm, "term");
    if (updating) {
      invoice.number = this.getInvoiceNumberOfRow(row, previousTerm?this.previousInvoiceColumn:undefined);
      let previousInvoiceInformation = this.getInvoiceRanges(invoice.number)
      invoice.note = "This invoice is an updated version of an invoice sent on " + previousInvoiceInformation.date.getValue().toLocaleString('en-NZ') + ", for $" + previousInvoiceInformation.amount.getValue() + ".";
      invoice.updated = true;
    }
    invoiceSheet.loadInvoice(invoice);

    // ----------------------------
    // Load the invoice number into the attendance sheet
    // ----------------------------
    this.sheet.getRange(activeRow, previousTerm?this.previousInvoiceColumn:this.currentInvoiceColumn).setValue(invoice.number);
    if (!updating) {
          this.sheet.getRange(activeRow,  (previousTerm?this.previousInvoiceColumn:this.currentInvoiceColumn)+1).setValue(invoice.getCosts().reduce( function(a, b){
        return a + b.price*b.quantity;
    }, 0));
    }


    SpreadsheetApp.flush(); //This ensures that the invoice is actaully loaded at this point

    // -----------------------------
    // Potentially send the invoice if needed
    // -----------------------------
    if(send) {
      invoiceSheet.sendInvoice(updating);
    }
  }

  updateSheetAfterInvoiceSent(invoiceNumber, totalCost, sentDate) {
    this.addInvoiceInfo(invoiceNumber, totalCost, sentDate);
    this.updateAttendanceToInvoiced(invoiceNumber);
  }

  clearSheetAfterClearingInvoiceSender(invoiceNumber) {
    this.clearAttendanceNotInvoiced(invoiceNumber);
    this.clearInvoiceNumber(invoiceNumber);
  }

  /**
   * This takes the terms final invoice information like sent date and amount.
   * This is to called once the invoice is sent and the information is final
   */
  addInvoiceInfo(invoiceNumber, amount, date) {
    let invoiceRanges = this.getInvoiceRanges(invoiceNumber, !this.invoiceCurrent(invoiceNumber));
    invoiceRanges.amount.setValue(amount)
    invoiceRanges.date.setValue(date)
    invoiceRanges.paidDate.clear()

    this.makeTempInvoicePermanent(this.getInvoiceRow(invoiceNumber))
  }

      /**
   * Get the pupil row for a particualr invoice number
   * @param {*} invoiceNumber Invoice number in question
   * @param {*} currentInvoiceColumn The current term invoice column in sheet
   * @returns The number of the row or -1 if it hasnt found a row
   */
    getInvoiceRow(invoiceNumber, currentInvoiceColumn = this.currentInvoiceColumn) {
      console.log("Using updated getInvoiceRow")
      let foundRowCurrent = this.sheet.getRange(3, this.currentInvoiceColumn, this.sheet.getMaxRows(), 1).createTextFinder(invoiceNumber).matchEntireCell(true).findNext();
      let foundRowPrevious = this.sheet.getRange(3, this.previousInvoiceColumn, this.sheet.getMaxRows(), 1).createTextFinder(invoiceNumber).matchEntireCell(true).findNext();
      if (foundRowCurrent == null && foundRowPrevious == null) {
        console.warn("Could not find a row for invoice number: " + invoiceNumber);
        return -1
      }
      else  {
        if (foundRowCurrent != null) return foundRowCurrent.getRowIndex()
        else return foundRowPrevious.getRowIndex()
      }
    }

  clearInvoiceNumber(invoiceNumber) {
    console.log("Trying to clear " + invoiceNumber + " from the database")
    try {
      let invoiceInfo = this.getInvoiceRanges(invoiceNumber, !this.invoiceCurrent(invoiceNumber));

      let previousInvoiceNumber = this.removeTempInvoice(this.getInvoiceRow(invoiceNumber))
      //Dont clear the number if the user was trying to update the a invoice but decided against it.
      if (invoiceInfo.date.isBlank()) {
        invoiceInfo.amount.clear({contentsOnly: true})
        invoiceInfo.number.clear({contentsOnly: true})
      }
      
      //Bring back the old invoice number if the user was going to create another non term length invoice.
      if (previousInvoiceNumber != null) {
        invoiceInfo.number.setValue(previousInvoiceNumber)
      }
    }
    catch(err) {
      console.warn("You have tried to clear a invoice number that couldnt be found.\n" + err);
    }
  }
  /**
   * This will check if the invoice number is from the current term or not.
   * You must konw that the invoice exists
   */
  invoiceCurrent(invoiceNumber) {
    let foundRowCurrent = this.sheet.getRange(3, this.currentInvoiceColumn, this.sheet.getMaxRows(), 1).createTextFinder(invoiceNumber).matchEntireCell(true).findNext();
    if (foundRowCurrent != null) return true
    else return false
  }

  // ---------------------------------------------------------------------------------------
  // ----------------------------- New pupils ----------------------------------------------
  // ---------------------------------------------------------------------------------------

  /**
   * Take all the needed variables and add a new pupil to the database.
   * @param {string} pupilName 
   * @param {string} parentName 
   * @param {string} email 
   * @param {number} costOfLesson 
   * @param {number} instrumentHire 
   * @param {string} billingCompany 
   * @param {string} tutor 
   * @param {number} phone 
   * @param {string} address 
   * @param {string} preferedDay 
   */
  addStudent(parentName, email, phone, address, pupilName, billingCompany,  preferedDay, lessonLength, costOfLesson, instrumentHire, tutor, instrument) {
    //Create the new row
    let newStudentRow = this.getInactiveRowNumber();
    this.sheet.insertRowAfter(newStudentRow-1);

    //Add in all the information
    this.sheet.getRange(newStudentRow, this.getColumn("Student Name")).setValue(pupilName);
    this.sheet.getRange(newStudentRow, this.getColumn("Guardian")).setValue(parentName);
    this.sheet.getRange(newStudentRow, this.getColumn("Email")).setValue(email);
    this.sheet.getRange(newStudentRow, this.getColumn("Lesson Cost")).setValue(costOfLesson);
    this.sheet.getRange(newStudentRow, this.getColumn("Hire cost")).setValue(instrumentHire);
    this.sheet.getRange(newStudentRow, this.getColumn("Pupils Billing Company")).setValue(billingCompany);
    this.sheet.getRange(newStudentRow, this.getColumn("Duration")).setValue(lessonLength);
    this.sheet.getRange(newStudentRow, this.getColumn("Teacher")).setValue(tutor);
    this.sheet.getRange(newStudentRow, this.getColumn("Phone")).setValue(phone);
    this.sheet.getRange(newStudentRow, this.getColumn("Suburb/Address")).setValue(address);
    this.sheet.getRange(newStudentRow, this.getColumn("Day")).setValue(preferedDay);
    this.sheet.getRange(newStudentRow, this.getColumn("Instrument")).setValue(instrument);
    
    //Set the status to active
    this.sheet.getRange(newStudentRow, this.getColumn("Status")).setValue("Active");

    //X out all of the weeks up to and including this week.
    let currentWeekNumber = this.getWeekNumber();
    if (currentWeekNumber < 1) {
      console.warn("Current week number is below 1: " + currentWeekNumber)
    }
    
    for (let i = 0; i < currentWeekNumber; i++) {
      this.sheet.getRange(newStudentRow, this.currentTermAttendanceColumnNum+i).setValue("X");
    }

    //Sort the sheet afterwards so the new pupil will be in the correct place
    this.clean();
  }

  confirmTrialLesson(row) {
    let attendanceRow = this.sheet.getRange(row, this.currentTermAttendanceColumnNum, 1, this.currentTermWeeks).getValues()[0];

    // Check if the pupil has a trial lesson booked.
    if (!attendanceRow.includes("T")) {
      throw "This pupil does not have a trial lesson booked."
    }

    let trialLessonWeek = attendanceRow.indexOf("T") + 1;
    let trialLessonDate = this.sheet.getRange(2, this.currentTermAttendanceColumnNum + trialLessonWeek - 1).getValue();

    let templateSheet = this.sheet.getParent().getSheetByName("Trial Confirmation template");
    let emailer = Emails.newEmailer(
      templateSheet.getRange(1,2).getValue(),
      templateSheet.getRange(2,2).getValue());

    emailer.sendEmail([this.sheet.getRange(row, this.getColumn("Email")).getValue()], {
      Parent_name: this.sheet.getRange(row, this.getColumn("Guardian")).getValue(),
      Student_name: this.sheet.getRange(row, this.getColumn("Student Name")).getValue(),
      Lesson_date: trialLessonDate.toLocaleDateString("en-NZ"),
      Tutor_name: this.sheet.getRange(row, this.getColumn("Teacher")).getValue()
    }
    );

    this.sheet.getParent().toast("Confirmation sent");
  }
}

function newAttendanceSheet(attendanceSheet) {
  return new AttendanceManager(attendanceSheet)
}

function TermAttendanceSheetName() {
  return AttendanceManager.sheetName();
}