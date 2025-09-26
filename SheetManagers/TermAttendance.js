class AttendanceManager extends AttendanceSheetManager {
  static sheetName(){return "Master sheet"}

  static getObjFromSS(SS) {
    return new AttendanceManager(SS.getSheetByName(AttendanceManager.sheetName()))
  }

  constructor(sheet, currentTerm) {
    super(sheet, currentTerm);
  }

  reset(nextTermDetails, nextTerm) {
    super.reset(nextTermDetails, nextTerm)

    // Reset the no of lessons and formula column

    let noLessons = this.getColumn("No. Lessons");

    this.sheet.getRange(3, noLessons, this.getInactiveRowNumber()-2, 1).clearContent();

    // Reset the formula column
    function getColumnLetter(columnNumber) {
        let letter = '';
        while (columnNumber > 0) {
            let modulo = (columnNumber - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            columnNumber = Math.floor((columnNumber - modulo) / 26);
        }
        return letter;
    }

    let billableAmount = this.getColumn("Billable Amount");
    let lessonCost = this.getColumn("Lesson Cost");

    // Convert column numbers to letters
    let billableAmountLetter = getColumnLetter(billableAmount);
    let lessonCostLetter = getColumnLetter(lessonCost);
    let noLessonsLetter = getColumnLetter(noLessons);

    // Set the formula
    let lastRow = this.getInactiveRowNumber() - 2;

    // Set the formula for each cell in the range
    for (let row = 3; row <= lastRow + 2; row++) {
        let formula = `= ${noLessonsLetter}${row} * ${lessonCostLetter}${row}`;
        this.sheet.getRange(row, billableAmount).setFormula(formula);
    }
  }

  clean() {
    // Get things that will be needed throughout the clean.
    // Some of these might need to be refactoed to be part of the object.
    let statusColumnNumber = this.getColumn("Status");
    let statusColumn = this.sheet.getRange(3, statusColumnNumber, this.getInactiveRowNumber()-2, 1);

    //Sort the active section by the status type
    // To do this I need to get the active range and then sort it using the sort function
    let activeRange = this.sheet.getRange(3, 1, this.getInactiveRowNumber()-3, this.sheet.getMaxColumns());

    activeRange.sort({column: this.getColumn("Teacher"), ascending: true})
  }

  getInactiveRowNumber() {
    return this.sheet.getLastRow();

  }

  // -------------------------------------------------------------------------------------------------------
  // --------- Invoices -----------------------------------------------------------------------------------
  // -------------------------------------------------------------------------------------------------------

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

  prepareInvoice(row, send = false, forTerm = true) {
    let invoiceSheet = newSheetManager(this, SpreadsheetApp.openById(this.databaseData.getVariable("Invoice Sender")).getSheetByName(this.databaseData.getVariable("Invoice Sender sheet name")));
    let ui = SpreadsheetApp.getUi();
    let activeRow = row;
    
    // --------------------------------
    // Checking for previous invoices
    // --------------------------------
    // ---- pupilName -----
    // Getting earlier as need to have this information for error message.
    let pupilName = this.sheet.getRange(activeRow, this.getColumn("Student Name")).getValue();

    // Check if invoice has already been sent.
    let invoiceNumberOfRow = this.getInvoiceNumberOfRow(activeRow)

    let updating = false;
    if (invoiceNumberOfRow != "" && !this.getInvoiceRanges(invoiceNumberOfRow).date.isBlank()) {
      if (forTerm) {
        let answer = ui.alert("It appears you have already made and sent an invoice for " + pupilName + " for the term.\nThe new invoice you create for this pupil will override the previous one you had.\nWould you like to continue with making a new one?", ui.ButtonSet.YES_NO)
        if (answer == ui.Button.NO) {
          return
        } else {
          updating = true
        }
      } else {
        let invoiceInformation = this.getInvoiceRanges(this.getInvoiceNumberOfRow(row))
        this.addTempInvoice(row, invoiceInformation.number.getValue(), !invoiceInformation.paidDate.isBlank())
        invoiceInformation.number.clear()
      }

    }

    // Check if the invoice sender is already occupied.
    if (invoiceSheet.isInvoiceLoaded()) {
      invoiceSheet.clearInvoice();
      this.sheet.getParent().toast("Invoice in the sender has been cleared");
    }

    // --------------------------------
    // Collecting information for invoice
    // --------------------------------

    // ---- parentName -----
    let parentName = this.sheet.getRange(activeRow, this.getColumn("Guardian")).getValue();

    let email = this.sheet.getRange(activeRow, this.getColumn("Email")).getValue();


    let billingCompany = this.sheet.getRange(activeRow, this.getColumn("Pupils Billing Company")).getValue();

    let invoiceTerm = this.currentTerm;

    let numberOfLessons = this.sheet.getRange(activeRow, this.getColumn("No. Lessons")).getValue();
    let costOfLesson = this.sheet.getRange(activeRow, this.getColumn("Lesson Cost")).getValue();

    // Check to make sure that all of the values are not empty
    console.log("Checking for empty values")

    let mandatoryVales = {
      "Parent Name": parentName,
      "Email": email,
      "Billing Company": billingCompany,
      "Pupil Name": pupilName,
      "Invoice Term": invoiceTerm,
      "Number of Lessons": numberOfLessons,
      "Cost of Lesson": !this.sheet.getRange(activeRow, this.getColumn("Lesson Cost")).isBlank(),
    }

    for (const [key, value] of Object.entries(mandatoryVales)) {
      if (!value) {
        ui.alert("Sorry the invoice for row " + row +" cannot be made as it is missing values. Please check all values and are present for the pupil. The value missing is: " + key + " Which is actaully '" + value + "'")
        return;
      }
    }

    let expectedTotalCost = this.sheet.getRange(activeRow, this.getColumn("Billable Amount")).getValue();

    if (numberOfLessons * costOfLesson != expectedTotalCost) {
      ui.alert("The total cost of the invoice does not match the expected total cost.\nExpected: $" + expectedTotalCost + "\nCalculated: $" + (numberOfLessons * costOfLesson) + "\nPlease check columns v - w and make sure that everything is correct and try again", ui.ButtonSet.OK)
      return
    }

    // -----------------------------
    // Create and load invoice into the invoice sheet
    // -----------------------------
    let invoice = newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, numberOfLessons, costOfLesson, billingCompany,invoiceTerm, "term");
    if (updating) {
      invoice.number = this.getInvoiceNumberOfRow(row);
      let previousInvoiceInformation = this.getInvoiceRanges(invoice.number)
      invoice.note = "This invoice is an updated version of an invoice sent on " + previousInvoiceInformation.date.getValue().toLocaleString('en-NZ') + ", for $" + previousInvoiceInformation.amount.getValue() + ".";
      invoice.updated = true;
    }
    invoiceSheet.loadInvoice(invoice);

    // ----------------------------
    // Load the invoice number into the attendance sheet
    // ----------------------------
    this.sheet.getRange(activeRow, this.currentInvoiceColumn).setValue(invoice.number);
    if (!updating) {
          this.sheet.getRange(activeRow,  (this.currentInvoiceColumn)+1).setValue(invoice.getCosts().reduce( function(a, b){
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

  updateSheetAfterInvoiceSent(invoiceNumber, totalCost, sentDate, numberOfLessons) {
    console.log(`Updating sheet after invoice ${invoiceNumber} is sent with cost: ${totalCost} and number of lessons ${numberOfLessons}` )
    this.addInvoiceInfo(invoiceNumber, totalCost, sentDate);
  }

  clearSheetAfterClearingInvoiceSender(invoiceNumber) {
    this.clearInvoiceNumber(invoiceNumber);
  }

  /**
   * This takes the terms final invoice information like sent date and amount.
   * This is to called once the invoice is sent and the information is final
   */
  addInvoiceInfo(invoiceNumber, amount, date) {
    let invoiceRanges = this.getInvoiceRanges(invoiceNumber);
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
      if (foundRowCurrent == null) {
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
      let invoiceInfo = this.getInvoiceRanges(invoiceNumber);

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

  // ---------------------------------------------------------------------------------------
  // ----------------------------- New pupils ----------------------------------------------
  // ---------------------------------------------------------------------------------------

  /**
   * Take all the needed variables and add a new pupil to the database.
   * @param {string} pupilName 
   * @param {string} parentName 
   * @param {string} email 
   * @param  costOfLesson 
   * @param {string} billingCompany 
   * @param {string} tutor 
   * @param {number} phone 
   * @param {string} address 
   * @param {string} preferedDay 
   */
  addStudent(parentName, email, phone, address, pupilName, billingCompany,  preferedDay, lessonLength, costOfLesson, tutor, instrument) {
    //Create the new row
    let newStudentRow = this.getInactiveRowNumber();
    this.sheet.insertRowAfter(newStudentRow-1);

    //Add in all the information
    this.sheet.getRange(newStudentRow, this.getColumn("Student Name")).setValue(pupilName);
    this.sheet.getRange(newStudentRow, this.getColumn("Guardian")).setValue(parentName);
    this.sheet.getRange(newStudentRow, this.getColumn("Email")).setValue(email);
    this.sheet.getRange(newStudentRow, this.getColumn("Lesson Cost")).setValue(costOfLesson);
    this.sheet.getRange(newStudentRow, this.getColumn("Pupils Billing Company")).setValue(billingCompany);
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

    let templateSS = (new DatabaseData(this.sheet.getParent())).getTemplateSS();
    let emailer = Emails.newEmailer(templateSS, "Trail lesson confirmation");

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