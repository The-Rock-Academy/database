class AttendanceManager extends DatabaseSheetManager {
  static sheetName(){return "Master sheet"}
  static numberOfColumnsBeforeAttendanceStart() {return 2}

  constructor(sheet, currentTerm) {
    super(sheet, currentTerm);

    this.currentTermAttendanceColumnNum = this.getColumn("Current " + this.currentTerm);
    this.currentTermWeeks = this.getCurrentTermWeeks();
    this.currentInvoiceColumn = this.getColumn("Current Invoice " + this.currentTerm, true)
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
    let columnOfNextTerm = this.currentTermAttendanceColumnNum+this.currentTermWeeks
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

    this.resetInvoiceColumns(nextTerm);

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
    let comments = this.sheet.getRange(row, this.getColumn("Current term comments")).getValue();
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
    let comments = this.sheet.getRange(row, this.getColumn("Current term comments"));
    
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
    this.getAttendanceRange(this.getInvoiceRow(invoiceNumber), true).forEach(range => {
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
      this.getAttendanceRange(attendanceRow, true).forEach(range => {
        if (range.getValue() == "I") range.clearContent();
      })
    }
  }

  // ------------------
  // Term length invoices
  // ------------------
  /**
   * Takes the row number and returns the attendance range.
   */
  getAttendanceRange(row, asArray = false) {
    if (asArray) {
      let attendanceRanges = new Array();
      let weeks = Array(this.currentTermWeeks).fill().map((element, index) => index)
      weeks.forEach(week => {
        attendanceRanges.push(this.sheet.getRange(row, this.currentTermAttendanceColumnNum + week))
      })
      return attendanceRanges
    } else {
      return this.sheet.getRange(row, this.currentTermAttendanceColumnNum, 1, this.currentTermWeeks)

    }
  }


  prepareInvoice(row, send = false, forTerm = true) {
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
    let currentDate = new Date();
    currentDate.setHours(0,0,0,0)
    if (currentDate.getDay() == 0) {
      currentDate.setDate(currentDate.getDate() + 1);
    } else if (currentDate.getDay() != 1) {
      currentDate.setDate(currentDate.getDate() + 8 - currentDate.getDay())
    }
    let nextMondayDate = currentDate;
    let termDates = this.sheet.getRange(2, this.currentTermAttendanceColumnNum, 1, this.currentTermWeeks).getValues();
    let weekNumber = termDates[0].findIndex(date => date.getTime() == nextMondayDate.getTime())


    //Count lessons up until end of term simple
    console.log("Week number is: " + weekNumber);
    let lessonsToInvoice = this.getAttendanceRange(activeRow, true).filter((range, index) => {
      console.log("Looking at range: " + range.getValue() + " which is index: " + index);
      return (range.isBlank() || range.getValue() == "P") && (range.getBackground() != "#c8c8c8") && (index < (forTerm ? 100 : weekNumber))
    });

    let trialLessons = this.getAttendanceRange(activeRow, true).filter((range, index) => {
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

    if (!(parentName && email && billingCompany && pupilName && costOfLesson && (chargedLessons || chargedLessons  == 0) &&this.currentTerm)) {
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
    let invoice = Invoices.newInvoice(this.databaseData.getVariable("Invoice Folder"), parentName, pupilName, email, chargedLessons, trialLessons, costOfLesson, instrumentHire, billingCompany,this.currentTerm, "term");
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
    this.sheet.getRange(activeRow, this.getColumn("Current Invoice " +this.currentTerm)).setValue(invoice.number);
    if (!updating) {
          this.sheet.getRange(activeRow, this.getColumn("Current Invoice " +this.currentTerm)+1).setValue(invoice.getCosts().reduce( function(a, b){
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
    let invoiceRanges = this.getInvoiceRanges(invoiceNumber);
    invoiceRanges.amount.setValue(amount)
    invoiceRanges.date.setValue(date)
    invoiceRanges.paidDate.clear()

    this.makeTempInvoicePermanent(this.getInvoiceRow(invoiceNumber))
  }

  clearInvoiceNumber(invoiceNumber) {
    console.log("Trying to clear " + invoiceNumber + " from the database")
    try {
      let invoiceInfo = this.getInvoiceRanges(invoiceNumber);

      let previousInvoiceNumber = this.removeTempInvoice(this.getInvoiceRow(invoiceNumber))
      //Dont clear the number if the user was trying to update the a invoice but decided against it.
      if (invoiceInfo.date.isBlank()) {
        console.log("Testing testing")
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

}

function newAttendanceSheet(attendanceSheet) {
  return new AttendanceManager(attendanceSheet)
}

function TermAttendanceSheetName() {
  return "Master sheet"
}