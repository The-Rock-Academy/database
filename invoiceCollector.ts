class InvoiceCollector {
    sent: number;
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    paid: number;
    static templateName: string = "Invoice Reminder";
    templateSS: GoogleAppsScript.Spreadsheet.Spreadsheet;

    constructor(sheet, sent, paid, parent_name, parent_email, pupil_name, invoice_number, invoice_amount, reminder_date) {
        this.sheet = sheet;
        this.sent = sent;
        this.paid = paid;
        this.templateSS= (new DatabaseData(this.sheet.getParent())).getTemplateSS();
        this.parent_name = parent_name;
        this.parent_email = parent_email;
        this.pupil_name = pupil_name;
        this.invoice_number = invoice_number;
        this.invoice_amount = invoice_amount;
        this.reminder_date = reminder_date;
    }
    
    sendReminders(range) {
        let startingRow = range.getRowIndex();
        let endingRow = range.getHeight() + range.getRowIndex();

        let sentDate = this.sheet.getRange(startingRow,this.sent, endingRow-startingRow, 1).getValues();
        let paidDate = this.sheet.getRange(startingRow,this.paid, endingRow-startingRow, 1).getValues();

        sentDate.forEach((element,index) => {
            if (element[0] != "" && paidDate[index][0] == "") {
                let invoiceNumber = this.sheet.getRange(startingRow+index,this.invoice_number).getValue();

                let invoiceFolder = newInvoiceFolder((new DatabaseData(this.sheet.getParent())).getVariable("Invoice Folder"))
                if (invoiceFolder.invoiceExists(invoiceNumber)) {
                    let invoicePDF = invoiceFolder.getInvoice(invoiceNumber);
                    
                    let emailer = Emails.newEmailer(this.templateSS, InvoiceCollector.templateName);
                    let data = this.sheet.getRange(startingRow+index,1,1,this.sheet.getLastColumn()).getValues()[0];
                    emailer.sendEmail([data[this.parent_email-1]], {
                        "parentName": data[this.parent_name-1],
                        "studentName": data[this.pupil_name-1],
                        "invoicePrice": data[this.invoice_amount-1],
                        "invoiceNumber": data[this.invoice_number-1],
                        "sentDate": (new Date(data[this.sent-1])).toLocaleDateString("en-NZ")
                    },[invoicePDF]);

                    this.sheet.getRange(startingRow+index,this.reminder_date).setValue((new Date()).toLocaleDateString("en-NZ"));
                } else {
                    console.warn("Invoice number: " + invoiceNumber + " does not exist in the invoice folder. Skipping.");
                }
            }
        });
    }
}