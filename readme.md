This is the main code the attendance and invoicing system.

By using clasp this is deployed to Google Apps Script library and then used in each of the spreadsheets.

# Message for freelancers

Note that this project has been poorly documented and cleaned. The systems works stably (error every few months) but the code is messy.
The main cause of messiness is that the code has evolved over many years without adequate refactoring and pruning of unused features.

Over the next week (by early Feb 2026) I will add some proper overview documentation to provide a basic guide for you.

# Guide to TRA systems

This may not be the best place to document this, but for now this is a basic guide to the systems that run for TRA to provide invoicing making, school holiday booking and attendance database management. To actually understand the system you will need to read through the code.

This system has been built up over a few years by [James Thompson](https://github.com/1jamesthompson1), he should be contacted for any questions about the system. Reachable at tech followed by the therockacademy.co.nz email domain.


## Intro to TRA

The rock academy (TRA) is a decentralized music school based in Wellington, New Zealand. TRA provides mobile music lessons (teachers go to students homes) as well as running band school sessions and school holiday programmes. All of this is managed through several Google Sheets, Forms and Files.

## Google Drive documents

The entire system is made up of several Google Sheets spreadsheets, a google form and this Apps Script library. Each spreadsheet has a small amount of code that is mainly defining functions to call the library code.

### Main Spreadsheets

These spreadsheets are "reset" once a school term (about 10 weeks) to start a new term. The resetting involves creating a copy (stored in the archive folder), clearing out old data and updating term dates.

| Spreadsheet name | Purpose |
|------------------|---------|
| Term Database | - Weekly mobile lessons attendance + invoicing<br>- School holiday programme bookings + invoicing<br>New pupil inquiry handling |
| Band School Database | - Band school attendance (updated by tutors) |
| Band School Database Invoicing | - Band school invoicing |

_You can find links to these spreadsheets from the "Term Database" spreadsheet. 'Data' tab_

### Auxiliary Spreadsheets

These spreadsheets are not reset each term, they contain data that is used across terms.

| Spreadsheet name | Purpose |
|------------------|---------|
| Email Templates | Email templates (powered by [this](https://github.com/The-Rock-Academy/emailTemplating)) used for invoicing, reminders, new pupils etc |
| Invoice Builder (spreadsheet) | A place where the invoice PDF generation happens. |

## Google Form

There is a Google Form that is used to collect school holiday programme bookings from parents. The form responses are stored in the "Term Database" spreadsheet. This form is embedded into the [TRA website](https://www.therockacademy.co.nz/school-holiday-programme-booking). There are two forms where the second form is sometimes used for when two weeks of holiday programme is running (normally over the Summer break).

## Code structure

Most of the code is in JavaScript however some of it is in TypeScript.

There are three parts to the code:
- [Invoices](./Invoices/): Code that handles the invoice building, sending and archiving.
- [SheetManagers](./SheetManagers/): Code that handles specific features of each spreadsheet tab (e.g SHP bookings, mobile lessons attendance etc). The main features for each tab are invoicing and resetting.
-  [Root level code](./): Code that handles the "Main database" as well as miscellaneous utilities used across the codebase.

# Contributing guide

If you are a developer who has been asked to contribute to this codebase, please reach out to James Thompson (tech followed by the therockacademy.co.nz email domain) for an introduction to the code and guidance on how to contribute.

## Deployment workflow

Each of the spreadsheets pulls this code from the Google Apps Script library. The spreadsheets uses a specific version of the library.
When a commit is made to the main branch of this repository, a new version of the library is creates [see workflow](./.github/workflows/clas-deploy.yml). This new version is not automatically used by the spreadsheets, the version used by each spreadsheet needs to be manually updated. Note the workflow does say it fails however this is a red herring as a new version is infact deployed.

## Development workflow

I use the clasp cli app with the command `clasp push -w` when developing. This means that the database library head is updated with the latest version of the code. Then for testing I have the spreadsheet pointed to the head version of the library.

There is a development folder which has some copies of the spreadsheet that are used for testing.