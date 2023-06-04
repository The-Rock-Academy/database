//Adding in a comment
function test() {
  let database = new Database(SpreadsheetApp.getActiveSpreadsheet())
  database.archive();
}

function helloWorld() {
  console.log("Hello World");
}


const DEBUG = true;

function debug(message) {
  if (DEBUG) {
    console.log(message)
  }
}
