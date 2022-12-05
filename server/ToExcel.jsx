// required modules in here: npm install mysql xlsx
const mysql = require("mysql"),
      xlsx = require("xlsx");
 
// connecting to database
const db = mysql.createConnection({});
 
// open the spreadsheet file
var workbook = xlsx.readFile("ticketsdata.xlsx"),
    worksheet = workbook.Sheets[workbook.SheetNames[0]],
    range = xlsx.utils.decode_range(worksheet["!ref"]);
 
// importing the spreadsheet
for (let row=range.s.r; row<=range.e.r; row++) {
  // reading the cells
  let data = [];
  for (let col=range.s.c; col<=range.e.c; col++) {
    let cell = worksheet[xlsx.utils.encode_cell({r:row, c:col})];
    data.push(cell.v);
  }
 
  // inserting into the database
  let sql = "INSERT INTO `tickets` (`ticketCode`, `is_iissued`, `issued_on`, `issued_to`) VALUES (?,?,?,?)";
  db.query(sql, data, (err, results, fields) => {
    if (err) { return console.error(err.message); }
    console.log("USER ID:" + results.insertId);
  });
}

db.end();




