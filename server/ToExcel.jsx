// required modules in here: npm install mysql xlsx
const xlsx = require("xlsx");
 
// connecting to database
function indertTicketsToDB(file, db)
{ 
  // open the spreadsheet file
  const workbook = xlsx.readFile("ticket_sheets.xlsx"),
    worksheet = workbook.Sheets[workbook.SheetNames[0]],
      range = xlsx.utils.decode_range(worksheet["!ref"]);
  
  // importing the spreadsheet
  for (let row=range.s.r; row<=range.e.r; row++) 
  {
    // reading the cells
    let data = [];
    for (let col=range.s.c; col<=range.e.c; col++) 
    {
      let cell = worksheet[xlsx.utils.encode_cell({r:row, c:col})];
      data.push(cell.v);
    }
    
    // inserting into the database
    let sql = "INSERT INTO `tickets` (`ticketCode`, `is_issued`, `issued_on`, `issued_to`) VALUES (?,?,?,?)";
    db.query(sql, data, (err, results, fields) => 
    {
      if (err) { return console.error(err.message);}
      console.log("USER ID:" + results.insertId);
    });
  }
}

exports.insertTicketsToDB = insertTicketsToDB;
