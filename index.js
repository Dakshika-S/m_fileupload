const express = require("express");
const app = express();
const port = 7000;
const mysql = require("mysql2");
const xlsx = require("xlsx");
const multer = require("multer");
const path = require("path");

//app.use(express.json());

//db connection
// const connection = mysql.createConnection({
//   user: "root",
//   host: "localhost",
//   password: "Dak@1999",
//   database: "my_db",
//   insecureAuth: true,
// });
// connection.connect(function (err) {
//   if (err) {
//     console.log("error in connection : " + err.message);
//   } else {
//     console.log("connected");
//   }
// });
const pool = mysql.createPool({
  host: "localhost",
  user: "root",
  password: "Dak@1999",
  database: "my_db",
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
});
app.use(express.json());
const promisePool = pool.promise();

const connectDatabase = () => {
  pool.getConnection((err, Connection) => {
    if (err) {
      console.error("error connecting to Mysql:" + err.message);
      return;
    }
    console.log("connected");
    Connection.release();
  });
};
connectDatabase();

//identifying file type -file filter
const excelFilter = (req, file, cb) => {
  if (
    file.mimetype.includes("excel") ||
    file.mimetype.includes("spreadsheetml") ||
    file.mimetype.includes(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
  ) {
    cb(null, true);
  } else {
    cb("please upload only excel file.", false);
  }
  // let filetypes = /xlsx|xls/;
  // let mimetype = filetypes.test(file.mimetype);
  // let extname = filetypes.test(path.extname(file.originalname));

  // if (mimetype && extname) {
  //   return cb(null, true);
  // }
  // return cb("Error: file upload only :" + filetypes);
};

//providing storge space and assignaing uniq name
var storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, "/uploads"));
  },
  filename: (req, file, cb) => {
    console.log(file.originalname);
    cb(null, `${Date.now()}-${file.originalname}`);
  },
});

//creating a multer instance
var uploadFile = multer({
  storage: storage, // specifiying above created storageengine(storage) to storage
  fileFilter: excelFilter,
}); //give function

app.post("/uploadfile", uploadFile.single("file"), async (req, res) => {
  // console.log("req.body.file is : " + req.body.file);
  // console.log("req.body.path is : " + req.body.path);
  console.log("req.file is : " + req.file);
  console.log("req.file.path is : " + req.file.path);
  try {
    if (!req.file && req.file.path) {
      // if (!req.body.file || !req.body.path) {
      throw new Error("file not provided");
    }

    //xls file reading
    //let workbook = xlsx.readFile(req.file.path); //reading existing book
    let workbook = xlsx.readFile(req.file.path); //reading existing book
    let worksheet = workbook.Sheets[workbook.SheetNames[0]]; // selecting 1st worksheet
    let range = xlsx.utils.decode_range(worksheet["!ref"]);

    let queryPromises = []; //new line
    for (let row = range.s.r; row <= range.e.r; row++) {
      let data = [];

      for (let col = range.s.c; col <= range.e.c; col++) {
        let cell = worksheet[xlsx.utils.encode_cell({ r: row, c: col })];
        data.push(cell.v);
      }
      console.log(data);

      //insert query
      let sql =
        "INSERT INTO `users` (`Name`, `Email`, `Address`, `Salary`) VALUES(?, ?, ?, ?)";

      pool.query(sql, data, (err, results, fields) => {
        if (err) {
          return console.error(err.message);
        }
        console.log("user ID: " + results.insertId);
      });
    }
    res.status(200).json({ message: "success" });
  } catch (err) {
    console.error("error" + err.message);
  }
  res.status(200).json({ message: "success" });
});

app.get("/", (req, res) => {
  res.send("hello !!!!");
});

app.listen(port, () => {
  console.log(`app listening on port ${port}`);
});
