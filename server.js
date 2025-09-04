const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const archiver = require("archiver");

const app = express();
const PORT = process.env.PORT || 3000;

// ensure uploads folder exists
const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);

// parse form bodies (for admin login)
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// serve static files from root (index.html)
app.use(express.static(__dirname));

// Excel file path
const EXCEL_FILE = path.join(UPLOAD_DIR, "data.xlsx");

// Multer storage: filename = rollno + original extension
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const rollno = (req.body.rollno || "").trim();
    const ext = path.extname(file.originalname) || "";
    const finalName = `${rollno}${ext}`;
    cb(null, finalName);
  }
});
const upload = multer({ storage });

// -------------------- Upload route --------------------
app.post("/upload", (req, res) => {
  upload.single("codefile")(req, res, (err) => {
    if (err) return res.send("❌ Upload error: " + err.message);

    if (!req.file) return res.send("❌ No file uploaded");

    const fileLink = `/uploads/${req.file.filename}`;
    const newEntry = {
      Name: req.body.name || "",
      RollNo: req.body.rollno || "",
      FileLink: fileLink
    };

    // read existing excel rows
    let data = [];
    if (fs.existsSync(EXCEL_FILE)) {
      try {
        const workbook = XLSX.readFile(EXCEL_FILE);
        const sheet = workbook.Sheets["Sheet1"];
        if (sheet) data = XLSX.utils.sheet_to_json(sheet);
      } catch (e) {
        console.error("Error reading excel:", e);
      }
    }

    // if a previous row for same roll exists, remove it (we overwrite file)
    data = data.filter(r => String(r.RollNo) !== String(newEntry.RollNo));
    data.push(newEntry);

    const worksheet = XLSX.utils.json_to_sheet(data);
