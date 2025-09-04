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
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, EXCEL_FILE);

    res.send(`
      ✅ File uploaded successfully!<br>
      Name: ${escapeHtml(newEntry.Name)}<br>
      Roll No: ${escapeHtml(newEntry.RollNo)}<br>
      <a href="${fileLink}" download>⬇ Download Your File</a>
    `);
  });
});

// -------------------- Check duplicate route --------------------
app.get("/check-file", (req, res) => {
  const rollno = (req.query.rollno || "").trim();
  if (!rollno) return res.json({ exists: false });
  try {
    const files = fs.readdirSync(UPLOAD_DIR);
    const exists = files.some(f => f.startsWith(rollno));
    return res.json({ exists: !!exists });
  } catch (e) {
    console.error(e);
    return res.json({ exists: false });
  }
});

// -------------------- Admin simple login --------------------
const ADMIN_PASSWORD = "admin123"; // change this before production

app.get("/admin", (req, res) => {
  res.send(`
    <h2>Admin Login</h2>
    <form method="post" action="/admin">
      <input type="password" name="password" placeholder="Enter Admin Password" required>
      <button type="submit">Login</button>
    </form>
  `);
});

app.post("/admin", (req, res) => {
  const password = req.body.password;
  if (password === ADMIN_PASSWORD) {
    res.send(`
      <h2>📊 Admin Panel</h2>
      <a href="/admin/excel" download>⬇ Download Excel Sheet</a><br><br>
      <a href="/admin/files" download>⬇ Download All Uploaded Files (ZIP)</a><br><br>
      <a href="/">Back to Upload Page</a>
    `);
  } else {
    res.send("❌ Wrong password!");
  }
});

// -------------------- Admin download excel --------------------
app.get("/admin/excel", (req, res) => {
  if (fs.existsSync(EXCEL_FILE)) {
    return res.download(EXCEL_FILE);
  } else {
    return res.send("📂 No data found yet!");
  }
});

// -------------------- Admin download all files as ZIP --------------------
app.get("/admin/files", (req, res) => {
  // create a fresh zip each time
  const zipPath = path.join(UPLOAD_DIR, "all_files.zip");

  // remove old zip if exists
  if (fs.existsSync(zipPath)) {
    try { fs.unlinkSync(zipPath); } catch (e) { /* ignore */ }
  }

  const output = fs.createWriteStream(zipPath);
  const archive = archiver("zip", { zlib: { level: 9 } });

  output.on("close", () => {
    res.download(zipPath, "all_uploaded_files.zip", (err) => {
      if (err) console.error("Error sending zip:", err);
    });
  });

  archive.on("error", (err) => {
    console.error("Archive error:", err);
    res.status(500).send("❌ Error creating ZIP");
  });

  archive.pipe(output);

  // Add files from UPLOAD_DIR (but exclude the ZIP itself)
  fs.readdirSync(UPLOAD_DIR).forEach(file => {
    if (file === "all_files.zip") return;
    const filePath = path.join(UPLOAD_DIR, file);
    if (fs.statSync(filePath).isFile()) {
      archive.file(filePath, { name: file });
    }
  });

  archive.finalize();
});

// serve uploaded files
app.use("/uploads", express.static(UPLOAD_DIR));

// simple escape for HTML output
function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));
