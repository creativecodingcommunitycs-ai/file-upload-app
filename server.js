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

// status.json to track submission ON/OFF
const STATUS_FILE = path.join(UPLOAD_DIR, "status.json");
function getStatus() {
  if (!fs.existsSync(STATUS_FILE)) return { acceptingSubmissions: true };
  return JSON.parse(fs.readFileSync(STATUS_FILE, "utf8"));
}
function setStatus(val) {
  fs.writeFileSync(STATUS_FILE, JSON.stringify({ acceptingSubmissions: val }));
}

// parse form bodies
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// serve static files
app.use(express.static(__dirname));

// Excel file path
const EXCEL_FILE = path.join(UPLOAD_DIR, "data.xlsx");

// Multer storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const rollno = (req.body.rollno || "").trim();
    const ext = path.extname(file.originalname) || "";
    cb(null, `${rollno}${ext}`);
  }
});
const upload = multer({ storage });

// -------------------- Upload route --------------------
app.post("/upload", (req, res) => {
  const status = getStatus();
  if (!status.acceptingSubmissions) return res.send("❌ Submissions are CLOSED by Admin.");

  upload.single("codefile")(req, res, (err) => {
    if (err) return res.send("❌ Upload error: " + err.message);
    if (!req.file) return res.send("❌ No file uploaded");

    const fileLink = `/uploads/${req.file.filename}`;
    const newEntry = {
      Name: req.body.name || "",
      RollNo: req.body.rollno || "",
      Batch: req.body.batch || "",
      FileLink: fileLink,
      DateTime: new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" })
    };

    let data = [];
    if (fs.existsSync(EXCEL_FILE)) {
      const workbook = XLSX.readFile(EXCEL_FILE);
      const sheet = workbook.Sheets["Sheet1"];
      if (sheet) data = XLSX.utils.sheet_to_json(sheet);
    }

    // overwrite old entry if RollNo exists
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
      Batch: ${escapeHtml(newEntry.Batch)}<br>
      Submitted At: ${newEntry.DateTime}<br>
      <a href="${fileLink}" download>⬇ Download Your File</a>
    `);
  });
});

// -------------------- Check duplicate --------------------
app.get("/check-file", (req, res) => {
  const rollno = (req.query.rollno || "").trim();
  if (!rollno) return res.json({ exists: false });
  try {
    const files = fs.readdirSync(UPLOAD_DIR);
    const exists = files.some(f => f.startsWith(rollno));
    return res.json({ exists: !!exists });
  } catch {
    return res.json({ exists: false });
  }
});

// -------------------- Admin login --------------------
const ADMIN_PASSWORD = "admin123";
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
  if (req.body.password === ADMIN_PASSWORD) return res.redirect("/admin/dashboard");
  res.send("❌ Wrong password!");
});

// -------------------- Admin dashboard --------------------
app.get("/admin/dashboard", (req, res) => {
  const status = getStatus();
  let submissions = [];

  if (fs.existsSync(EXCEL_FILE)) {
    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheet = workbook.Sheets["Sheet1"];
    if (sheet) submissions = XLSX.utils.sheet_to_json(sheet);
  }

  const total = submissions.length;

  let rows = submissions.map(r => `
    <tr>
      <td>${escapeHtml(r.Name)}</td>
      <td>${escapeHtml(r.RollNo)}</td>
      <td>${escapeHtml(r.Batch)}</td>
      <td><a href="${r.FileLink}" target="_blank">View File</a></td>
      <td>${escapeHtml(r.DateTime)}</td>
      <td><a class="delete-btn" href="/admin/delete?rollno=${encodeURIComponent(r.RollNo)}" onclick="return confirm('Delete this submission?')">🗑 Delete</a></td>
    </tr>
  `).join("");

  // Chart: Batch-wise submissions
  const countByBatch = {};
  submissions.forEach(s => {
    if (!countByBatch[s.Batch]) countByBatch[s.Batch] = 0;
    countByBatch[s.Batch]++;
  });
  const batchLabels = Object.keys(countByBatch);
  const batchValues = Object.values(countByBatch);

  res.send(`
  <!DOCTYPE html>
  <html>
  <head>
    <title>Admin Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
      body {font-family: Arial;background: linear-gradient(135deg, #6a11cb, #2575fc);margin:0;padding:0;display:flex;justify-content:center;align-items:flex-start;min-height:100vh;}
      .container {background:#fff;margin:30px;padding:25px;border-radius:12px;box-shadow:0 8px 20px rgba(0,0,0,0.2);width:95%;max-width:1100px;}
      table{width:100%;border-collapse:collapse;margin-top:15px;max-height:400px;overflow-y:auto;display:block;}
      thead, tbody {display:table;width:100%;table-layout:fixed;}
      th,td{border:1px solid #ddd;padding:10px;text-align:center;}
      th{background:#2575fc;color:#fff;position:sticky;top:0;}
      tr:nth-child(even){background:#f9f9f9;}
      button,.btn{background:#2575fc;color:#fff;border:none;padding:8px 14px;margin:5px;border-radius:6px;font-weight:bold;cursor:pointer;text-decoration:none;}
      .btn:hover,button:hover{background:#1a5bd9;}
      .delete-btn{background:#e74c3c;color:#fff;padding:6px 10px;border-radius:6px;text-decoration:none;}
    </style>
  </head>
  <body>
    <div class="container">
      <h2>📊 Admin Dashboard</h2>
      <p>Status: <b>${status.acceptingSubmissions ? "✅ OPEN" : "⛔ CLOSED"}</b></p>
      <p>Total Submissions: <b>${total}</b></p>
      <form method="post" action="/admin/toggle">
        <button type="submit">${status.acceptingSubmissions ? "Close Submissions" : "Open Submissions"}</button>
      </form>
      
      <h3>🔍 Search by Roll No</h3>
      <form method="get" action="/admin/search">
        <input type="text" name="rollno" placeholder="Enter Roll No" required>
        <button type="submit">Search</button>
      </form>

      <h3>📂 All Submissions</h3>
      <table>
        <thead>
          <tr><th>Name</th><th>Roll No</th><th>Batch</th><th>File</th><th>DateTime</th><th>Action</th></tr>
        </thead>
        <tbody>
          ${rows || "<tr><td colspan='6'>No submissions yet</td></tr>"}
        </tbody>
      </table>

      <canvas id="batchChart"></canvas>
      <br>
      <a class="btn" href="/admin/excel" download>⬇ Excel</a>
      <a class="btn" href="/admin/files" download>⬇ All Files (ZIP)</a>
      <a class="btn" href="/">⬅ Back</a>
    </div>
    <script>
      const ctx2 = document.getElementById('batchChart').getContext('2d');
      new Chart(ctx2, {
        type: 'bar',
        data: {
          labels: ${JSON.stringify(batchLabels)},
          datasets: [{
            label: 'Submissions by Batch',
            data: ${JSON.stringify(batchValues)},
            backgroundColor: '#36a2eb'
          }]
        },
        options: {scales: {y: {beginAtZero: true}}}
      });
    </script>
  </body>
  </html>
  `);
});

// -------------------- Toggle submissions --------------------
app.post("/admin/toggle", (req, res) => {
  const status = getStatus();
  setStatus(!status.acceptingSubmissions);
  res.redirect("/admin/dashboard");
});

// -------------------- Search by roll no --------------------
app.get("/admin/search", (req, res) => {
  const rollno = (req.query.rollno || "").trim();
  if (!rollno) return res.send("❌ Enter roll number.");
  let submissions = [];
  if (fs.existsSync(EXCEL_FILE)) {
    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheet = workbook.Sheets["Sheet1"];
    if (sheet) submissions = XLSX.utils.sheet_to_json(sheet);
  }
  const record = submissions.find(r => String(r.RollNo) === rollno);
  if (record) {
    res.send(`
      <h2>🔍 Search Result</h2>
      <p><b>Name:</b> ${escapeHtml(record.Name)}</p>
      <p><b>Roll No:</b> ${escapeHtml(record.RollNo)}</p>
      <p><b>Batch:</b> ${escapeHtml(record.Batch)}</p>
      <p><b>File:</b> <a href="${record.FileLink}" target="_blank">View File</a></p>
      <p><b>Submitted At:</b> ${escapeHtml(record.DateTime)}</p>
      <br><a href="/admin/dashboard">⬅ Back</a>
    `);
  } else {
    res.send(`❌ No record for Roll No: ${escapeHtml(rollno)}<br><a href="/admin/dashboard">⬅ Back</a>`);
  }
});

// -------------------- Delete submission --------------------
app.get("/admin/delete", (req, res) => {
  const rollno = (req.query.rollno || "").trim();
  if (!rollno) return res.send("❌ No Roll No provided.");
  let submissions = [];
  if (fs.existsSync(EXCEL_FILE)) {
    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheet = workbook.Sheets["Sheet1"];
    if (sheet) submissions = XLSX.utils.sheet_to_json(sheet);
  }
  const record = submissions.find(r => String(r.RollNo) === rollno);
  if (!record) return res.send("❌ Not found<br><a href='/admin/dashboard'>Back</a>");
  const filePath = path.join(__dirname, record.FileLink);
  if (fs.existsSync(filePath)) { try { fs.unlinkSync(filePath); } catch {} }
  submissions = submissions.filter(r => String(r.RollNo) !== rollno);
  const worksheet = XLSX.utils.json_to_sheet(submissions);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, EXCEL_FILE);
  res.send(`✅ Deleted Roll No: ${escapeHtml(rollno)}<br><a href="/admin/dashboard">Back</a>`);
});

// -------------------- Admin downloads --------------------
app.get("/admin/excel", (req, res) => {
  if (fs.existsSync(EXCEL_FILE)) return res.download(EXCEL_FILE);
  res.send("📂 No data yet!");
});
app.get("/admin/files", (req, res) => {
  const zipPath = path.join(UPLOAD_DIR, "all_files.zip");
  if (fs.existsSync(zipPath)) { try { fs.unlinkSync(zipPath); } catch {} }
  const output = fs.createWriteStream(zipPath);
  const archive = archiver("zip", { zlib: { level: 9 } });
  output.on("close", () => res.download(zipPath, "all_uploaded_files.zip"));
  archive.on("error", err => res.status(500).send("❌ Error creating ZIP"));
  archive.pipe(output);
  fs.readdirSync(UPLOAD_DIR).forEach(file => {
    if (["all_files.zip", "status.json", "data.xlsx"].includes(file)) return;
    const filePath = path.join(UPLOAD_DIR, file);
    if (fs.statSync(filePath).isFile()) archive.file(filePath, { name: file });
  });
  archive.finalize();
});

// serve uploaded files
app.use("/uploads", express.static(UPLOAD_DIR));

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));
