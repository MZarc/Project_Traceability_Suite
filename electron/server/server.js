const express = require('express');
const Database = require('better-sqlite3');
const cors = require('cors');
const bodyParser = require('body-parser');
const path = require('path');
const cron = require('node-cron');
const XLSX = require('xlsx');

const ExcelJS = require('exceljs');
const fs = require('fs-extra');

const app = express();
const port = 5000;

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

// Determine writable data path
const args = process.argv.slice(2);
const baseDirArg = args.find(a => a.startsWith('--base-dir='));
const baseDir = baseDirArg ? baseDirArg.split('=')[1] : __dirname;

const dbPath = path.join(baseDir, 'database.sqlite');
const backupDir = path.join(baseDir, 'backups');

const db = new Database(dbPath);

try {
  db.pragma('journal_mode = WAL');
  db.pragma('synchronous = NORMAL');
  db.pragma('temp_store = MEMORY');
  db.pragma('cache_size = -64000');
} catch (e) {
  console.log('[DB PRAGMA] Fallback applied:', e.message);
}

const STAGE_STRUCTURE = [
  { name: 'Sheet metal', sub: ['status', 'Due date', 'Date completion'] },
  { name: 'Plating / powder coating', sub: ['status', 'Due date', 'Date completion'] },
  { name: 'Knurling / Pressing / Gasket', sub: ['status', 'Due date', 'Date completion'] },
  { name: 'Kitting', sub: ['status', 'Due date', 'Date completion'] },
  { name: 'Procurement / ordersheet', sub: ['status', 'Due date', 'Date completion'] },
  { name: 'SMT', sub: ['status', 'Plan given date', 'Date completion'] },
  { name: 'TH', sub: ['status', 'Plan given date', 'Date completion'] },
  { name: 'Dispatch / FG', sub: ['status', 'Due date', 'Date completion'] },
];

const STAGE_STRUCTURE_NORMALIZED = STAGE_STRUCTURE.map(s => ({
  ...s,
  display: s.name.toUpperCase()
}));

async function runAutomatedBackup() {
  console.log('[BACKUP] Starting automated daily backup...');
  try {
    await fs.ensureDir(backupDir);

    const projects = db.prepare('SELECT * FROM projects ORDER BY id ASC').all();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Meet Mistry';
    workbook.lastModifiedBy = 'Meet Mistry';
    workbook.created = new Date();
    workbook.modified = new Date();
    const worksheet = workbook.addWorksheet('Daily Backup');

    // 1. Columns Definition
    const mainCols = [
      { header: 'SR.NO', width: 8 },
      { header: 'PROJECT NAME', width: 30 },
      { header: 'CUSTOMER', width: 25 },
      { header: 'QTY', width: 10 },
      { header: 'PRIORITY', width: 15 },
      { header: 'DATA RECEIVED', width: 18 },
      { header: 'DATA SHARED', width: 18 },
    ];
    STAGE_STRUCTURE_NORMALIZED.forEach(stage => {
      mainCols.push({ header: stage.display, width: 15 }, { header: stage.display, width: 15 }, { header: stage.display, width: 15 });
    });
    mainCols.push({ header: 'REMARKS', width: 40 });

    const header1 = worksheet.addRow(mainCols.map(c => c.header));
    const subHeader = [...Array(7).fill(null)];
    STAGE_STRUCTURE_NORMALIZED.forEach(stage => {
      subHeader.push('STATUS', stage.name === 'SMT' || stage.name === 'TH' ? 'PLAN GIVEN DATE' : 'DUE DATE', 'COMPLETION DATE');
    });
    subHeader.push(null);
    const header2 = worksheet.addRow(subHeader);

    // 2. Head Styling & Merges
    [header1, header2].forEach((row, idx) => {
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 10 };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
    });
    for(let i=1; i<=7; i++) worksheet.mergeCells(1, i, 2, i);
    let stCol = 8;
    STAGE_STRUCTURE_NORMALIZED.forEach(() => {
      worksheet.mergeCells(1, stCol, 1, stCol+2);
      stCol += 3;
    });
    worksheet.mergeCells(1, stCol, 2, stCol);

    // 3. Data Rows
    projects.forEach(p => {
      const stages = JSON.parse(p.stages || '{}');
      const rowData = [
        String(p.sr_no || '').toUpperCase(),
        String(p.project_name || '').toUpperCase(),
        String(p.customer || '').toUpperCase(),
        p.qty,
        String(p.priority || 'MEDIUM').toUpperCase(),
        String(p.due_date || '').toUpperCase(),
        String(p.completion_date || '').toUpperCase()
      ];
      STAGE_STRUCTURE_NORMALIZED.forEach(s => {
        const plannedKey = s.name === 'SMT' || s.name === 'TH' ? `${s.name}_Plan given date` : `${s.name}_Due date`;
        rowData.push(
          String(stages[`${s.name}_status`] || 'PENDING').toUpperCase(),
          String(stages[plannedKey] || '').toUpperCase(),
          String(stages[`${s.name}_Date completion`] || '').toUpperCase()
        );
      });
      rowData.push(String(p.remarks || '').toUpperCase());
      const r = worksheet.addRow(rowData);
      r.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        if (colNum > 7 && (colNum-8) % 3 === 0) {
          const s = String(cell.value).toLowerCase();
          let c = 'FFF1F5F9';
          if (s.includes('done') || s.includes('compl')) c = 'FFD1FAE5';
          else if (s.includes('hold')) c = 'FFFEE2E2';
          else if (s.includes('run')) c = 'FFE0F2FE';
          else if (s === 'pending') c = 'FFFEF3C7';
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c } };
        }
      });
    });

    const fileName = `PMS_Backup_${new Date().toISOString().split('T')[0]}.xlsx`;
    const filePath = path.join(backupDir, fileName);
    await workbook.xlsx.writeFile(filePath);
    console.log(`[BACKUP] Saved: ${fileName}`);

    // 4. Cleanup Logic (Keep 3 latest)
    const files = await fs.readdir(backupDir);
    const backups = files.filter(f => f.startsWith('PMS_Backup_') && f.endsWith('.xlsx'))
                        .map(f => ({ name: f, time: fs.statSync(path.join(backupDir, f)).mtimeMs }))
                        .sort((a, b) => b.time - a.time);

    if (backups.length > 3) {
      for (let i = 3; i < backups.length; i++) {
        await fs.remove(path.join(backupDir, backups[i].name));
        console.log(`[BACKUP] Rotated old version: ${backups[i].name}`);
      }
    }
  } catch (err) {
    console.error('[BACKUP] Error during backup cycle:', err);
  }
}

// Scheduled Backup: Every day at midnight
cron.schedule('0 0 * * *', runAutomatedBackup);
// System Boot Backup
runAutomatedBackup();

// Schedualed Backup every day at midnight

// Create tables
db.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sr_no TEXT,
    customer TEXT,
    project_name TEXT,
    qty INTEGER,
    priority TEXT,
    due_date TEXT,
    completion_date TEXT,
    remarks TEXT,
    stages JSON DEFAULT '{}',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS dynamic_columns (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE,
    type TEXT DEFAULT 'text',
    display_order INTEGER
  );

  CREATE TABLE IF NOT EXISTS todos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task TEXT NOT NULL,
    priority TEXT DEFAULT 'Medium',
    category TEXT DEFAULT 'Production',
    due_date TEXT,
    due_time TEXT,
    completed INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    address TEXT,
    credibility TEXT DEFAULT 'Medium',
    gst_no TEXT,
    remarks TEXT,
    contacts JSON DEFAULT '[]',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  CREATE INDEX IF NOT EXISTS idx_projects_sr_no ON projects(sr_no);
  CREATE INDEX IF NOT EXISTS idx_projects_customer ON projects(customer);
  CREATE INDEX IF NOT EXISTS idx_projects_priority ON projects(priority);
  CREATE INDEX IF NOT EXISTS idx_projects_due_date ON projects(due_date);
  CREATE INDEX IF NOT EXISTS idx_todos_completed ON todos(completed);
  CREATE INDEX IF NOT EXISTS idx_customers_name ON customers(name);
`);

// Customer Database Endpoints
app.get('/api/customers', (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM customers ORDER BY name ASC').all();
    const formatted = rows.map(r => ({
      ...r,
      contacts: typeof r.contacts === 'string' ? JSON.parse(r.contacts || '[]') : r.contacts
    }));
    res.json(formatted);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/customers', (req, res) => {
  try {
    const { name, address, credibility, gst_no, remarks, contacts } = req.body;
    if (!name || !name.trim()) {
      return res.status(400).json({ error: 'Customer name is required' });
    }
    const stmt = db.prepare(`
      INSERT INTO customers (name, address, credibility, gst_no, remarks, contacts)
      VALUES (?, ?, ?, ?, ?, ?)
    `);
    const info = stmt.run(
      name.trim(),
      address || '',
      credibility || 'Medium',
      gst_no || '',
      remarks || '',
      JSON.stringify(contacts || [])
    );
    const newCustomer = db.prepare('SELECT * FROM customers WHERE id = ?').get(info.lastInsertRowid);
    newCustomer.contacts = JSON.parse(newCustomer.contacts || '[]');
    res.json(newCustomer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/customers/:id', (req, res) => {
  try {
    const { id } = req.params;
    const { name, address, credibility, gst_no, remarks, contacts } = req.body;
    const stmt = db.prepare(`
      UPDATE customers 
      SET name = ?, address = ?, credibility = ?, gst_no = ?, remarks = ?, contacts = ?
      WHERE id = ?
    `);
    stmt.run(
      name.trim(),
      address || '',
      credibility || 'Medium',
      gst_no || '',
      remarks || '',
      JSON.stringify(contacts || []),
      id
    );
    const updated = db.prepare('SELECT * FROM customers WHERE id = ?').get(id);
    if (updated) {
      updated.contacts = JSON.parse(updated.contacts || '[]');
    }
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/customers/:id', (req, res) => {
  try {
    const { id } = req.params;
    db.prepare('DELETE FROM customers WHERE id = ?').run(id);
    res.json({ success: true, id });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// To-Do API Endpoints
app.get('/api/todos', (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM todos ORDER BY completed ASC, created_at DESC').all();
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/todos', (req, res) => {
  const { task, priority, category, due_date, due_time } = req.body;
  try {
    const stmt = db.prepare('INSERT INTO todos (task, priority, category, due_date, due_time) VALUES (?, ?, ?, ?, ?)');
    const info = stmt.run(task, priority || 'Medium', category || 'Production', due_date || '', due_time || '');
    res.json({ id: info.lastInsertRowid });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/todos/:id', (req, res) => {
  const { id } = req.params;
  const { task, priority, category, due_date, due_time, completed } = req.body;
  try {
    const stmt = db.prepare('UPDATE todos SET task = ?, priority = ?, category = ?, due_date = ?, due_time = ?, completed = ? WHERE id = ?');
    stmt.run(task, priority, category, due_date, due_time, completed ? 1 : 0, id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/todos/:id', (req, res) => {
  try {
    db.prepare('DELETE FROM todos WHERE id = ?').run(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// API Endpoints
app.get('/api/projects', (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT *
      FROM projects
      ORDER BY
        CASE
          WHEN trim(sr_no) = '' THEN 1
          ELSE 0
        END,
        CAST(sr_no AS REAL) ASC,
        id ASC
    `).all();
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/projects', (req, res) => {
  const { sr_no, customer, project_name, qty, priority, due_date, completion_date, remarks, stages } = req.body;
  try {
    const stmt = db.prepare(`
      INSERT INTO projects (sr_no, customer, project_name, qty, priority, due_date, completion_date, remarks, stages)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const info = stmt.run(sr_no, customer, project_name, qty, priority, due_date, completion_date, remarks, JSON.stringify(stages || {}));
    res.json({ id: info.lastInsertRowid });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/projects/:id', (req, res) => {
  const { id } = req.params;
  const { sr_no, customer, project_name, qty, priority, due_date, completion_date, remarks, stages } = req.body;
  try {
    const stmt = db.prepare(`
      UPDATE projects 
      SET sr_no = ?, customer = ?, project_name = ?, qty = ?, priority = ?, due_date = ?, completion_date = ?, remarks = ?, stages = ?
      WHERE id = ?
    `);
    stmt.run(sr_no, customer, project_name, qty, priority, due_date, completion_date, remarks, JSON.stringify(stages || {}), id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/projects/:id', (req, res) => {
  try {
    db.prepare('DELETE FROM projects WHERE id = ?').run(req.params.id);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/projects-clear', (req, res) => {
  try {
    db.prepare('DELETE FROM projects').run();
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Dynamic Columns Management
app.get('/api/columns', (req, res) => {
  const rows = db.prepare('SELECT * FROM dynamic_columns ORDER BY display_order').all();
  res.json(rows);
});

app.post('/api/columns', (req, res) => {
  const { name, type } = req.body;
  try {
    const stmt = db.prepare('INSERT INTO dynamic_columns (name, type, display_order) VALUES (?, ?, ?)');
    stmt.run(name, type || 'text', 99);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/import-excel', (req, res) => {
  const { projects, mappings } = req.body;
  
  try {
    const insertStmt = db.prepare(`
      INSERT INTO projects (sr_no, customer, project_name, qty, priority, due_date, completion_date, remarks, stages)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const transaction = db.transaction((data) => {
      for (const row of data) {
        const mapped = {};
        for (const [sysKey, exKey] of Object.entries(mappings)) {
          mapped[sysKey] = row[exKey];
        }

        // Handle dynamic stages
        const stages = {};
        for (const [key, val] of Object.entries(row)) {
          if (!Object.values(mappings).includes(key)) {
            stages[key] = val;
            db.prepare(`
              INSERT INTO dynamic_columns (name, type, display_order)
              VALUES (?, 'text', 99)
              ON CONFLICT(name) DO NOTHING
            `).run(key);
          }
        }

        insertStmt.run(
          mapped.sr_no || '',
          mapped.customer || '',
          mapped.project_name || '',
          parseInt(mapped.qty) || 0,
          mapped.priority || '',
          mapped.due_date || '',
          mapped.completion_date || '',
          mapped.remarks || '',
          JSON.stringify(stages)
        );
      }
    });

    transaction(projects);
    res.json({ success: true, count: projects.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
