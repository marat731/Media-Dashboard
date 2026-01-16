const express = require('express');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const { nanoid } = require('nanoid');

const app = express();
const PORT = process.env.PORT || 3000;

// Initialize SQLite database
const db = new Database(process.env.DATABASE_PATH || './dashboards.db');

// Create tables
db.exec(`
  CREATE TABLE IF NOT EXISTS dashboards (
    id TEXT PRIMARY KEY,
    share_id TEXT UNIQUE,
    name TEXT,
    data TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

// Middleware
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files are allowed'));
    }
  }
});

// Parse Excel file and extract data
function parseExcelFile(buffer, filename) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Extract brand name from first row
  const headerRow = rawData[0] || [];
  const brandMatch = headerRow[0]?.match(/Details:\s*(.+)/);
  const brandName = brandMatch ? brandMatch[1].trim() : filename.replace(/\.(xlsx|xls)$/i, '').replace(/ - Spend by Property/i, '');

  // Find the data rows (starting after "Media Owner" header)
  let dataStartIndex = -1;
  let totalRow = null;

  for (let i = 0; i < rawData.length; i++) {
    if (rawData[i][0] === 'Media Owner') {
      dataStartIndex = i + 1;
      break;
    }
  }

  if (dataStartIndex === -1) {
    throw new Error(`Could not find data in file: ${filename}`);
  }

  // Extract property data
  const properties = [];
  let totalSpend = 0;
  let yoyChange = '';
  let propertyCount = 0;

  for (let i = dataStartIndex; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row[0] || row[0] === '') break;

    const mediaOwner = row[0];
    const property = row[1];
    const type = row[2] || 'B2C';
    const tvSpend = parseFloat(row[3]) || 0;
    const digitalSpend = parseFloat(row[4]) || 0;
    const printSpend = parseFloat(row[5]) || 0;
    const printPages = parseFloat(row[6]) || 0;
    const totalPropertySpend = parseFloat(row[7]) || 0;
    const trend = row[8] || '';

    if (mediaOwner === 'Total') {
      totalSpend = totalPropertySpend;
      yoyChange = trend;
      continue;
    }

    propertyCount++;
    properties.push({
      property: property || mediaOwner,
      owner: mediaOwner,
      type,
      tvSpend,
      digitalSpend,
      printSpend,
      printPages,
      spend: totalPropertySpend,
      trend
    });
  }

  // Calculate stats
  const newProperties = properties.filter(p => p.trend.includes('NEW')).length;
  const topChannel = properties.length > 0 ? properties[0] : null;

  // Find fastest growing (highest positive percentage, excluding NEW)
  let fastestGrowing = null;
  let highestGrowth = -Infinity;

  for (const prop of properties) {
    if (prop.trend.includes('NEW') || prop.trend.includes('-')) continue;
    const match = prop.trend.match(/\+?([\d,]+)%/);
    if (match) {
      const growth = parseFloat(match[1].replace(',', ''));
      if (growth > highestGrowth) {
        highestGrowth = growth;
        fastestGrowing = prop;
      }
    }
    // Handle "Over X%" format
    const overMatch = prop.trend.match(/Over\s+([\d,]+)/i);
    if (overMatch) {
      const growth = parseFloat(overMatch[1].replace(',', ''));
      if (growth > highestGrowth) {
        highestGrowth = growth;
        fastestGrowing = prop;
      }
    }
  }

  return {
    name: brandName,
    totalSpend,
    yoyChange,
    properties: propertyCount,
    newProperties,
    topChannel: topChannel ? { name: topChannel.property, trend: topChannel.trend } : null,
    fastestGrowing: fastestGrowing ? { name: fastestGrowing.property, trend: fastestGrowing.trend } : null,
    data: properties
  };
}

// API Routes

// Create new dashboard
app.post('/api/dashboards', upload.array('files', 5), (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const brands = [];
    for (const file of req.files) {
      const brandData = parseExcelFile(file.buffer, file.originalname);
      brands.push(brandData);
    }

    const dashboardId = nanoid(10);
    const dashboardName = req.body.name || `Dashboard ${new Date().toLocaleDateString()}`;

    const stmt = db.prepare(`
      INSERT INTO dashboards (id, name, data)
      VALUES (?, ?, ?)
    `);

    stmt.run(dashboardId, dashboardName, JSON.stringify(brands));

    res.json({
      success: true,
      dashboardId,
      redirect: `/dashboard/${dashboardId}`
    });
  } catch (error) {
    console.error('Error creating dashboard:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get dashboard data
app.get('/api/dashboards/:id', (req, res) => {
  try {
    const stmt = db.prepare('SELECT * FROM dashboards WHERE id = ? OR share_id = ?');
    const dashboard = stmt.get(req.params.id, req.params.id);

    if (!dashboard) {
      return res.status(404).json({ error: 'Dashboard not found' });
    }

    // Check if accessing via share_id (view-only mode)
    const isViewOnly = dashboard.share_id === req.params.id;

    res.json({
      id: dashboard.id,
      name: dashboard.name,
      brands: JSON.parse(dashboard.data),
      shareId: dashboard.share_id,
      isViewOnly,
      createdAt: dashboard.created_at
    });
  } catch (error) {
    console.error('Error fetching dashboard:', error);
    res.status(500).json({ error: error.message });
  }
});

// Generate share link
app.post('/api/dashboards/:id/share', (req, res) => {
  try {
    const stmt = db.prepare('SELECT * FROM dashboards WHERE id = ?');
    const dashboard = stmt.get(req.params.id);

    if (!dashboard) {
      return res.status(404).json({ error: 'Dashboard not found' });
    }

    // Generate share ID if not exists
    let shareId = dashboard.share_id;
    if (!shareId) {
      shareId = nanoid(12);
      const updateStmt = db.prepare('UPDATE dashboards SET share_id = ? WHERE id = ?');
      updateStmt.run(shareId, req.params.id);
    }

    res.json({
      success: true,
      shareId,
      shareUrl: `/view/${shareId}`
    });
  } catch (error) {
    console.error('Error generating share link:', error);
    res.status(500).json({ error: error.message });
  }
});

// Serve upload page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve dashboard page
app.get('/dashboard/:id', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

// Serve view-only page (same as dashboard but will detect view-only mode)
app.get('/view/:shareId', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

app.listen(PORT, () => {
  console.log(`Dashboard platform running on port ${PORT}`);
});
