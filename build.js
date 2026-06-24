/**
 * Fleet Dashboard Build Script v2
 * Usage: node build.js <client-folder>
 * Example: node build.js menengai
 */

const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const http = require('http');

const CLIENTS_DIR = path.join(__dirname, 'clients');
const OUTPUT_DIR  = path.join(__dirname, 'output');
const SUMMARY_PREFIX = 'Summary';
const TRIP_PREFIX    = 'Trip';

function log(msg)  { console.log(`  ✔  ${msg}`); }
function warn(msg) { console.warn(`  ⚠  ${msg}`); }
function err(msg)  { console.error(`  ✘  ${msg}`); process.exit(1); }

function formatMonth(m) {
  const [y, mo] = m.split('-');
  return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+mo-1]+' '+y;
}

function toMonthStr(val) {
  if (val instanceof Date) {
    const rounded = new Date(Math.round(val.getTime() / 86400000) * 86400000);
    const y = rounded.getUTCFullYear();
    const m = String(rounded.getUTCMonth()+1).padStart(2,'0');
    return `${y}-${m}`;
  }
  
  if (typeof val === 'number') {
    const d = new Date(Math.round((val-25569)*86400*1000));
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}`;
  }
  if (typeof val === 'string') { const d=new Date(val); if(!isNaN(d)) return d.toISOString().slice(0,7); }
  return null;
}

function toDayStr(val) {
  if (val instanceof Date) {
    // Round to nearest day to absorb floating-point serial drift from source data
    const rounded = new Date(Math.round(val.getTime() / 86400000) * 86400000);
    const y = rounded.getUTCFullYear();
    const m = String(rounded.getUTCMonth()+1).padStart(2,'0');
    const d = String(rounded.getUTCDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'number') {
    const d = new Date(Math.round((val-25569)*86400*1000));
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`;
  }
  if (typeof val === 'string') { const d=new Date(val); if(!isNaN(d)) return d.toISOString().slice(0,10); }
  return null;
}

function normalizeCategory(raw) {
  if (!raw) return 'UNKNOWN';
  const s = String(raw).trim().toUpperCase();
  if (s.includes('MORL FARM') || s.startsWith('MORL')) return 'FARM';
  if (s.includes('RAIMDF')) return 'RAIMDF';
  if (s === 'TR' || s === 'WL') return 'RAIMDF';
  return s;
}

function extractCategoryFromLabel(label) {
  label = String(label||'').trim();
  if (label.toUpperCase().includes('MORL FARM')) return 'FARM';
  if (label.toUpperCase().includes('RAIMDF')) return 'RAIMDF';
  const parts = label.split(/\s*-\s*/);
  if (parts.length >= 2) {
    const cat = parts[0].trim().toUpperCase();
    if (cat === 'TR' || cat === 'WL') return 'RAIMDF';
    if (cat === 'MENENGAI') return 'UNKNOWN';
    return cat;
  }
  return 'UNKNOWN';
}

// ── Read Summary ──
function readSummary(filePath) {
  log(`Reading summary: ${path.basename(filePath)}`);
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = wb.SheetNames.find(s => s === 'Sheet1') || wb.SheetNames[1];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:0 });
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const [vehicle, category, mileage, duration] = rows[i];
    if (!vehicle || typeof vehicle !== 'string') continue;
    const month = toMonthStr(duration);
    if (!month) continue;
    data.push({ vehicle: vehicle.trim(), category: normalizeCategory(category), mileage: parseFloat(mileage)||0, month });
  }
  log(`  → ${data.length} records across ${new Set(data.map(d=>d.month)).size} months`);
  return data;
}

// ── Read Trip File (Utilization + Trips sheets) ──
function readTrip(filePath) {
  log(`Reading trip file: ${path.basename(filePath)}`);
  const wb = XLSX.readFile(filePath, { cellDates: true });

  // ── Utilization sheet ──
  const utilSheet = wb.SheetNames.find(s => s.toLowerCase().includes('utilization')) || wb.SheetNames[0];
  const utilRows = XLSX.utils.sheet_to_json(wb.Sheets[utilSheet], { header:1, defval:0 });
  const header = utilRows[0];

  const dayColIdxs = [], dayDates = [];
  let totalColIdx = -1;
  header.forEach((h, i) => {
    if (h instanceof Date || (typeof h === 'number' && h > 40000)) {
      const iso = toDayStr(h);
      if (iso) { dayColIdxs.push(i); dayDates.push(iso); }
    }
    if (typeof h === 'string' && h.toLowerCase().includes('total')) totalColIdx = i;
  });

  const month = dayDates.length ? dayDates[0].slice(0,7) : null;
  if (!month) { warn(`Could not detect month in ${path.basename(filePath)}`); return null; }

  const dailyMap = {};
  for (let i = 1; i < utilRows.length; i++) {
    const row = utilRows[i];
    const label = row[0];
    if (!label || typeof label !== 'string') continue;
    if (label.toLowerCase().includes('grand total')) continue;
    const days = {};
    dayColIdxs.forEach((ci, di) => { days[dayDates[di]] = parseFloat(row[ci])||0; });
    const total = totalColIdx>=0 ? (parseFloat(row[totalColIdx])||0) : Object.values(days).reduce((s,v)=>s+v,0);
    dailyMap[label.trim()] = {
      vehicle: label.trim(),
      category: extractCategoryFromLabel(label),
      days,
      total: parseFloat(total.toFixed(2)),
      active_days: Object.values(days).filter(v=>v>0).length,
      zero_days: dayDates.length - Object.values(days).filter(v=>v>0).length,
      trip_count: 0,
      avg_trip_km: 0
    };
  }

  // ── Trips sheet ──
  const tripsSheet = wb.SheetNames.find(s => s.toLowerCase() === 'trips');
  if (tripsSheet) {
    const tripRows = XLSX.utils.sheet_to_json(wb.Sheets[tripsSheet], { defval: '' });
    const subTrips = tripRows.filter(r => {
      const n = String(r['№']||'');
      return n.includes('.');
    });
    const tripStats = {};
    subTrips.forEach(r => {
      const grouping = String(r['Grouping']||'').trim();
      const dist = parseFloat(r['Distance Driven'])||0;
      if (!tripStats[grouping]) tripStats[grouping] = { count: 0, total: 0 };
      tripStats[grouping].count++;
      tripStats[grouping].total += dist;
    });
    Object.entries(tripStats).forEach(([vehicle, stats]) => {
      if (dailyMap[vehicle]) {
        dailyMap[vehicle].trip_count = stats.count;
        dailyMap[vehicle].avg_trip_km = parseFloat((stats.total/stats.count).toFixed(2));
      }
    });
    log(`  → Trip stats loaded for ${Object.keys(tripStats).length} vehicles`);
  }

  const vehicles = Object.values(dailyMap);
  log(`  → ${vehicles.length} vehicles · month: ${month}`);
  return { month, vehicles };
}

// ── Anomaly Detection ──
function detectAnomalies(monthlyVehicles, monthlyByCategory, allMonths) {
  const anomalies = [];

  // 1. Idle 3+ consecutive months
  monthlyVehicles.forEach(v => {
    let streak = 0, maxStreak = 0;
    allMonths.forEach(m => {
      if ((v.monthly[m]||0) === 0) { streak++; maxStreak = Math.max(maxStreak, streak); }
      else streak = 0;
    });
    if (maxStreak >= 3) anomalies.push({ type:'idle_streak', vehicle:v.vehicle, category:v.category, detail:`${maxStreak} consecutive idle months` });
  });

  // 2. Category drop 20%+ month on month
  Object.entries(monthlyByCategory).forEach(([cat, monthData]) => {
    allMonths.forEach((m, i) => {
      if (i === 0) return;
      const prev = monthData[allMonths[i-1]]||0;
      const curr = monthData[m]||0;
      if (prev > 0 && curr < prev * 0.8) {
        const drop = Math.round((1-curr/prev)*100);
        anomalies.push({ type:'category_drop', category:cat, month:m, detail:`${drop}% drop vs previous month`, vehicle:null });
      }
    });
  });

  // 3. Active last month now idle
  if (allMonths.length >= 2) {
    const lastM = allMonths[allMonths.length-1];
    const prevM = allMonths[allMonths.length-2];
    monthlyVehicles.forEach(v => {
      if ((v.monthly[prevM]||0) > 0 && (v.monthly[lastM]||0) === 0) {
        anomalies.push({ type:'newly_idle', vehicle:v.vehicle, category:v.category, detail:`Active in ${formatMonth(prevM)}, idle in ${formatMonth(lastM)}` });
      }
    });
  }
  return anomalies;
}

// ── Build Data ──
function buildData(clientName, summaryData, tripFiles) {
  const vehicleMap = {};
  summaryData.forEach(({ vehicle, category, mileage, month }) => {
    if (!vehicleMap[vehicle]) vehicleMap[vehicle] = { vehicle, category, monthly:{}, total:0, trip_count:0, avg_trip_km:0 };
    vehicleMap[vehicle].monthly[month] = parseFloat(mileage.toFixed(2));
    vehicleMap[vehicle].total = parseFloat((vehicleMap[vehicle].total + mileage).toFixed(2));
  });

  // Add trip stats to monthly vehicles
  tripFiles.forEach(({ vehicles }) => {
    vehicles.forEach(dv => {
      if (vehicleMap[dv.vehicle] && dv.trip_count > 0) {
        vehicleMap[dv.vehicle].trip_count = dv.trip_count;
        vehicleMap[dv.vehicle].avg_trip_km = dv.avg_trip_km;
      }
    });
  });

  const monthlyVehicles = Object.values(vehicleMap);
  const allMonths = [...new Set(summaryData.map(d=>d.month))].sort();
  const monthlyFleet = {};
  allMonths.forEach(m => {
    monthlyFleet[m] = parseFloat(summaryData.filter(d=>d.month===m).reduce((s,d)=>s+d.mileage,0).toFixed(2));
  });

  const allCats = [...new Set([
    ...summaryData.map(d=>d.category),
    ...tripFiles.flatMap(t=>t.vehicles.map(v=>v.category))
  ])].filter(c=>c!=='UNKNOWN').sort();
  allCats.push('UNKNOWN');

  const monthlyByCat = {};
  allCats.forEach(cat => {
    monthlyByCat[cat] = {};
    allMonths.forEach(m => {
      monthlyByCat[cat][m] = parseFloat(summaryData.filter(d=>d.category===cat&&d.month===m).reduce((s,d)=>s+d.mileage,0).toFixed(2));
    });
  });

  const dailyVehicles = tripFiles.flatMap(t=>t.vehicles);
  const dailyMonths = tripFiles.map(t=>t.month);
  const dailyFleet = {};
  tripFiles.forEach(({ vehicles }) => {
    if (!vehicles.length) return;
    const allDayKeys = new Set();
    vehicles.forEach(v => Object.keys(v.days).forEach(d => allDayKeys.add(d)));
    allDayKeys.forEach(d => {
      dailyFleet[d] = parseFloat(vehicles.reduce((s,v)=>s+(v.days[d]||0),0).toFixed(2));
    });
  });

  const anomalies = detectAnomalies(monthlyVehicles, monthlyByCat, allMonths);

  return {
    client: clientName,
    generated: new Date().toISOString().slice(0,16).replace('T',' '),
    categories: allCats,
    months: allMonths,
    daily_months: dailyMonths,
    monthly_vehicles: monthlyVehicles,
    daily_vehicles: dailyVehicles,
    daily_fleet_totals: dailyFleet,
    monthly_by_category: monthlyByCat,
    monthly_fleet_totals: monthlyFleet,
    anomalies
  };
}

// ── Inject into Template ──
function buildHTML(data, templatePath) {
  let html = fs.readFileSync(templatePath, 'utf8');
  const dataJson = JSON.stringify(data);
  if (html.includes('{data_json}')) {
    html = html.replace('{data_json}', dataJson);
  } else {
    html = html.replace(
      /const RAW = [\s\S]*?;(\s*<\/script>)/,
      `const RAW = ${dataJson};$1`
    );
  }
  html = html.replace(/<title>.*?<\/title>/, `<title>${data.client} Fleet Dashboard</title>`);
  return html;
}
// ── Main ──
function main() {
  const clientArg = process.argv[2];
  if (!clientArg) err('Usage: node build.js <client-name>');
  const clientDir = path.join(CLIENTS_DIR, clientArg);
  if (!fs.existsSync(clientDir)) err(`Client folder not found: ${clientDir}`);

  console.log(`\n🚀 Building dashboard for: ${clientArg}\n`);

  const files = fs.readdirSync(clientDir).filter(f => f.endsWith('.xlsx'));
  const summaryFiles = files.filter(f => f.startsWith(SUMMARY_PREFIX));
  const tripFileNames = files.filter(f => f.startsWith(TRIP_PREFIX));

  if (!summaryFiles.length) err(`No Summary*.xlsx found in ${clientDir}`);
  if (summaryFiles.length > 1) warn(`Multiple summary files — using: ${summaryFiles[0]}`);

  log(`Found summary: ${summaryFiles[0]}`);
  log(`Found ${tripFileNames.length} trip file(s): ${tripFileNames.join(', ')}`);

  const summaryData = readSummary(path.join(clientDir, summaryFiles[0]));
  const tripFiles = tripFileNames.map(f => readTrip(path.join(clientDir, f))).filter(Boolean);

  const clientName = clientArg.charAt(0).toUpperCase() + clientArg.slice(1);
  const data = buildData(clientName, summaryData, tripFiles);

  log(`Months: ${data.months.join(', ')}`);
  log(`Daily months: ${data.daily_months.join(', ') || 'none'}`);
  log(`Monthly vehicles: ${data.monthly_vehicles.length}`);
  log(`Daily vehicles: ${data.daily_vehicles.length}`);
  log(`Anomalies detected: ${data.anomalies.length}`);

  const templatePath = path.join(__dirname, 'template', 'dashboard.html');
  if (!fs.existsSync(templatePath)) err(`Template not found: ${templatePath}`);

  const html = buildHTML(data, templatePath);
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const outName = `${clientName}_Dashboard.html`;
  const outPath = path.join(OUTPUT_DIR, outName);
  fs.writeFileSync(outPath, html, 'utf8');

  console.log(`\n✅ Dashboard built!\n`);
  console.log(`   📄 Output: output\\${outName}`);
  console.log(`   📦 Size:   ${(html.length/1024).toFixed(1)} KB\n`);
}

main();

// ── Watch Mode ──
if (process.argv.includes('--watch')) {
  const clientArg = process.argv[2];
  const watchDir = path.join(CLIENTS_DIR, clientArg);
  console.log(`\n👁  Watching for changes in clients\\${clientArg}\\ ...\n`);
  console.log(`    Drop files in anytime — dashboard rebuilds automatically.`);
  console.log(`    Press Ctrl+C to stop.\n`);
  let debounce = null;
  fs.watch(watchDir, (eventType, filename) => {
    if (!filename || !filename.endsWith('.xlsx')) return;
    clearTimeout(debounce);
    debounce = setTimeout(() => {
      console.log(`\n📂 Change detected: ${filename}`);
      try { main(); } catch(e) { console.error('Build failed:', e.message); }
    }, 1500);
  });
}

// ── Server Mode ──
if (process.argv.includes('--serve')) {
  const clientArg = process.argv[2];
  const port = 3000;
  const server = http.createServer((req, res) => {
    const clientName = clientArg.charAt(0).toUpperCase() + clientArg.slice(1);
    const filePath = path.join(OUTPUT_DIR, `${clientName}_Dashboard.html`);
    fs.readFile(filePath, (err, data) => {
      if (err) { res.writeHead(404); res.end('Dashboard not built yet'); return; }
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(data);
    });
  });
  server.listen(port, () => {
    console.log(`\n🌐 Dashboard live at: http://localhost:${port}`);
    console.log(`   Opening in browser...\n`);
    const { exec } = require('child_process');
    exec(`start http://localhost:${port}`);
  });
}