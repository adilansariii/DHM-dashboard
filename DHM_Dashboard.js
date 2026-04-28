// ============================================================
// STATE
// ============================================================
let DATA = {};
let SORT = {};
let PAGE = {};

const PAGE_SIZE = 50;

// ============================================================
// DRAG & DROP / FILE INPUT
// ============================================================
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) processFile(f);
});
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) processFile(e.target.files[0]);
});

function processFile(file) {
  document.getElementById('parse-error').style.display = 'none';
  document.getElementById('loading').style.display = 'flex';
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:'array', cellDates:true });
      parseWorkbook(wb, file.name);
    } catch(err) {
      document.getElementById('loading').style.display = 'none';
      document.getElementById('parse-error').style.display = 'block';
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// PARSE WORKBOOK
// ============================================================
function parseWorkbook(wb, filename) {
  const gs = n => wb.SheetNames.find(s => s.toLowerCase().replace(/\s/g,'').includes(n.toLowerCase().replace(/\s/g,'')));

  // Helper: get sheet rows as array of arrays
  function sheetRows(shName) {
    const sh = wb.Sheets[shName];
    if (!sh) return [];
    return XLSX.utils.sheet_to_json(sh, { header:1, defval:'' });
  }

  // --- Dashboard sheet (KPIs) ---
  const dashName = gs('Dashboard');
  const dashData = sheetRows(dashName || wb.SheetNames[0]);
  DATA.kpis = parseDashboardKPIs(dashData);

  // --- Account Wise % Count ---
  // UK: rows 2–32 (header row 1, col indices 0=Account,1=Total,2=Unhealthy,3=Healthy,5=Healthy%)
  // US: rows 37–43 (header row 36, col indices 0=Account,1=Total,2=NotConnected,3=Healthy,5=Healthy%)
  DATA.accountsUK = [];
  DATA.accountsUS = [];
  const accRows = sheetRows(gs('Account Wise'));
  let inUS = false;
  for (let i=0; i<accRows.length; i++) {
    const r = accRows[i];
    const cell0 = String(r[0]||'').trim();
    if (cell0 === 'Account') { if (DATA.accountsUK.length > 0) inUS = true; continue; }
    if (!cell0 || cell0 === 'Grand Total') continue;
    const total = toNum(r[1]);
    if (isNaN(total) || total === 0) continue;
    const healthy = toNum(r[3]);
    const unhealthy = total - (isNaN(healthy) ? 0 : healthy);
    const pctH = toNum(r[5]);
    const obj = { account: cell0, total, unhealthy: isNaN(unhealthy)?0:unhealthy, healthy: isNaN(healthy)?0:healthy, pct: isNaN(pctH)?0:pctH };
    if (!inUS) DATA.accountsUK.push(obj);
    else DATA.accountsUS.push(obj);
  }

  // --- UK Server (sheet: 'UK Sever') ---
  // Header at row index 9: col1=Account, col2=Location, col3=Building, col4=Floor
  //   col5=TotalDevice, col6=UnhealthyDevices, col9=DeviceOffline%, col10=Issue, col11=Status, col14=Action, col15=SiteType
  DATA.floorsUK = [];
  const ukRows = sheetRows(gs('UK Sever') || gs('UK Server'));
  for (let i=0; i<ukRows.length; i++) {
    const r = ukRows[i];
    const acc = String(r[1]||'').trim();
    if (!acc || acc==='Account' || acc==='Grand Total') continue;
    const total = toNum(r[5]);
    if (isNaN(total)) continue;
    DATA.floorsUK.push({
      account: acc,
      location: String(r[2]||'').trim(),
      building: String(r[3]||'').trim(),
      floor: String(r[4]||'').trim(),
      total,
      unhealthy: toNum(r[6]) || 0,
      band: String(r[9]||'').trim(),
      issue: String(r[10]||'').trim(),
      status: String(r[11]||'').trim(),
      siteType: String(r[15]||'').trim(),
    });
  }

  // --- US Server ---
  // Header at row index 10: col1=Account, col2=Location, col3=Building, col4=Floor
  //   col5=TotalDevices, col6=NotUpdated48hrs, col8=HealthyDevices, col9=DeviceOffline%, col10=Issue, col11=Status
  DATA.floorsUS = [];
  const usRows = sheetRows(gs('US Server'));
  for (let i=0; i<usRows.length; i++) {
    const r = usRows[i];
    const acc = String(r[1]||'').trim();
    if (!acc || acc==='Account' || acc==='Grand Total') continue;
    const total = toNum(r[5]);
    if (isNaN(total) || total === 0) continue;
    const notConnected = toNum(r[6]);
    DATA.floorsUS.push({
      account: acc,
      location: String(r[2]||'').trim(),
      building: String(r[3]||'').trim(),
      floor: String(r[4]||'').trim(),
      total,
      unhealthy: isNaN(notConnected) ? 0 : notConnected,
      band: String(r[9]||'').trim(),
      issue: String(r[10]||'').trim(),
      status: String(r[11]||'').trim(),
    });
  }

  // --- MS Report ---
  // Summary: row0=UK header, row1=UK values, row3=US header, row4=US values
  // Data header at row9, data starts at row10
  // Cols: 0=Account,1=Location,2=Building,3=Floor,4=MSName,5=MSID,6=LastConnected,7=DateExtracted,8=DaysOffline,11=overallStatus,13=Status,14=Comment
  DATA.msRows = [];
  DATA.msSummary = { ukTotal:0, ukOnline:0, ukOffline:0, usTotal:0, usOnline:0, usOffline:0 };
  const msRows = sheetRows(gs('MS Report'));
  if (msRows.length > 0) {
    // Row 0 = UK label, Row 1 = UK values
    DATA.msSummary.ukTotal  = toNum(msRows[1]?.[0]) || 0;
    DATA.msSummary.ukOnline  = toNum(msRows[1]?.[1]) || 0;
    DATA.msSummary.ukOffline = toNum(msRows[1]?.[2]) || 0;
    // Row 3 = US label, Row 4 = US values
    DATA.msSummary.usTotal  = toNum(msRows[4]?.[0]) || 0;
    DATA.msSummary.usOnline  = toNum(msRows[4]?.[1]) || 0;
    DATA.msSummary.usOffline = toNum(msRows[4]?.[2]) || 0;
    // Data rows start at row 10 (header at row 9)
    for (let i=10; i<msRows.length; i++) {
      const d = msRows[i];
      const acc = String(d[0]||'').trim();
      if (!acc) continue;
      const daysRaw = toNum(d[8]);
      const overallStatus = String(d[11]||'').toLowerCase().trim();
      const isOffline = overallStatus === 'danger';
      DATA.msRows.push({
        account: acc,
        location: String(d[1]||'').trim(),
        building: String(d[2]||'').trim(),
        floor: String(d[3]||'').trim(),
        msName: String(d[4]||'').trim(),
        daysOffline: isNaN(daysRaw) ? 0 : daysRaw,
        status: isOffline ? 'Offline' : 'Online',
        ticketStatus: String(d[13]||'').trim(),
        comment: String(d[14]||'').trim(),
      });
    }
  }

  // --- Mastercard Devices ---
  // Header at row0: 0=Account,1=Location,2=Building,3=BatterySwap,4=IoTDevice,5=FreespaceDevice,6=TotalDevices,7=NotConnected,8=PctOffline,9=TicketLink,10=Comment,11=priority
  // Data rows start at row1
  DATA.mcRows = [];
  const mcRows = sheetRows(gs('Mastercard Devices'));
  for (let i=1; i<mcRows.length; i++) {
    const r = mcRows[i];
    const loc = String(r[1]||'').trim();
    if (!loc || loc === ' ') continue;
    const total = toNum(r[6]);
    if (isNaN(total) || total === 0) continue;
    const pctOff = toNum(r[8]);
    DATA.mcRows.push({
      account: String(r[0]||'').trim(),
      location: loc,
      building: String(r[2]||'').trim(),
      batSwap: String(r[3]||'').trim(),
      total,
      notConnected: toNum(r[7]) || 0,
      pct: isNaN(pctOff) ? 0 : pctOff,
      priority: String(r[11]||'').trim(),
      comment: String(r[10]||'').trim(),
    });
  }

  // --- Extract report date from filename ---
  const dateMatch = filename.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (dateMatch) {
    document.getElementById('report-date').textContent = `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`;
  } else {
    document.getElementById('report-date').textContent = new Date().toLocaleDateString('en-GB');
  }

  document.getElementById('loading').style.display = 'none';
  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';

  populateFilters();
  renderOverview();
  renderAccountsTable();
  renderFloorTable('uk');
  renderFloorTable('us');
  renderMSTable();
  renderMCTable();
}

// ============================================================
// PARSE KPIs FROM DASHBOARD SHEET
// ============================================================
function parseDashboardKPIs(rows) {
  const kpi = {};
  // Based on exact layout validation:
  // Row 3:  col1='Total Devices Under DHM Monitoring', col3=value | col6='Total Floors...', col8=value | col10='Floor Summary - UK'
  // Row 6:  col1='Healthy Devices', col3=value | col6='Healthy Floors', col8=value | col10='Floor Health' header
  // Row 7:  col10='Above 60%', col11=count
  // Row 8:  col10='Between 02 to 05%', col11=count
  // Row 11: col14='Issue - Status UK', col15='Count' | col17='Issue - Status US', col18='Count'
  // Row 12 onward: Issue UK at col14/15, Issue US at col17/18

  // Total devices
  if (rows[3]) {
    kpi.totalDevices = toNum(rows[3][3]);
    kpi.totalFloors  = toNum(rows[3][8]);
  }
  if (rows[6]) {
    kpi.healthyDevices   = toNum(rows[6][3]);
    kpi.healthyFloors    = toNum(rows[6][8]);
  }
  if (rows[8]) {
    kpi.unhealthyDevices = toNum(rows[8][3]);
    kpi.unhealthyFloors  = toNum(rows[8][8]);
  }
  // UK / US device totals (rows 14)
  if (rows[14]) {
    kpi.totalUK = toNum(rows[14][1]);
    kpi.totalUS = toNum(rows[20][1]);
  }

  // Floor Summary UK — col10=label, col11=count — rows 7–12
  kpi.floorSummaryUK = {};
  const floorLabels = ['Above 60%','Between 02 to 05%','Between 05 to 10%','Between 10 to 25%','Between 25 to 60%','Healthy Floors'];
  for (let i=3; i<rows.length; i++) {
    const r = rows[i];
    const label = String(r[10]||'').trim();
    if (floorLabels.includes(label)) kpi.floorSummaryUK[label] = toNum(r[11]);
    // Stop after we've got them all
    if (Object.keys(kpi.floorSummaryUK).length >= 6) break;
  }

  // Issue Status UK (col14=label, col15=count) from row 12 onward until empty
  kpi.issueUK = {};
  kpi.issueUS = {};
  let foundIssueHeader = false;
  for (let i=0; i<rows.length; i++) {
    const r = rows[i];
    if (String(r[14]||'').trim() === 'Issue - Status UK') { foundIssueHeader = true; continue; }
    if (!foundIssueHeader) continue;
    // UK issues at col14/15, US issues at col17/18
    const ukLabel = String(r[14]||'').trim();
    const ukCount = toNum(r[15]);
    if (ukLabel && !isNaN(ukCount) && ukCount > 0) kpi.issueUK[ukLabel] = ukCount;
    const usLabel = String(r[17]||'').trim();
    const usCount = toNum(r[18]);
    if (usLabel && !isNaN(usCount) && usCount > 0) kpi.issueUS[usLabel] = usCount;
    // Stop when both columns are blank for 3 rows
    if (!ukLabel && !usLabel) break;
  }

  // Device types — look for 'Device Type' header
  kpi.deviceTypes = [];
  for (let i=0; i<rows.length; i++) {
    if (String(rows[i][10]||'').trim() === 'Device Type') {
      for (let k=i+1; k<Math.min(i+8,rows.length); k++) {
        const dt = String(rows[k][10]||'').trim();
        if (!dt || dt==='Total') continue;
        const total = toNum(rows[k][11]);
        const offline = toNum(rows[k][12]);
        if (!isNaN(total) && total > 0) kpi.deviceTypes.push({ type:dt, total, offline: isNaN(offline)?0:offline });
      }
      break;
    }
  }

  // MS Issues
  kpi.msIssues = [];
  for (let i=0; i<rows.length; i++) {
    if (String(rows[i][13]||'').trim() === 'Media Server' || String(rows[i][14]||'').trim() === 'MS Issues') {
      for (let k=i+1; k<Math.min(i+12,rows.length); k++) {
        const acc = String(rows[k][13]||'').trim();
        const cnt = toNum(rows[k][14]);
        if (acc && !isNaN(cnt) && cnt > 0) kpi.msIssues.push({ account:acc, count:cnt });
      }
      break;
    }
  }

  kpi.healthyPct = (kpi.totalDevices && kpi.healthyDevices) ? kpi.healthyDevices / kpi.totalDevices : 0;
  kpi.unhealthyPct = 1 - kpi.healthyPct;
  return kpi;
}

// ============================================================
// RENDER OVERVIEW
// ============================================================
function renderOverview() {
  const k = DATA.kpis;

  // KPI cards
  const healthColor = k.healthyPct >= 0.95 ? 'healthy' : k.healthyPct >= 0.85 ? 'warning' : 'danger';
  document.getElementById('kpi-grid').innerHTML = `
    <div class="kpi-card ${healthColor}">
      <div class="kpi-label">Total Devices</div>
      <div class="kpi-value">${fmt(k.totalDevices)}</div>
      <div class="kpi-sub">Under DHM Monitoring</div>
    </div>
    <div class="kpi-card ${healthColor}">
      <div class="kpi-label">Device Health</div>
      <div class="kpi-value">${pct(k.healthyPct)}</div>
      <div class="kpi-sub">${fmt(k.healthyDevices)} healthy · ${fmt(k.unhealthyDevices)} unhealthy</div>
    </div>
    <div class="kpi-card neutral">
      <div class="kpi-label">UK Devices</div>
      <div class="kpi-value">${fmt(k.totalUK)}</div>
      <div class="kpi-sub">UK Portal</div>
    </div>
    <div class="kpi-card neutral">
      <div class="kpi-label">US Devices</div>
      <div class="kpi-value">${fmt(k.totalUS)}</div>
      <div class="kpi-sub">US Portal</div>
    </div>
    <div class="kpi-card neutral">
      <div class="kpi-label">Total Floors</div>
      <div class="kpi-value">${fmt(k.totalFloors)}</div>
      <div class="kpi-sub">${fmt(k.healthyFloors)} healthy · ${fmt(k.unhealthyFloors)} unhealthy</div>
    </div>
    <div class="kpi-card ${DATA.msSummary.ukOffline + DATA.msSummary.usOffline > 10 ? 'warning' : 'healthy'}">
      <div class="kpi-label">Media Servers Offline</div>
      <div class="kpi-value">${fmt(DATA.msSummary.ukOffline + DATA.msSummary.usOffline)}</div>
      <div class="kpi-sub">of ${fmt(DATA.msSummary.ukTotal + DATA.msSummary.usTotal)} total deployed</div>
    </div>
  `;

  // Global device bars (UK + US accounts combined)
  const allAccounts = [...DATA.accountsUK, ...DATA.accountsUS]
    .sort((a,b) => b.total - a.total).slice(0,8);
  document.getElementById('global-device-bars').innerHTML = allAccounts.map(a => {
    const h = isNaN(a.pct) ? 0 : a.pct;
    const cls = h >= 0.95 ? 'fill-green' : h >= 0.80 ? 'fill-amber' : 'fill-red';
    return `<div class="progress-wrap">
      <div class="progress-label"><span>${a.account}</span><span>${pct(h)}</span></div>
      <div class="progress-bar"><div class="progress-fill ${cls}" style="width:${Math.min(h*100,100)}%"></div></div>
    </div>`;
  }).join('');

  // Floor buckets UK
  const fuk = k.floorSummaryUK || {};
  document.getElementById('floor-buckets-uk').innerHTML = `
    <div class="bucket-card healthy"><div class="bucket-num">${fuk['Healthy Floors']||0}</div><div class="bucket-label">Healthy</div></div>
    <div class="bucket-card warning"><div class="bucket-num">${fuk['Between 02 to 05%']||0}</div><div class="bucket-label">2–5% offline</div></div>
    <div class="bucket-card warning"><div class="bucket-num">${fuk['Between 05 to 10%']||0}</div><div class="bucket-label">5–10% offline</div></div>
    <div class="bucket-card danger"><div class="bucket-num">${fuk['Between 10 to 25%']||0}</div><div class="bucket-label">10–25% offline</div></div>
    <div class="bucket-card danger"><div class="bucket-num">${fuk['Between 25 to 60%']||0}</div><div class="bucket-label">25–60% offline</div></div>
    <div class="bucket-card danger"><div class="bucket-num">${fuk['Above 60%']||0}</div><div class="bucket-label">Above 60%</div></div>
  `;

  // Issue lists
  renderIssueList('issue-list-uk', k.issueUK||{});
  renderIssueList('issue-list-us', k.issueUS||{});

  // Device types
  const dt = k.deviceTypes||[];
  document.getElementById('device-type-panel').innerHTML = dt.length ? dt.map(d => {
    const pctH = d.total ? (d.total-d.offline)/d.total : 0;
    const cls = pctH >= 0.95 ? 'fill-green' : pctH >= 0.80 ? 'fill-amber' : 'fill-red';
    return `<div class="progress-wrap">
      <div class="progress-label"><span>${d.type}</span><span>${fmt(d.total)} total · ${fmt(d.offline)} offline</span></div>
      <div class="progress-bar"><div class="progress-fill ${cls}" style="width:${Math.min(pctH*100,100)}%"></div></div>
    </div>`;
  }).join('') : '<div class="empty">No device type data</div>';

  // MS summary
  const ms = DATA.msSummary;
  const ukPct = ms.ukTotal ? ms.ukOnline/ms.ukTotal : 0;
  const usPct = ms.usTotal ? ms.usOnline/ms.usTotal : 0;
  document.getElementById('ms-summary-panel').innerHTML = `
    <div class="progress-wrap">
      <div class="progress-label"><span>UK Media Servers</span><span>${ms.ukOnline}/${ms.ukTotal} online</span></div>
      <div class="progress-bar"><div class="progress-fill ${ukPct>=0.95?'fill-green':'fill-amber'}" style="width:${ukPct*100}%"></div></div>
    </div>
    <div class="progress-wrap">
      <div class="progress-label"><span>US Media Servers</span><span>${ms.usOnline}/${ms.usTotal} online</span></div>
      <div class="progress-bar"><div class="progress-fill ${usPct>=0.95?'fill-green':'fill-amber'}" style="width:${usPct*100}%"></div></div>
    </div>
    ${(k.msIssues||[]).length > 0 ? `<div style="margin-top:16px" class="section-title">Accounts with MS Issues</div>
    ${k.msIssues.map(m=>`<div class="progress-wrap"><div class="progress-label"><span>${m.account}</span><span>${m.count}</span></div></div>`).join('')}` : ''}
  `;
}

function renderIssueList(elId, data) {
  const entries = Object.entries(data).filter(([k,v])=>v>0).sort((a,b)=>b[1]-a[1]);
  const max = entries.length ? entries[0][1] : 1;
  document.getElementById(elId).innerHTML = entries.length
    ? entries.map(([label, count]) => `
      <div class="issue-row">
        <span class="issue-name">${label}</span>
        <div class="issue-bar-bg"><div class="issue-bar-fill" style="width:${(count/max*100).toFixed(0)}%"></div></div>
        <span class="issue-count">${count}</span>
      </div>`).join('')
    : '<div class="empty">No issue data</div>';
}

// ============================================================
// ACCOUNTS TABLE
// ============================================================
let accSortKey='total', accSortDir=-1;
function sortTable(table, key) {
  if (table==='accounts') { if(accSortKey===key) accSortDir*=-1; else { accSortKey=key; accSortDir=-1; } renderAccountsTable(); }
  if (table==='ms') { if(msSortKey===key) msSortDir*=-1; else { msSortKey=key; msSortDir=-1; } renderMSTable(); }
  if (table==='mc') { if(mcSortKey===key) mcSortDir*=-1; else { mcSortKey=key; mcSortDir=-1; } renderMCTable(); }
  if (table==='fluk') { if(flukSortKey===key) flukSortDir*=-1; else { flukSortKey=key; flukSortDir=-1; } renderFloorTable('uk'); }
}

let accPage=1;
function renderAccountsTable() {
  const search = document.getElementById('acc-search').value.toLowerCase();
  const region = document.getElementById('acc-region').value;
  const hf = document.getElementById('acc-health-filter').value;

  let rows = [];
  if (region !== 'us') DATA.accountsUK.forEach(a => rows.push({...a, region:'UK'}));
  if (region !== 'uk') DATA.accountsUS.forEach(a => rows.push({...a, region:'US'}));

  if (search) rows = rows.filter(r => r.account.toLowerCase().includes(search));
  if (hf==='critical') rows = rows.filter(r => r.pct < 0.80);
  if (hf==='warning') rows = rows.filter(r => r.pct >= 0.80 && r.pct < 0.95);
  if (hf==='good') rows = rows.filter(r => r.pct >= 0.95);

  const keyMap = {account:'account',region:'region',total:'total',healthy:'healthy',unhealthy:'unhealthy',pct:'pct'};
  rows.sort((a,b)=>{
    const va=a[keyMap[accSortKey]]||0, vb=b[keyMap[accSortKey]]||0;
    return typeof va==='string' ? va.localeCompare(vb)*accSortDir : (va-vb)*accSortDir;
  });

  const total = rows.length;
  const pages = Math.ceil(total/PAGE_SIZE);
  if (accPage > pages) accPage=1;
  const slice = rows.slice((accPage-1)*PAGE_SIZE, accPage*PAGE_SIZE);

  document.getElementById('accounts-tbody').innerHTML = slice.map(r => {
    const h = isNaN(r.pct) ? 0 : r.pct;
    const cls = h >= 0.95 ? 'fill-green' : h >= 0.80 ? 'fill-amber' : 'fill-red';
    const badgeCls = h >= 0.95 ? 'badge-green' : h >= 0.80 ? 'badge-amber' : 'badge-red';
    return `<tr>
      <td><strong>${r.account}</strong></td>
      <td><span class="badge ${r.region==='UK'?'badge-blue':'badge-grey'}">${r.region}</span></td>
      <td class="mono">${fmt(r.total)}</td>
      <td class="mono">${fmt(r.healthy)}</td>
      <td class="mono">${fmt(r.unhealthy)}</td>
      <td>
        <div class="health-cell">
          <div class="health-bar-mini"><div class="health-bar-mini-fill ${cls}" style="width:${Math.min(h*100,100)}%"></div></div>
          <span class="badge ${badgeCls}">${pct(h)}</span>
        </div>
      </td>
    </tr>`;
  }).join('') || '<tr><td colspan="6" class="empty">No results</td></tr>';

  renderPagination('acc-pagination', pages, accPage, p => { accPage=p; renderAccountsTable(); });
}

// ============================================================
// FLOOR TABLES
// ============================================================
let flukSortKey='unhealthy', flukSortDir=-1;
let flusSortKey='unhealthy', flusSortDir=-1;
let flukPage=1, flusPage=1;

function renderFloorTable(region) {
  const isUK = region==='uk';
  const prefix = isUK ? 'fluk' : 'flus';
  const searchEl = document.getElementById(`${prefix}-search`);
  const accEl = document.getElementById(`${prefix}-account`);
  const healthEl = document.getElementById(`${prefix}-health`);
  const tbodyId = `${prefix}-tbody`;
  const paginationId = `${prefix}-pagination`;

  const search = searchEl ? searchEl.value.toLowerCase() : '';
  const accFilter = accEl ? accEl.value : 'all';
  const healthFilter = healthEl ? healthEl.value : 'all';

  let rows = isUK ? [...DATA.floorsUK] : [...DATA.floorsUS];
  if (search) rows = rows.filter(r =>
    (r.account+r.location+r.building+r.floor).toLowerCase().includes(search));
  if (accFilter !== 'all') rows = rows.filter(r => r.account === accFilter);
  if (healthFilter !== 'all') rows = rows.filter(r => r.band === healthFilter);

  // sort
  const sk = isUK ? flukSortKey : flusSortKey;
  const sd = isUK ? flukSortDir : flusSortDir;
  const keyMap = {account:'account', unhealthy:'unhealthy', total:'total'};
  if (keyMap[sk]) rows.sort((a,b)=>{
    const va=a[keyMap[sk]]||0, vb=b[keyMap[sk]]||0;
    return typeof va==='string'?va.localeCompare(vb)*sd:(va-vb)*sd;
  });

  const total = rows.length;
  const pages = Math.ceil(total/PAGE_SIZE)||1;
  let curPage = isUK ? flukPage : flusPage;
  if (curPage > pages) curPage=1;
  if (isUK) flukPage=curPage; else flusPage=curPage;

  const slice = rows.slice((curPage-1)*PAGE_SIZE, curPage*PAGE_SIZE);

  document.getElementById(tbodyId).innerHTML = slice.map(r => {
    const bandCls = r.band==='Healthy Floors' ? 'badge-green'
      : r.band.includes('02')||r.band.includes('05') ? 'badge-amber' : 'badge-red';
    const statusCls = getStatusBadge(r.status);
    return `<tr>
      <td><strong>${r.account}</strong></td>
      <td>${r.location}</td>
      <td>${r.building}</td>
      <td>${r.floor}</td>
      <td class="mono">${fmt(r.total)}</td>
      <td class="mono">${r.unhealthy > 0 ? `<span style="color:var(--red)">${fmt(r.unhealthy)}</span>` : '0'}</td>
      <td><span class="badge ${bandCls}">${r.band||'—'}</span></td>
      ${isUK ? `<td><span style="font-size:0.75rem;color:var(--text-muted)">${r.issue||'—'}</span></td>` : ''}
      <td><span class="badge ${statusCls}" style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.status||'—'}</span></td>
    </tr>`;
  }).join('') || `<tr><td colspan="${isUK?9:8}" class="empty">No results</td></tr>`;

  renderPagination(paginationId, pages, curPage, p => {
    if (isUK) flukPage=p; else flusPage=p;
    renderFloorTable(region);
  });
}

// ============================================================
// MS TABLE
// ============================================================
let msSortKey='days', msSortDir=-1, msPage=1;

function renderMSTable() {
  const search = document.getElementById('ms-search').value.toLowerCase();
  const statusFilter = document.getElementById('ms-status-filter').value;

  let rows = [...DATA.msRows];
  if (search) rows = rows.filter(r =>
    (r.account+r.location+r.floor+r.msName).toLowerCase().includes(search));
  if (statusFilter==='offline') rows = rows.filter(r => r.status==='Offline');
  if (statusFilter==='online') rows = rows.filter(r => r.status==='Online');

  const keyMap = {days:'daysOffline'};
  if (keyMap[msSortKey]) rows.sort((a,b)=>(a[keyMap[msSortKey]]-b[keyMap[msSortKey]])*msSortDir);

  const total = rows.length;
  const pages = Math.ceil(total/PAGE_SIZE)||1;
  if (msPage > pages) msPage=1;
  const slice = rows.slice((msPage-1)*PAGE_SIZE, msPage*PAGE_SIZE);

  // global stats
  const offline = DATA.msRows.filter(r=>r.status==='Offline').length;
  const online = DATA.msRows.length - offline;
  document.getElementById('ms-global-stats').innerHTML = `
    <div class="global-stat"><div class="global-stat-label">Total MS</div><div class="global-stat-val">${DATA.msSummary.ukTotal + DATA.msSummary.usTotal}</div></div>
    <div class="global-stat"><div class="global-stat-label">UK Online</div><div class="global-stat-val" style="color:var(--green)">${DATA.msSummary.ukOnline}</div></div>
    <div class="global-stat"><div class="global-stat-label">UK Offline</div><div class="global-stat-val" style="color:var(--red)">${DATA.msSummary.ukOffline}</div></div>
    <div class="global-stat"><div class="global-stat-label">US Online</div><div class="global-stat-val" style="color:var(--green)">${DATA.msSummary.usOnline}</div></div>
    <div class="global-stat"><div class="global-stat-label">US Offline</div><div class="global-stat-val" style="color:var(--red)">${DATA.msSummary.usOffline}</div></div>
  `;

  document.getElementById('ms-tbody').innerHTML = slice.map(r => {
    const daysCls = r.daysOffline === 0 ? 'days-ok' : r.daysOffline <= 7 ? 'days-warn' : 'days-crit';
    const statusCls = r.status==='Offline' ? 'ms-status-offline' : 'ms-status-online';
    const tsCls = getStatusBadge(r.ticketStatus);
    return `<tr>
      <td><strong>${r.account}</strong></td>
      <td>${r.location}</td>
      <td>${r.floor}</td>
      <td style="font-size:0.75rem;max-width:200px">${r.msName}</td>
      <td><span class="days-badge ${daysCls}">${r.daysOffline}d</span></td>
      <td class="${statusCls}">${r.status}</td>
      <td><span class="badge ${tsCls}">${r.ticketStatus||'—'}</span></td>
      <td style="font-size:0.72rem;color:var(--text-muted);max-width:220px">${r.comment||'—'}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="8" class="empty">No results</td></tr>';

  renderPagination('ms-pagination', pages, msPage, p => { msPage=p; renderMSTable(); });
}

// ============================================================
// MASTERCARD TABLE
// ============================================================
let mcSortKey='pct', mcSortDir=-1, mcPage=1;

function renderMCTable() {
  const search = document.getElementById('mc-search').value.toLowerCase();
  const pf = document.getElementById('mc-priority-filter').value;

  let rows = [...DATA.mcRows];
  if (search) rows = rows.filter(r =>
    (r.location+r.building).toLowerCase().includes(search));
  if (pf !== 'all') rows = rows.filter(r => r.priority.toLowerCase() === pf.toLowerCase());

  const keyMap = {location:'location',total:'total',offline:'notConnected',pct:'pct',priority:'priority'};
  rows.sort((a,b)=>{
    const va=a[keyMap[mcSortKey]]||0, vb=b[keyMap[mcSortKey]]||0;
    return typeof va==='string'?va.localeCompare(vb)*mcSortDir:(va-vb)*mcSortDir;
  });

  // global stats
  const totalMCDevices = DATA.mcRows.reduce((s,r)=>s+r.total,0);
  const totalMCOffline = DATA.mcRows.reduce((s,r)=>s+(r.notConnected||0),0);
  const highPriority = DATA.mcRows.filter(r=>r.priority==='High').length;
  document.getElementById('mc-global-stats').innerHTML = `
    <div class="global-stat"><div class="global-stat-label">Total Locations</div><div class="global-stat-val">${DATA.mcRows.length}</div></div>
    <div class="global-stat"><div class="global-stat-label">Total Devices</div><div class="global-stat-val">${fmt(totalMCDevices)}</div></div>
    <div class="global-stat"><div class="global-stat-label">Not Connected</div><div class="global-stat-val" style="color:var(--red)">${fmt(totalMCOffline)}</div></div>
    <div class="global-stat"><div class="global-stat-label">High Priority</div><div class="global-stat-val" style="color:var(--amber)">${highPriority}</div></div>
    <div class="global-stat"><div class="global-stat-label">Overall Offline %</div><div class="global-stat-val" style="color:${totalMCOffline/totalMCDevices>0.15?'var(--red)':'var(--amber)'}">${totalMCDevices?pct(totalMCOffline/totalMCDevices):'—'}</div></div>
  `;

  const total = rows.length;
  const pages = Math.ceil(total/PAGE_SIZE)||1;
  if (mcPage>pages) mcPage=1;
  const slice = rows.slice((mcPage-1)*PAGE_SIZE, mcPage*PAGE_SIZE);

  document.getElementById('mc-tbody').innerHTML = slice.map(r => {
    const pctVal = isNaN(r.pct)?0:r.pct;
    const pctCls = pctVal >= 0.30 ? 'badge-red' : pctVal >= 0.10 ? 'badge-amber' : 'badge-green';
    const prCls = `mc-priority-${r.priority}`;
    return `<tr>
      <td><strong>${r.location}</strong></td>
      <td>${r.building}</td>
      <td class="mono">${fmt(r.total)}</td>
      <td class="mono">${r.notConnected>0?`<span style="color:var(--red)">${fmt(r.notConnected)}</span>`:'0'}</td>
      <td><span class="badge ${pctCls}">${pct(pctVal)}</span></td>
      <td style="font-size:0.75rem;color:var(--text-muted)">${r.batSwap||'—'}</td>
      <td class="${prCls}">${r.priority||'—'}</td>
      <td style="font-size:0.72rem;color:var(--text-muted);max-width:220px">${r.comment||'—'}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="8" class="empty">No results</td></tr>';

  renderPagination('mc-pagination', pages, mcPage, p => { mcPage=p; renderMCTable(); });
}

// ============================================================
// POPULATE FILTERS
// ============================================================
function populateFilters() {
  // UK account dropdown
  const ukAccounts = [...new Set(DATA.floorsUK.map(r=>r.account))].sort();
  const flukAcc = document.getElementById('fluk-account');
  flukAcc.innerHTML = '<option value="all">All Accounts</option>' +
    ukAccounts.map(a=>`<option value="${a}">${a}</option>`).join('');

  // US account dropdown
  const usAccounts = [...new Set(DATA.floorsUS.map(r=>r.account))].sort();
  const flusAcc = document.getElementById('flus-account');
  flusAcc.innerHTML = '<option value="all">All Accounts</option>' +
    usAccounts.map(a=>`<option value="${a}">${a}</option>`).join('');
}

// ============================================================
// HELPERS
// ============================================================
function fmt(n) {
  if (n===undefined||n===null||isNaN(n)) return '—';
  return Number(n).toLocaleString();
}
function pct(n) {
  if (n===undefined||n===null||isNaN(n)) return '—';
  return (n*100).toFixed(1)+'%';
}
function toNum(v) {
  if (v===null||v===undefined||v==='') return NaN;
  const n = parseFloat(String(v).replace(/,/g,''));
  return isNaN(n) ? NaN : n;
}
function getStatusBadge(status) {
  if (!status) return 'badge-grey';
  const s = status.toLowerCase();
  if (s.includes('hold')) return 'badge-amber';
  if (s.includes('customer')||s.includes('account team')||s.includes('project manager')) return 'badge-blue';
  if (s.includes('factory')||s.includes('partnership')) return 'badge-amber';
  if (s.includes('resolved')) return 'badge-green';
  if (s.includes('dhm')||s.includes('service desk')||s.includes('field')) return 'badge-grey';
  return 'badge-grey';
}

function renderPagination(elId, pages, current, onPage) {
  if (pages <= 1) { document.getElementById(elId).innerHTML=''; return; }
  let html = `<span class="page-info">${current}/${pages}</span>`;
  html += `<button class="page-btn" onclick="(${onPage.toString()})(${Math.max(1,current-1)})">‹</button>`;
  for (let p=1; p<=pages; p++) {
    if (pages<=7 || Math.abs(p-current)<=1 || p===1 || p===pages) {
      html += `<button class="page-btn${p===current?' active':''}" onclick="(${onPage.toString()})(${p})">${p}</button>`;
    } else if (Math.abs(p-current)===2) {
      html += `<span style="color:var(--text-muted);padding:0 4px">…</span>`;
    }
  }
  html += `<button class="page-btn" onclick="(${onPage.toString()})(${Math.min(pages,current+1)})">›</button>`;
  document.getElementById(elId).innerHTML = html;
}

function showTab(id) {
  document.querySelectorAll('.tab').forEach((t,i)=>{
    const ids=['overview','accounts','floors-uk','floors-us','media-servers','mastercard'];
    t.classList.toggle('active', ids[i]===id);
  });
  document.querySelectorAll('.tab-panel').forEach(p=>{
    p.classList.toggle('active', p.id===`panel-${id}`);
  });
}

function resetDashboard() {
  DATA = {};
  document.getElementById('upload-screen').style.display='flex';
  document.getElementById('dashboard').style.display='none';
  document.getElementById('file-input').value='';
  document.getElementById('parse-error').style.display='none';
}
