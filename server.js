const express = require('express');
const http = require('http');
const WebSocket = require('ws');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });
const DATA_FILE = path.join(__dirname, 'data.json');

function loadData() {
  if (fs.existsSync(DATA_FILE)) {
    try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); } catch(e) {}
  }
  return {};
}
function saveData(data) {
  try { fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2)); } catch(e) {
    console.warn('Ephemeral FS - data in memory only');
  }
}

let db = loadData();

function broadcast(msg) {
  const str = JSON.stringify(msg);
  wss.clients.forEach(c => { if (c.readyState === WebSocket.OPEN) c.send(str); });
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/expenses', (req, res) => res.json(db));

app.post('/api/expenses', (req, res) => {
  const { monthKey, expense } = req.body;
  if (!db[monthKey]) db[monthKey] = [];
  expense.id = Date.now() + Math.random();
  expense.createdAt = new Date().toISOString();
  db[monthKey].push(expense);
  saveData(db);
  broadcast({ type: 'ADD', monthKey, expense });
  res.json({ ok: true, expense });
});

app.delete('/api/expenses/:monthKey/:id', (req, res) => {
  const { monthKey, id } = req.params;
  if (db[monthKey]) {
    db[monthKey] = db[monthKey].filter(e => String(e.id) !== String(id));
    saveData(db);
    broadcast({ type: 'DELETE', monthKey, id });
  }
  res.json({ ok: true });
});

app.delete('/api/expenses/:monthKey', (req, res) => {
  const { monthKey } = req.params;
  delete db[monthKey];
  saveData(db);
  broadcast({ type: 'CLEAR', monthKey });
  res.json({ ok: true });
});

app.get('/api/export/:monthKey', (req, res) => {
  const { monthKey } = req.params;
  const [year, month] = monthKey.split('-');
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const monthName = MONTHS[parseInt(month)] || month;
  const exps = db[monthKey] || [];
  const wb = XLSX.utils.book_new();
  const rows = [['Date','Description','Category','Paid By','Payment Mode','Split','Amount','Dev Share','Vatsal Share']];
  let devPaid=0,vatsalPaid=0,devOwes=0,vatsalOwes=0;
  const catTotals={}, modeTotals={};
  exps.forEach(e => {
    let ds=0,vs=0;
    if(e.split==='equal'){ds=e.amt/2;vs=e.amt/2;}
    else if(e.split==='dev'){ds=e.amt;}else{vs=e.amt;}
    if(e.payer==='dev') devPaid+=e.amt; else vatsalPaid+=e.amt;
    devOwes+=ds; vatsalOwes+=vs;
    catTotals[e.cat]=(catTotals[e.cat]||0)+e.amt;
    modeTotals[e.mode]=(modeTotals[e.mode]||0)+e.amt;
    rows.push([e.date,e.desc,e.cat,e.payer==='dev'?'Dev':'Vatsal',e.mode||'',
      e.split==='equal'?'Equal':e.split==='dev'?'Only Dev':'Only Vatsal',
      e.amt,Math.round(ds*100)/100,Math.round(vs*100)/100]);
  });
  rows.push([],[' SUMMARY']);
  rows.push(['Total','',' ',' ',' ',' ',devPaid+vatsalPaid]);
  rows.push(['Dev Paid','',' ',' ',' ',' ',devPaid]);
  rows.push(['Vatsal Paid','',' ',' ',' ',' ',vatsalPaid]);
  const bal = devPaid - devOwes;
  rows.push([bal>=0?'Vatsal owes Dev':'Dev owes Vatsal','',' ',' ',' ',' ',Math.abs(Math.round(bal*100)/100)]);
  const ws1 = XLSX.utils.aoa_to_sheet(rows);
  ws1['!cols']=[{wch:12},{wch:30},{wch:14},{wch:10},{wch:14},{wch:14},{wch:14},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws1, `${monthName} ${year}`);
  const ws2 = XLSX.utils.aoa_to_sheet([['Category','Total'],...Object.entries(catTotals).sort((a,b)=>b[1]-a[1])]);
  XLSX.utils.book_append_sheet(wb, ws2, 'By Category');
  const ws3 = XLSX.utils.aoa_to_sheet([['Mode','Total'],...Object.entries(modeTotals).sort((a,b)=>b[1]-a[1])]);
  XLSX.utils.book_append_sheet(wb, ws3, 'By Mode');
  const buf = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename=Expenses_${monthName}_${year}.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

wss.on('connection', (ws) => {
  broadcast({ type: 'CLIENTS', count: wss.clients.size });
  ws.on('close', () => broadcast({ type: 'CLIENTS', count: wss.clients.size }));
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅  SplitRoom running on port ${PORT}\n`);
});
