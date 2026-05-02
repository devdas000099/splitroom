require('dotenv').config();
const express = require('express');
const http = require('http');
const WebSocket = require('ws');
const mongoose = require('mongoose');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

// ── MongoDB Schema ──
const expenseSchema = new mongoose.Schema({
  monthKey: { type: String, required: true, index: true },
  id: { type: Number, required: true },
  desc: String,
  amt: Number,
  cat: String,
  payer: String,
  split: String,
  mode: String,
  date: String,
  createdAt: { type: Date, default: Date.now }
});
const Expense = mongoose.model('Expense', expenseSchema);

// ── Connect MongoDB ──
const MONGO_URI = process.env.MONGO_URI;
if (!MONGO_URI) {
  console.error('\n❌  MONGO_URI not set! Add it to your .env file or Render environment variables.\n');
  process.exit(1);
}

mongoose.connect(MONGO_URI)
  .then(() => console.log('✅  MongoDB connected!'))
  .catch(err => { console.error('MongoDB error:', err); process.exit(1); });

// ── Broadcast ──
function broadcast(msg) {
  const str = JSON.stringify(msg);
  wss.clients.forEach(c => { if (c.readyState === WebSocket.OPEN) c.send(str); });
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// GET all expenses grouped by monthKey
app.get('/api/expenses', async (req, res) => {
  try {
    const all = await Expense.find({}).lean();
    const grouped = {};
    all.forEach(e => {
      if (!grouped[e.monthKey]) grouped[e.monthKey] = [];
      grouped[e.monthKey].push(e);
    });
    res.json(grouped);
  } catch(err) { res.status(500).json({ error: err.message }); }
});

// POST add expense
app.post('/api/expenses', async (req, res) => {
  try {
    const { monthKey, expense } = req.body;
    expense.id = Date.now() + Math.random();
    expense.monthKey = monthKey;
    const doc = await Expense.create(expense);
    broadcast({ type: 'ADD', monthKey, expense: doc.toObject() });
    res.json({ ok: true, expense: doc });
  } catch(err) { res.status(500).json({ error: err.message }); }
});

// DELETE one expense
app.delete('/api/expenses/:monthKey/:id', async (req, res) => {
  try {
    const { monthKey, id } = req.params;
    await Expense.deleteOne({ monthKey, id: parseFloat(id) });
    broadcast({ type: 'DELETE', monthKey, id });
    res.json({ ok: true });
  } catch(err) { res.status(500).json({ error: err.message }); }
});

// DELETE whole month
app.delete('/api/expenses/:monthKey', async (req, res) => {
  try {
    const { monthKey } = req.params;
    await Expense.deleteMany({ monthKey });
    broadcast({ type: 'CLEAR', monthKey });
    res.json({ ok: true });
  } catch(err) { res.status(500).json({ error: err.message }); }
});

// Excel export
app.get('/api/export/:monthKey', async (req, res) => {
  try {
    const { monthKey } = req.params;
    const [year, month] = monthKey.split('-');
    const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const monthName = MONTHS[parseInt(month)] || month;
    const exps = await Expense.find({ monthKey }).lean();

    const wb = XLSX.utils.book_new();
    const rows = [['Date','Description','Category','Paid By','Payment Mode','Split','Amount (₹)','Dev Share (₹)','Vatsal Share (₹)']];
    let devPaid=0,vPaid=0,devOwes=0,vOwes=0;
    const catTotals={}, modeTotals={};

    exps.forEach(e => {
      let ds=0,vs=0;
      if(e.split==='equal'){ds=e.amt/2;vs=e.amt/2;}
      else if(e.split==='dev'){ds=e.amt;}
      else{vs=e.amt;}
      if(e.payer==='dev') devPaid+=e.amt; else vPaid+=e.amt;
      devOwes+=ds; vOwes+=vs;
      catTotals[e.cat]=(catTotals[e.cat]||0)+e.amt;
      modeTotals[e.mode]=(modeTotals[e.mode]||0)+e.amt;
      rows.push([e.date,e.desc,e.cat,e.payer==='dev'?'Dev':'Vatsal',e.mode||'—',
        e.split==='equal'?'Equal':e.split==='dev'?'Only Dev':'Only Vatsal',
        e.amt, Math.round(ds*100)/100, Math.round(vs*100)/100]);
    });

    rows.push([],[' ── SUMMARY ──']);
    rows.push(['Total Spent','','','','','', devPaid+vPaid]);
    rows.push(['Dev Paid','','','','','', devPaid]);
    rows.push(['Vatsal Paid','','','','','', vPaid]);
    rows.push(['Dev Share Owed','','','','','', Math.round(devOwes*100)/100]);
    rows.push(['Vatsal Share Owed','','','','','', Math.round(vOwes*100)/100]);
    const bal = devPaid - devOwes;
    rows.push([bal>=0?'Vatsal owes Dev':'Dev owes Vatsal','','','','','', Math.abs(Math.round(bal*100)/100)]);

    const ws1 = XLSX.utils.aoa_to_sheet(rows);
    ws1['!cols']=[{wch:12},{wch:30},{wch:14},{wch:10},{wch:14},{wch:14},{wch:14},{wch:16},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws1, `${monthName} ${year}`);

    const ws2 = XLSX.utils.aoa_to_sheet([['Category','Total (₹)'],...Object.entries(catTotals).sort((a,b)=>b[1]-a[1])]);
    ws2['!cols']=[{wch:18},{wch:18}];
    XLSX.utils.book_append_sheet(wb, ws2, 'By Category');

    const ws3 = XLSX.utils.aoa_to_sheet([['Payment Mode','Total (₹)'],...Object.entries(modeTotals).sort((a,b)=>b[1]-a[1])]);
    ws3['!cols']=[{wch:18},{wch:18}];
    XLSX.utils.book_append_sheet(wb, ws3, 'By Payment Mode');

    const buf = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename=Expenses_${monthName}_${year}.xlsx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch(err) { res.status(500).json({ error: err.message }); }
});

// WebSocket
wss.on('connection', (ws) => {
  broadcast({ type: 'CLIENTS', count: wss.clients.size });
  ws.on('close', () => broadcast({ type: 'CLIENTS', count: wss.clients.size }));
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n🚀  SplitRoom — Dev & Vatsal`);
  console.log(`   Running on port ${PORT}`);
  console.log(`   http://localhost:${PORT}\n`);
});
