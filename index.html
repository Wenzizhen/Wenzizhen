<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>互動式工程進度計算器</title>
<style>
  body{font-family: system-ui,-apple-system,'Segoe UI',Roboto,'Noto Sans TC',sans-serif;margin:18px;background:#f7f7fb;color:#111}
  h1{font-size:20px;margin-bottom:6px}
  .controls{display:flex;gap:12px;align-items:center;margin-bottom:8px;flex-wrap:wrap}
  table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d6dbe6}
  th,td{padding:6px 8px;text-align:left;font-size:13px;border:1px solid #d6dbe6}
  input[type=date], input[type=text], input[type=number]{width:100%;box-sizing:border-box;padding:6px;border:1px solid #d6dbe6;border-radius:4px}
  button{padding:8px 10px;border-radius:6px;border:0;background:#2563eb;color:#fff;cursor:pointer}
  button.secondary{background:#6b7280}
  .small{padding:6px 8px;font-size:13px}
  .note{font-size:12px;color:#374151}
  .gantt{margin-top:12px;overflow:auto;background:#fff;padding:10px;border:1px solid #e2e8f0;position:relative;font-size:12px}
  .gantt table{border-collapse:collapse;width:100%}
  .gantt th, .gantt td{border:1px solid #d6dbe6;text-align:center;padding:2px 4px;font-size:12px;white-space:nowrap}
  .gantt th.date{font-weight:500;background:#fafafa}
  .gantt td.task-name{text-align:left;padding-left:6px;font-weight:600;white-space:nowrap}
  .gantt td.bar{background:#4a90e2;color:#fff;font-weight:600;vertical-align:middle;text-align:center}
  .gantt td.empty { color: #fff; background: transparent; }
  .gantt th, .gantt td { min-width: 16px; height: 20px; }

  .summary-table{margin-top:8px;border-collapse:collapse;width:100%}
  .summary-table th, .summary-table td{border:1px solid #d6dbe6;padding:6px 8px;font-size:13px;text-align:left}
  .summary-table th:nth-child(3),
  .summary-table td:nth-child(3){ width: 90px; max-width:90px; }
  .summary-table th:nth-child(4),
  .summary-table td:nth-child(4){ width: 60px; max-width:60px; text-align:left; }
  .summary-table th:nth-child(5),
  .summary-table td:nth-child(5){ width: 90px; max-width:90px; }
</style>
</head>
<body>
<h1>互動式工程進度計算器</h1>

<div class="controls">
  <button id="addRow" class="small secondary">＋ 新增工項</button>
  <button id="compute" class="small">計算</button>
  <button id="exportXlsx" class="small">匯出 Excel</button>
  <button id="exportGantt" class="small secondary">匯出甘特圖</button>
  <input type="file" id="importFile" class="small" accept=".xlsx" />
</div>

<div class="note">提示：在「前置工項」欄位輸入對應工項的編號，或留空。</div>

<table id="taskTable">
  <thead>
    <tr>
      <th style="width:46px">編號</th>
      <th>工項</th>
      <th>剩餘數量</th>
      <th>單位</th>
      <th>每週工率</th>
      <th>開始施作日</th>
      <th>預估工期(天)</th>
      <th>前置工項(編號)</th>
      <th>備註</th>
      <th>操作</th>
    </tr>
  </thead>
  <tbody></tbody>
</table>

<div class="gantt" id="resultArea">
  <h3>計算結果</h3>
  <div id="summary"></div>
  <div id="ganttContainer"></div>
</div>

<script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<script>
const tbody = document.querySelector('#taskTable tbody');

function addRow(data, insertBeforeTr){
  const tr = document.createElement('tr');
  const index = insertBeforeTr ? Array.from(tbody.children).indexOf(insertBeforeTr) + 1 : tbody.children.length + 1;
  tr.innerHTML = `
    <td class="idx">${index}</td>
    <td><input type="text" class="task" value="${data?.task||''}"/></td>
    <td><input type="number" class="qty" value="${data?.qty||''}"/></td>
    <td><input type="text" class="unit" value="${data?.unit||''}"/></td>
    <td><input type="number" class="rate_week" value="${data?.rate_week||''}"/></td>
    <td><input type="date" class="start" value="${data?.start||''}"/></td>
    <td><input type="number" class="days" value="${data?.days||''}"/></td>
    <td><input type="text" class="dep" value="${data?.dep||''}"/></td>
    <td><input type="text" class="note" value="${data?.note||''}"/></td>
    <td class="row-actions">
      <button class="insert small secondary">插入一列</button>
      <button class="del small secondary">刪除</button>
    </td>
  `;
  if(insertBeforeTr) tbody.insertBefore(tr, insertBeforeTr);
  else tbody.appendChild(tr);

  tr.querySelector('.del').addEventListener('click', ()=>{ tr.remove(); refreshIndices(); });
  tr.querySelector('.insert').addEventListener('click', ()=>{ addRow({}, tr); refreshIndices(); });
  refreshIndices();
}

function refreshIndices(){ Array.from(tbody.children).forEach((tr,i)=>{ tr.querySelector('.idx').textContent = i+1; }); }
document.getElementById('addRow').addEventListener('click', ()=>addRow({}));

function parseRows(){
  return Array.from(tbody.children).map((tr,i)=>({
    id:i+1,
    task:tr.querySelector('.task').value.trim(),
    qty:tr.querySelector('.qty').value? Number(tr.querySelector('.qty').value): null,
    unit:tr.querySelector('.unit').value.trim(),
    rate_week:tr.querySelector('.rate_week').value? Number(tr.querySelector('.rate_week').value): null,
    start:tr.querySelector('.start').value || null,
    days:tr.querySelector('.days').value? Number(tr.querySelector('.days').value): null,
    dep:tr.querySelector('.dep').value.trim() || null,
    note:tr.querySelector('.note').value.trim() || ''
  }));
}

function compute(){
  const rows=parseRows();
  rows.forEach(r=>{
    if(r.rate_week && r.qty){
      const perDay = r.rate_week/7;
      r.computed_days = Math.ceil(r.qty / perDay);
      r.note = r.note || `每週${r.rate_week}，每日${perDay.toFixed(2)}`;
    } else if(r.days){
      r.computed_days = r.days;
    } else {
      r.computed_days = 0;
    }
  });

  const endById = {};
  rows.forEach(r=>{
    let startDate;
    if(r.start){
      const [y,m,d] = r.start.split('-').map(Number);
      startDate = new Date(y,m-1,d);
    } else if(r.dep){
      const deps = r.dep.split(',').map(x=>Number(x.trim())).filter(Boolean);
      if(deps.length>0){
        const depEnds = deps.map(id=>endById[id]).filter(Boolean);
        if(depEnds.length>0){
          const latest = new Date(Math.max(...depEnds.map(d=>d.getTime())));
          latest.setDate(latest.getDate()+1);
          startDate = latest;
        } else startDate = new Date();
      } else startDate = new Date();
    } else startDate = new Date();
    r.actual_start = startDate;
    const end = new Date(startDate);
    end.setDate(end.getDate() + r.computed_days - 1);
    r.end = end;
    r.duration_days = r.computed_days;
    endById[r.id] = r.end;
  });

  renderTable(rows);
  renderGantt(rows);
  window._lastComputed = rows;
}

function formatDateYMD(d){
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
}

function renderTable(rows){
  const summary=document.getElementById('summary');
  summary.innerHTML='';
  const tbl=document.createElement('table');
  tbl.className = 'summary-table';
  tbl.innerHTML=`<thead><tr><th>編號</th><th>工項</th><th>開始施作日</th><th>預估工期</th><th>預估完成日</th><th>前置工項</th><th>備註</th></tr></thead>`;
  const tb=document.createElement('tbody');
  rows.forEach(r=>{
    const tr=document.createElement('tr');
    const deps = r.dep ? r.dep.split(',').map(id=>{
      const found = rows.find(x=>x.id==id);
      return found ? `${id}-${found.task}` : id;
    }).join(',') : '';
    tr.innerHTML = `<td>${r.id}</td><td>${r.task}</td><td>${formatDateYMD(r.actual_start)}</td><td>${r.duration_days}</td><td>${formatDateYMD(r.end)}</td><td>${deps}</td><td>${r.note||''}</td>`;
    tb.appendChild(tr);
  });
  tbl.appendChild(tb);
  summary.appendChild(tbl);
}

function renderGantt(rows){
  const container = document.getElementById('ganttContainer');
  container.innerHTML = '';
  if(!rows || rows.length===0){ container.innerHTML = '<div class="note">尚無資料</div>'; return; }

  const oneDayMs = 24*60*60*1000;
  const dateSet = new Set();
  rows.forEach(r=>{
    for(let i=0;i<r.duration_days;i++){
      const dt = new Date(r.actual_start.getTime() + i*oneDayMs);
      dateSet.add(dt.toDateString());
    }
  });
  const datesArray = Array.from(dateSet).map(s=>new Date(s)).sort((a,b)=>a-b);

  const tbl = document.createElement('table');
  tbl.className = 'gantt-table';

  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  const thBlank = document.createElement('th'); thBlank.textContent = ''; trHead.appendChild(thBlank);
  datesArray.forEach(d=>{
    const th = document.createElement('th'); th.className='date';
    th.textContent = `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  tbl.appendChild(thead);

  const tbodyG = document.createElement('tbody');
  rows.forEach(r=>{
    const tr = document.createElement('tr');
    const tdName = document.createElement('td'); tdName.className='task-name'; tdName.textContent = r.task; tr.appendChild(tdName);

    let i = 0;
    while(i < datesArray.length){
      const cur = datesArray[i];
      const curTime = new Date(cur.toDateString()).getTime();
      const taskStartTime = new Date(r.actual_start.toDateString()).getTime();
      const taskEndTime = new Date(r.end.toDateString()).getTime();

      if(curTime >= taskStartTime && curTime <= taskEndTime){
        let j = i;
        while(j < datesArray.length){
          const jt = new Date(datesArray[j].toDateString()).getTime();
          if(jt >= taskStartTime && jt <= taskEndTime) j++;
          else break;
        }
        const colspan = j - i;
        const td = document.createElement('td');
        td.className = 'bar';
        td.colSpan = colspan;
        td.textContent = `${r.duration_days}天`;
        tr.appendChild(td);
        i = j;
      } else {
        const td = document.createElement('td');
        td.className = 'empty';
        td.textContent = '';
        tr.appendChild(td);
        i++;
      }
    }

    tbodyG.appendChild(tr);
  });
  tbl.appendChild(tbodyG);
  container.appendChild(tbl);
}

document.getElementById('compute').addEventListener('click', compute);

function exportXlsx(){
  const rows = window._lastComputed || parseRows();
  const aoa = [["編號","工項","剩餘數量","單位","每週工率","實際開始日","預估工期","預估完成日","前置工項","備註"]];
  rows.forEach(r=>{
    const fmt = d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    const deps = r.dep ? r.dep.split(',').map(id=>{
      const found = rows.find(x=>x.id==id);
      return found ? `${id}-${found.task}` : id;
    }).join(',') : '';
    aoa.push([r.id, r.task, r.qty||'', r.unit||'', r.rate_week||'', fmt(r.actual_start), r.duration_days||'', fmt(r.end), deps, r.note||'']);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '進度');
  const wbout = XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const blob = new Blob([wbout],{type:'application/octet-stream'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = '工程進度分析.xlsx'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}
document.getElementById('exportXlsx').addEventListener('click', exportXlsx);

document.getElementById('exportGantt').addEventListener('click', ()=>{
  const container = document.getElementById('ganttContainer');
  if(!container.innerHTML.trim()) return alert('尚無甘特圖');
  html2canvas(container, {scale:2,useCORS:true,backgroundColor:'#fff',width:container.scrollWidth,height:container.scrollHeight,scrollX:0,scrollY:-container.offsetTop})
    .then(canvas=>{
      canvas.toBlob(function(blob){
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = '甘特圖.jpg';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }, 'image/jpeg', 0.95);
    });
});

document.getElementById('importFile').addEventListener('change', function(e){
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = function(evt){
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data,{type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws,{header:1});
    tbody.innerHTML = '';
    json.slice(1).forEach(row=>{
      addRow({
        task: row[1] || '',
        qty: row[2] || '',
        unit: row[3] || '',
        rate_week: row[4] || '',
        start: row[5] || '',
        days: row[6] || '',
        dep: row[8] || '',
        note: row[9] || ''
      });
    });
    refreshIndices();
  };
  reader.readAsArrayBuffer(file);
});
</script>
</body>
</html>
