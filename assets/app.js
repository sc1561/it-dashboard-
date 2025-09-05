/* =========================================================
   لوحة تقارير ودرجات — متعدد القوالب (Read-Only)
   ========================================================= */
(() => {
  const fileInput     = document.getElementById('fileInput');
  const dropzone      = document.getElementById('dropzone');
  const loader        = document.getElementById('loader');

  const stageFilter   = document.getElementById('stageFilter');
  const gradeFilter   = document.getElementById('gradeFilter');
  const sectionFilter = document.getElementById('sectionFilter');
  const sheetFilter   = document.getElementById('sheetFilter');
  const sourceFilter  = document.getElementById('sourceFilter');
  const searchInput   = document.getElementById('searchInput');

  const exportPdfBtn  = document.getElementById('exportPdfBtn');
  const exportXlsxBtn = document.getElementById('exportXlsxBtn');
  const printBtn      = document.getElementById('printBtn');
  const shareBtn      = document.getElementById('shareBtn');
  const resetBtn      = document.getElementById('resetBtn');

  const tableHead     = document.getElementById('tableHead');
  const tableBody     = document.getElementById('tableBody');
  const stats         = document.getElementById('stats');
  const readmeLink    = document.getElementById('readmeLink');

  let aggregated = [];
  let displayed  = [];
  let allSheets  = new Set();
  let allSources = new Set();

  const COLS = { grade:'grade', section:'section', name:'name', score:'score' };

  const TEMPLATES = [
    {
      id: 'stage_5_6',
      label: 'الصفوف 5-6',
      match: (fileName, sheetName) =>
        /5-6|5_6|الصفوف\s*5|الصفوف\s*6/i.test(fileName) || /5|6/.test(sheetName),
      map: { 'الصف':'grade', 'الشعبة':'section', 'الاسم':'name', 'الدرجة':'score' },
      aliases: { 'الدرجة النهائية':'الدرجة', 'اسم الطالب':'الاسم' }
    },
    {
      id: 'stage_7_10',
      label: 'الصفوف 7-10',
      match: (fileName, sheetName) =>
        /7-10|7_10|الصفوف\s*7|الصفوف\s*8|الصفوف\s*9|الصفوف\s*10/i.test(fileName) || /7|8|9|10/.test(sheetName),
      map: { 'الصف':'grade', 'الشعبة':'section', 'الاسم':'name', 'الدرجة':'score' },
      aliases: { 'المحصلة':'الدرجة', 'الطالب':'الاسم' }
    },
    {
      id: 'stage_11_12',
      label: 'الصفان 11-12',
      match: (fileName, sheetName) =>
        /11|12|الصفين\s*11|الصفين\s*12/i.test(fileName) || /11|12/.test(sheetName),
      map: { 'الصف':'grade', 'الشعبة':'section', 'الاسم':'name', 'الدرجة':'score' },
      aliases: { 'المجموع':'الدرجة', 'اسم المتعلم':'الاسم' }
    },
    {
      id: 'primary_cycle',
      label: 'الحلقة الأولى',
      match: (fileName, sheetName) =>
        /الحلقة\s*الأولى|الاولى|الأولى/i.test(fileName) || /الحلقة/.test(sheetName),
      map: { 'الصف':'grade', 'الشعبة':'section', 'الاسم':'name', 'الدرجة':'score' },
      aliases: { 'علامة':'الدرجة' }
    }
  ];

  const showLoader = (v=true) => loader.classList.toggle('show', v);
  const s = (v) => (v===undefined || v===null) ? '' : String(v);

  function findTemplate(fileName, sheetName, headers){
    const viaName = TEMPLATES.find(t => t.match(fileName, sheetName));
    if (viaName) return viaName;
    const H = new Set(headers.map(h => s(h).trim()));
    const guesses = TEMPLATES.filter(t => Object.keys(t.map).some(k => H.has(k)));
    return guesses[0] || null;
  }

  function normalizeHeaders(headers, template){
    const out = headers.map(h => s(h));
    if (!template?.aliases) return out;
    return out.map(h => (template.aliases[h] ? template.aliases[h] : h));
  }

  function mapRowToCanonical(row, template){
    const canonical = {};
    Object.keys(row).forEach(k => { canonical[k] = row[k]; });
    if (template?.map){
      Object.entries(template.map).forEach(([src, dst]) => {
        if (row.hasOwnProperty(src)) canonical[dst] = row[src];
      });
    }
    canonical[COLS.grade]   = canonical[COLS.grade]   ?? row['الصف']    ?? row['صف']    ?? '';
    canonical[COLS.section] = canonical[COLS.section] ?? row['الشعبة']  ?? row['شعبة']  ?? '';
    canonical[COLS.name]    = canonical[COLS.name]    ?? row['الاسم']   ?? row['اسم']   ?? row['اسم الطالب'] ?? '';
    canonical[COLS.score]   = canonical[COLS.score]   ?? row['الدرجة']  ?? row['علامة'] ?? row['المجموع'] ?? row['المحصلة'] ?? '';
    return canonical;
  }

  async function parseFile(file){
    const buf = await file.arrayBuffer();
    let wb;
    try{ wb = XLSX.read(buf, { type:'array', cellDates:true, raw:false, WTF:true }); }
    catch(e){
      const txt = await file.text();
      wb = XLSX.read(txt, { type:'string' });
    }
    const fileName = file.name;
    const result = [];
    (wb.SheetNames || []).forEach(sheetName => {
      const ws = wb.Sheets[sheetName];
      if (!ws) return;
      let rows = XLSX.utils.sheet_to_json(ws, { defval:'', raw:false, blankrows:false });
      if (!rows.length) return;
      const headers = Object.keys(rows[0]);
      const template = findTemplate(fileName, sheetName, headers);
      const normHeaders = normalizeHeaders(headers, template);
      rows = rows.map(r => {
        const o = {};
        normHeaders.forEach((h, i) => { o[h] = r[headers[i]]; });
        return o;
      });
      const unified = rows.map(r => {
        const u = mapRowToCanonical(r, template);
        u.__sheet  = sheetName;
        u.__source = fileName;
        u.__template = template ? template.id : 'unknown';
        return u;
      });
      result.push(...unified);
      allSheets.add(sheetName);
      allSources.add(fileName);
    });
    return result;
  }

  async function handleFiles(fileList){
    if (!fileList?.length) return;
    showLoader(true);
    try{
      aggregated = [];
      allSheets.clear(); allSources.clear();
      for (const file of fileList){
        const part = await parseFile(file);
        aggregated.push(...part);
      }
  updateFiltersByStage();
  applyFiltersAndRender();
      stats.textContent = `تم تحميل ${fileList.length} ملف(ات) — السجلات: ${aggregated.length}`;
    }catch(e){
      console.error(e);
      alert('تعذر قراءة بعض الملفات. تأكد من الصيغة أو احفظ كـ .xlsx ثم أعد الرفع.');
    }finally{
      showLoader(false);
    }
  }

  function uniqueValues(rows, key){
    const vals = new Set();
    rows.forEach(r => { if (s(r[key])) vals.add(s(r[key])); });
    return Array.from(vals).filter(Boolean).sort((a,b)=> (''+a).localeCompare((''+b),'ar', {numeric:true}));
  }

  function fillSelect(selectEl, options){
    const current = selectEl.value || '';
    selectEl.innerHTML = '<option value="">الكل</option>';
    options.forEach(v=>{
      const opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      selectEl.appendChild(opt);
    });
    if ([...selectEl.options].some(o=>o.value===current)) selectEl.value = current;
  }

  function applyFiltersAndRender(){
  const stage = stageFilter.value.trim();
  const g = gradeFilter.value.trim();
  const se = sectionFilter.value.trim();
  const sh = sheetFilter.value.trim();
  const so = sourceFilter.value.trim();
  const q = searchInput.value.trim().toLowerCase();

  const byStage = stage ? (r => r.__template === stage) : ()=>true;
  const byG  = g  ? (r => s(r.grade)   === g)  : ()=>true;
  const bySe = se ? (r => s(r.section) === se) : ()=>true;
  const bySh = sh ? (r => s(r.__sheet) === sh) : ()=>true;
  const bySo = so ? (r => s(r.__source)=== so) : ()=>true;
  const byQ  = q  ? (r => s(r.name).toLowerCase().includes(q)) : ()=>true;

  displayed = aggregated.filter(r => byStage(r) && byG(r) && bySe(r) && bySh(r) && bySo(r) && byQ(r));
  renderTable(displayed);
  const filters = [stage&&`المرحلة=${stageFilter.options[stageFilter.selectedIndex].text}`,g&&`الصف=${g}`, se&&`الشعبة=${se}`, sh&&`الورقة=${sh}`, so&&`المصدر=${so}`].filter(Boolean).join(', ');
  stats.textContent = `عدد السجلات: ${displayed.length} من ${aggregated.length}${filters?` (${filters})`:''}${q?` — بحث: "${q}"`:''}`;

  // تحديث الفلاتر الأخرى عند تغيير المرحلة
  updateFiltersByStage();
  function updateFiltersByStage() {
    const stage = stageFilter.value.trim();
    let filtered = aggregated;
    if (stage) filtered = aggregated.filter(r => r.__template === stage);
    fillSelect(sheetFilter, uniqueValues(filtered, '__sheet'));
    fillSelect(sourceFilter, uniqueValues(filtered, '__source'));
    fillSelect(gradeFilter, uniqueValues(filtered, 'grade'));
    fillSelect(sectionFilter, uniqueValues(filtered, 'section'));
  }
  stageFilter.addEventListener('change', () => {
    updateFiltersByStage();
    applyFiltersAndRender();
  });
  }

  function renderTable(rows){
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    if (!rows.length){
      tableHead.innerHTML = '<tr><th>لا توجد بيانات مطابقة</th></tr>';
      return;
    }
    const baseKeys = ['grade','section','name','score','__sheet','__source','__template'];
    const extraKeys = Array.from(rows.reduce((acc, r)=>{
      Object.keys(r).forEach(k=>{ if (!baseKeys.includes(k) && !k.startsWith('__')) acc.add(k); });
      return acc;
    }, new Set()));
    const keys = [...baseKeys, ...extraKeys];

    const headRow = document.createElement('tr');
    keys.forEach(k=>{
      const th = document.createElement('th');
      th.textContent = k;
      th.style.cursor = 'pointer';
      th.title = 'انقر للفرز';
      th.addEventListener('click', () => {
        const asc = th.dataset.sort !== 'asc';
        rows.sort((a,b)=>{
          const va = s(a[k]), vb = s(b[k]);
          return asc ? va.localeCompare(vb, 'ar', {numeric:true}) : vb.localeCompare(va, 'ar', {numeric:true});
        });
        [...headRow.children].forEach(h => delete h.dataset.sort);
        th.dataset.sort = asc ? 'asc' : 'desc';
        renderTable(rows);
      });
      headRow.appendChild(th);
    });
    tableHead.appendChild(headRow);

    const frag = document.createDocumentFragment();
    rows.forEach(r=>{
      const tr = document.createElement('tr');
      keys.forEach(k=>{
        const td = document.createElement('td');
        td.textContent = s(r[k]);
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    });
    tableBody.appendChild(frag);
  }

  sheetFilter.addEventListener('change', applyFiltersAndRender);
  sourceFilter.addEventListener('change', applyFiltersAndRender);
  gradeFilter.addEventListener('change', applyFiltersAndRender);
  sectionFilter.addEventListener('change', applyFiltersAndRender);

  let t=null;
  searchInput.addEventListener('input', ()=>{ clearTimeout(t); t=setTimeout(applyFiltersAndRender, 200); });
  resetBtn.addEventListener('click', ()=>{
    [sheetFilter, sourceFilter, gradeFilter, sectionFilter].forEach(el => el.value='');
    searchInput.value=''; applyFiltersAndRender();
  });

  shareBtn.addEventListener('click', async ()=>{
    const summary = `لوحة متعددة القوالب — السجلات المعروضة: ${displayed.length} من ${aggregated.length}`;
    if (navigator.share){ try{ await navigator.share({ title:'لوحة الدرجات', text:summary, url:location.href }); }catch(e){} }
    else { navigator.clipboard?.writeText(summary); alert('تم نسخ الملخص للحافظة.'); }
  });

  exportXlsxBtn.addEventListener('click', ()=>{
    if (!displayed.length) return alert('لا توجد بيانات للتصدير.');
    const ws = XLSX.utils.json_to_sheet(displayed);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Results');
    const filename = `Results_Merged_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, filename, { compression:true });
  });

  exportPdfBtn.addEventListener('click', ()=>{
    if (!displayed.length) return alert('لا توجد بيانات للتصدير.');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation:'l', unit:'pt', format:'a4' });
    doc.setFont('helvetica','bold');
    doc.text('لوحة الدرجات — متعدد القوالب', 40, 40);
    doc.setFont('helvetica','normal');
    doc.text(`سجلات: ${displayed.length}`, 40, 60);
    const keys = Object.keys(displayed[0]);
    const head = [keys];
    const body = displayed.map(r => keys.map(k => s(r[k])));
    doc.autoTable({ head, body, startY:80, styles:{ font:'helvetica', fontSize:9 }, headStyles:{ fillColor:[31,143,160] } });
    doc.save(`Results_Merged.pdf`);
  });

  printBtn.addEventListener('click', ()=> window.print());
  ['dragenter','dragover'].forEach(evt => dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.add('hover'); }));
  ['dragleave','drop'].forEach(evt => dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.remove('hover'); }));
  dropzone.addEventListener('drop', e=>{ const files = e.dataTransfer?.files; if (files?.length) handleFiles(files); });
  dropzone.addEventListener('click', ()=> fileInput.click());
  fileInput.addEventListener('change', e => handleFiles(e.target.files));
  readmeLink.addEventListener('click', (e)=>{ e.preventDefault(); alert('انظر README.md داخل المستودع.'); });
})();
