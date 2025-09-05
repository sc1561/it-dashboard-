/* =========================================================
   لوحة تقارير ودرجات — للعرض فقط
   - لا يوجد خادم. كل المعالجة تتم داخل المتصفح.
   - استخدم SheetJS لقراءة/كتابة Excel.
   - تصدير: PDF (jsPDF+AutoTable) / Excel / طباعة.
   ========================================================= */

(() => {
  // عناصر DOM
  const fileInput     = document.getElementById('fileInput');
  const dropzone      = document.getElementById('dropzone');
  const loader        = document.getElementById('loader');

  const gradeFilter   = document.getElementById('gradeFilter');
  const sectionFilter = document.getElementById('sectionFilter');
  const sheetFilter   = document.getElementById('sheetFilter');
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

  // حالة التطبيق (in-memory فقط)
  let workbook = null;
  let sheetNames = [];
  let rawDataBySheet = {};  // {sheetName: Array<Object>}
  let currentSheet = '';
  let displayedData = [];   // بيانات الجدول الحالية بعد الفلترة
  let originalData = [];    // نسخة من بيانات الورقة المختارة قبل الفلترة

  // مفاتيح أعمدة متوقعة (للفلاتر والبحث)
  const COL_GRADE   = 'الصف';
  const COL_SECTION = 'الشعبة';
  const COL_NAME    = 'الاسم';

  // أدوات مساعدة
  const showLoader = (v=true) => loader.classList.toggle('show', v);
  const text = (v) => (v===undefined || v===null) ? '' : String(v);

  // قراءة ملف Excel داخل المتصفح (Read-Only)
  async function handleFile(file){
    if(!file) return;
    showLoader(true);
    try {
      const data = await file.arrayBuffer();
      workbook = XLSX.read(data, { type: 'array', cellDates: true });
      sheetNames = workbook.SheetNames || [];
      rawDataBySheet = {};

      // حوّل كل ورقة إلى JSON
      sheetNames.forEach(name => {
        const ws = workbook.Sheets[name];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rawDataBySheet[name] = json;
      });

      // ملء قائمة أوراق العمل
      fillSheetFilter(sheetNames);
      // عيّن الورقة الحالية
      currentSheet = sheetNames[0] || '';
      originalData = rawDataBySheet[currentSheet] || [];
      updateFiltersFromData(originalData);
      applyFiltersAndRender();

      stats.textContent = `تم تحميل الملف: ${file.name} — الأوراق: ${sheetNames.length}`;
    } catch (err) {
      console.error(err);
      alert('تعذر قراءة الملف. تأكد أن الصيغة صحيحة (.xlsx/.xls)');
    } finally {
      showLoader(false);
    }
  }

  // تعبئة قائمة الأوراق
  function fillSheetFilter(names){
    sheetFilter.innerHTML = '<option value="">(أول ورقة)</option>';
    names.forEach(n => {
      const opt = document.createElement('option');
      opt.value = n; opt.textContent = n;
      sheetFilter.appendChild(opt);
    });
  }

  // استنتاج خيارات الفلاتر (الصف والشعبة) من البيانات
  function updateFiltersFromData(rows){
    // اجلب القيم الفريدة
    const grades = new Set();
    const sections = new Set();

    rows.forEach(r => {
      if (r[COL_GRADE]) grades.add(text(r[COL_GRADE]));
      if (r[COL_SECTION]) sections.add(text(r[COL_SECTION]));
    });

    fillSelect(gradeFilter, [...grades]);
    fillSelect(sectionFilter, [...sections]);
  }

  function fillSelect(selectEl, options){
    const current = selectEl.value || '';
    selectEl.innerHTML = '<option value="">الكل</option>';
    options.sort((a,b)=> (''+a).localeCompare((''+b),'ar')).forEach(v=>{
      const opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      selectEl.appendChild(opt);
    });
    // حاول الحفاظ على الاختيار السابق إن وُجِد ضمن الخيارات
    if ([...selectEl.options].some(o=>o.value===current)) {
      selectEl.value = current;
    }
  }

  // فلترة + بحث + عرض
  function applyFiltersAndRender(){
    const g = gradeFilter.value.trim();
    const s = sectionFilter.value.trim();
    const q = searchInput.value.trim();

    const byGrade   = g ? (r => text(r[COL_GRADE]) === g) : (()=>true);
    const bySection = s ? (r => text(r[COL_SECTION]) === s) : (()=>true);
    const byQuery   = q ? (r => text(r[COL_NAME]).toLowerCase().includes(q.toLowerCase())) : (()=>true);

    displayedData = originalData.filter(r => byGrade(r) && bySection(r) && byQuery(r));

    renderTable(displayedData);
    stats.textContent = `عدد السجلات: ${displayedData.length} من ${originalData.length}${
      g || s ? ` (فلاتر: ${[g&&`الصف=${g}`, s&&`الشعبة=${s}`].filter(Boolean).join(', ')})` : ''
    }${q ? ` — بحث: "${q}"` : ''}`;
  }

  // رسم الجدول ديناميكيًا
  function renderTable(rows){
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (!rows.length){
      tableHead.innerHTML = '<tr><th>لا توجد بيانات مطابقة</th></tr>';
      return;
    }

    // رأس الجدول من مفاتيح أول صف
    const keys = Object.keys(rows[0]);
    const headRow = document.createElement('tr');
    keys.forEach(k=>{
      const th = document.createElement('th');
      th.textContent = k;
      headRow.appendChild(th);
    });
    tableHead.appendChild(headRow);

    // جسم الجدول
    const frag = document.createDocumentFragment();
    rows.forEach(r=>{
      const tr = document.createElement('tr');
      keys.forEach(k=>{
        const td = document.createElement('td');
        td.textContent = text(r[k]);
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    });
    tableBody.appendChild(frag);
  }

  // تغيير ورقة العمل
  sheetFilter.addEventListener('change', () => {
    const val = sheetFilter.value;
    currentSheet = val || (sheetNames[0] || '');
    originalData = rawDataBySheet[currentSheet] || [];
    updateFiltersFromData(originalData);
    applyFiltersAndRender();
  });

  // فلاتر/بحث
  [gradeFilter, sectionFilter].forEach(el => el.addEventListener('change', applyFiltersAndRender));
  // منع استدعاء كثيف أثناء الكتابة
  let searchTimer = null;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(applyFiltersAndRender, 200);
  });

  // إعادة التعيين
  resetBtn.addEventListener('click', ()=>{
    gradeFilter.value = '';
    sectionFilter.value = '';
    searchInput.value = '';
    applyFiltersAndRender();
  });

  // مشاركة النتائج (Web Share API إن توفرت)
  shareBtn.addEventListener('click', async ()=>{
    const summary = `لوحة درجات تقنية المعلومات — السجلات المعروضة: ${displayedData.length} من ${originalData.length}` +
      (currentSheet ? ` — ورقة: ${currentSheet}` : '');
    if (navigator.share) {
      try{
        await navigator.share({ title:'لوحة الدرجات', text: summary, url: location.href });
      }catch(e){ /* المستخدم ألغى */ }
    } else {
      navigator.clipboard?.writeText(summary);
      alert('تم نسخ الملخص للحافظة. يمكنك لصقه في شبكات التواصل.');
    }
  });

  // تصدير إلى Excel (الصفوف المفلترة فقط)
  exportXlsxBtn.addEventListener('click', ()=>{
    if (!displayedData.length) return alert('لا توجد بيانات للتصدير.');
    const ws = XLSX.utils.json_to_sheet(displayedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (currentSheet || 'نتائج'));
    const filename = `نتائج_تقنية_المعلومات_${currentSheet || 'Sheet'}_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, filename, { compression:true });
  });

  // تصدير إلى PDF
  exportPdfBtn.addEventListener('click', ()=>{
    if (!displayedData.length) return alert('لا توجد بيانات للتصدير.');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation:'l', unit:'pt', format:'a4' });
    doc.setFont('helvetica','bold');
    doc.text('لوحة الدرجات — تقنية المعلومات', 40, 40);
    doc.setFont('helvetica','normal');
    doc.text(`ورقة: ${currentSheet || 'أول ورقة'} | سجلات: ${displayedData.length}`, 40, 60);

    // بناء الأعمدة والصفوف
    const keys = Object.keys(displayedData[0]);
    const head = [keys];
    const body = displayedData.map(r => keys.map(k => (r[k] ?? '') + ''));

    doc.autoTable({
      head, body,
      startY: 80,
      styles: { font: 'helvetica', fontSize: 9 },
      headStyles: { fillColor: [31, 143, 160] }
    });
    doc.save(`نتائج_تقنية_المعلومات_${currentSheet || 'Sheet'}.pdf`);
  });

  // طباعة
  printBtn.addEventListener('click', ()=>{
    window.print();
  });

  // تعاملات السحب-والإفلات
  ['dragenter','dragover'].forEach(evt =>
    dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.add('hover'); })
  );
  ['dragleave','drop'].forEach(evt =>
    dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.remove('hover'); })
  );
  dropzone.addEventListener('drop', e=>{
    const file = e.dataTransfer?.files?.[0];
    if (file) handleFile(file);
  });
  dropzone.addEventListener('click', ()=> fileInput.click());
  fileInput.addEventListener('change', e => handleFile(e.target.files?.[0]));

  // رابط README (إرشادي فقط داخل المشروع)
  readmeLink.addEventListener('click', (e)=>{
    e.preventDefault();
    alert('افتح ملف README.md داخل المستودع للاطلاع على طريقة التشغيل والنشر.');
  });

})();
