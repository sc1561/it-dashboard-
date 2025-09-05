/* =========================================================
   لوحة تقارير ودرجات — للعرض فقط (Read-Only)
   - لا يوجد خادم. كل المعالجة تتم داخل المتصفح.
   - SheetJS (xlsx) لقراءة/كتابة Excel + CSV.
   - تصدير: PDF (jsPDF+AutoTable) / Excel / طباعة.
   - تدعيم ملفات: .xlsx / .xlsm / .xls / .csv
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
  let displayedData = [];   // بيانات الجدول الحالية بعد الفلترة/الفرز/البحث
  let originalData = [];    // نسخة من بيانات الورقة المختارة قبل الفلترة

  // مفاتيح أعمدة متوقعة (للفلاتر والبحث)
  const COL_GRADE   = 'الصف';
  const COL_SECTION = 'الشعبة';
  const COL_NAME    = 'الاسم';

  // أدوات مساعدة
  const showLoader = (v=true) => loader.classList.toggle('show', v);
  const text = (v) => (v===undefined || v===null) ? '' : String(v);

  /* =========================================================
     قراءة ملف Excel/CSV داخل المتصفح (Read-Only)
     ========================================================= */
  async function handleFile(file){
    if(!file) return;

    showLoader(true);

    // فحص الامتداد
    const name = file.name || '';
    const lower = name.toLowerCase();
    const isCSV  = lower.endsWith('.csv');
    const isXLSX = lower.endsWith('.xlsx') || lower.endsWith('.xlsm');
    const isXLS  = lower.endsWith('.xls');

    try {
      const buf = await file.arrayBuffer();

      // قراءة أساسية
      workbook = XLSX.read(buf, {
        type: 'array',
        cellDates: true,
        raw: false,
        WTF: true   // يعطي رسائل تشخيصية في الكونسول إذا الملف معطوب
      });

      // معالجة خاصة لـ CSV عند الحاجة
      if (isCSV && (!workbook.SheetNames || workbook.SheetNames.length === 0)) {
        const textCsv = await file.text();
        workbook = XLSX.read(textCsv, { type: 'string' });
      }

      sheetNames = workbook.SheetNames || [];
      if (!sheetNames.length) {
        throw new Error('لم يتم العثور على أوراق عمل داخل الملف.');
      }

      // تحويل أوراق العمل إلى JSON
      rawDataBySheet = {};
      sheetNames.forEach(sName => {
        const ws = workbook.Sheets[sName];
        const json = XLSX.utils.sheet_to_json(ws, {
          defval: '',
          raw: false,
          blankrows: false
        });
        rawDataBySheet[sName] = json;
      });

      const firstSheet = sheetNames[0];
      if (!rawDataBySheet[firstSheet]?.length) {
        throw new Error('تم فتح الملف بنجاح لكن لم يتم العثور على صفوف بيانات. تأكد أن الصف الأول يحوي عناوين أعمدة واضحة.');
      }

      // ملء قائمة الأوراق واختيار الأولى
      fillSheetFilter(sheetNames);
      currentSheet = firstSheet;
      originalData = rawDataBySheet[currentSheet] || [];
      updateFiltersFromData(originalData);
      applyFiltersAndRender();

      stats.textContent = `تم تحميل الملف: ${name} — الأوراق: ${sheetNames.length}`;
    } catch (err) {
      console.error(err);

      // رسائل مساعدة حسب الحالة
      let hint = 'تعذر قراءة الملف. تأكد أن الصيغة صحيحة وأن الملف غير محمي بكلمة مرور.';
      if (isXLS) {
        hint += '\nملاحظة: ملفات ‎.xls‎ القديمة قد لا تُقرأ دائمًا. افتحها في Excel ثم احفظها كـ ‎.xlsx‎ وأعد الرفع.';
      }
      if (!isCSV && !isXLS && !isXLSX) {
        hint += '\nالامتدادات المدعومة: ‎.xlsx / .xlsm / .xls / .csv';
      }
      alert(hint);
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

  // استنتاج خيارات الفلاتر (الصف/الشعبة)
  function updateFiltersFromData(rows){
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
    if ([...selectEl.options].some(o=>o.value===current)) {
      selectEl.value = current;
    }
  }

  /* =========================================================
     فلترة + بحث + عرض
     ========================================================= */
  function applyFiltersAndRender(){
    const g = gradeFilter.value.trim();
    const s = sectionFilter.value.trim();
    const q = searchInput.value.trim();

    const byGrade   = g ? (r => text(r[COL_GRADE]) === g) : (()=>true);
    const bySection = s ? (r => text(r[COL_SECTION]) === s) : (()=>true);
    const byQuery   = q ? (r => text(r[COL_NAME]).toLowerCase().includes(q.toLowerCase())) : (()=>true);

    displayedData = (originalData || []).filter(r => byGrade(r) && bySection(r) && byQuery(r));

    renderTable(displayedData);
    stats.textContent = `عدد السجلات: ${displayedData.length} من ${originalData.length}${
      g || s ? ` (فلاتر: ${[g&&`الصف=${g}`, s&&`الشعبة=${s}`].filter(Boolean).join(', ')})` : ''
    }${q ? ` — بحث: "${q}"` : ''}`;
  }

  // رسم الجدول + تمكين الفرز بالنقر على رؤوس الأعمدة
  function renderTable(rows){
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (!rows.length){
      tableHead.innerHTML = '<tr><th>لا توجد بيانات مطابقة</th></tr>';
      return;
    }

    const keys = Object.keys(rows[0]);

    // رأس الجدول
    const headRow = document.createElement('tr');
    keys.forEach(k=>{
      const th = document.createElement('th');
      th.textContent = k;
      th.style.cursor = 'pointer';
      th.title = 'انقر للفرز';
      th.addEventListener('click', () => {
        const asc = th.dataset.sort !== 'asc';
        displayedData.sort((a,b)=>{
          const va = (a[k] ?? '') + '', vb = (b[k] ?? '') + '';
          return asc ? va.localeCompare(vb, 'ar', {numeric:true})
                     : vb.localeCompare(va, 'ar', {numeric:true});
        });
        // ضبط مؤشرات الفرز البصرية (بسيطة)
        [...headRow.children].forEach(h => delete h.dataset.sort);
        th.dataset.sort = asc ? 'asc' : 'desc';
        renderTable(displayedData);
      });
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

  /* =========================================================
     تبديل الورقة + فلاتر/بحث + أزرار
     ========================================================= */
  sheetFilter.addEventListener('change', () => {
    const val = sheetFilter.value;
    const newSheet = val || (sheetNames[0] || '');
    currentSheet = newSheet;
    originalData = rawDataBySheet[currentSheet] || [];
    updateFiltersFromData(originalData);
    applyFiltersAndRender();
  });

  [gradeFilter, sectionFilter].forEach(el => el.addEventListener('change', applyFiltersAndRender));

  // تقليل ضغط التصفية أثناء الكتابة
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

    // الأعمدة والصفوف
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
