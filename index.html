async function handleFile(file){
  if(!file) return;

  showLoader(true);

  // 1) فحص الامتداد أولًا
  const name = file.name || '';
  const lower = name.toLowerCase();
  const isCSV  = lower.endsWith('.csv');
  const isXLSX = lower.endsWith('.xlsx') || lower.endsWith('.xlsm');
  const isXLS  = lower.endsWith('.xls');

  try {
    const buf = await file.arrayBuffer();

    // 2) خيارات قراءة مرنة
    /** ملاحظات:
     * - type:'array' يناسب معظم الصيغ الحديثة.
     * - cellDates:true يحفظ التواريخ ككائن Date.
     * - WTF:true يعطي أخطاء تشخيصية لملفات معطوبة.
     */
    workbook = XLSX.read(buf, {
      type: 'array',
      cellDates: true,
      raw: false,
      WTF: true
    });

    // 3) معالجة CSV خاصّة (أحيانًا أفضل كمسار منفصل)
    if (isCSV && (!workbook.SheetNames || workbook.SheetNames.length === 0)) {
      const text = await file.text();
      const csvWB = XLSX.read(text, { type: 'string' });
      workbook = csvWB;
    }

    // 4) تحقّق من وجود أوراق
    sheetNames = workbook.SheetNames || [];
    if (!sheetNames.length) {
      throw new Error('لم يتم العثور على أوراق عمل داخل الملف.');
    }

    // 5) تحويل كل ورقة إلى JSON مع defval لتجنب undefined
    rawDataBySheet = {};
    sheetNames.forEach(name => {
      const ws = workbook.Sheets[name];
      // defval:'' يضمن عدم وجود قيم undefined
      const json = XLSX.utils.sheet_to_json(ws, {
        defval: '',
        raw: false,      // يحوّل الأرقام/التواريخ لنصوص مقروءة عند اللزوم
        blankrows: false
      });
      rawDataBySheet[name] = json;
    });

    // 6) إنذار مبكر لو لم توجد صفوف فعلية
    const firstSheet = sheetNames[0];
    if (!rawDataBySheet[firstSheet]?.length) {
      throw new Error('تم فتح الملف بنجاح لكن لم يتم العثور على صفوف بيانات. تأكد أن الصف الأول يحوي عناوين أعمدة واضحة.');
    }

    // 7) تحديث القوائم والعرض
    fillSheetFilter(sheetNames);
    currentSheet = sheetNames[0] || '';
    originalData = rawDataBySheet[currentSheet] || [];
    updateFiltersFromData(originalData);
    applyFiltersAndRender();

    stats.textContent = `تم تحميل الملف: ${name} — الأوراق: ${sheetNames.length}`;
  } catch (err) {
    console.error(err);

    // 8) رسائل خطأ مفيدة بحسب الحالة
    let hint = 'تعذر قراءة الملف. تأكد أن الصيغة صحيحة وأن الملف غير محمي بكلمة مرور.';
    if (isXLS) {
      hint += '\nملاحظة: ملفات ‎.xls‎ القديمة قد تحتوي تنسيقًا لا يُقرأ دائمًا. جرّب حفظه كـ ‎.xlsx‎ من Excel ثم أعد رفعه.';
    }
    if (!isCSV && !isXLS && !isXLSX) {
      hint += '\nالامتدادات المدعومة: ‎.xlsx / .xlsm / .xls / .csv';
    }
    alert(hint);
  } finally {
    showLoader(false);
  }
}
