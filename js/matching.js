const Matching = {

/* ============================================================
   1-A) مطابقة الأصناف — وضع "رفع جميع المخازن"
   ============================================================ */
async runItemsMatchAllStores(files){
  const startTime = performance.now();
  console.log('═══════════════════════════════════════');
  console.log('📊 عدد الملفات (المخازن):', files.length);

  if(!files.length) throw new Error('❌ يرجى رفع ملف واحد على الأقل');

  // ===== 1) قراءة كل الملفات =====
  const storesData = [];
  for(let i = 0; i < files.length; i++){
    const f = files[i];
    const storeName = f.name.replace(/\.(xlsx|xls|csv)$/i, '').trim();
    try{
      const d = await Excel.readFileAsArray(f);
      const rows = d.rows;
      const headers = d.headers;

      if(headers.length < 3){
        console.warn(`⚠️ ${storeName}: الملف لا يحتوي على 3 أعمدة — تخطي`);
        continue;
      }

      const cName  = headers[0];
      const cPrice = headers[1];
      const cDisc  = headers[2];

      const items = rows.map(r=>({
        name:  String(r[cName]  ?? '').trim(),
        price: toNum(r[cPrice]),
        disc:  toNum(r[cDisc])
      })).filter(x => x.name);

      storesData.push({ name: storeName, items });
      console.log(`✅ ${storeName}: ${items.length} صنف`);
    }catch(err){
      console.error(`❌ ${storeName}: ${err.message}`);
    }
  }

  if(!storesData.length) throw new Error('❌ لم يتم قراءة أي ملف بنجاح');

  // ===== 2) تنظيف وتطبيع =====
  const cleanPrice = (str)=>{
    if(str == null) return '';
    let s = String(str).trim();
    s = s.replace(/\/\s*\d+(\.\d+)?\s*(جنيه|جنية|ج|جم|جنيهًا)\s*$/gi, '');
    s = s.replace(/ق\s*\/\s*\d+(\.\d+)?\s*(جنيه|جنية|ج|جم|جنيهًا)?\s*$/gi, '');
    s = s.replace(/\/\s*\d+(\.\d+)?\s*(جنيه|جنية|ج|جم|جنيهًا)?/gi, '');
    s = s.replace(/\b\d+(\.\d+)?\s*(جنيه|جنية|جنيهًا)\b/gi, '');
    return s.trim();
  };

  const removeExtras = (str)=>{
    if(!str) return '';
    let s = str;
    const extras = ['س ج','سج','س ق','سق','جديد','جديده','جديدة','قديم','قديمه','قديمة',
      'اقراص','قرص','كبسول','كبسوله','امبول','امبولات','امبوله','عرض','عروض','سعر','جنيه','جنية'];
    extras.forEach(w => {
      s = s.replace(new RegExp(`(^|\\s)${w}(\\s|$)`, 'gi'), ' ');
    });
    return s.replace(/\s+/g, ' ').trim();
  };

  const normalizeText = (str)=>{
    if(str == null) return '';
    let s = String(str).trim().toLowerCase();
    s = s.replace(/[أإآٱ]/g,'ا').replace(/[ىئ]/g,'ي').replace(/ؤ/g,'و').replace(/ة/g,'ه');
    s = s.replace(/[\u064B-\u0652\u0640]/g,'');
    s = s.replace(/[^\w\u0600-\u06FF\s]/g,' ');
    return s.replace(/\s+/g,' ').trim();
  };

  const extractConc = (str)=>{
    const s = String(str).toLowerCase();
    const conc = [];
    [/(\d+(?:\.\d+)?)\s*(?:مجم|مج|مليجرام|mg)/gi,
     /(\d+(?:\.\d+)?)\s*(?:مل|مللي|مليلتر|ml)/gi,
     /(\d+(?:\.\d+)?)\s*(?:جم|جرام|g)/gi].forEach(regex => {
      let m; while((m = regex.exec(s)) !== null) conc.push(parseFloat(m[1]));
    });
    return conc.sort((a,b)=>a-b);
  };

  const concMatch = (a, b)=>{
    if(!a.length || !b.length) return true;
    if(a.length !== b.length) return false;
    for(let i = 0; i < a.length; i++) if(Math.abs(a[i]-b[i]) > 0.01) return false;
    return true;
  };

  const stripNumbers = (s)=> s.replace(/\d+(?:\.\d+)?/g, ' ').replace(/\s+/g, ' ').trim();

  const fullClean = (str)=>{
    return removeExtras(normalizeText(cleanPrice(str)));
  };

  // ===== 3) تجميع الأصناف من كل المخازن =====
  console.log('🔧 تجميع الأصناف من جميع المخازن...');
  const itemsMap = {}; // cleanName → { rawName, price, stores: Set, discounts: [], conc }

  storesData.forEach(store=>{
    store.items.forEach(it=>{
      const clean = fullClean(it.name);
      if(!clean) return;
      const noNum = stripNumbers(clean);
      if(!noNum) return;

      // مفتاح: النص بدون أرقام (للتجميع الذكي)
      const key = noNum;

      if(!itemsMap[key]){
        itemsMap[key] = {
          rawName: it.name,
          cleanName: clean,
          price: it.price,
          conc: extractConc(it.name),
          stores: new Set(),
          discounts: [],
          storeData: {}
        };
      }

      itemsMap[key].stores.add(store.name);
      itemsMap[key].discounts.push(it.disc);
      itemsMap[key].storeData[store.name] = it.disc;
    });
  });

  console.log(`✅ تم تجميع ${Object.keys(itemsMap).length} صنف فريد`);

  // ===== 4) بناء النتيجة =====
  const result = Object.values(itemsMap).map(item=>{
    const storesArr = Array.from(item.stores);
    const repeat = storesArr.length;

    // أعلى خصم
    let maxDisc = -Infinity;
    let maxDiscStore = '—';
    storesArr.forEach(s=>{
      const d = item.storeData[s];
      if(d > maxDisc){ maxDisc = d; maxDiscStore = s; }
    });
    if(maxDisc === -Infinity) maxDisc = 0;

    return {
      name: item.rawName,
      cleanName: item.cleanName,
      price: item.price,
      discount: item.discounts[0] || 0,
      stores: storesArr.join(' | '),
      storesArr: storesArr,
      discounts: item.discounts.join(' | '),
      repeat,
      maxDiscount: maxDisc,
      maxDiscountStore: maxDiscStore
    };
  });

  result.sort((a, b) => b.maxDiscount - a.maxDiscount);

  const endTime = performance.now();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  console.log(`✅ النتيجة: ${result.length} صنف فريد | ⏱️ ${duration} ثانية`);

  return { rows: result, total: result.length, storesCount: storesData.length };
},

/* ============================================================
   1-B) مطابقة الأصناف — وضع "رفع مخزن واحد"
   - ملف 1 (الفرع): [الكود] [الصنف] [السعر]
   - ملف 2 (المطابقة): [الصنف] [السعر] [الخصم]
   - الناتج: الكود | الصنف | السعر | الخصم | التطابق | النسبة | صنف التطابق
   ============================================================ */
async runItemsMatch(fileBranch, fileMatch){
  const startTime = performance.now();

  const d1 = await Excel.readFileAsArray(fileBranch);
  const d2 = await Excel.readFileAsArray(fileMatch);

  const branchRows = d1.rows, branchHeaders = d1.headers;
  const matchRows = d2.rows, matchHeaders = d2.headers;

  console.log('═══════════════════════════════════════');
  console.log('📊 صفوف الفرع:', branchRows.length);
  console.log('📊 صفوف المطابقة:', matchRows.length);

  if(branchHeaders.length < 3) throw new Error('❌ ملف الفرع يجب أن يحتوي على 3 أعمدة');
  if(matchHeaders.length < 3) throw new Error('❌ ملف المطابقة يجب أن يحتوي على 3 أعمدة');

  const bCode  = branchHeaders[0];
  const bName  = branchHeaders[1];
  const bPrice = branchHeaders[2];

  const mName  = matchHeaders[0];
  const mPrice = matchHeaders[1];
  const mDisc  = matchHeaders[2];

  // ===== نفس دوال التنظيف =====
  const cleanPrice = (str)=>{
    if(str == null) return '';
    let s = String(str).trim();
    s = s.replace(/\/\s*\d+(\.\d+)?\s*(جنيه|جنية|ج|جم|جنيهًا)\s*$/gi, '');
    s = s.replace(/ق\s*\/\s*\d+(\.\d+)?\s*(جنيه|جنية|ج|جم|جنيهًا)?\s*$/gi, '');
    s = s.replace(/\/\s*\d+(\.\d+)?\s*(جنيه|جنية|ج|جم|جنيهًا)?/gi, '');
    s = s.replace(/\b\d+(\.\d+)?\s*(جنيه|جنية|جنيهًا)\b/gi, '');
    return s.trim();
  };

  const removeExtras = (str)=>{
    if(!str) return '';
    let s = str;
    const extras = ['س ج','سج','س ق','سق','جديد','جديده','جديدة','قديم','قديمه','قديمة',
      'اقراص','قرص','كبسول','كبسوله','امبول','امبولات','امبوله','عرض','عروض','سعر','جنيه','جنية'];
    extras.forEach(w => {
      s = s.replace(new RegExp(`(^|\\s)${w}(\\s|$)`, 'gi'), ' ');
    });
    return s.replace(/\s+/g, ' ').trim();
  };

  const normalizeText = (str)=>{
    if(str == null) return '';
    let s = String(str).trim().toLowerCase();
    s = s.replace(/[أإآٱ]/g,'ا').replace(/[ىئ]/g,'ي').replace(/ؤ/g,'و').replace(/ة/g,'ه');
    s = s.replace(/[\u064B-\u0652\u0640]/g,'');
    s = s.replace(/[^\w\u0600-\u06FF\s]/g,' ');
    return s.replace(/\s+/g,' ').trim();
  };

  const extractConc = (str)=>{
    const s = String(str).toLowerCase();
    const conc = [];
    [/(\d+(?:\.\d+)?)\s*(?:مجم|مج|مليجرام|mg)/gi,
     /(\d+(?:\.\d+)?)\s*(?:مل|مللي|مليلتر|ml)/gi,
     /(\d+(?:\.\d+)?)\s*(?:جم|جرام|g)/gi].forEach(regex => {
      let m; while((m = regex.exec(s)) !== null) conc.push(parseFloat(m[1]));
    });
    return conc.sort((a,b)=>a-b);
  };

  const concMatch = (a, b)=>{
    if(!a.length || !b.length) return true;
    if(a.length !== b.length) return false;
    for(let i = 0; i < a.length; i++) if(Math.abs(a[i]-b[i]) > 0.01) return false;
    return true;
  };

  const stripNumbers = (s)=> s.replace(/\d+(?:\.\d+)?/g, ' ').replace(/\s+/g, ' ').trim();
  const fullClean = (str)=> removeExtras(normalizeText(cleanPrice(str)));

  const levenshtein = (a, b)=>{
    if(a === b) return 0;
    const la = a.length, lb = b.length;
    if(!la) return lb; if(!lb) return la;
    if(Math.abs(la-lb) > Math.max(la,lb)*0.4) return Math.max(la,lb);
    let prev = new Array(lb+1), curr = new Array(lb+1);
    for(let j=0;j<=lb;j++) prev[j]=j;
    for(let i=1;i<=la;i++){
      curr[0]=i;
      for(let j=1;j<=lb;j++){
        curr[j] = a[i-1]===b[j-1] ? prev[j-1] : Math.min(prev[j-1],curr[j-1],prev[j])+1;
      }
      [prev,curr] = [curr,prev];
    }
    return prev[lb];
  };

  // بناء فهرس المطابقة
  const exactIndex = {};
  const noNumIndex = {};
  const wordTokens = [];

  matchRows.forEach(r=>{
    const raw = String(r[mName] ?? '').trim();
    if(!raw) return;
    const clean = fullClean(raw);
    if(!clean) return;
    const noNum = stripNumbers(clean);
    if(!exactIndex[clean]) exactIndex[clean] = r;
    if(noNum){
      if(!noNumIndex[noNum]) noNumIndex[noNum] = [];
      noNumIndex[noNum].push(r);
    }
    wordTokens.push({
      clean, noNum,
      words: clean.split(' ').filter(w=>w.length>=2),
      conc: extractConc(raw),
      row: r
    });
  });

  // المطابقة
  const result = branchRows.map(r=>{
    const code  = String(r[bCode] ?? '').trim();
    const name  = String(r[bName] ?? '').trim();
    const price = toNum(r[bPrice]);
    const clean = fullClean(name);
    const noNum = stripNumbers(clean);
    const words = clean.split(' ').filter(w=>w.length>=2);
    const conc = extractConc(name);

    let matchRow = null;
    let matchType = 'none';
    let score = 0;
    let matched = false;

    // 1) تام
    if(exactIndex[clean]){
      matchRow = exactIndex[clean]; matched = true; matchType = 'exact'; score = 1;
    }

    // 2) بدون أرقام
    if(!matchRow && noNum && noNumIndex[noNum]){
      for(const cand of noNumIndex[noNum]){
        const candConc = extractConc(String(cand[mName] ?? ''));
        if(!concMatch(conc, candConc)) continue;
        const candClean = fullClean(String(cand[mName] ?? ''));
        const dist = levenshtein(clean, candClean);
        const maxLen = Math.max(clean.length, candClean.length);
        const s = maxLen ? (1 - dist / maxLen) : 1;
        if(s >= 0.85){ matchRow = cand; matched = true; matchType = 'noNumbers'; score = s; break; }
      }
    }

    // 3) كلمات
    if(!matchRow && words.length >= 2){
      const firstWord = words[0];
      const candidates = wordTokens.filter(x => x.words[0] === firstWord || x.words.some(w=>w===firstWord));
      let best = null, bestScore = 0;
      for(const x of candidates){
        if(!concMatch(conc, x.conc)) continue;
        if(Math.abs(x.words.length - words.length) > 2) continue;

        const setX = new Set(x.words);
        let common = 0;
        for(const w of words) if(setX.has(w)) common++;
        const wordRatio = common / Math.max(words.length, x.words.length);

        if(wordRatio >= 0.85){ if(wordRatio > bestScore){ bestScore = wordRatio; best = x.row; } continue; }
        if(wordRatio >= 0.5){
          const dist = levenshtein(clean, x.clean);
          const maxLen = Math.max(clean.length, x.clean.length);
          const s = maxLen ? (1 - dist / maxLen) : 0;
          if(s > bestScore){ bestScore = s; best = x.row; }
        }
      }
      if(best && bestScore >= 0.85){
        matchRow = best; matched = true;
        matchType = bestScore >= 0.95 ? 'high' : (bestScore >= 0.90 ? 'medium' : 'low');
        score = bestScore;
      }
    }

    const discount = matchRow ? toNum(matchRow[mDisc]) : 0;
    const matchName = matchRow ? String(matchRow[mName] ?? '') : '';

    return {
      code, name, price, discount, matched, matchType,
      percentage: Math.round(score * 100),
      matchName, cleanedName: clean
    };
  });

  const matchCount = result.filter(r=>r.matched).length;
  const ratio = result.length ? ((matchCount/result.length)*100).toFixed(1) : 0;
  const typeStats = {};
  result.filter(r=>r.matched).forEach(r=>{ typeStats[r.matchType] = (typeStats[r.matchType]||0)+1; });

  const endTime = performance.now();
  console.log(`✅ المطابقة: ${matchCount}/${result.length} (${ratio}%)`, typeStats);
  console.log(`⏱️ ${((endTime-startTime)/1000).toFixed(2)} ثانية`);

  return { rows: result, ratio, matchCount, total: result.length };
},
  /* ============================================================
     2) فرق الجملة
     ============================================================ */
  async runWholeDiff(fileLink, fileWholeDiff, fileWhole){
    const d1 = await Excel.readFileAsArray(fileLink);
    const d2 = await Excel.readFileAsArray(fileWholeDiff);
    const d3 = await Excel.readFileAsArray(fileWhole);

    const link = d1.rows, linkHeaders = d1.headers;
    const wholeDiff = d2.rows, wholeDiffHeaders = d2.headers;
    const whole = d3.rows, wholeHeaders = d3.headers;

    console.log('═══════════════════════════════════════');
    console.log('📊 صفوف الربط:', link.length);
    console.log('📊 صفوف فرق الجملة:', wholeDiff.length);
    console.log('📊 صفوف الجملة:', whole.length);
    console.log('═══════════════════════════════════════');

    if(linkHeaders.length < 4){
      throw new Error(`❌ ملف الربط يجب أن يحتوي على 4 أعمدة على الأقل: [كود طيبة، كود الفرع، الصنف، الخصم]. الموجود: ${linkHeaders.length}`);
    }
    if(wholeDiffHeaders.length < 4){
      throw new Error(`❌ ملف فرق الجملة يجب أن يحتوي على 4 أعمدة: [الكود، المورد، الصنف، الخصم]. الموجود: ${wholeDiffHeaders.length}`);
    }
    if(wholeHeaders.length < 3){
      throw new Error(`❌ ملف الجملة يجب أن يحتوي على 3 أعمدة: [الكود، الخصم، الصنف]. الموجود: ${wholeHeaders.length}`);
    }

    const cLink = { taibaCode: linkHeaders[0], branchCode: linkHeaders[1], name: linkHeaders[2], discount: linkHeaders[3] };
    const cWholeDiff = { code: wholeDiffHeaders[0], supplier: wholeDiffHeaders[1], name: wholeDiffHeaders[2], discount: wholeDiffHeaders[3] };
    const cWhole = { code: wholeHeaders[0], discount: wholeHeaders[1], name: wholeHeaders[2] };

    const wholeDiffMap = {};
    wholeDiff.forEach(r=>{ const k = normalize(r[cWholeDiff.code]); if(k && !wholeDiffMap[k]) wholeDiffMap[k] = r; });

    const wholeMap = {};
    whole.forEach(r=>{ const k = normalize(r[cWhole.code]); if(k && !wholeMap[k]) wholeMap[k] = r; });

    const result = link.map(r=>{
      const taibaCode = String(r[cLink.taibaCode] ?? '').trim();
      const branchCode = String(r[cLink.branchCode] ?? '').trim();
      const linkName = String(r[cLink.name] ?? '').trim();
      const linkDiscount = toNum(r[cLink.discount]);

      const wdRow = wholeDiffMap[normalize(branchCode)];
      const name = wdRow ? String(wdRow[cWholeDiff.name] ?? linkName) : linkName;
      const supplier = wdRow ? String(wdRow[cWholeDiff.supplier] ?? '') : '';
      const discount = wdRow ? toNum(wdRow[cWholeDiff.discount]) : linkDiscount;

      const wRow = wholeMap[normalize(taibaCode)];
      const wholeName = wRow ? String(wRow[cWhole.name] ?? '') : '';
      const wholeDiscount = wRow ? toNum(wRow[cWhole.discount]) : 0;

      const diff = wholeDiscount - discount;

      return { taibaCode, branchCode, name, discount, supplier, wholeDiscount, wholeName, diff,
        matchedWholeDiff: !!wdRow, matchedWhole: !!wRow };
    });

    result.sort((a, b) => b.diff - a.diff);
    return result;
  },

  /* ============================================================
     3) فرق الفاكسات — مطابقة تامة فقط + التكرار والمخازن والخصومات
     ============================================================ */
  async runFaxDiff(fileFaxDiff, fileFax){
    const d1 = await Excel.readFileAsArray(fileFaxDiff);
    const d2 = await Excel.readFileAsArray(fileFax);

    const diffRows = d1.rows, diffHeaders = d1.headers;
    const faxRows = d2.rows, faxHeaders = d2.headers;

    if(diffHeaders.length < 5){
      throw new Error(`❌ ملف فرق الفاكسات يجب أن يحتوي على 5 أعمدة: [الصنف، الاسم، الكمية، الكاتب، الخصم]. الموجود: ${diffHeaders.length}`);
    }
    if(faxHeaders.length < 3){
      throw new Error(`❌ ملف الفاكسات يجب أن يحتوي على 3 أعمدة: [الصنف، خصم الفاكسات، المورد]. الموجود: ${faxHeaders.length}`);
    }

    const cDiff = { name: diffHeaders[0], store: diffHeaders[1], qty: diffHeaders[2], writer: diffHeaders[3], discount: diffHeaders[4] };
    const cFax = { name: faxHeaders[0], discount: faxHeaders[1], supplier: faxHeaders[2] };

    // فهرس الفاكسات (تجميع التكرار والمخازن والخصومات)
    const faxGroups = {};

    faxRows.forEach(r=>{
      const rawName = String(r[cFax.name] ?? '').trim();
      if(!rawName) return;
      const normName = normalize(rawName);
      if(!normName) return;

      const discount = toNum(r[cFax.discount]);
      const supplier = String(r[cFax.supplier] ?? '').trim() || '—';

      if(!faxGroups[normName]){
        faxGroups[normName] = {
          rawName: rawName,
          stores: new Set(),
          storeDiscounts: {}
        };
      }

      faxGroups[normName].stores.add(supplier);

      if(!faxGroups[normName].storeDiscounts[supplier]){
        faxGroups[normName].storeDiscounts[supplier] = [];
      }
      faxGroups[normName].storeDiscounts[supplier].push(discount);
    });

    console.log(`🔎 فهرس الفاكسات: ${Object.keys(faxGroups).length} صنف فريد`);

    // بناء النتيجة مع بيانات التكرار
    const result = diffRows.map(r=>{
      const name = String(r[cDiff.name] ?? '').trim();
      const storeName = String(r[cDiff.store] ?? '').trim();
      const qty = toNum(r[cDiff.qty]);
      const writer = String(r[cDiff.writer] ?? '').trim();
      const discount = toNum(r[cDiff.discount]);

      const faxGroup = faxGroups[normalize(name)] || null;
      const matched = !!faxGroup;

      let faxDiscount = null;
      let supplier = '_____';
      let repeat = 0;
      let storesText = '—';
      let discountsText = '—';

      if(faxGroup){
        const storesArr = Array.from(faxGroup.stores);
        repeat = storesArr.length;
        supplier = storesArr.join(' | ') || '_____';
        storesText = storesArr.join(' | ') || '—';

        const allDiscounts = [];
        storesArr.forEach(store => {
          const discounts = faxGroup.storeDiscounts[store] || [];
          discounts.forEach(d => allDiscounts.push(d));
        });
        discountsText = allDiscounts.join(' | ') || '—';

        faxDiscount = allDiscounts.length ? allDiscounts[0] : null;
      }

      const diff = (faxDiscount != null) ? (faxDiscount - discount) : null;

      return {
        name, storeName, qty, writer, discount, faxDiscount, supplier, diff, matched,
        repeat, stores: storesText, discounts: discountsText
      };
    });

    result.sort((a, b) => (b.diff ?? -Infinity) - (a.diff ?? -Infinity));
    return result;
  },

  /* ============================================================
     4) مصفر الفاكسات — 4 ملفات + التكرار والمخازن والخصومات
     ============================================================ */
  async runFaxFilter(fileReceived, fileOrders, fileCollector, fileFax){
    const dCol = await Excel.readFileAsArray(fileCollector);
    const dFax = await Excel.readFileAsArray(fileFax);

    const collector = dCol.rows, colHeaders = dCol.headers;
    const faxRows = dFax.rows, faxHeaders = dFax.headers;

    const dRec = fileReceived ? await Excel.readFileAsArray(fileReceived) : { headers: [], rows: [] };
    const dOrd = fileOrders ? await Excel.readFileAsArray(fileOrders) : { headers: [], rows: [] };
    const received = dRec.rows, recHeaders = dRec.headers;
    const orders = dOrd.rows, ordHeaders = dOrd.headers;

    console.log('═══════════════════════════════════════');
    console.log('📊 صفوف المجمع:', collector.length);
    console.log('📊 صفوف الفاكسات:', faxRows.length);
    console.log('📊 صفوف الوارد:', received.length);
    console.log('📊 صفوف الطلبات:', orders.length);
    console.log('═══════════════════════════════════════');

    if(colHeaders.length < 6){
      throw new Error(`❌ ملف المجمع يجب أن يحتوي على 6 أعمدة: [الكود، الاسم، س البيع، المرجح، ك مخزون، ك صافي بيع]. الموجود: ${colHeaders.length}`);
    }
    if(faxHeaders.length < 3){
      throw new Error(`❌ ملف الفاكسات يجب أن يحتوي على 3 أعمدة: [الصنف، الخصم، المورد]. الموجود: ${faxHeaders.length}`);
    }
    if(received.length && recHeaders.length < 3){
      throw new Error(`❌ ملف الوارد يجب أن يحتوي على 3 أعمدة: [الكود، الكمية، الصنف]`);
    }
    if(orders.length && ordHeaders.length < 3){
      throw new Error(`❌ ملف الطلبات يجب أن يحتوي على 3 أعمدة: [الكود، الكمية، الصنف]`);
    }

    const cCol = { code: colHeaders[0], name: colHeaders[1], salePrice: colHeaders[2], weighted: colHeaders[3], stock: colHeaders[4], netSale: colHeaders[5] };
    const cFax = { name: faxHeaders[0], discount: faxHeaders[1], supplier: faxHeaders[2] };
    const cRec = received.length ? { code: recHeaders[0], qty: recHeaders[1], name: recHeaders[2] } : null;
    const cOrd = orders.length ? { code: ordHeaders[0], qty: ordHeaders[1], name: ordHeaders[2] } : null;

    // فهرس الوارد والطلبات
    const recMap = {};
    if(cRec){
      received.forEach(r=>{
        const k = normalize(r[cRec.code]);
        if(k) recMap[k] = (recMap[k] || 0) + toNum(r[cRec.qty]);
      });
    }

    const ordMap = {};
    if(cOrd){
      orders.forEach(r=>{
        const k = normalize(r[cOrd.code]);
        if(k) ordMap[k] = (ordMap[k] || 0) + toNum(r[cOrd.qty]);
      });
    }

    // فهرس الفاكسات (تجميع التكرار والمخازن والخصومات)
    const faxGroups = {};

    faxRows.forEach(r=>{
      const rawName = String(r[cFax.name] ?? '').trim();
      if(!rawName) return;
      const normName = normalize(rawName);
      if(!normName) return;

      const discount = toNum(r[cFax.discount]);
      const supplier = String(r[cFax.supplier] ?? '').trim() || '—';

      if(!faxGroups[normName]){
        faxGroups[normName] = {
          rawName: rawName,
          stores: new Set(),
          storeDiscounts: {}
        };
      }

      faxGroups[normName].stores.add(supplier);

      if(!faxGroups[normName].storeDiscounts[supplier]){
        faxGroups[normName].storeDiscounts[supplier] = [];
      }
      faxGroups[normName].storeDiscounts[supplier].push(discount);
    });

    console.log(`🔎 فهرس الفاكسات: ${Object.keys(faxGroups).length} صنف فريد`);

    // بناء النتيجة
    const result = collector.map(r=>{
      const code = String(r[cCol.code] ?? '').trim();
      const name = String(r[cCol.name] ?? '').trim();
      const salePrice = toNum(r[cCol.salePrice]);
      const weighted = toNum(r[cCol.weighted]);
      const stock = toNum(r[cCol.stock]);
      const netSale = toNum(r[cCol.netSale]);

      const recQty = recMap[normalize(code)] || 0;
      const ordQty = ordMap[normalize(code)] || 0;

      const faxGroup = faxGroups[normalize(name)] || null;
      const matchedFax = !!faxGroup;

      let faxDiscount = null;
      let supplier = '_____';
      let repeat = 0;
      let storesText = '—';
      let discountsText = '—';

      if(faxGroup){
        const storesArr = Array.from(faxGroup.stores);
        repeat = storesArr.length;
        supplier = storesArr.join(' | ') || '_____';
        storesText = storesArr.join(' | ') || '—';

        const allDiscounts = [];
        storesArr.forEach(store => {
          const discounts = faxGroup.storeDiscounts[store] || [];
          discounts.forEach(d => allDiscounts.push(d));
        });
        discountsText = allDiscounts.join(' | ') || '—';

        faxDiscount = allDiscounts.length ? allDiscounts[0] : null;
      }

      const totalQty = stock + recQty + ordQty;
      const required = netSale > totalQty ? (netSale - totalQty) : 0;

      return {
        code, name, salePrice, weighted, stock, recQty, ordQty, totalQty, netSale, required,
        supplier: supplier || '_____',
        faxDiscount: faxDiscount != null ? faxDiscount : '_____',
        matchedFax,
        repeat,
        stores: storesText,
        discounts: discountsText
      };
    });

    const total = result.length;
    const withFax = result.filter(r=>r.matchedFax).length;
    console.log(`✅ ربط الفاكسات: ${withFax}/${total}`);

    const masfar = result.filter(r => r.required !== 0 && r.totalQty <= 10);
    const هيسفر = result.filter(r => r.required > 5 && r.totalQty >= 11 && r.totalQty <= 2000);

    console.log(`📊 مصفر: ${masfar.length} صف`);
    console.log(`📊 هيصفر: ${هيسفر.length} صف`);

    return { all: result, masfar, هيسفر };
  },

  /* ============================================================
     5) المصفر العام — 3 ملفات
     ============================================================ */
  async runGeneralFilter(fileCollector, fileReceived, fileOrders){
    const dCol = await Excel.readFileAsArray(fileCollector);
    const collector = dCol.rows, colHeaders = dCol.headers;

    const dRec = fileReceived ? await Excel.readFileAsArray(fileReceived) : { headers: [], rows: [] };
    const dOrd = fileOrders ? await Excel.readFileAsArray(fileOrders) : { headers: [], rows: [] };
    const received = dRec.rows, recHeaders = dRec.headers;
    const orders = dOrd.rows, ordHeaders = dOrd.headers;

    if(colHeaders.length < 6){
      throw new Error(`❌ ملف المجمع يجب أن يحتوي على 6 أعمدة. الموجود: ${colHeaders.length}`);
    }

    const cCol = { code: colHeaders[0], name: colHeaders[1], salePrice: colHeaders[2], weighted: colHeaders[3], stock: colHeaders[4], netSale: colHeaders[5] };
    const cRec = received.length ? { code: recHeaders[0], qty: recHeaders[1], name: recHeaders[2] } : null;
    const cOrd = orders.length ? { code: ordHeaders[0], qty: ordHeaders[1], name: ordHeaders[2] } : null;

    const recMap = {};
    if(cRec) received.forEach(r=>{ const k = normalize(r[cRec.code]); if(k) recMap[k] = (recMap[k] || 0) + toNum(r[cRec.qty]); });
    const ordMap = {};
    if(cOrd) orders.forEach(r=>{ const k = normalize(r[cOrd.code]); if(k) ordMap[k] = (ordMap[k] || 0) + toNum(r[cOrd.qty]); });

    const result = collector.map(r=>{
      const code = String(r[cCol.code] ?? '').trim();
      const name = String(r[cCol.name] ?? '').trim();
      const salePrice = toNum(r[cCol.salePrice]);
      const weighted = toNum(r[cCol.weighted]);
      const stock = toNum(r[cCol.stock]);
      const netSale = toNum(r[cCol.netSale]);

      const recQty = recMap[normalize(code)] || 0;
      const ordQty = ordMap[normalize(code)] || 0;

      const totalQty = stock + recQty + ordQty;
      const required = netSale > totalQty ? (netSale - totalQty) : 0;

      return { code, name, salePrice, weighted, stock, recQty, ordQty, totalQty, netSale, required };
    });

    const masfar = result.filter(r => r.required !== 0 && r.totalQty <= 10);
    const هيسفر = result.filter(r => r.required > 5 && r.totalQty >= 11 && r.totalQty <= 2000);

    return { all: result, masfar, هيسفر };
  },

  /* ============================================================
     6) أوفر استوك — 4 ملفات
     ============================================================ */
  async runOverStock(fileReceived, fileCollector, fileBalances, fileOverStock){
    const dCol = await Excel.readFileAsArray(fileCollector);
    const collector = dCol.rows, colHeaders = dCol.headers;

    const dOS = await Excel.readFileAsArray(fileOverStock);
    const osRows = dOS.rows, osHeaders = dOS.headers;

    const dRec = fileReceived ? await Excel.readFileAsArray(fileReceived) : { headers: [], rows: [] };
    const received = dRec.rows, recHeaders = dRec.headers;

    const dBal = fileBalances ? await Excel.readFileAsArray(fileBalances) : { headers: [], rows: [] };
    const balances = dBal.rows, balHeaders = dBal.headers;

    console.log('═══════════════════════════════════════');
    console.log('📊 صفوف المجمع:', collector.length);
    console.log('📊 صفوف أوفر استوك:', osRows.length);
    console.log('📊 صفوف الوارد:', received.length);
    console.log('📊 صفوف الأرصدة:', balances.length);
    console.log('═══════════════════════════════════════');

    if(colHeaders.length < 6){
      throw new Error(`❌ ملف المجمع يجب أن يحتوي على 6 أعمدة. الموجود: ${colHeaders.length}`);
    }
    if(osHeaders.length < 6){
      throw new Error(`❌ ملف أوفر استوك يجب أن يحتوي على 6 أعمدة. الموجود: ${osHeaders.length}`);
    }

    const cCol = { code: colHeaders[0], name: colHeaders[1], salePrice: colHeaders[2], weighted: colHeaders[3], stock: colHeaders[4], netSale: colHeaders[5] };
    const cOS = { code: osHeaders[0], name: osHeaders[1], name2: osHeaders[2], writer: osHeaders[3], orders: osHeaders[4], discount: osHeaders[5] };
    const cRec = received.length ? { code: recHeaders[0], qty: recHeaders[1], name: recHeaders[2] } : null;
    const cBal = balances.length ? { code: balHeaders[0], weighted: balHeaders[1], name: balHeaders[2] } : null;

    const recMap = {};
    if(cRec) received.forEach(r=>{ const k = normalize(r[cRec.code]); if(k) recMap[k] = (recMap[k] || 0) + toNum(r[cRec.qty]); });

    const colMap = {};
    collector.forEach(r=>{
      const code = String(r[cCol.code] ?? '').trim();
      const k = normalize(code);
      if(!k) return;

      const name = String(r[cCol.name] ?? '').trim();
      const salePrice = toNum(r[cCol.salePrice]);
      const weighted = toNum(r[cCol.weighted]);
      const stock = toNum(r[cCol.stock]);
      const netSale = toNum(r[cCol.netSale]);

      const recQty = recMap[k] || 0;
      const totalQty = stock + recQty;
      const required = netSale > totalQty ? (netSale - totalQty) : 0;

      colMap[k] = { code, name, salePrice, weighted, stock, recQty, totalQty, netSale, required };
    });

    const balMap = {};
    if(cBal) balances.forEach(r=>{ const k = normalize(r[cBal.code]); if(k) balMap[k] = toNum(r[cBal.weighted]); });

    const result = osRows.map(r=>{
      const code = String(r[cOS.code] ?? '').trim();
      const name = String(r[cOS.name] ?? '').trim();
      const name2 = String(r[cOS.name2] ?? '').trim();
      const writer = String(r[cOS.writer] ?? '').trim();
      const orders = toNum(r[cOS.orders]);
      const discount = toNum(r[cOS.discount]);

      const k = normalize(code);
      const colData = colMap[k] || null;

      const stock = colData ? colData.totalQty : 0;
      const netSale = colData ? colData.netSale : 0;
      const required = colData ? colData.required : 0;
      const weighted = colData ? colData.weighted : 0;

      const overStock = orders > required ? (orders - required) : 0;
      const lastDiscount = balMap[k] || 0;
      const diff = weighted - discount;

      let analysis = '';
      if(required === 0 && stock === 0) analysis = 'استوك عادي';
      else if(required === 0 && stock > netSale) analysis = 'استوك للصنف';
      else if(Math.abs(stock - netSale) <= Math.max(5, netSale * 0.1) && weighted > discount) analysis = 'استوك للصنف';
      else if(overStock > 0) analysis = 'أوفر استوك';
      else analysis = 'مراجعة';

      return { code, name, name2, writer, orders, discount, stock, netSale, required, overStock, lastDiscount, weighted, diff, analysis, matchedCol: !!colData };
    });

    const filtered = result.filter(r => r.diff > -2);
    return filtered;
  },

  /* ============================================================
     اكتشاف أعمدة دقيق
     ============================================================ */
  detectColumns(sample, keysMap){
    const cols = Object.keys(sample || {});
    const result = {};
    const normCol = (c)=> normalize(String(c));

    Object.entries(keysMap).forEach(([target, keys])=>{
      if(!keys || !keys.length){ result[target] = null; return; }
      let found = cols.find(c => keys.some(k => normCol(c) === normalize(k)));
      if(found){ result[target] = found; return; }
      found = cols.find(c => keys.some(k => normCol(c).includes(normalize(k))));
      if(found){ result[target] = found; return; }
      result[target] = null;
    });

    const used = new Set();
    Object.keys(keysMap).forEach(k=>{
      if(result[k] && used.has(result[k])) result[k] = null;
      else if(result[k]) used.add(result[k]);
    });

    return result;
  },

  guessColumns(sample, codeKeys, nameKeys, priceKeys, discountKeys){
    const cols = Object.keys(sample || {});
    if(!cols.length) return { code:'', name:'', price:'', discount:'' };

    const normCol = (c)=> normalize(String(c));
    const used = new Set();

    const find = (keys)=>{
      if(!keys || !keys.length) return null;
      let found = cols.find(c => !used.has(c) && keys.some(k => normCol(c) === normalize(k)));
      if(found) return found;
      found = cols.find(c => !used.has(c) && keys.some(k => normCol(c).includes(normalize(k))));
      return found || null;
    };

    const code = find(codeKeys) || cols[0];
    used.add(code);
    const name = find(nameKeys) || cols.find(c=>!used.has(c)) || '';
    used.add(name);
    const price = find(priceKeys) || cols.find(c=>!used.has(c)) || '';
    used.add(price);
    const discount = find(discountKeys) || cols.find(c=>!used.has(c)) || '';
    used.add(discount);

    return { code, name, price, discount };
  }
};
