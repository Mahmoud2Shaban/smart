const Matching = {

  /* ============================================================
     1) مطابقة الأصناف
     ============================================================ */
  async runItemsMatch(fileBranch, fileMatch){
    const rowsBranch = await Excel.readFile(fileBranch);
    const rowsMatch = await Excel.readFile(fileMatch);
    if(!rowsBranch.length || !rowsMatch.length) throw new Error('أحد الملفين فارغ');

    const mapBranch = this.detectColumns(rowsBranch[0], {
      code:     ['كود','كود الصنف','الكود','code','item code','itemcode','رقم الصنف'],
      name:     ['صنف','اسم الصنف','الصنف','name','item name','itemname','الاسم','اسم المنتج','المنتج'],
      price:    ['سعر','السعر','price','سعر البيع','unit price','سعر الجمهور'],
      discount: []
    });

    const mapMatch = this.detectColumns(rowsMatch[0], {
      code:     [],
      name:     ['صنف','اسم الصنف','الصنف','name','item name','itemname','الاسم','اسم المنتج','المنتج'],
      price:    ['سعر','السعر','price','سعر البيع','unit price','سعر الجمهور'],
      discount: ['خصم','الخصم','discount','خصم الصنف','نسبة الخصم','discount %','disc','قيمة الخصم']
    });

    console.log('📋 أعمدة الفرع:', mapBranch);
    console.log('📋 أعمدة المطابقة:', mapMatch);

    if(!mapBranch.name) throw new Error('❌ لم يُعثر على "الصنف" في ملف الفرع. الأعمدة: ' + Object.keys(rowsBranch[0]).join(' | '));
    if(!mapMatch.name) throw new Error('❌ لم يُعثر على "الصنف" في ملف المطابقة. الأعمدة: ' + Object.keys(rowsMatch[0]).join(' | '));
    if(!mapMatch.discount){
      const cols = Object.keys(rowsMatch[0]);
      mapMatch.discount = cols[cols.length - 1];
    }

    const exactIndex = {}, normIndex = {}, partialIndex = [], wordIndex = [];

    rowsMatch.forEach(r=>{
      const rawName = String(r[mapMatch.name] ?? '').trim();
      if(!rawName) return;
      const normName = normalize(rawName);
      if(!normName) return;
      if(!exactIndex[rawName]) exactIndex[rawName] = r;
      if(!normIndex[normName]) normIndex[normName] = r;
      partialIndex.push({ normName, row: r });
      wordIndex.push({ normName, words: normName.split(' ').filter(w=>w.length>=2), row: r });
    });

    console.log(`🔎 فهرس المطابقة: ${Object.keys(normIndex).length} صنف`);

    const wordMatchScore = (wordsA, wordsB)=>{
      if(!wordsA.length || !wordsB.length) return 0;
      let common = 0;
      wordsA.forEach(w=>{ if(wordsB.includes(w)) common++; });
      return common / Math.max(wordsA.length, wordsB.length);
    };

    const levenshtein = (a, b)=>{
      if(!a.length) return b.length;
      if(!b.length) return a.length;
      const m = [];
      for(let i=0;i<=b.length;i++) m[i]=[i];
      for(let j=0;j<=a.length;j++) m[0][j]=j;
      for(let i=1;i<=b.length;i++){
        for(let j=1;j<=a.length;j++){
          m[i][j] = b[i-1]===a[j-1] ? m[i-1][j-1]
            : Math.min(m[i-1][j-1]+1, m[i][j-1]+1, m[i-1][j]+1);
        }
      }
      return m[b.length][a.length];
    };

    const result = rowsBranch.map(r=>{
      const code = r[mapBranch.code] ?? '';
      const rawName = String(r[mapBranch.name] ?? '').trim();
      const price = toNum(r[mapBranch.price]);
      const normName = normalize(rawName);
      const words = normName.split(' ').filter(w=>w.length>=2);

      let matchRow = null, matchType = '', score = 0;

      if(exactIndex[rawName]){ matchRow = exactIndex[rawName]; matchType = 'exact'; score = 1; }
      else if(normIndex[normName]){ matchRow = normIndex[normName]; matchType = 'normalized'; score = 1; }
      else if(normName.length >= 4){
        const candidates = partialIndex.filter(x =>
          x.normName === normName || x.normName.includes(normName) || normName.includes(x.normName)
        );
        if(candidates.length){
          candidates.sort((a,b)=> Math.abs(b.normName.length - normName.length) - Math.abs(a.normName.length - normName.length));
          matchRow = candidates[0].row;
          matchType = 'partial';
          score = 0.9;
        }
      }
      if(!matchRow && words.length >= 2){
        let best = null, bestScore = 0;
        wordIndex.forEach(x=>{
          const s = wordMatchScore(words, x.words);
          if(s > bestScore){ bestScore = s; best = x.row; }
        });
        if(bestScore >= 0.7){ matchRow = best; matchType = 'word'; score = bestScore; }
      }
      if(!matchRow && normName.length >= 6){
        let best = null, bestDist = Infinity;
        const maxDist = Math.max(2, Math.floor(normName.length * 0.15));
        wordIndex.forEach(x=>{
          if(Math.abs(x.normName.length - normName.length) > maxDist) return;
          const d = levenshtein(normName, x.normName);
          if(d < bestDist){ bestDist = d; best = x.row; }
        });
        if(bestDist <= maxDist){ matchRow = best; matchType = 'fuzzy'; score = 1 - (bestDist / normName.length); }
      }

      let discount = 0, matched = false, matchPrice = null;
      if(matchRow){
        matchPrice = toNum(matchRow[mapMatch.price]);
        matched = true;
        discount = toNum(matchRow[mapMatch.discount]);
      }

      return { code, name: rawName, price, discount, matched, matchPrice, matchType,
        matchName: matchRow ? String(matchRow[mapMatch.name] ?? '') : '', score };
    });

    const matchCount = result.filter(r=>r.matched).length;
    const ratio = result.length ? ((matchCount/result.length)*100).toFixed(1) : 0;

    const typeStats = {};
    result.filter(r=>r.matched).forEach(r=>{ typeStats[r.matchType] = (typeStats[r.matchType]||0) + 1; });
    console.log(`✅ المطابقة: ${matchCount}/${result.length} (${ratio}%)`, typeStats);

    return { rows: result, ratio, matchCount, total: result.length };
  },

  /* ============================================================
     2) فرق الجملة — ربط صارم بالأعمدة حسب الترتيب
     
     الملف 1 (الربط):     [كود طيبة] [كود الفرع] [الصنف] [الخصم]
     الملف 2 (فرق الجملة): [الكود] [المورد] [الصنف] [الخصم]
     الملف 3 (الجملة):     [الكود] [الخصم] [الصنف]
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
    console.log('📋 رأس الربط:', linkHeaders);
    console.log('📋 رأس فرق الجملة:', wholeDiffHeaders);
    console.log('📋 رأس الجملة:', wholeHeaders);
    console.log('═══════════════════════════════════════');

    if(linkHeaders.length < 4){
      throw new Error(`❌ ملف الربط يجب أن يحتوي على 4 أعمدة على الأقل: [كود طيبة، كود الفرع، الصنف، الخصم]. الموجود: ${linkHeaders.length} (${linkHeaders.join(' | ')})`);
    }
    if(wholeDiffHeaders.length < 4){
      throw new Error(`❌ ملف فرق الجملة يجب أن يحتوي على 4 أعمدة: [الكود، المورد، الصنف، الخصم]. الموجود: ${wholeDiffHeaders.length} (${wholeDiffHeaders.join(' | ')})`);
    }
    if(wholeHeaders.length < 3){
      throw new Error(`❌ ملف الجملة يجب أن يحتوي على 3 أعمدة: [الكود، الخصم، الصنف]. الموجود: ${wholeHeaders.length} (${wholeHeaders.join(' | ')})`);
    }

    const cLink = {
      taibaCode:  linkHeaders[0],
      branchCode: linkHeaders[1],
      name:       linkHeaders[2],
      discount:   linkHeaders[3]
    };

    const cWholeDiff = {
      code:     wholeDiffHeaders[0],
      supplier: wholeDiffHeaders[1],
      name:     wholeDiffHeaders[2],
      discount: wholeDiffHeaders[3]
    };

    const cWhole = {
      code:     wholeHeaders[0],
      discount: wholeHeaders[1],
      name:     wholeHeaders[2]
    };

    console.log('✅ الربط:', cLink);
    console.log('✅ فرق الجملة:', cWholeDiff);
    console.log('✅ الجملة:', cWhole);

    const wholeDiffMap = {};
    wholeDiff.forEach(r=>{
      const k = normalize(r[cWholeDiff.code]);
      if(k && !wholeDiffMap[k]) wholeDiffMap[k] = r;
    });

    const wholeMap = {};
    whole.forEach(r=>{
      const k = normalize(r[cWhole.code]);
      if(k && !wholeMap[k]) wholeMap[k] = r;
    });

    console.log(`🔎 فهرس فرق الجملة: ${Object.keys(wholeDiffMap).length} كود`);
    console.log(`🔎 فهرس الجملة: ${Object.keys(wholeMap).length} كود`);

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

      return {
        taibaCode,
        branchCode,
        name,
        discount,
        supplier,
        wholeDiscount,
        wholeName,
        diff,
        matchedWholeDiff: !!wdRow,
        matchedWhole: !!wRow
      };
    });

    result.sort((a, b) => b.diff - a.diff);

    const matchedBoth = result.filter(r=>r.matchedWholeDiff && r.matchedWhole).length;
    const matchedWD = result.filter(r=>r.matchedWholeDiff).length;
    const matchedW = result.filter(r=>r.matchedWhole).length;
    console.log(`✅ ربط فرق الجملة: ${matchedWD}/${result.length}`);
    console.log(`✅ ربط الجملة: ${matchedW}/${result.length}`);
    console.log(`✅ ربط كامل (الاثنين): ${matchedBoth}/${result.length}`);

    return result;
  },

  /* ============================================================
     3) فرق الفاكسات — مطابقة تامة فقط
     
     الملف 1 (فرق الفاكسات): [الصنف] [الاسم] [الكمية] [الكاتب] [الخصم]
     الملف 2 (الفاكسات):      [الصنف] [خصم الفاكسات] [المورد]
     ============================================================ */
  async runFaxDiff(fileFaxDiff, fileFax){
    const d1 = await Excel.readFileAsArray(fileFaxDiff);
    const d2 = await Excel.readFileAsArray(fileFax);

    const diffRows = d1.rows;
    const diffHeaders = d1.headers;
    const faxRows = d2.rows;
    const faxHeaders = d2.headers;

    console.log('═══════════════════════════════════════');
    console.log('📊 صفوف فرق الفاكسات:', diffRows.length);
    console.log('📊 صفوف الفاكسات:', faxRows.length);
    console.log('📋 رأس فرق الفاكسات:', diffHeaders);
    console.log('📋 رأس الفاكسات:', faxHeaders);
    console.log('═══════════════════════════════════════');

    if(diffHeaders.length < 5){
      throw new Error(`❌ ملف فرق الفاكسات يجب أن يحتوي على 5 أعمدة على الأقل: [الصنف، الاسم، الكمية، الكاتب، الخصم]. الموجود: ${diffHeaders.length} (${diffHeaders.join(' | ')})`);
    }
    if(faxHeaders.length < 3){
      throw new Error(`❌ ملف الفاكسات يجب أن يحتوي على 3 أعمدة: [الصنف، خصم الفاكسات، المورد]. الموجود: ${faxHeaders.length} (${faxHeaders.join(' | ')})`);
    }

    const cDiff = {
      name:     diffHeaders[0],
      store:    diffHeaders[1],
      qty:      diffHeaders[2],
      writer:   diffHeaders[3],
      discount: diffHeaders[4]
    };

    const cFax = {
      name:     faxHeaders[0],
      discount: faxHeaders[1],
      supplier: faxHeaders[2]
    };

    console.log('✅ فرق الفاكسات:', cDiff);
    console.log('✅ الفاكسات:', cFax);

    const faxNormMap = {};

    faxRows.forEach(r=>{
      const rawName = String(r[cFax.name] ?? '').trim();
      if(!rawName) return;
      const normName = normalize(rawName);
      if(!normName) return;
      if(!faxNormMap[normName]) faxNormMap[normName] = r;
    });

    console.log(`🔎 فهرس الفاكسات (تام فقط): ${Object.keys(faxNormMap).length} صنف فريد`);

    const result = diffRows.map(r=>{
      const rawName = String(r[cDiff.name] ?? '').trim();
      const name = rawName;
      const storeName = String(r[cDiff.store] ?? '').trim();
      const qty = toNum(r[cDiff.qty]);
      const writer = String(r[cDiff.writer] ?? '').trim();
      const discount = toNum(r[cDiff.discount]);
      const normName = normalize(rawName);

      const faxRow = faxNormMap[normName] || null;
      const matched = !!faxRow;

      let faxDiscount = null;
      let supplier = '_____';

      if(faxRow){
        faxDiscount = toNum(faxRow[cFax.discount]);
        supplier = String(faxRow[cFax.supplier] ?? '_____') || '_____';
      }

      const diff = (faxDiscount != null) ? (faxDiscount - discount) : null;

      return {
        name,
        storeName,
        qty,
        writer,
        discount,
        faxDiscount,
        supplier,
        diff,
        matched
      };
    });

    result.sort((a, b) => (b.diff ?? -Infinity) - (a.diff ?? -Infinity));

    const matchedCount = result.filter(r=>r.matched).length;
    const unmatchedCount = result.length - matchedCount;
    console.log(`✅ مطابقة تامة: ${matchedCount}/${result.length}`);
    console.log(`❌ غير موجود في ملف الفاكسات: ${unmatchedCount}`);

    return result;
  },

  /* ============================================================
     4) مصفر الفاكسات — ربط صارم بالأعمدة حسب الترتيب
     
     الملف 1 (الوارد):   [الكود] [الكمية] [الصنف]
     الملف 2 (الطلبات):  [الكود] [الكمية] [الصنف]
     الملف 3 (المجمع):   [الكود] [الاسم] [س البيع] [المرجح] [ك مخزون] [ك صافي بيع]
     الملف 4 (الفاكسات): [الصنف] [الخصم] [المورد]
     
     ⚠️ شرط هيصفر: المطلوب > 5 (وليس > 0)
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
    console.log('📋 رأس المجمع:', colHeaders);
    console.log('📋 رأس الفاكسات:', faxHeaders);
    console.log('📋 رأس الوارد:', recHeaders);
    console.log('📋 رأس الطلبات:', ordHeaders);
    console.log('═══════════════════════════════════════');

    if(colHeaders.length < 6){
      throw new Error(`❌ ملف المجمع يجب أن يحتوي على 6 أعمدة: [الكود، الاسم، س البيع، المرجح، ك مخزون، ك صافي بيع]. الموجود: ${colHeaders.length} (${colHeaders.join(' | ')})`);
    }
    if(faxHeaders.length < 3){
      throw new Error(`❌ ملف الفاكسات يجب أن يحتوي على 3 أعمدة: [الصنف، الخصم، المورد]. الموجود: ${faxHeaders.length} (${faxHeaders.join(' | ')})`);
    }
    if(received.length && recHeaders.length < 3){
      throw new Error(`❌ ملف الوارد يجب أن يحتوي على 3 أعمدة: [الكود، الكمية، الصنف]. الموجود: ${recHeaders.length}`);
    }
    if(orders.length && ordHeaders.length < 3){
      throw new Error(`❌ ملف الطلبات يجب أن يحتوي على 3 أعمدة: [الكود، الكمية، الصنف]. الموجود: ${ordHeaders.length}`);
    }

    const cCol = {
      code:     colHeaders[0],
      name:     colHeaders[1],
      salePrice: colHeaders[2],
      weighted: colHeaders[3],
      stock:    colHeaders[4],
      netSale:  colHeaders[5]
    };

    const cFax = {
      name:     faxHeaders[0],
      discount: faxHeaders[1],
      supplier: faxHeaders[2]
    };

    const cRec = received.length ? {
      code: recHeaders[0],
      qty:  recHeaders[1],
      name: recHeaders[2]
    } : null;

    const cOrd = orders.length ? {
      code: ordHeaders[0],
      qty:  ordHeaders[1],
      name: ordHeaders[2]
    } : null;

    console.log('✅ المجمع:', cCol);
    console.log('✅ الفاكسات:', cFax);
    if(cRec) console.log('✅ الوارد:', cRec);
    if(cOrd) console.log('✅ الطلبات:', cOrd);

    // فهرس الوارد (الكود → مجموع الكميات)
    const recMap = {};
    received.forEach(r=>{
      const k = normalize(r[cRec.code]);
      if(k) recMap[k] = (recMap[k] || 0) + toNum(r[cRec.qty]);
    });

    // فهرس الطلبات (الكود → مجموع الكميات)
    const ordMap = {};
    orders.forEach(r=>{
      const k = normalize(r[cOrd.code]);
      if(k) ordMap[k] = (ordMap[k] || 0) + toNum(r[cOrd.qty]);
    });

    // فهرس الفاكسات (الصنف → صف)
    const faxMap = {};
    faxRows.forEach(r=>{
      const k = normalize(r[cFax.name]);
      if(k && !faxMap[k]) faxMap[k] = r;
    });

    console.log(`🔎 فهرس الوارد: ${Object.keys(recMap).length} كود`);
    console.log(`🔎 فهرس الطلبات: ${Object.keys(ordMap).length} كود`);
    console.log(`🔎 فهرس الفاكسات: ${Object.keys(faxMap).length} صنف`);

    const result = collector.map(r=>{
      const code = String(r[cCol.code] ?? '').trim();
      const name = String(r[cCol.name] ?? '').trim();
      const salePrice = toNum(r[cCol.salePrice]);
      const weighted = toNum(r[cCol.weighted]);
      const stock = toNum(r[cCol.stock]);
      const netSale = toNum(r[cCol.netSale]);

      const recQty = recMap[normalize(code)] || 0;
      const ordQty = ordMap[normalize(code)] || 0;

      const fRow = faxMap[normalize(name)] || null;
      const faxDiscount = fRow ? toNum(fRow[cFax.discount]) : null;
      const supplier = fRow ? String(fRow[cFax.supplier] ?? '') : '';

      const totalQty = stock + recQty + ordQty;
      const required = netSale > totalQty ? (netSale - totalQty) : 0;

      return {
        code,
        name,
        salePrice,
        weighted,
        stock,
        recQty,
        ordQty,
        totalQty,
        netSale,
        required,
        supplier: supplier || '_____',
        faxDiscount: faxDiscount != null ? faxDiscount : '_____',
        matchedFax: !!fRow
      };
    });

    const total = result.length;
    const withFax = result.filter(r=>r.matchedFax).length;
    console.log(`✅ ربط الفاكسات: ${withFax}/${total}`);

    // === مصفر الفاكسات: المطلوب ≠ 0 + إجمالي الكمية ≤ 10 ===
    const masfar = result.filter(r => r.required !== 0 && r.totalQty <= 10);

    // === هيصفر الفاكسات: المطلوب > 5 + إجمالي الكمية بين 11 و 2000 ===
    const هيسفر = result.filter(r => r.required > 5 && r.totalQty >= 11 && r.totalQty <= 2000);

    console.log(`📊 مصفر: ${masfar.length} صف`);
    console.log(`📊 هيصفر: ${هيسفر.length} صف`);

    return { all: result, masfar, هيسفر };
  },

  /* ============================================================
   5) المصفر العام — ربط صارم بالأعمدة حسب الترتيب
   ============================================================ */
  async runGeneralFilter(fileCollector, fileReceived, fileOrders){
    // === قراءة المجمع (إلزامي) ===
    const dCol = await Excel.readFileAsArray(fileCollector);
    const collector = dCol.rows, colHeaders = dCol.headers;

    // === قراءة الوارد والطلبات (اختياري) ===
    const dRec = fileReceived ? await Excel.readFileAsArray(fileReceived) : { headers: [], rows: [] };
    const dOrd = fileOrders ? await Excel.readFileAsArray(fileOrders) : { headers: [], rows: [] };
    const received = dRec.rows, recHeaders = dRec.headers;
    const orders = dOrd.rows, ordHeaders = dOrd.headers;

    console.log('═══════════════════════════════════════');
    console.log('📊 صفوف المجمع:', collector.length);
    console.log('📊 صفوف الوارد:', received.length);
    console.log('📊 صفوف الطلبات:', orders.length);
    console.log('📋 رأس المجمع:', colHeaders);
    console.log('📋 رأس الوارد:', recHeaders);
    console.log('📋 رأس الطلبات:', ordHeaders);
    console.log('═══════════════════════════════════════');

    if(colHeaders.length < 6){
      throw new Error(`❌ ملف المجمع يجب أن يحتوي على 6 أعمدة: [الكود، الاسم، س البيع، المرجح، ك مخزون، ك صافي بيع]. الموجود: ${colHeaders.length} (${colHeaders.join(' | ')})`);
    }
    if(received.length && recHeaders.length < 3){
      throw new Error(`❌ ملف الوارد يجب أن يحتوي على 3 أعمدة: [الكود، الكمية، الصنف]. الموجود: ${recHeaders.length} (${recHeaders.join(' | ')})`);
    }
    if(orders.length && ordHeaders.length < 3){
      throw new Error(`❌ ملف الطلبات يجب أن يحتوي على 3 أعمدة: [الكود، الكمية، الصنف]. الموجود: ${ordHeaders.length} (${ordHeaders.join(' | ')})`);
    }

    const cCol = {
      code:      colHeaders[0],
      name:      colHeaders[1],
      salePrice: colHeaders[2],
      weighted:  colHeaders[3],
      stock:     colHeaders[4],
      netSale:   colHeaders[5]
    };

    const cRec = received.length ? { code: recHeaders[0], qty: recHeaders[1], name: recHeaders[2] } : null;
    const cOrd = orders.length ? { code: ordHeaders[0], qty: ordHeaders[1], name: ordHeaders[2] } : null;

    console.log('✅ المجمع:', cCol);
    if(cRec) console.log('✅ الوارد:', cRec);
    if(cOrd) console.log('✅ الطلبات:', cOrd);

    // فهرس الوارد (الكود → مجموع الكميات)
    const recMap = {};
    if(cRec){
      received.forEach(r=>{
        const k = normalize(r[cRec.code]);
        if(k) recMap[k] = (recMap[k] || 0) + toNum(r[cRec.qty]);
      });
    }

    // فهرس الطلبات (الكود → مجموع الكميات)
    const ordMap = {};
    if(cOrd){
      orders.forEach(r=>{
        const k = normalize(r[cOrd.code]);
        if(k) ordMap[k] = (ordMap[k] || 0) + toNum(r[cOrd.qty]);
      });
    }

    console.log(`🔎 فهرس الوارد: ${Object.keys(recMap).length} كود`);
    console.log(`🔎 فهرس الطلبات: ${Object.keys(ordMap).length} كود`);

    // بناء الصفوف الناتجة
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

    console.log(`✅ إجمالي الأصناف: ${result.length}`);

    // === المصفر: المطلوب ≠ 0 + إجمالي الكمية ≤ 10 ===
    const masfar = result.filter(r => r.required !== 0 && r.totalQty <= 10);

    // === هيصفر: المطلوب > 5 + إجمالي الكمية بين 11 و 2000 ===
    const هيسفر = result.filter(r => r.required > 5 && r.totalQty >= 11 && r.totalQty <= 2000);

    console.log(`📊 المصفر: ${masfar.length} صف`);
    console.log(`📊 هيصفر: ${هيسفر.length} صف`);

    return { all: result, masfar, هيسفر };
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