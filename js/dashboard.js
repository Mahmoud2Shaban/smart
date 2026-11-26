const Dashboard = {

  renderBranchPage(branchKey){
    const container = document.getElementById('page-branch');
    container.classList.remove('hidden');
    container.classList.add('active');
    document.getElementById('page-home').classList.add('hidden');
    document.getElementById('page-admin').classList.add('hidden');

    const b = BRANCHES[branchKey] || { name: branchKey, icon: 'fa-store' };
    const isTaiba = branchKey === 'taiba';

    container.innerHTML = `
      <div class="branch-header">
        <h1><i class="fas ${b.icon}"></i> ${b.name}</h1>
        <p><i class="fas fa-circle-info"></i> لوحة إدارة ومتابعة العمليات للفرع</p>
      </div>
      <div class="sub-tabs" id="branchTabs">
        <button class="active" data-view="branch-home"><i class="fas fa-house"></i> الرئيسية</button>
        <button data-view="items-match"><i class="fas fa-table-cells-large"></i> مطابقة الأصناف</button>
        ${isTaiba ? '' : `
        <button data-view="whole-diff"><i class="fas fa-people-arrows"></i> فرق الجملة</button>
        <button data-view="fax-diff"><i class="fas fa-receipt"></i> فرق الفاكسات</button>
        <button data-view="fax-filter"><i class="fas fa-filter"></i> مصفر الفاكسات</button>
        <button data-view="general-filter"><i class="fas fa-layer-group"></i> المصفر</button>
        <button data-view="over-stock"><i class="fas fa-boxes-packing"></i> أوفر استوك</button>`}
      </div>
      <div id="branchView"></div>
    `;

    container.querySelectorAll('#branchTabs button').forEach(btn=>{
      btn.onclick = ()=>{
        container.querySelectorAll('#branchTabs button').forEach(x=>x.classList.remove('active'));
        btn.classList.add('active');
        this.renderBranchView(branchKey, btn.dataset.view);
      };
    });
    this.renderBranchView(branchKey, 'branch-home');
  },

  renderBranchView(branchKey, view){
    const el = document.getElementById('branchView');
    if(!el) return;
    const dataKey = branchKey;

    if(view === 'branch-home'){
      el.innerHTML = `
        <div class="features-grid">
          <div class="feature-card"><i class="fas fa-table-cells-large"></i><h3>مطابقة الأصناف</h3><p>رفع ملفين ومطابقة الأصناف والأسعار</p></div>
          ${branchKey === 'taiba' ? '' : `
          <div class="feature-card"><i class="fas fa-people-arrows"></i><h3>فرق الجملة</h3><p>ربط 3 ملفات وحساب فرق الجملة</p></div>
          <div class="feature-card"><i class="fas fa-receipt"></i><h3>فرق الفاكسات</h3><p>حساب الفروق بين ملفي الفاكسات</p></div>
          <div class="feature-card"><i class="fas fa-filter"></i><h3>مصفر الفاكسات</h3><p>حساب المطلوب والمصفر والفاكسات</p></div>
          <div class="feature-card"><i class="fas fa-layer-group"></i><h3>المصفر العام</h3><p>حساب المصفر العام للفرع</p></div>
          <div class="feature-card"><i class="fas fa-boxes-packing"></i><h3>أوفر استوك</h3><p>تحليل ذكي للأوفر استوك</p></div>`}
        </div>`;
      return;
    }

    /* 1) مطابقة الأصناف */
    if(view === 'items-match'){
      el.innerHTML = `
        <div class="card">
          <h3><i class="fas fa-table-cells-large"></i> مطابقة الأصناف</h3>
          ${this.uploadZone('zoneBranch','ملف الفرع (كود، صنف، سعر)', ['كود','صنف','سعر'], 'fa-store')}
          ${this.uploadZone('zoneMatch','ملف المطابقة (صنف، سعر، خصم)', ['صنف','خصم'], 'fa-clipboard-check')}
          <div class="action-row">
            <button class="btn-primary" id="btnRunMatch"><i class="fas fa-play"></i> المتابعة والمطابقة</button>
            <button class="btn-ghost" id="btnClearMatch"><i class="fas fa-trash-can"></i> حذف البيانات</button>
          </div>
          <div class="toolbar">
            <input type="text" id="searchMatch" placeholder="🔍 بحث عن الصنف بالاسم...">
            <select id="filterMatch">
              <option value="all">كل الأصناف</option>
              <option value="matched">المطابق فقط</option>
              <option value="unmatched">غير المطابق فقط</option>
            </select>
            <button class="btn-primary" id="btnExportMatch" style="width:auto"><i class="fas fa-file-excel"></i> تحميل Excel</button>
          </div>
          <div id="matchResults"></div>
        </div>`;

      this.wireUploadZones(['zoneBranch','zoneMatch']);
      let currentRows = [];

      const render = ()=>{
        const q = normalize(document.getElementById('searchMatch').value);
        const f = document.getElementById('filterMatch').value;
        let rows = currentRows.filter(r => (!q || normalize(r.name).includes(q)));
        if(f === 'matched') rows = rows.filter(r => r.matched);
        if(f === 'unmatched') rows = rows.filter(r => !r.matched);

        const matchedCount = currentRows.filter(r=>r.matched).length;
        const unmatchedCount = currentRows.length - matchedCount;
        const ratio = currentRows.length ? ((matchedCount / currentRows.length) * 100).toFixed(1) : 0;
        const matchTypeLabel = { exact:'تام', normalized:'مطبّع', partial:'جزئي', word:'كلمات', fuzzy:'ضبابي', '':'' };

        if(!rows.length){
          document.getElementById('matchResults').innerHTML = `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>${currentRows.length ? 'لا توجد نتائج مطابقة للبحث' : 'لا توجد بيانات — قم برفع الملفين واضغط "المتابعة والمطابقة"'}</p>
            </div>`;
          return;
        }

        let tableRows = '';
        rows.forEach((r,i)=>{
          tableRows += `<tr>
            <td>${i+1}</td><td>${r.code}</td><td>${r.name}</td><td>${r.price}</td>
            <td><b style="color:${r.discount>0?'#166534':'#64748b'}">${r.discount}</b></td>
            <td style="font-size:12px;color:#64748b">${r.matchName || '—'}</td>
            <td style="font-size:11px;color:#64748b">${matchTypeLabel[r.matchType] || '—'}</td>
            <td><span class="${r.matched?'badge-match':'badge-miss'}"><i class="fas ${r.matched?'fa-circle-check':'fa-circle-xmark'}"></i>${r.matched?'مطابق':'غير مطابق'}</span></td>
          </tr>`;
        });

        document.getElementById('matchResults').innerHTML = `
          <div class="stats-grid" style="margin:15px 0">
            <div class="stat-card"><i class="fas fa-list-ol"></i><div class="stat-num">${currentRows.length}</div><div class="stat-label">إجمالي الصفوف</div></div>
            <div class="stat-card"><i class="fas fa-circle-check" style="color:var(--success)"></i><div class="stat-num" style="color:var(--success)">${matchedCount}</div><div class="stat-label">مطابق</div></div>
            <div class="stat-card"><i class="fas fa-circle-xmark" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${unmatchedCount}</div><div class="stat-label">غير مطابق</div></div>
            <div class="stat-card"><i class="fas fa-percent"></i><div class="stat-num">${ratio}%</div><div class="stat-label">نسبة المطابقة</div></div>
          </div>
          <div style="margin:10px 0;color:var(--muted);font-size:13px"><i class="fas fa-info-circle"></i> يُعرض ${rows.length} من ${currentRows.length} صف</div>
          <div class="table-wrap"><table>
            <thead><tr><th>#</th><th>الكود</th><th>الصنف (الفرع)</th><th>السعر</th><th>الخصم</th><th>اسم المطابقة</th><th>نوع</th><th>الحالة</th></tr></thead>
            <tbody>${tableRows}</tbody>
          </table></div>`;
      };

      document.getElementById('searchMatch').oninput = render;
      document.getElementById('filterMatch').onchange = render;

      document.getElementById('btnRunMatch').onclick = async ()=>{
        const btn = document.getElementById('btnRunMatch');
        const z1 = document.getElementById('zoneBranch');
        const z2 = document.getElementById('zoneMatch');
        if(!z1._file){ showToast('❌ ملف الفرع غير صالح','error'); z1.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        if(!z2._file){ showToast('❌ ملف المطابقة غير صالح','error'); z2.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        setLoading(btn, true);
        try{
          const { rows } = await Matching.runItemsMatch(z1._file, z2._file);
          currentRows = rows;
          App.branchesData[dataKey] = { ...(App.branchesData[dataKey]||{}), itemsMatch: rows };
          App.saveData();
          render();
          showToast('تم المطابقة بنجاح','success');
        }catch(err){ showToast(err.message || 'حدث خطأ','error'); }
        setLoading(btn, false);
      };

      document.getElementById('btnExportMatch').onclick = ()=>{
        if(!currentRows.length){ showToast('لا توجد بيانات','error'); return; }
        const out = currentRows.map(r=>({
          'الكود': r.code, 'الصنف': r.name, 'السعر': r.price, 'الخصم': r.discount,
          'اسم المطابقة': r.matchName || '', 'سعر المطابقة': r.matchPrice ?? '',
          'الحالة': r.matched ? 'مطابق' : 'غير مطابق'
        }));
        Excel.download(out, 'مطابقة_الأصناف.xlsx', 'مطابقة الأصناف');
      };

      document.getElementById('btnClearMatch').onclick = async ()=>{
        if(await confirmModal('حذف البيانات','حذف بيانات المطابقة؟')){
          currentRows = [];
          document.getElementById('matchResults').innerHTML = '';
          Excel.resetAllZones();
          if(App.branchesData[dataKey]) delete App.branchesData[dataKey].itemsMatch;
          App.saveData();
          showToast('تم الحذف','success');
        }
      };
      return;
    }

    /* 2) فرق الجملة */
    if(view === 'whole-diff'){
      el.innerHTML = `
        <div class="card">
          <h3><i class="fas fa-people-arrows"></i> فرق الجملة</h3>
          ${this.uploadZone('zoneLink','الملف 1: الربط (كود طيبة، كود الفرع، الصنف، الخصم)', null, 'fa-link')}
          ${this.uploadZone('zoneWholeDiff','الملف 2: فرق الجملة (الكود، المورد، الصنف، الخصم)', null, 'fa-people-arrows')}
          ${this.uploadZone('zoneWhole','الملف 3: الجملة (الكود، الخصم، الصنف)', null, 'fa-boxes-stacked')}
          <div class="action-row">
            <button class="btn-primary" id="btnRunWhole"><i class="fas fa-calculator"></i> حساب الفرق</button>
            <button class="btn-ghost" id="btnClearWhole"><i class="fas fa-trash-can"></i> حذف البيانات</button>
          </div>
          <div class="toolbar">
            <input type="text" id="searchWhole" placeholder="🔍 بحث بالصنف أو الكود أو المورد...">
            <select id="sortWhole">
              <option value="desc">⬇️ من الأكبر إلى الأصغر (الفرق)</option>
              <option value="asc">⬆️ من الأصغر إلى الأكبر (الفرق)</option>
              <option value="all">📋 بدون ترتيب (الكل)</option>
            </select>
            <button class="btn-primary" id="exportWhole" style="width:auto"><i class="fas fa-file-excel"></i> تحميل Excel</button>
          </div>
          <div id="wholeResults"></div>
        </div>`;

      this.wireUploadZones(['zoneLink','zoneWholeDiff','zoneWhole']);
      let currentRows = [];
      let baseRows = [];

      const render = ()=>{
        const q = normalize(document.getElementById('searchWhole').value);
        const sort = document.getElementById('sortWhole').value;

        let rows = baseRows.filter(r =>
          !q || normalize(r.name || '').includes(q) || normalize(r.taibaCode || '').includes(q) ||
          normalize(r.branchCode || '').includes(q) || normalize(r.supplier || '').includes(q) ||
          normalize(r.wholeName || '').includes(q)
        );

        if(sort === 'desc') rows = [...rows].sort((a,b) => b.diff - a.diff);
        else if(sort === 'asc') rows = [...rows].sort((a,b) => a.diff - b.diff);

        const total = baseRows.length;
        const positive = baseRows.filter(r=>r.diff > 0).length;
        const negative = baseRows.filter(r=>r.diff < 0).length;
        const zero = baseRows.filter(r=>r.diff === 0).length;
        const matchedBoth = baseRows.filter(r=>r.matchedWholeDiff && r.matchedWhole).length;

        if(!rows.length){
          document.getElementById('wholeResults').innerHTML = `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>${baseRows.length ? 'لا توجد نتائج مطابقة للبحث' : 'لا توجد بيانات — قم برفع الملفات الثلاثة واضغط "حساب الفرق"'}</p>
            </div>`;
          return;
        }

        let tableRows = '';
        rows.forEach((r, i)=>{
          const diffVal = (typeof r.diff === 'number') ? r.diff : 0;
          let diffClass = '', diffIcon = 'fa-equals';
          if(diffVal > 0){ diffClass = 'badge-match'; diffIcon = 'fa-arrow-trend-up'; }
          else if(diffVal < 0){ diffClass = 'badge-miss'; diffIcon = 'fa-arrow-trend-down'; }

          tableRows += `<tr>
            <td>${i+1}</td>
            <td>${r.taibaCode || ''}</td>
            <td>${r.branchCode || ''}</td>
            <td>${r.name || ''}</td>
            <td>${r.discount ?? ''}</td>
            <td>${r.supplier || '—'}</td>
            <td>${r.wholeDiscount ?? ''}</td>
            <td>${r.wholeName || '—'}</td>
            <td><span class="${diffClass}" style="min-width:80px;justify-content:center;display:inline-flex;align-items:center;gap:4px"><i class="fas ${diffIcon}"></i> ${diffVal}</span></td>
          </tr>`;
        });

        document.getElementById('wholeResults').innerHTML = `
          <div class="stats-grid" style="margin:15px 0;grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">
            <div class="stat-card"><i class="fas fa-list-ol"></i><div class="stat-num">${total}</div><div class="stat-label">إجمالي</div></div>
            <div class="stat-card"><i class="fas fa-arrow-trend-up" style="color:var(--success)"></i><div class="stat-num" style="color:var(--success)">${positive}</div><div class="stat-label">موجب</div></div>
            <div class="stat-card"><i class="fas fa-arrow-trend-down" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${negative}</div><div class="stat-label">سالب</div></div>
            <div class="stat-card"><i class="fas fa-equals" style="color:var(--muted)"></i><div class="stat-num" style="color:var(--muted)">${zero}</div><div class="stat-label">صفر</div></div>
            <div class="stat-card"><i class="fas fa-link" style="color:var(--primary)"></i><div class="stat-num" style="color:var(--primary)">${matchedBoth}</div><div class="stat-label">ربط كامل</div></div>
          </div>
          <div style="margin:10px 0;color:var(--muted);font-size:13px"><i class="fas fa-info-circle"></i> يُعرض ${rows.length} من ${total} صف</div>
          <div class="table-wrap" style="max-height:70vh;overflow:auto">
            <table>
              <thead><tr><th>#</th><th>كود طيبة</th><th>كود الفرع</th><th>الصنف</th><th>الخصم</th><th>المورد</th><th>خصم الجملة</th><th>صنف الجملة</th><th>الفرق</th></tr></thead>
              <tbody>${tableRows}</tbody>
            </table>
          </div>`;
      };

      document.getElementById('searchWhole').oninput = render;
      document.getElementById('sortWhole').onchange = render;

      document.getElementById('btnRunWhole').onclick = async ()=>{
        const btn = document.getElementById('btnRunWhole');
        const z1 = document.getElementById('zoneLink');
        const z2 = document.getElementById('zoneWholeDiff');
        const z3 = document.getElementById('zoneWhole');
        if(!z1._file){ showToast('❌ الملف 1 غير صالح','error'); z1.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        if(!z2._file){ showToast('❌ الملف 2 غير صالح','error'); z2.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        if(!z3._file){ showToast('❌ الملف 3 غير صالح','error'); z3.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        setLoading(btn, true);
        try{
          const rows = await Matching.runWholeDiff(z1._file, z2._file, z3._file);
          baseRows = [...rows];
          currentRows = [...rows].sort((a,b) => b.diff - a.diff);
          App.branchesData[dataKey] = { ...(App.branchesData[dataKey]||{}), wholeDiff: currentRows };
          App.saveData();
          render();
          showToast('تم حساب الفرق','success');
        }catch(err){ showToast(err.message || 'حدث خطأ','error'); }
        setLoading(btn, false);
      };

      document.getElementById('btnClearWhole').onclick = async ()=>{
        if(await confirmModal('حذف','حذف البيانات؟')){
          currentRows = []; baseRows = [];
          document.getElementById('wholeResults').innerHTML = '';
          Excel.resetAllZones();
          if(App.branchesData[dataKey]) delete App.branchesData[dataKey].wholeDiff;
          App.saveData();
          showToast('تم الحذف','success');
        }
      };

      document.getElementById('exportWhole').onclick = ()=>{
        if(!baseRows.length){ showToast('لا توجد بيانات','error'); return; }
        const sort = document.getElementById('sortWhole').value;
        let rows = [...baseRows];
        if(sort === 'desc') rows.sort((a,b) => b.diff - a.diff);
        else if(sort === 'asc') rows.sort((a,b) => a.diff - b.diff);
        const out = rows.map(r=>({
          'كود طيبة': r.taibaCode || '', 'كود الفرع': r.branchCode || '',
          'الصنف': r.name || '', 'الخصم': r.discount ?? '',
          'المورد': r.supplier || '', 'خصم الجملة': r.wholeDiscount ?? '',
          'صنف الجملة': r.wholeName || '', 'الفرق': r.diff ?? 0
        }));
        Excel.download(out, 'فرق_الجملة.xlsx', 'فرق الجملة');
      };
      return;
    }

    /* 3) فرق الفاكسات */
    if(view === 'fax-diff'){
      el.innerHTML = `
        <div class="card">
          <h3><i class="fas fa-receipt"></i> فرق الفاكسات</h3>
          ${this.uploadZone('zoneFaxDiff','الملف 1: فرق الفاكسات (الصنف، الاسم، الكمية، الكاتب، الخصم)', null, 'fa-receipt')}
          ${this.uploadZone('zoneFax','الملف 2: الفاكسات (الصنف، خصم الفاكسات، المورد)', null, 'fa-file-invoice')}
          <div class="action-row">
            <button class="btn-primary" id="btnRunFax"><i class="fas fa-handshake"></i> فرق الفاكسات</button>
            <button class="btn-ghost" id="btnClearFax"><i class="fas fa-trash-can"></i> حذف البيانات</button>
          </div>
          <div class="toolbar">
            <input type="text" id="searchFax" placeholder="🔍 بحث...">
            <select id="sortFax">
              <option value="desc">⬇️ الأكبر إلى الأصغر (الفرق)</option>
              <option value="asc">⬆️ الأصغر إلى الأكبر (الفرق)</option>
              <option value="all">📋 بدون ترتيب</option>
            </select>
            <button class="btn-primary" id="exportFax" style="width:auto"><i class="fas fa-file-excel"></i> تحميل Excel</button>
          </div>
          <div id="faxResults"></div>
        </div>`;

      this.wireUploadZones(['zoneFaxDiff','zoneFax']);
      let currentRows = [];
      let baseRows = [];

      const render = ()=>{
        const q = normalize(document.getElementById('searchFax').value);
        const sort = document.getElementById('sortFax').value;

        let rows = baseRows.filter(r =>
          !q || normalize(r.name || '').includes(q) || normalize(r.storeName || '').includes(q) ||
          normalize(r.supplier || '').includes(q) || normalize(r.writer || '').includes(q)
        );

        if(sort === 'desc') rows = [...rows].sort((a,b) => (b.diff ?? -Infinity) - (a.diff ?? -Infinity));
        else if(sort === 'asc') rows = [...rows].sort((a,b) => (a.diff ?? Infinity) - (b.diff ?? Infinity));

        const total = baseRows.length;
        const matched = baseRows.filter(r=>r.matched).length;
        const unmatched = total - matched;
        const positive = baseRows.filter(r=>r.diff != null && r.diff > 0).length;
        const negative = baseRows.filter(r=>r.diff != null && r.diff < 0).length;
        const zero = baseRows.filter(r=>r.diff === 0).length;

        if(!rows.length){
          document.getElementById('faxResults').innerHTML = `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>${baseRows.length ? 'لا توجد نتائج' : 'لا توجد بيانات — ارفع الملفين واضغط "فرق الفاكسات"'}</p>
            </div>`;
          return;
        }

        let tableRows = '';
        rows.forEach((r, i)=>{
          let diffVal = '_____', diffClass = '', diffIcon = 'fa-minus';
          if(r.diff != null){
            diffVal = r.diff;
            if(r.diff > 0){ diffClass = 'badge-match'; diffIcon = 'fa-arrow-trend-up'; }
            else if(r.diff < 0){ diffClass = 'badge-miss'; diffIcon = 'fa-arrow-trend-down'; }
            else { diffIcon = 'fa-equals'; }
          }

          tableRows += `<tr>
            <td>${i+1}</td><td>${r.name || ''}</td>
            <td style="color:var(--muted)">${r.storeName || '—'}</td>
            <td>${r.qty !== '' && r.qty != null ? r.qty : '—'}</td>
            <td>${r.writer || '—'}</td>
            <td>${r.discount ?? '—'}</td>
            <td>${r.faxDiscount != null ? r.faxDiscount : '_____'}</td>
            <td>${r.supplier || '_____'}</td>
            <td><span class="${diffClass}" style="min-width:80px;justify-content:center;display:inline-flex;align-items:center;gap:4px"><i class="fas ${diffIcon}"></i> ${diffVal}</span></td>
          </tr>`;
        });

        document.getElementById('faxResults').innerHTML = `
          <div class="stats-grid" style="margin:15px 0">
            <div class="stat-card"><i class="fas fa-list-ol"></i><div class="stat-num">${total}</div><div class="stat-label">إجمالي</div></div>
            <div class="stat-card"><i class="fas fa-circle-check" style="color:var(--success)"></i><div class="stat-num" style="color:var(--success)">${matched}</div><div class="stat-label">مطابق</div></div>
            <div class="stat-card"><i class="fas fa-circle-xmark" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${unmatched}</div><div class="stat-label">غير موجود</div></div>
            <div class="stat-card"><i class="fas fa-arrow-trend-up" style="color:var(--success)"></i><div class="stat-num" style="color:var(--success)">${positive}</div><div class="stat-label">موجب</div></div>
            <div class="stat-card"><i class="fas fa-arrow-trend-down" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${negative}</div><div class="stat-label">سالب</div></div>
            <div class="stat-card"><i class="fas fa-equals" style="color:var(--muted)"></i><div class="stat-num" style="color:var(--muted)">${zero}</div><div class="stat-label">صفر</div></div>
          </div>
          <div style="margin:10px 0;color:var(--muted);font-size:13px"><i class="fas fa-info-circle"></i> يُعرض ${rows.length} من ${total} صف</div>
          <div class="table-wrap" style="max-height:70vh;overflow:auto">
            <table>
              <thead><tr><th>#</th><th>الصنف</th><th>الاسم</th><th>الكمية</th><th>الكاتب</th><th>الخصم</th><th>خصم الفاكسات</th><th>المورد</th><th>الفرق</th></tr></thead>
              <tbody>${tableRows}</tbody>
            </table>
          </div>`;
      };

      document.getElementById('searchFax').oninput = render;
      document.getElementById('sortFax').onchange = render;

      document.getElementById('btnRunFax').onclick = async ()=>{
        const btn = document.getElementById('btnRunFax');
        const z1 = document.getElementById('zoneFaxDiff');
        const z2 = document.getElementById('zoneFax');
        if(!z1._file){ showToast('❌ الملف 1 غير صالح','error'); z1.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        if(!z2._file){ showToast('❌ الملف 2 غير صالح','error'); z2.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        setLoading(btn, true);
        try{
          const rows = await Matching.runFaxDiff(z1._file, z2._file);
          baseRows = [...rows];
          currentRows = [...rows].sort((a,b) => (b.diff ?? -Infinity) - (a.diff ?? -Infinity));
          App.branchesData[dataKey] = { ...(App.branchesData[dataKey]||{}), faxDiff: currentRows };
          App.saveData();
          render();
          showToast('تم حساب فرق الفاكسات','success');
        }catch(err){ showToast(err.message || 'حدث خطأ','error'); }
        setLoading(btn, false);
      };

      document.getElementById('btnClearFax').onclick = async ()=>{
        if(await confirmModal('حذف','حذف البيانات؟')){
          currentRows = []; baseRows = [];
          document.getElementById('faxResults').innerHTML = '';
          Excel.resetAllZones();
          if(App.branchesData[dataKey]) delete App.branchesData[dataKey].faxDiff;
          App.saveData();
          showToast('تم الحذف','success');
        }
      };

      document.getElementById('exportFax').onclick = ()=>{
        if(!baseRows.length){ showToast('لا توجد بيانات','error'); return; }
        const sort = document.getElementById('sortFax').value;
        let rows = [...baseRows];
        if(sort === 'desc') rows.sort((a,b) => (b.diff ?? -Infinity) - (a.diff ?? -Infinity));
        else if(sort === 'asc') rows.sort((a,b) => (a.diff ?? Infinity) - (b.diff ?? Infinity));

        const out = rows.map(r=>({
          'الصنف': r.name || '', 'الاسم': r.storeName || '',
          'الكمية': r.qty !== '' && r.qty != null ? r.qty : '',
          'الكاتب': r.writer || '', 'الخصم': r.discount ?? '',
          'خصم الفاكسات': r.faxDiscount != null ? r.faxDiscount : '_____',
          'المورد': r.supplier || '_____',
          'الفرق': r.diff != null ? r.diff : '_____'
        }));
        Excel.download(out, 'فرق_الفاكسات.xlsx', 'فرق الفاكسات');
      };
      return;
    }

        /* ============================================================
      4) مصفر الفاكسات — مع التكرار والمخازن والخصومات
      ============================================================ */
    if(view === 'fax-filter'){
      el.innerHTML = `
        <div class="card">
          <h3><i class="fas fa-filter"></i> مصفر الفاكسات</h3>
          ${this.uploadZone('zoneRec','الملف 1: الوارد (الكود، الكمية، الصنف) — اختياري', null, 'fa-arrow-down-to-line')}
          ${this.uploadZone('zoneOrd','الملف 2: الطلبات (الكود، الكمية، الصنف) — اختياري', null, 'fa-arrow-up-from-line')}
          ${this.uploadZone('zoneCol','الملف 3: المجمع (الكود، الاسم، س البيع، المرجح، ك مخزون، ك صافي بيع) — إلزامي', null, 'fa-layer-group')}
          ${this.uploadZone('zoneFaxF','الملف 4: الفاكسات (الصنف، الخصم، المورد) — إلزامي', null, 'fa-file-invoice')}
          <div class="action-row">
            <button class="btn-primary" id="btnRunFaxFilter"><i class="fas fa-calculator"></i> تشغيل المصفر</button>
            <button class="btn-ghost" id="btnClearFaxFilter"><i class="fas fa-trash-can"></i> حذف البيانات</button>
          </div>
          <div class="toolbar">
            <button class="btn-primary" id="btnShowAll" style="width:auto;background:#64748b"><i class="fas fa-list"></i> كل البيانات</button>
            <button class="btn-primary" id="btnShowMasfar" style="width:auto;background:#ef4444"><i class="fas fa-filter"></i> مصفر</button>
            <button class="btn-primary" id="btnShowHaysfar" style="width:auto;background:#f59e0b"><i class="fas fa-filter"></i> هيصفر</button>
            <button class="btn-primary" id="exportFaxFilter" style="width:auto"><i class="fas fa-file-excel"></i> تحميل Excel</button>
          </div>
          <div class="toolbar">
            <input type="text" id="searchFaxFilter" placeholder="🔍 بحث بالصنف أو الكود أو المورد...">
            <select id="filterRepeat">
              <option value="all">كل التكرارات</option>
              <option value="0">غير موجود في الفاكسات</option>
              <option value="1">مخزن واحد</option>
              <option value="2">مخزنين</option>
              <option value="3">3 مخازن</option>
              <option value="4+">4 مخازن فأكثر</option>
            </select>
          </div>
          <div id="faxFilterResults"></div>
        </div>`;

      this.wireUploadZones(['zoneRec','zoneOrd','zoneCol','zoneFaxF']);
      let data = { all: [], masfar: [], هيسفر: [] };
      let currentView = 'masfar';

      const cols = [
        { label:'الكود', key:'code' },
        { label:'الاسم', key:'name' },
        { label:'س البيع', key:'salePrice' },
        { label:'المرجح', key:'weighted' },
        { label:'ك مخزون', key:'stock' },
        { label:'ك الوارد', key:'recQty' },
        { label:'ك الطلبات', key:'ordQty' },
        { label:'إجمالي الكمية', key:'totalQty' },
        { label:'ك صافي بيع', key:'netSale' },
        { label:'المطلوب', key:'required' },
        { label:'المورد', key:'supplier' },
        { label:'خصم الفاكسات', key:'faxDiscount' },
        { label:'التكرار', key:'repeat' },
        { label:'المخازن', key:'stores' },
        { label:'الخصومات', key:'discounts' }
      ];

      const renderTable = (rows, title, color) => {
        if(!rows.length){
          return `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>${title} — لا توجد بيانات</p>
            </div>`;
        }

        let tableRows = '';
        rows.forEach((r, i)=>{
          tableRows += `<tr>
            <td>${i+1}</td>
            ${cols.map(c=>{
              let val = r[c.key];
              if(val === null || val === undefined) val = '_____';

              // عمود المطلوب - تلوين
              if(c.key === 'required' && val > 0){
                return `<td><span class="badge-miss"><i class="fas fa-arrow-trend-up"></i> ${val}</span></td>`;
              }
              if(c.key === 'required' && val === 0){
                return `<td><span class="badge-match"><i class="fas fa-equals"></i> 0</span></td>`;
              }

              // عمود التكرار - تلوين
              if(c.key === 'repeat'){
                let badge = '';
                if(val === 0) badge = '<span class="badge-miss"><i class="fas fa-circle-xmark"></i> 0</span>';
                else if(val === 1) badge = `<span class="badge-match">${val}</span>`;
                else badge = `<span class="badge-match" style="background:#dbeafe;color:#1e40af">${val}</span>`;
                return `<td>${badge}</td>`;
              }

              // عمود المخازن - يظهر بـ tooltip
              if(c.key === 'stores'){
                return `<td style="max-width:250px;white-space:normal;font-size:12px" title="${val}">${val}</td>`;
              }

              // عمود الخصومات
              if(c.key === 'discounts'){
                return `<td style="max-width:200px;white-space:normal;font-size:12px" title="${val}">${val}</td>`;
              }

              return `<td>${val}</td>`;
            }).join('')}
          </tr>`;
        });

        return `
          <h3 style="margin:15px 0;color:${color}"><i class="fas fa-table"></i> ${title} (${rows.length} صف)</h3>
          <div class="table-wrap" style="max-height:70vh;overflow:auto">
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  ${cols.map(c=>`<th>${c.label}</th>`).join('')}
                </tr>
              </thead>
              <tbody>${tableRows}</tbody>
            </table>
          </div>`;
      };

      const getFilteredRows = (rows) => {
        const q = normalize(document.getElementById('searchFaxFilter')?.value || '');
        const repeatFilter = document.getElementById('filterRepeat')?.value || 'all';

        let filtered = rows.filter(r => {
          // فلترة البحث
          if(q){
            const matchSearch =
              normalize(r.code || '').includes(q) ||
              normalize(r.name || '').includes(q) ||
              normalize(r.supplier || '').includes(q) ||
              normalize(r.stores || '').includes(q);
            if(!matchSearch) return false;
          }

          // فلترة التكرار
          if(repeatFilter !== 'all'){
            const repeat = r.repeat || 0;
            if(repeatFilter === '0' && repeat !== 0) return false;
            if(repeatFilter === '1' && repeat !== 1) return false;
            if(repeatFilter === '2' && repeat !== 2) return false;
            if(repeatFilter === '3' && repeat !== 3) return false;
            if(repeatFilter === '4+' && repeat < 4) return false;
          }

          return true;
        });

        return filtered;
      };

      const render = () => {
        if(!data.all.length && !data.masfar.length && !data.هيسفر.length){
          document.getElementById('faxFilterResults').innerHTML = `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>لا توجد بيانات — ارفع الملفات واضغط "تشغيل المصفر"</p>
            </div>`;
          return;
        }

        const total = data.all.length;
        const masfarCount = data.masfar.length;
        const haysfarCount = data.هيسفر.length;
        const withFax = data.all.filter(r=>r.matchedFax).length;
        const repeated = data.all.filter(r=>r.repeat > 1).length;

        let html = `
          <div class="stats-grid" style="margin:15px 0">
            <div class="stat-card"><i class="fas fa-list-ol"></i><div class="stat-num">${total}</div><div class="stat-label">إجمالي الأصناف</div></div>
            <div class="stat-card"><i class="fas fa-circle-check" style="color:var(--success)"></i><div class="stat-num" style="color:var(--success)">${withFax}</div><div class="stat-label">مرتبط بالفاكسات</div></div>
            <div class="stat-card"><i class="fas fa-copy" style="color:var(--primary)"></i><div class="stat-num" style="color:var(--primary)">${repeated}</div><div class="stat-label">في أكثر من مخزن</div></div>
            <div class="stat-card"><i class="fas fa-filter" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${masfarCount}</div><div class="stat-label">مصفر</div></div>
            <div class="stat-card"><i class="fas fa-filter" style="color:var(--accent)"></i><div class="stat-num" style="color:var(--accent)">${haysfarCount}</div><div class="stat-label">هيصفر</div></div>
          </div>`;

        // تطبيق الفلاتر
        if(currentView === 'all'){
          const filtered = getFilteredRows(data.all);
          html += renderTable(filtered, 'جميع البيانات', 'var(--primary)');
        } else if(currentView === 'masfar'){
          const filtered = getFilteredRows(data.masfar);
          html += renderTable(filtered, 'مصفر (المطلوب ≠ 0 + إجمالي ≤ 10)', 'var(--danger)');
        } else if(currentView === 'haysfar'){
          const filtered = getFilteredRows(data.هيسفر);
          html += renderTable(filtered, 'هيصفر (المطلوب > 5 + إجمالي 11–2000)', 'var(--accent)');
        }

        document.getElementById('faxFilterResults').innerHTML = html;

        // ربط أزرار التبديل
        ['btnShowAll','btnShowMasfar','btnShowHaysfar'].forEach(id=>{
          const btn = document.getElementById(id);
          if(btn){
            btn.onclick = () => {
              currentView = id === 'btnShowAll' ? 'all' : id === 'btnShowMasfar' ? 'masfar' : 'haysfar';
              render();
            };
          }
        });
      };

      render();

      // ربط البحث والفلتر
      document.getElementById('searchFaxFilter').oninput = render;
      document.getElementById('filterRepeat').onchange = render;

      document.getElementById('btnRunFaxFilter').onclick = async ()=>{
        const btn = document.getElementById('btnRunFaxFilter');
        const zRec = document.getElementById('zoneRec');
        const zOrd = document.getElementById('zoneOrd');
        const zCol = document.getElementById('zoneCol');
        const zFax = document.getElementById('zoneFaxF');

        if(!zCol._file){ showToast('❌ ملف المجمع إلزامي','error'); zCol.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        if(!zFax._file){ showToast('❌ ملف الفاكسات إلزامي','error'); zFax.scrollIntoView({behavior:'smooth',block:'center'}); return; }

        setLoading(btn, true);
        try{
          data = await Matching.runFaxFilter(
            zRec._file, zOrd._file, zCol._file, zFax._file
          );
          App.branchesData[dataKey] = { ...(App.branchesData[dataKey]||{}), faxFilter: data };
          App.saveData();
          currentView = 'masfar';
          render();
          showToast('تم حساب المصفر بنجاح','success');
        }catch(err){
          console.error('❌ خطأ:', err);
          showToast(err.message || 'حدث خطأ','error');
        }
        setLoading(btn, false);
      };

      document.getElementById('btnClearFaxFilter').onclick = async ()=>{
        if(await confirmModal('حذف','حذف البيانات؟')){
          data = { all: [], masfar: [], هيسفر: [] };
          document.getElementById('faxFilterResults').innerHTML = '';
          Excel.resetAllZones();
          if(App.branchesData[dataKey]) delete App.branchesData[dataKey].faxFilter;
          App.saveData();
          showToast('تم الحذف','success');
          render();
        }
      };

      document.getElementById('exportFaxFilter').onclick = ()=>{
        let exportRows = [];
        let title = 'مصفر الفاكسات';

        if(currentView === 'all'){ exportRows = getFilteredRows(data.all); title = 'جميع البيانات'; }
        else if(currentView === 'masfar'){ exportRows = getFilteredRows(data.masfar); title = 'مصفر الفاكسات'; }
        else if(currentView === 'haysfar'){ exportRows = getFilteredRows(data.هيسفر); title = 'هيصفر الفاكسات'; }

        if(!exportRows.length){ showToast('لا توجد بيانات للتحميل','error'); return; }

        const out = exportRows.map(r=>{
          const o = {};
          cols.forEach(c => { o[c.label] = r[c.key] ?? ''; });
          return o;
        });

        Excel.download(out, `${title}.xlsx`, title);
      };
      return;
}

    /* 5) المصفر العام */
    if(view === 'general-filter'){
      el.innerHTML = `
        <div class="card">
          <h3><i class="fas fa-layer-group"></i> المصفر العام</h3>
          ${this.uploadZone('zoneRecG','الملف 1: الوارد (الكود، الكمية، الصنف) — اختياري', null, 'fa-arrow-down-to-line')}
          ${this.uploadZone('zoneOrdG','الملف 2: الطلبات (الكود، الكمية، الصنف) — اختياري', null, 'fa-arrow-up-from-line')}
          ${this.uploadZone('zoneColG','الملف 3: المجمع (الكود، الاسم، س البيع، المرجح، ك مخزون، ك صافي بيع) — إلزامي', null, 'fa-layer-group')}
          <div class="action-row">
            <button class="btn-primary" id="btnRunGen"><i class="fas fa-calculator"></i> ربط الملفات</button>
            <button class="btn-ghost" id="btnClearGen"><i class="fas fa-trash-can"></i> حذف البيانات</button>
          </div>
          <div class="toolbar">
            <button class="btn-primary" id="btnGenAll" style="width:auto;background:#64748b"><i class="fas fa-list"></i> كل البيانات</button>
            <button class="btn-primary" id="btnGenMasfar" style="width:auto;background:#ef4444"><i class="fas fa-filter"></i> المصفر</button>
            <button class="btn-primary" id="btnGenHaysfar" style="width:auto;background:#f59e0b"><i class="fas fa-filter"></i> هيصفر</button>
            <button class="btn-primary" id="exportGen" style="width:auto"><i class="fas fa-file-excel"></i> تحميل Excel</button>
          </div>
          <div id="genResults"></div>
        </div>`;

      this.wireUploadZones(['zoneRecG','zoneOrdG','zoneColG']);
      let data = { all: [], masfar: [], هيسفر: [] };
      let currentView = 'all';

      const cols = [
        { label:'الكود', key:'code' }, { label:'الاسم', key:'name' },
        { label:'س البيع', key:'salePrice' }, { label:'المرجح', key:'weighted' },
        { label:'ك مخزون', key:'stock' }, { label:'ك الوارد', key:'recQty' },
        { label:'ك الطلبات', key:'ordQty' }, { label:'إجمالي الكمية', key:'totalQty' },
        { label:'ك صافي بيع', key:'netSale' }, { label:'المطلوب', key:'required' }
      ];

      const renderTable = (rows, title, color) => {
        if(!rows.length) return `<div style="text-align:center;padding:40px;color:var(--muted)"><i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i><p>${title} — لا توجد بيانات</p></div>`;

        let tableRows = '';
        rows.forEach((r, i)=>{
          tableRows += `<tr><td>${i+1}</td>${cols.map(c=>{
            let val = r[c.key];
            if(val === null || val === undefined) val = '';
            if(c.key === 'required' && val > 0) return `<td><span class="badge-miss"><i class="fas fa-arrow-trend-up"></i> ${val}</span></td>`;
            if(c.key === 'required' && val === 0) return `<td><span class="badge-match"><i class="fas fa-equals"></i> 0</span></td>`;
            return `<td>${val}</td>`;
          }).join('')}</tr>`;
        });

        return `<h3 style="margin:15px 0;color:${color}"><i class="fas fa-table"></i> ${title} (${rows.length} صف)</h3>
          <div class="table-wrap" style="max-height:70vh;overflow:auto">
            <table><thead><tr><th>#</th>${cols.map(c=>`<th>${c.label}</th>`).join('')}</tr></thead>
            <tbody>${tableRows}</tbody></table>
          </div>`;
      };

      const render = () => {
        if(!data.all.length && !data.masfar.length && !data.هيسفر.length){
          document.getElementById('genResults').innerHTML = `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>لا توجد بيانات — ارفع الملفات واضغط "ربط الملفات"</p>
            </div>`;
          return;
        }

        const total = data.all.length;
        const masfarCount = data.masfar.length;
        const haysfarCount = data.هيسفر.length;

        let html = `<div class="stats-grid" style="margin:15px 0">
          <div class="stat-card"><i class="fas fa-list-ol"></i><div class="stat-num">${total}</div><div class="stat-label">إجمالي</div></div>
          <div class="stat-card"><i class="fas fa-filter" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${masfarCount}</div><div class="stat-label">المصفر</div></div>
          <div class="stat-card"><i class="fas fa-filter" style="color:var(--accent)"></i><div class="stat-num" style="color:var(--accent)">${haysfarCount}</div><div class="stat-label">هيصفر</div></div>
        </div>`;

        if(currentView === 'all') html += renderTable(data.all, 'جميع البيانات', 'var(--primary)');
        else if(currentView === 'masfar') html += renderTable(data.masfar, 'المصفر (المطلوب ≠ 0 + إجمالي ≤ 10)', 'var(--danger)');
        else if(currentView === 'haysfar') html += renderTable(data.هيسفر, 'هيصفر (المطلوب > 5 + إجمالي 11–2000)', 'var(--accent)');

        document.getElementById('genResults').innerHTML = html;

        ['btnGenAll','btnGenMasfar','btnGenHaysfar'].forEach(id=>{
          const btn = document.getElementById(id);
          if(btn) btn.onclick = () => {
            currentView = id === 'btnGenAll' ? 'all' : id === 'btnGenMasfar' ? 'masfar' : 'haysfar';
            render();
          };
        });
      };

      render();

      document.getElementById('btnRunGen').onclick = async ()=>{
        const btn = document.getElementById('btnRunGen');
        const zRec = document.getElementById('zoneRecG');
        const zOrd = document.getElementById('zoneOrdG');
        const zCol = document.getElementById('zoneColG');

        if(!zCol._file){ showToast('❌ ملف المجمع إلزامي','error'); zCol.scrollIntoView({behavior:'smooth',block:'center'}); return; }

        setLoading(btn, true);
        try{
          data = await Matching.runGeneralFilter(zCol._file, zRec._file, zOrd._file);
          App.branchesData[dataKey] = { ...(App.branchesData[dataKey]||{}), generalFilter: data };
          App.saveData();
          currentView = 'all';
          render();
          showToast('تم ربط الملفات','success');
        }catch(err){ showToast(err.message || 'حدث خطأ','error'); }
        setLoading(btn, false);
      };

      document.getElementById('btnClearGen').onclick = async ()=>{
        if(await confirmModal('حذف','حذف البيانات؟')){
          data = { all: [], masfar: [], هيسفر: [] };
          document.getElementById('genResults').innerHTML = '';
          Excel.resetAllZones();
          if(App.branchesData[dataKey]) delete App.branchesData[dataKey].generalFilter;
          App.saveData();
          showToast('تم الحذف','success');
          render();
        }
      };

      document.getElementById('exportGen').onclick = ()=>{
        let exportRows = [], title = 'المصفر';
        if(currentView === 'all'){ exportRows = data.all; title = 'جميع البيانات'; }
        else if(currentView === 'masfar'){ exportRows = data.masfar; title = 'المصفر'; }
        else if(currentView === 'haysfar'){ exportRows = data.haysفر; title = 'هيصفر'; }

        if(!exportRows.length){ showToast('لا توجد بيانات','error'); return; }
        const out = exportRows.map(r=>{ const o = {}; cols.forEach(c => { o[c.label] = r[c.key] ?? ''; }); return o; });
        Excel.download(out, `${title}.xlsx`, title);
      };
      return;
    }

    /* 6) أوفر استوك */
    if(view === 'over-stock'){
      el.innerHTML = `
        <div class="card">
          <h3><i class="fas fa-boxes-packing"></i> أوفر استوك</h3>
          ${this.uploadZone('zoneRecOS','الملف 1: الوارد (الكود، الكمية، الصنف) — اختياري', null, 'fa-arrow-down-to-line')}
          ${this.uploadZone('zoneColOS','الملف 2: المجمع (الكود، الاسم، س البيع، المرجح، ك مخزون، ك صافي بيع) — إلزامي', null, 'fa-layer-group')}
          ${this.uploadZone('zoneBalOS','الملف 3: الأرصدة (الكود، المرجح، الصنف) — اختياري', null, 'fa-coins')}
          ${this.uploadZone('zoneOS','الملف 4: أوفر استوك (الكود، الاسم، الصنف، الكاتب، الطلبات، الخصم) — إلزامي', null, 'fa-boxes-packing')}
          <div class="action-row">
            <button class="btn-primary" id="btnRunOverStock"><i class="fas fa-calculator"></i> أوفر استوك</button>
            <button class="btn-ghost" id="btnClearOS"><i class="fas fa-trash-can"></i> حذف البيانات</button>
          </div>
          <div class="toolbar">
            <input type="text" id="searchOS" placeholder="🔍 بحث بالكود أو الاسم أو الصنف...">
            <select id="filterAnalysisOS">
              <option value="all">كل التحليلات</option>
              <option value="استوك عادي">استوك عادي</option>
              <option value="استوك للصنف">استوك للصنف</option>
              <option value="أوفر استوك">أوفر استوك</option>
              <option value="مراجعة">مراجعة</option>
            </select>
            <button class="btn-primary" id="exportOS" style="width:auto"><i class="fas fa-file-excel"></i> تحميل Excel</button>
          </div>
          <div id="osResults"></div>
        </div>`;

      this.wireUploadZones(['zoneRecOS','zoneColOS','zoneBalOS','zoneOS']);
      let currentRows = [];

      const render = ()=>{
        const q = normalize(document.getElementById('searchOS').value);
        const f = document.getElementById('filterAnalysisOS').value;

        let rows = currentRows.filter(r =>
          !q || normalize(r.code || '').includes(q) || normalize(r.name || '').includes(q) ||
          normalize(r.name2 || '').includes(q) || normalize(r.writer || '').includes(q)
        );

        if(f !== 'all') rows = rows.filter(r => r.analysis === f);

        const total = currentRows.length;
        const normalCount = currentRows.filter(r=>r.analysis === 'استوك عادي').length;
        const forItemCount = currentRows.filter(r=>r.analysis === 'استوك للصنف').length;
        const overCount = currentRows.filter(r=>r.analysis === 'أوفر استوك').length;
        const reviewCount = currentRows.filter(r=>r.analysis === 'مراجعة').length;

        if(!rows.length){
          document.getElementById('osResults').innerHTML = `
            <div style="text-align:center;padding:40px;color:var(--muted)">
              <i class="fas fa-inbox" style="font-size:50px;margin-bottom:15px"></i>
              <p>${currentRows.length ? 'لا توجد نتائج مطابقة للبحث' : 'لا توجد بيانات — ارفع الملفات واضغط "أوفر استوك"'}</p>
            </div>`;
          return;
        }

        let tableRows = '';
        rows.forEach((r, i)=>{
          let diffClass = '';
          if(r.diff > 0) diffClass = 'badge-match';
          else if(r.diff < 0) diffClass = 'badge-miss';

          let analysisClass = '';
          if(r.analysis === 'أوفر استوك') analysisClass = 'badge-miss';
          else if(r.analysis === 'استوك للصنف') analysisClass = 'badge-match';

          tableRows += `<tr>
            <td>${i+1}</td><td>${r.code || ''}</td><td>${r.name || ''}</td>
            <td>${r.name2 || ''}</td><td>${r.writer || ''}</td>
            <td>${r.stock ?? 0}</td><td>${r.netSale ?? 0}</td>
            <td>${r.orders ?? 0}</td><td>${r.required ?? 0}</td>
            <td><b style="color:${r.overStock > 0 ? '#166534' : '#64748b'}">${r.overStock}</b></td>
            <td>${r.discount ?? 0}</td><td>${r.lastDiscount ?? 0}</td>
            <td>${r.weighted ?? 0}</td>
            <td><span class="${diffClass}">${r.diff}</span></td>
            <td><span class="${analysisClass}" style="font-size:11px">${r.analysis}</span></td>
          </tr>`;
        });

        document.getElementById('osResults').innerHTML = `
          <div class="stats-grid" style="margin:15px 0">
            <div class="stat-card"><i class="fas fa-list-ol"></i><div class="stat-num">${total}</div><div class="stat-label">إجمالي</div></div>
            <div class="stat-card"><i class="fas fa-box" style="color:var(--muted)"></i><div class="stat-num" style="color:var(--muted)">${normalCount}</div><div class="stat-label">استوك عادي</div></div>
            <div class="stat-card"><i class="fas fa-star" style="color:var(--success)"></i><div class="stat-num" style="color:var(--success)">${forItemCount}</div><div class="stat-label">استوك للصنف</div></div>
            <div class="stat-card"><i class="fas fa-boxes-packing" style="color:var(--danger)"></i><div class="stat-num" style="color:var(--danger)">${overCount}</div><div class="stat-label">أوفر استوك</div></div>
            <div class="stat-card"><i class="fas fa-triangle-exclamation" style="color:var(--accent)"></i><div class="stat-num" style="color:var(--accent)">${reviewCount}</div><div class="stat-label">مراجعة</div></div>
          </div>
          <div style="margin:10px 0;color:var(--muted);font-size:13px"><i class="fas fa-info-circle"></i> يُعرض ${rows.length} من ${total} صف</div>
          <div class="table-wrap" style="max-height:70vh;overflow:auto">
            <table>
              <thead>
                <tr>
                  <th>#</th><th>الكود</th><th>الاسم</th><th>الصنف</th><th>الكاتب</th>
                  <th>ك مخزون</th><th>ك البيع</th><th>الطلبات</th><th>المطلوب</th>
                  <th>أوفر استوك</th><th>الخصم</th><th>آخر خصم</th><th>المرجح</th>
                  <th>الفرق</th><th>التحليل</th>
                </tr>
              </thead>
              <tbody>${tableRows}</tbody>
            </table>
          </div>`;
      };

      document.getElementById('searchOS').oninput = render;
      document.getElementById('filterAnalysisOS').onchange = render;

      document.getElementById('btnRunOverStock').onclick = async ()=>{
        const btn = document.getElementById('btnRunOverStock');
        const zRec = document.getElementById('zoneRecOS');
        const zCol = document.getElementById('zoneColOS');
        const zBal = document.getElementById('zoneBalOS');
        const zOS = document.getElementById('zoneOS');

        if(!zCol._file){ showToast('❌ ملف المجمع إلزامي','error'); zCol.scrollIntoView({behavior:'smooth',block:'center'}); return; }
        if(!zOS._file){ showToast('❌ ملف أوفر استوك إلزامي','error'); zOS.scrollIntoView({behavior:'smooth',block:'center'}); return; }

        setLoading(btn, true);
        try{
          currentRows = await Matching.runOverStock(zRec._file, zCol._file, zBal._file, zOS._file);
          App.branchesData[dataKey] = { ...(App.branchesData[dataKey]||{}), overStock: currentRows };
          App.saveData();
          render();
          showToast('تم حساب أوفر استوك','success');
        }catch(err){ showToast(err.message || 'حدث خطأ','error'); }
        setLoading(btn, false);
      };

      document.getElementById('btnClearOS').onclick = async ()=>{
        if(await confirmModal('حذف','حذف بيانات أوفر استوك؟')){
          currentRows = [];
          document.getElementById('osResults').innerHTML = '';
          Excel.resetAllZones();
          if(App.branchesData[dataKey]) delete App.branchesData[dataKey].overStock;
          App.saveData();
          showToast('تم الحذف','success');
        }
      };

      document.getElementById('exportOS').onclick = ()=>{
        if(!currentRows.length){ showToast('لا توجد بيانات','error'); return; }
        const out = currentRows.map(r=>({
          'الكود': r.code, 'الاسم': r.name, 'الصنف': r.name2, 'الكاتب': r.writer,
          'ك مخزون': r.stock, 'ك البيع': r.netSale, 'الطلبات': r.orders, 'المطلوب': r.required,
          'أوفر استوك': r.overStock, 'الخصم': r.discount, 'آخر خصم': r.lastDiscount,
          'المرجح': r.weighted, 'الفرق': r.diff, 'التحليل': r.analysis
        }));
        Excel.download(out, 'أوفر_استوك.xlsx', 'أوفر استوك');
      };
      return;
    }
  },

  uploadZone(id, label, requiredColumns = null, icon = 'fa-cloud-arrow-up'){
    const dataReq = requiredColumns ? `data-required="${requiredColumns.join(',')}"` : '';
    return `<div class="upload-zone" id="${id}" ${dataReq}>
      <i class="fas ${icon}"></i>
      <p>${label}</p>
      <div class="file-name"></div>
      <div class="file-info"></div>
      <div class="file-error"></div>
      <input type="file" accept=".xlsx,.xls,.csv">
    </div>`;
  },

  wireUploadZones(ids){
    ids.forEach(id => {
      const zone = document.getElementById(id);
      if(!zone){ console.error('❌ صندوق الرفع غير موجود:', id); return; }
      const req = zone.dataset.required ? zone.dataset.required.split(',') : null;
      Excel.setupUpload(id, req);
    });
  },

  buildTable(rows, cols, keys){
    if(!rows.length) return '';
    return `<div class="table-wrap"><table>
      <thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r=>`<tr>${keys.map(k=>`<td>${r[k] ?? ''}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`;
  },

  tableSizeControls(tableWrapId){
    return `
      <div style="display:flex;gap:6px;justify-content:flex-end;margin-bottom:8px;flex-wrap:wrap">
        <button class="btn-ghost" style="width:auto;padding:6px 12px;margin:0" onclick="Dashboard.resizeTable('${tableWrapId}','small')" title="جدول صغير">
          <i class="fas fa-compress"></i> صغير
        </button>
        <button class="btn-ghost" style="width:auto;padding:6px 12px;margin:0" onclick="Dashboard.resizeTable('${tableWrapId}','medium')" title="جدول متوسط">
          <i class="fas fa-arrows-alt-h"></i> متوسط
        </button>
        <button class="btn-ghost" style="width:auto;padding:6px 12px;margin:0" onclick="Dashboard.resizeTable('${tableWrapId}','large')" title="جدول كبير">
          <i class="fas fa-expand"></i> كبير
        </button>
        <button class="btn-ghost" style="width:auto;padding:6px 12px;margin:0" onclick="Dashboard.resizeTable('${tableWrapId}','full')" title="ملء الشاشة">
          <i class="fas fa-expand-arrows-alt"></i> ملء الشاشة
        </button>
      </div>
    `;
  },

  resizeTable(tableWrapId, size){
    const wrap = document.getElementById(tableWrapId);
    if(!wrap) return;
    const sizes = {
      small:  { height: '300px', minWidth: '900px'  },
      medium: { height: '500px', minWidth: '1100px' },
      large:  { height: '70vh',  minWidth: '1400px' },
      full:   { height: '90vh',  minWidth: '100%'   }
    };
    const s = sizes[size] || sizes.medium;
    wrap.style.maxHeight = s.height;
    const table = wrap.querySelector('table');
    if(table) table.style.minWidth = s.minWidth;
    const labels = { small:'صغير', medium:'متوسط', large:'كبير', full:'ملء الشاشة' };
    showToast(`حجم الجدول: ${labels[size] || 'متوسط'}`, 'info');
  },

  renderAdmin(){
    const el = document.getElementById('page-admin');
    el.classList.remove('hidden');
    el.classList.add('active');
    document.getElementById('page-home').classList.add('hidden');
    document.getElementById('page-branch').classList.add('hidden');

    el.querySelectorAll('.tab-btn').forEach(btn=>{
      btn.onclick = ()=>{
        el.querySelectorAll('.tab-btn').forEach(x=>x.classList.remove('active'));
        btn.classList.add('active');
        el.querySelectorAll('.tab-panel').forEach(p=>p.classList.add('hidden'));
        document.getElementById(btn.dataset.tab).classList.remove('hidden');
      };
    });
    this.renderAdminHome();
    this.renderAdminUsers();
    this.renderAdminPerms();
  },

  renderAdminHome(){
    const el = document.getElementById('admin-home');
    const branchesCount = Object.keys(BRANCHES).length;
    let totalItems = 0, totalWhole = 0, totalFax = 0, totalFilter = 0, totalOS = 0;
    Object.values(App.branchesData).forEach(b=>{
      totalItems += (b.itemsMatch||[]).length;
      totalWhole += (b.wholeDiff||[]).length;
      totalFax += (b.faxDiff||[]).length;
      totalFilter += ((b.faxFilter?.masfar||[]).length + (b.faxFilter?.هيسفر||[]).length);
      totalOS += (b.overStock||[]).length;
    });
    el.innerHTML = `
      <div class="stats-grid">
        <div class="stat-card"><i class="fas fa-store"></i><div class="stat-num">${branchesCount}</div><div class="stat-label">عدد الفروع</div></div>
        <div class="stat-card"><i class="fas fa-table-cells-large"></i><div class="stat-num">${totalItems}</div><div class="stat-label">مطابقات الأصناف</div></div>
        <div class="stat-card"><i class="fas fa-people-arrows"></i><div class="stat-num">${totalWhole}</div><div class="stat-label">فرق الجملة</div></div>
        <div class="stat-card"><i class="fas fa-receipt"></i><div class="stat-num">${totalFax}</div><div class="stat-label">فرق الفاكسات</div></div>
        <div class="stat-card"><i class="fas fa-filter"></i><div class="stat-num">${totalFilter}</div><div class="stat-label">مصفر الفاكسات</div></div>
        <div class="stat-card"><i class="fas fa-boxes-packing"></i><div class="stat-num">${totalOS}</div><div class="stat-label">أوفر استوك</div></div>
      </div>
      <div class="card"><h3><i class="fas fa-chart-line"></i> تحليل بياني لكل الفروع</h3>
        <div class="table-wrap" style="margin-top:15px"><table>
          <thead><tr><th>الفرع</th><th>مطابقة</th><th>فرق الجملة</th><th>فرق الفاكسات</th><th>مصفر</th><th>أوفر استوك</th></tr></thead>
          <tbody>
          ${Object.entries(BRANCHES).map(([k,v])=>{
            const b = App.branchesData[k]||{};
            return `<tr><td><i class="fas ${v.icon}"></i> ${v.name}</td><td>${(b.itemsMatch||[]).length}</td><td>${(b.wholeDiff||[]).length}</td><td>${(b.faxDiff||[]).length}</td><td>${((b.faxFilter?.masfar||[]).length + (b.faxFilter?.هيسفر||[]).length)}</td><td>${(b.overStock||[]).length}</td></tr>`;
          }).join('')}
          </tbody>
        </table></div>
      </div>`;
  },

  renderAdminUsers(){
    const el = document.getElementById('admin-users');
    if(!App.users || typeof App.users !== 'object') App.users = {};
    const usersArr = Object.values(App.users).filter(u => u && u.user);
    el.innerHTML = `
      <div class="card"><h3><i class="fas fa-user-plus"></i> إنشاء مستخدم جديد</h3>
        <div class="input-group"><i class="fas fa-store"></i><input id="newUserBranch" placeholder="الفرع (smart / san / taiba / dawa / rawi / admin)"></div>
        <div class="input-group"><i class="fas fa-user"></i><input id="newUserName" placeholder="اسم المستخدم"></div>
        <div class="input-group"><i class="fas fa-key"></i><input id="newUserPass" placeholder="الرقم السري"></div>
        <button class="btn-primary" id="btnCreateUser"><i class="fas fa-plus"></i> إنشاء</button>
        <div style="margin-top:20px">
          <h3 style="margin-bottom:10px"><i class="fas fa-users"></i> المستخدمون (${usersArr.length})</h3>
          <div class="table-wrap"><table>
            <thead><tr><th>المستخدم</th><th>الفرع</th><th>الدور</th><th>الحالة</th><th>إجراءات</th></tr></thead>
            <tbody>
            ${usersArr.length ? usersArr.map(u=>`<tr>
              <td><i class="fas fa-user"></i> ${u.user}</td><td>${u.branch || '—'}</td>
              <td>${u.role === 'developer' ? 'مطور' : u.role === 'admin' ? 'أدمن' : 'فرع'}</td>
              <td>${u.disabled ? '<i class="fas fa-circle-pause"></i> موقوف' : '<i class="fas fa-circle-check"></i> نشط'}</td>
              <td>${u.role === 'developer' ? '—' : `<div class="row-actions">
                <button class="${u.disabled ? 'btn-ok' : 'btn-warn'}" onclick="Dashboard.toggleUser('${u.user}')"><i class="fas ${u.disabled?'fa-check':'fa-pause'}"></i>${u.disabled?'تفعيل':'إيقاف'}</button>
                <button class="btn-del" onclick="Dashboard.deleteUser('${u.user}')"><i class="fas fa-trash"></i> حذف</button>
              </div>`}</td>
            </tr>`).join('') : `<tr><td colspan="5" style="text-align:center">لا يوجد مستخدمون</td></tr>`}
            </tbody>
          </table></div>
        </div>
      </div>`;

    document.getElementById('btnCreateUser').onclick = ()=>{
      const branch = document.getElementById('newUserBranch').value.trim().toLowerCase();
      const name = document.getElementById('newUserName').value.trim();
      const pass = document.getElementById('newUserPass').value.trim();
      if(!branch || !name || !pass){ showToast('يرجى ملء كل الحقول','error'); return; }
      const validBranches = ['smart','san','taiba','dawa','rawi','admin'];
      if(!validBranches.includes(branch)){ showToast('اسم الفرع غير صحيح','error'); return; }
      if(Object.values(App.users).some(u => u && typeof u.user === 'string' && u.user.toLowerCase() === name.toLowerCase())){
        showToast('اسم المستخدم موجود بالفعل','error'); return;
      }
      const role = branch === 'admin' ? 'admin' : 'branch';
      App.users[name] = { user: name, pass, role, branch };
      App.saveUsers();
      showToast('تم إنشاء المستخدم بنجاح','success');
      document.getElementById('newUserBranch').value = '';
      document.getElementById('newUserName').value = '';
      document.getElementById('newUserPass').value = '';
      this.renderAdminUsers();
      this.renderAdminPerms();
    };
  },

  toggleUser(name){
    if(!App.users || typeof App.users !== 'object') App.users = {};
    const u = App.users[name];
    if(!u) return;
    if(u.role === 'developer'){ showToast('لا يمكن إيقاف حساب المطور','error'); return; }
    u.disabled = !u.disabled;
    App.saveUsers();
    this.renderAdminUsers();
    this.renderAdminPerms();
    showToast('تم تحديث حالة المستخدم','success');
  },

  async deleteUser(name){
    if(!App.users || typeof App.users !== 'object') App.users = {};
    const u = App.users[name];
    if(!u) return;
    if(u.role === 'developer'){ showToast('لا يمكن حذف المطور','error'); return; }
    if(App.currentUser && App.currentUser.user === name){ showToast('لا يمكنك حذف حسابك','error'); return; }
    const ok = await confirmModal('حذف المستخدم', `حذف "${name}" نهائياً؟`);
    if(!ok) return;
    delete App.users[name];
    App.saveUsers();
    this.renderAdminUsers();
    this.renderAdminPerms();
    showToast(`تم حذف "${name}"`,'success');
  },

  renderAdminPerms(){
    const el = document.getElementById('admin-perms');
    if(!App.users || typeof App.users !== 'object') App.users = {};
    const usersArr = Object.values(App.users).filter(u => u && u.user);
    el.innerHTML = `
      <div class="card"><h3><i class="fas fa-user-lock"></i> الصلاحيات</h3>
        <div class="table-wrap"><table>
          <thead><tr><th>المستخدم</th><th>الفرع</th><th>الحالة</th><th>إجراءات</th></tr></thead>
          <tbody>
          ${usersArr.length ? usersArr.map(u=>`<tr>
            <td><i class="fas fa-user"></i> ${u.user}</td><td>${u.branch || '—'}</td>
            <td>${u.disabled ? '<i class="fas fa-circle-pause"></i> موقوف' : '<i class="fas fa-circle-check"></i> نشط'}</td>
            <td>${u.role === 'developer' ? '—' : `<div class="row-actions">
              <button class="${u.disabled ? 'btn-ok' : 'btn-warn'}" onclick="Dashboard.toggleUser('${u.user}')"><i class="fas ${u.disabled?'fa-check':'fa-pause'}"></i>${u.disabled?'تفعيل':'إيقاف'}</button>
              <button class="btn-del" onclick="Dashboard.deleteUser('${u.user}')"><i class="fas fa-trash"></i> حذف</button>
            </div>`}</td>
          </tr>`).join('') : `<tr><td colspan="4" style="text-align:center">لا يوجد مستخدمون</td></tr>`}
          </tbody>
        </table></div>
      </div>`;
  },

  renderDevHome(){
    const el = document.getElementById('page-branch');
    el.classList.remove('hidden');
    el.classList.add('active');
    document.getElementById('page-home').classList.add('hidden');
    document.getElementById('page-admin').classList.add('hidden');

    el.innerHTML = `
      <div class="branch-header">
        <h1><i class="fas fa-code"></i> لوحة المطور</h1>
        <p><i class="fas fa-user-tie"></i> Developer M.SH — متابعة كاملة لكل الفروع</p>
      </div>
      <div class="stats-grid">
        ${Object.entries(BRANCHES).map(([k,v])=>{
          const b = App.branchesData[k]||{};
          return `<div class="stat-card">
            <i class="fas ${v.icon}"></i>
            <div class="stat-num" style="font-size:16px">${v.name}</div>
            <div class="stat-label">
              <i class="fas fa-table-cells-large"></i> ${(b.itemsMatch||[]).length} |
              <i class="fas fa-people-arrows"></i> ${(b.wholeDiff||[]).length} |
              <i class="fas fa-receipt"></i> ${(b.faxDiff||[]).length} |
              <i class="fas fa-boxes-packing"></i> ${(b.overStock||[]).length}
            </div>
          </div>`;
        }).join('')}
      </div>
      <div class="card"><h3><i class="fas fa-database"></i> متابعة البيانات لكل فرع</h3>
        <div class="table-wrap"><table>
          <thead><tr><th>الفرع</th><th>مطابقة</th><th>فرق الجملة</th><th>فرق الفاكسات</th><th>مصفر</th><th>أوفر استوك</th></tr></thead>
          <tbody>
          ${Object.entries(BRANCHES).map(([k,v])=>{
            const b = App.branchesData[k]||{};
            return `<tr><td><i class="fas ${v.icon}"></i> ${v.name}</td><td>${(b.itemsMatch||[]).length}</td><td>${(b.wholeDiff||[]).length}</td><td>${(b.faxDiff||[]).length}</td><td>${((b.faxFilter?.masfar||[]).length + (b.faxFilter?.هيسفر||[]).length)}</td><td>${(b.overStock||[]).length}</td></tr>`;
          }).join('')}
          </tbody>
        </table></div>
      </div>
      <div style="text-align:center;margin-top:30px;color:var(--muted);font-weight:600"><i class="fas fa-code"></i> Developer M.SH</div>`;
  }
};
