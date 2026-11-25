const Excel = {
  async readFile(file){
    return new Promise((resolve, reject)=>{
      const reader = new FileReader();
      reader.onload = (e)=>{
        try{
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, {type:'array'});
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, {defval:''});
          resolve(rows);
        }catch(err){ reject('تعذر قراءة الملف: ' + err.message); }
      };
      reader.onerror = ()=> reject('تعذر قراءة الملف');
      reader.readAsArrayBuffer(file);
    });
  },

  /* ============================================================
     readFileAsArray — يقرأ الملف مع ضمان ترتيب الأعمدة
     يعيد: { headers: [أسماء الأعمدة بالترتيب], rows: [صفوف ككائنات] }
     ============================================================ */
  async readFileAsArray(file){
    return new Promise((resolve, reject)=>{
      const reader = new FileReader();
      reader.onload = (e)=>{
        try{
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, {type:'array'});
          const ws = wb.Sheets[wb.SheetNames[0]];
          // header: 1 → يعيد مصفوفة من المصفوفات (الصف الأول = رأس الأعمدة)
          const rawRows = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
          if(!rawRows.length){ resolve({ headers: [], rows: [] }); return; }
          const headers = rawRows[0].map(h => String(h).trim());
          const rows = rawRows.slice(1).map(arr => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = arr[i] !== undefined ? arr[i] : ''; });
            return obj;
          });
          resolve({ headers, rows });
        }catch(err){ reject('تعذر قراءة الملف: ' + err.message); }
      };
      reader.onerror = ()=> reject('تعذر قراءة الملف');
      reader.readAsArrayBuffer(file);
    });
  },

  download(rows, filename, headerTitle){
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      [headerTitle],
      ['Developer M.SH'],
      []
    ]);
    XLSX.utils.sheet_add_json(ws, rows, {origin:'A4'});
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const colCount = Math.max(headers.length, 8);
    ws['!merges'] = [
      { s:{r:0,c:0}, e:{r:0,c:colCount-1} },
      { s:{r:1,c:0}, e:{r:1,c:colCount-1} }
    ];
    if(ws['A1']) ws['A1'].s = {
      font: { bold: true, sz: 20, color: { rgb: 'FFFFFF' }, name: 'Calibri' },
      fill: { fgColor: { rgb: '0F766E' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    if(ws['A2']) ws['A2'].s = {
      font: { bold: true, italic: true, sz: 14, color: { rgb: 'F59E0B' }, name: 'Calibri' },
      fill: { fgColor: { rgb: '0F172A' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    ws['!rows'] = [{ hpt: 35 }, { hpt: 28 }, { hpt: 12 }];
    headers.forEach((_, idx)=>{
      const cell = XLSX.utils.encode_cell({ r:3, c:idx });
      if(ws[cell]) ws[cell].s = {
        font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' }, name: 'Calibri' },
        fill: { fgColor: { rgb: '14B8A6' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: {
          top: { style: 'thin', color: { rgb: '0F766E' } },
          bottom: { style: 'thin', color: { rgb: '0F766E' } },
          left: { style: 'thin', color: { rgb: 'FFFFFF' } },
          right: { style: 'thin', color: { rgb: 'FFFFFF' } }
        }
      };
    });
    for(let i = 4; i < 4 + rows.length; i++){
      headers.forEach((_, idx)=>{
        const cell = XLSX.utils.encode_cell({ r: i, c: idx });
        if(ws[cell]) ws[cell].s = {
          font: { sz: 11, name: 'Calibri' },
          alignment: { horizontal: 'right', vertical: 'center' },
          border: { bottom: { style: 'hair', color: { rgb: 'E2E8F0' } } }
        };
      });
    }
    ws['!cols'] = headers.map(h => ({ wch: Math.max(14, Math.min(40, h.length + 6)) }));
    XLSX.utils.book_append_sheet(wb, ws, 'Report');
    XLSX.writeFile(wb, filename);
    showToast('تم تحميل الملف بنجاح','success');
  },

  setupUpload(zoneId, requiredColumns = null){
    const zone = document.getElementById(zoneId);
    if(!zone){ console.error('❌ صندوق الرفع غير موجود:', zoneId); return null; }
    const input = zone.querySelector('input[type=file]');
    if(!input){ console.error('❌ لا يوجد input[type=file] داخل:', zoneId); return null; }

    if(zone._clickHandler) zone.removeEventListener('click', zone._clickHandler);
    if(input._changeHandler) input.removeEventListener('change', input._changeHandler);

    zone._clickHandler = (e)=>{
      if(e.target.closest('.zone-clear')) return;
      if(e.target === input) return;
      input.click();
    };
    input._changeHandler = (e)=>{
      const file = e.target.files && e.target.files[0];
      if(file) this._setFile(zone, file, requiredColumns);
    };

    zone.ondragover = (e)=>{ e.preventDefault(); zone.classList.add('dragover'); };
    zone.ondragleave = ()=> zone.classList.remove('dragover');
    zone.ondrop = (e)=>{
      e.preventDefault();
      zone.classList.remove('dragover');
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if(file) this._setFile(zone, file, requiredColumns);
    };

    zone.addEventListener('click', zone._clickHandler);
    input.addEventListener('change', input._changeHandler);
    zone._requiredColumns = requiredColumns;
    console.log('✅ تم ربط صندوق الرفع:', zoneId);
    return zone;
  },

  async _setFile(zone, file, requiredColumns){
    const errEl = zone.querySelector('.file-error');
    const nameEl = zone.querySelector('.file-name');
    const infoEl = zone.querySelector('.file-info');
    const iconEl = zone.querySelector(':scope > i');

    if(errEl) errEl.innerHTML = '';
    if(nameEl) nameEl.innerHTML = '';
    if(infoEl) infoEl.innerHTML = '';
    if(iconEl) iconEl.className = 'fas fa-cloud-arrow-up';
    zone.dataset.file = '';
    zone._file = null;
    zone.classList.remove('zone-valid','zone-invalid');

    if(!/\.(xlsx|xls|csv)$/i.test(file.name)){
      this._showZoneError(zone, '<i class="fas fa-circle-xmark"></i> نوع الملف غير مدعوم (يجب أن يكون <b>xlsx</b> أو <b>xls</b>)');
      showToast('نوع الملف غير مدعوم','error');
      return;
    }
    if(file.size > 10 * 1024 * 1024){
      this._showZoneError(zone, '<i class="fas fa-weight-hanging"></i> حجم الملف كبير جداً (الحد الأقصى 10 ميجا)');
      showToast('حجم الملف كبير جداً','error');
      return;
    }

    try{
      const rows = await this.readFile(file);
      if(!rows || rows.length === 0){
        this._showZoneError(zone, '<i class="fas fa-folder-open"></i> الملف فارغ ولا يحتوي على أي صفوف بيانات');
        return;
      }
      const columns = Object.keys(rows[0]);
      if(requiredColumns && requiredColumns.length){
        const normCol = (c)=> normalize(String(c));
        const normalized = columns.map(c => normCol(c));
        const missing = requiredColumns.filter(req => {
          const nreq = normalize(req);
          return !normalized.some(c => c === nreq || c.includes(nreq) || nreq.includes(c));
        });
        if(missing.length){
          const msg = `
            <div><i class="fas fa-circle-exclamation"></i> <b>الملف لا يحتوي على الأعمدة المطلوبة:</b></div>
            <div style="margin-top:4px">${missing.map(m=>`<span class="chip chip-danger"><i class="fas fa-xmark"></i> ${m}</span>`).join(' ')}</div>
            <div style="margin-top:8px"><i class="fas fa-table-columns"></i> <b>الأعمدة المتاحة:</b></div>
            <div style="margin-top:4px">${columns.map(c=>`<span class="chip chip-info"><i class="fas fa-check"></i> ${c}</span>`).join(' ')}</div>
          `;
          this._showZoneError(zone, msg);
          showToast('الملف لا يحتوي على الأعمدة المطلوبة','error');
          return;
        }
      }
      const emptyRows = rows.filter(r => Object.values(r).every(v => !v || String(v).trim() === ''));
      if(emptyRows.length === rows.length){
        this._showZoneError(zone, '<i class="fas fa-ban"></i> جميع الصفوف فارغة');
        return;
      }

      zone.classList.add('zone-valid');
      zone.dataset.file = file.name;
      zone._file = file;
      if(iconEl) iconEl.className = 'fas fa-file-circle-check';
      if(nameEl) nameEl.innerHTML = `<i class="fas fa-file-excel"></i> ${file.name}`;
      if(infoEl) infoEl.innerHTML = `<i class="fas fa-list-ol"></i> ${rows.length} صف | <i class="fas fa-table-columns"></i> ${columns.length} عمود`;

      if(!zone.querySelector('.zone-clear')){
        const btn = document.createElement('button');
        btn.className = 'zone-clear';
        btn.innerHTML = '<i class="fas fa-xmark"></i>';
        btn.title = 'حذف الملف';
        btn.onclick = (e)=>{
          e.stopPropagation();
          zone._file = null;
          zone.dataset.file = '';
          zone.classList.remove('zone-valid','zone-invalid');
          if(iconEl) iconEl.className = 'fas fa-cloud-arrow-up';
          if(nameEl) nameEl.innerHTML = '';
          if(errEl) errEl.innerHTML = '';
          if(infoEl) infoEl.innerHTML = '';
          const inp = zone.querySelector('input[type=file]');
          if(inp) inp.value = '';
          showToast('تم حذف الملف','info');
        };
        zone.appendChild(btn);
      }
      showToast('تم رفع الملف: '+file.name,'success');
    }catch(err){
      this._showZoneError(zone, `<i class="fas fa-triangle-exclamation"></i> ${err.message || 'تعذر قراءة الملف'}`);
    }
  },

  _showZoneError(zone, msg){
    zone.classList.add('zone-invalid');
    zone.classList.remove('zone-valid');
    const iconEl = zone.querySelector(':scope > i');
    if(iconEl) iconEl.className = 'fas fa-file-circle-xmark';
    const errEl = zone.querySelector('.file-error');
    if(errEl) errEl.innerHTML = msg;
    zone._file = null;
    zone.dataset.file = '';
    const inp = zone.querySelector('input[type=file]');
    if(inp) inp.value = '';
  },

  resetAllZones(){
    document.querySelectorAll('.upload-zone').forEach(z=>{
      z._file = null;
      z.dataset.file = '';
      z.classList.remove('zone-valid','zone-invalid');
      const iconEl = z.querySelector(':scope > i');
      if(iconEl) iconEl.className = 'fas fa-cloud-arrow-up';
      const nameEl = z.querySelector('.file-name');
      const errEl = z.querySelector('.file-error');
      const infoEl = z.querySelector('.file-info');
      if(nameEl) nameEl.innerHTML = '';
      if(errEl) errEl.innerHTML = '';
      if(infoEl) infoEl.innerHTML = '';
      const inp = z.querySelector('input[type=file]');
      if(inp) inp.value = '';
      const clr = z.querySelector('.zone-clear');
      if(clr) clr.remove();
    });
  }
};