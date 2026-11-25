const BRANCHES = {
  smart:{name:'سمارت فارما', icon:'fa-microchip'},
  san:{name:'صن فارما', icon:'fa-capsules'},
  taiba:{name:'طيبة القصاصين', icon:'fa-hospital'},
  dawa:{name:'الدواء فارما', icon:'fa-pills'},
  rawi:{name:'الراوي فارما', icon:'fa-mortar-pestle'}
};

const Auth = {
  pendingBranch: null,

  init(){
    if(!App.users || typeof App.users !== 'object') App.users = {};
    if(!App.users['developer']){
      App.users['developer'] = { user:'developer', pass:'msh2000', role:'developer', branch:'dev' };
      App.saveUsers();
    }
    document.getElementById('loginBtn').onclick = ()=> this.login();
    document.getElementById('loginBack').onclick = ()=> this.cancelLogin();
    document.getElementById('loginUser').addEventListener('keypress', (e)=>{
      if(e.key === 'Enter') document.getElementById('loginPass').focus();
    });
    document.getElementById('loginPass').addEventListener('keypress', (e)=>{
      if(e.key === 'Enter') this.login();
    });
  },

  requestLogin(branchKey){
    this.pendingBranch = branchKey;
    const name = BRANCHES[branchKey]
      ? BRANCHES[branchKey].name
      : (branchKey === 'admin' ? 'إنشاء حساب جديد' : branchKey === 'developer' ? 'لوحة المطور' : 'تسجيل الدخول');
    document.getElementById('loginBranchName').textContent = name;
    document.getElementById('loginUser').value = '';
    document.getElementById('loginPass').value = '';
    document.getElementById('loginScreen').classList.remove('hidden');
    document.getElementById('loginUser').focus();
  },

  cancelLogin(){
    this.pendingBranch = null;
    document.getElementById('loginScreen').classList.add('hidden');
  },

  login(){
    const u = document.getElementById('loginUser').value.trim();
    const p = document.getElementById('loginPass').value.trim();
    if(!u || !p){ showToast('يرجى إدخال اسم المستخدم والرقم السري','error'); return; }
    if(!App.users || typeof App.users !== 'object') App.users = {};
    const found = Object.values(App.users).find(x =>
      x && typeof x.user === 'string' &&
      x.user.toLowerCase() === u.toLowerCase() && x.pass === p
    );
    if(!found){ showToast('اسم المستخدم أو الرقم السري غير صحيح','error'); return; }
    if(found.disabled){ showToast('تم إيقاف صلاحية هذا المستخدم','error'); return; }
    const target = this.pendingBranch || found.branch;
    if(found.role === 'developer'){
    } else if(found.role === 'admin'){
      if(target !== 'admin' && target !== 'developer'){
        showToast('هذا المستخدم ليس لديه صلاحية على هذا الفرع','error');
        return;
      }
    } else {
      if(found.branch !== target){
        showToast('هذا المستخدم غير مصرح له بالدخول لهذا الفرع','error');
        return;
      }
    }
    App.currentUser = found;
    App.currentBranch = target;
    document.getElementById('loginScreen').classList.add('hidden');
    bindSidebarNav();
    goToPage(target);
    showToast('تم تسجيل الدخول بنجاح','success');
  }
};

Auth.init();
