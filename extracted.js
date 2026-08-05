
/* =====================================================================
   MICROSOFT SSO (Entra ID / Azure AD) — MSAL.js configuratie
   =====================================================================
   Om dit aan jouw organisatie te koppelen:
   1. Ga naar Azure Portal > Microsoft Entra ID > App registrations > New registration.
   2. Kies "Single-page application (SPA)" als platform en zet de Redirect URI
      op de URL waar dit bestand wordt gehost (bv. https://intranet.bedrijf.nl/hr-applicatie.html).
      Lokaal testen vanaf file:// werkt NIET met redirect-flows — host het bestand
      via een lokale server (bv. `npx serve`) of op een echte URL.
   3. Kopieer de "Application (client) ID" en "Directory (tenant) ID" hieronder.
   4. Geen client secret nodig — SPA's gebruiken PKCE, volledig veilig client-side.
   ===================================================================== */
  const msalConfig = {
    auth: {
      clientId: 'VUL-HIER-JE-APPLICATION-CLIENT-ID-IN',
      authority: 'https://login.microsoftonline.com/VUL-HIER-JE-TENANT-ID-IN',
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
      cacheLocation: 'sessionStorage',
      storeAuthStateInCookie: false
    }
  };

  const loginRequest = { scopes: ['User.Read'] };

  let msalInstance = null;
  let activeAccount = null;

  const loginScreen = document.getElementById('loginScreen');
  const appRoot = document.getElementById('appRoot');
  const msLoginBtn = document.getElementById('msLoginBtn');
  const loginError = document.getElementById('loginError');
  const userInfo = document.getElementById('userInfo');
  const userAvatar = document.getElementById('userAvatar');
  const userNameEl = document.getElementById('userName');
  const userEmailEl = document.getElementById('userEmail');
  const logoutBtn = document.getElementById('logoutBtn');

  function showLoginError(msg){
    loginError.textContent = msg;
    loginError.style.display = 'block';
  }

  function initialsFromName(name){
    return name.split(' ').filter(Boolean).slice(0,2).map(n=>n[0].toUpperCase()).join('');
  }

  function showApp(account){
    activeAccount = account;
    const name = account.name || account.username;
    userNameEl.textContent = name;
    userEmailEl.textContent = account.username || '';
    userAvatar.textContent = initialsFromName(name) || '?';
    userInfo.style.display = 'flex';
    loginScreen.style.display = 'none';
    appRoot.style.display = 'block';
    startEmployeeApp();
  }

  async function initMsal(){
    if(msalConfig.auth.clientId === 'VUL-HIER-JE-APPLICATION-CLIENT-ID-IN'){
      showLoginError('SSO is nog niet geconfigureerd: vul clientId en tenantId in bovenaan het script (msalConfig).');
      return;
    }
    try{
      msalInstance = new msal.PublicClientApplication(msalConfig);
      await msalInstance.initialize();
      const response = await msalInstance.handleRedirectPromise();
      const account = response ? response.account : msalInstance.getAllAccounts()[0];
      if(account){
        msalInstance.setActiveAccount(account);
        showApp(account);
      }
    }catch(err){
      showLoginError('Inloggen mislukt: ' + (err && err.message ? err.message : err));
    }
  }

  msLoginBtn.addEventListener('click', async ()=>{
    loginError.style.display = 'none';
    if(!msalInstance){
      if(msalConfig.auth.clientId === 'VUL-HIER-JE-APPLICATION-CLIENT-ID-IN'){
        showLoginError('SSO is nog niet geconfigureerd: vul clientId en tenantId in bovenaan het script (msalConfig).');
        return;
      }
      msalInstance = new msal.PublicClientApplication(msalConfig);
      await msalInstance.initialize();
    }
    try{
      const result = await msalInstance.loginPopup(loginRequest);
      msalInstance.setActiveAccount(result.account);
      showApp(result.account);
    }catch(err){
      showLoginError('Inloggen mislukt: ' + (err && err.message ? err.message : err));
    }
  });

  logoutBtn.addEventListener('click', ()=>{
    if(msalInstance && activeAccount){
      msalInstance.logoutPopup({ account: activeAccount }).catch(()=>{});
    }
    appRoot.style.display = 'none';
    userInfo.style.display = 'none';
    loginScreen.style.display = 'flex';
  });

  initMsal();

  let appStarted = false;
  function startEmployeeApp(){
  if(appStarted) return; // voorkom dubbele event listeners bij opnieuw inloggen
  appStarted = true;
  const STORAGE_KEY = 'hrapp_employees_v1';

  const seed = [
    {id: cryptoId(), name:'Lotte de Vries', email:'lotte.devries@bedrijf.nl', role:'HR Manager', dept:'HR', date:'2021-03-01', status:'Actief'},
    {id: cryptoId(), name:'Sem Bakker', email:'sem.bakker@bedrijf.nl', role:'Software Engineer', dept:'Engineering', date:'2022-06-15', status:'Actief'},
    {id: cryptoId(), name:'Anouk Jansen', email:'anouk.jansen@bedrijf.nl', role:'Marketing Specialist', dept:'Marketing', date:'2020-11-20', status:'Actief'},
    {id: cryptoId(), name:'Daan Visser', email:'daan.visser@bedrijf.nl', role:'Sales Manager', dept:'Sales', date:'2019-01-10', status:'Inactief'},
    {id: cryptoId(), name:'Fleur Smit', email:'fleur.smit@bedrijf.nl', role:'Financieel Analist', dept:'Finance', date:'2023-02-05', status:'Actief'}
  ];

  function cryptoId(){
    return 'e_' + Math.random().toString(36).slice(2,10) + Date.now().toString(36);
  }

  function loadEmployees(){
    try{
      const raw = localStorage.getItem(STORAGE_KEY);
      if(raw) return JSON.parse(raw);
    }catch(e){}
    saveEmployees(seed);
    return seed;
  }

  function saveEmployees(list){
    try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); }catch(e){}
  }

  let employees = loadEmployees();
  let editingId = null;

  const tableBody = document.getElementById('tableBody');
  const emptyState = document.getElementById('emptyState');
  const statsEl = document.getElementById('stats');
  const searchInput = document.getElementById('searchInput');
  const deptFilter = document.getElementById('deptFilter');
  const statusFilter = document.getElementById('statusFilter');
  const overlay = document.getElementById('overlay');
  const form = document.getElementById('employeeForm');
  const modalTitle = document.getElementById('modalTitle');
  const toast = document.getElementById('toast');

  function initials(name){
    return name.split(' ').filter(Boolean).slice(0,2).map(n=>n[0].toUpperCase()).join('');
  }

  function formatDate(d){
    if(!d) return '&mdash;';
    const parts = d.split('-');
    if(parts.length!==3) return d;
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }

  function showToast(msg){
    toast.textContent = msg;
    toast.classList.add('show');
    setTimeout(()=>toast.classList.remove('show'), 2200);
  }

  function updateDeptFilterOptions(){
    const depts = [...new Set(employees.map(e=>e.dept))].sort();
    const current = deptFilter.value;
    deptFilter.innerHTML = '<option value="">Alle afdelingen</option>' +
      depts.map(d=>`<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join('');
    if(depts.includes(current)) deptFilter.value = current;
  }

  function escapeHtml(str){
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function renderStats(){
    const total = employees.length;
    const active = employees.filter(e=>e.status==='Actief').length;
    const inactive = total - active;
    const depts = new Set(employees.map(e=>e.dept)).size;
    statsEl.innerHTML = `
      <div class="stat-card"><div class="num">${total}</div><div class="label">Totaal werknemers</div></div>
      <div class="stat-card"><div class="num">${active}</div><div class="label">Actief</div></div>
      <div class="stat-card"><div class="num">${inactive}</div><div class="label">Inactief</div></div>
      <div class="stat-card"><div class="num">${depts}</div><div class="label">Afdelingen</div></div>
    `;
  }

  function getFiltered(){
    const q = searchInput.value.trim().toLowerCase();
    const dept = deptFilter.value;
    const status = statusFilter.value;
    return employees.filter(e=>{
      const matchesQ = !q || e.name.toLowerCase().includes(q) || e.email.toLowerCase().includes(q);
      const matchesDept = !dept || e.dept === dept;
      const matchesStatus = !status || e.status === status;
      return matchesQ && matchesDept && matchesStatus;
    });
  }

  function render(){
    renderStats();
    updateDeptFilterOptions();
    const list = getFiltered();
    if(list.length === 0){
      tableBody.innerHTML = '';
      emptyState.style.display = 'block';
      return;
    }
    emptyState.style.display = 'none';
    tableBody.innerHTML = list.map(e=>`
      <tr>
        <td>
          <div class="name-cell">
            <span class="avatar">${escapeHtml(initials(e.name))}</span>
            <span>${escapeHtml(e.name)}</span>
          </div>
        </td>
        <td>${escapeHtml(e.role)}</td>
        <td>${escapeHtml(e.dept)}</td>
        <td>${escapeHtml(e.email)}</td>
        <td>${formatDate(e.date)}</td>
        <td><span class="badge ${e.status==='Actief'?'badge-active':'badge-inactive'}">${escapeHtml(e.status)}</span></td>
        <td>
          <div class="actions-cell">
            <button class="btn-secondary btn-small" data-edit="${e.id}">Bewerken</button>
            <button class="btn-danger btn-small" data-delete="${e.id}">Verwijderen</button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  function openModal(emp){
    editingId = emp ? emp.id : null;
    modalTitle.textContent = emp ? 'Werknemer bewerken' : 'Werknemer toevoegen';
    document.getElementById('empId').value = emp ? emp.id : '';
    document.getElementById('empName').value = emp ? emp.name : '';
    document.getElementById('empEmail').value = emp ? emp.email : '';
    document.getElementById('empRole').value = emp ? emp.role : '';
    document.getElementById('empDept').value = emp ? emp.dept : '';
    document.getElementById('empDate').value = emp ? emp.date : '';
    document.getElementById('empStatus').value = emp ? emp.status : 'Actief';
    clearErrors();
    overlay.classList.add('open');
    document.getElementById('empName').focus();
  }

  function closeModal(){
    overlay.classList.remove('open');
    form.reset();
    editingId = null;
  }

  function clearErrors(){
    document.querySelectorAll('.error-text').forEach(el=>el.style.display='none');
    document.querySelectorAll('input,select').forEach(el=>el.style.borderColor = 'var(--border)');
  }

  function showError(id){
    document.getElementById('err-'+id).style.display = 'block';
    document.getElementById(id).style.borderColor = 'var(--danger)';
  }

  function validate(){
    clearErrors();
    let valid = true;
    const name = document.getElementById('empName').value.trim();
    const email = document.getElementById('empEmail').value.trim();
    const role = document.getElementById('empRole').value.trim();
    const dept = document.getElementById('empDept').value;

    if(!name){ showError('empName'); valid = false; }
    if(!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){ showError('empEmail'); valid = false; }
    if(!role){ showError('empRole'); valid = false; }
    if(!dept){ showError('empDept'); valid = false; }
    return valid;
  }

  // Event listeners
  document.getElementById('addBtn').addEventListener('click', ()=>openModal(null));
  document.getElementById('cancelBtn').addEventListener('click', closeModal);
  overlay.addEventListener('click', (e)=>{ if(e.target === overlay) closeModal(); });

  form.addEventListener('submit', (e)=>{
    e.preventDefault();
    if(!validate()) return;

    const data = {
      name: document.getElementById('empName').value.trim(),
      email: document.getElementById('empEmail').value.trim(),
      role: document.getElementById('empRole').value.trim(),
      dept: document.getElementById('empDept').value,
      date: document.getElementById('empDate').value,
      status: document.getElementById('empStatus').value
    };

    if(editingId){
      const idx = employees.findIndex(e=>e.id === editingId);
      if(idx > -1) employees[idx] = {...employees[idx], ...data};
      showToast('Werknemer bijgewerkt.');
    } else {
      employees.push({id: cryptoId(), ...data});
      showToast('Werknemer toegevoegd.');
    }
    saveEmployees(employees);
    closeModal();
    render();
  });

  tableBody.addEventListener('click', (e)=>{
    const editId = e.target.getAttribute('data-edit');
    const delId = e.target.getAttribute('data-delete');
    if(editId){
      const emp = employees.find(x=>x.id === editId);
      if(emp) openModal(emp);
    }
    if(delId){
      const emp = employees.find(x=>x.id === delId);
      if(emp && confirm(`Weet je zeker dat je ${emp.name} wilt verwijderen?`)){
        employees = employees.filter(x=>x.id !== delId);
        saveEmployees(employees);
        showToast('Werknemer verwijderd.');
        render();
      }
    }
  });

  searchInput.addEventListener('input', render);
  deptFilter.addEventListener('change', render);
  statusFilter.addEventListener('change', render);

  document.addEventListener('keydown', (e)=>{
    if(e.key === 'Escape' && overlay.classList.contains('open')) closeModal();
  });

  render();
  } // einde startEmployeeApp
