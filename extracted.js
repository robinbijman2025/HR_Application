
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
      clientId: '522b8929-f96a-452e-a8de-0271d803eee1',
      authority: 'https://login.microsoftonline.com/aada8fcb-ea3f-43ad-9ca9-62fdaf03d5e8',
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
    startEmployeeApp(name || account.username || 'Onbekende gebruiker');
  }

  async function initMsal(){
    if(msalConfig.auth.clientId === 'VUL-HIER-JE-APPLICATION-CLIENT-ID-IN'){
      showLoginError('SSO is nog niet geconfigureerd: vul clientId en tenantId in bovenaan het script (msalConfig).');
      return;
    }
    try{
      msalInstance = new msal.PublicClientApplication(msalConfig);
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
  function startEmployeeApp(currentUserName){
  if(appStarted) return; // voorkom dubbele event listeners bij opnieuw inloggen
  appStarted = true;

  const STORAGE_KEY = 'hrapp_employees_v2';

  /* ===================================================================
     Veldconfiguratie per categorie van het personeelsdossier
     =================================================================== */
  const FIELDS = [
    // Persoonlijke gegevens
    {key:'firstName', label:'Voornaam', cat:'personal', type:'text', required:true},
    {key:'lastName', label:'Achternaam', cat:'personal', type:'text', required:true},
    {key:'birthDate', label:'Geboortedatum', cat:'personal', type:'date'},
    {key:'gender', label:'Geslacht', cat:'personal', type:'select', options:['','Man','Vrouw','Anders','Zeg ik liever niet']},
    {key:'nationality', label:'Nationaliteit', cat:'personal', type:'text'},
    // Contactgegevens
    {key:'email', label:'E-mailadres', cat:'contact', type:'email', required:true},
    {key:'phone', label:'Telefoonnummer', cat:'contact', type:'tel'},
    {key:'street', label:'Straat en huisnummer', cat:'contact', type:'text'},
    {key:'postalCode', label:'Postcode', cat:'contact', type:'text'},
    {key:'city', label:'Plaats', cat:'contact', type:'text'},
    {key:'emergencyContactName', label:'Noodcontact (naam)', cat:'contact', type:'text'},
    {key:'emergencyContactPhone', label:'Noodcontact (telefoon)', cat:'contact', type:'tel'},
    // Contractinformatie
    {key:'contractType', label:'Type contract', cat:'contract', type:'select', options:['','Onbepaalde tijd','Bepaalde tijd','Stage','Oproepcontract','Uitzendkracht']},
    {key:'startDate', label:'In dienst sinds', cat:'contract', type:'date'},
    {key:'endDate', label:'Einddatum contract', cat:'contract', type:'date'},
    {key:'hoursPerWeek', label:'Uren per week', cat:'contract', type:'number'},
    {key:'probationEndDate', label:'Proeftijd tot', cat:'contract', type:'date'},
    // Salarisgegevens
    {key:'grossSalary', label:'Bruto maandsalaris (€)', cat:'salary', type:'number'},
    {key:'salaryScale', label:'Salarisschaal', cat:'salary', type:'text'},
    {key:'lastRaiseDate', label:'Datum laatste verhoging', cat:'salary', type:'date'},
    // Functie en afdeling
    {key:'role', label:'Functie', cat:'function', type:'text', required:true},
    {key:'dept', label:'Afdeling', cat:'function', type:'select', options:['','Sales','Marketing','Engineering','HR','Finance','Operations','Support'], required:true},
    {key:'manager', label:'Leidinggevende', cat:'function', type:'text'},
    {key:'status', label:'Status', cat:'function', type:'select', options:['Actief','Inactief']}
  ];
  const FIELDS_BY_KEY = Object.fromEntries(FIELDS.map(f=>[f.key,f]));

  function emptyEmployee(){
    const e = {id:null};
    FIELDS.forEach(f=>{ e[f.key] = f.key==='status' ? 'Actief' : ''; });
    e.history = [];
    return e;
  }

  function fullName(e){
    return `${e.firstName||''} ${e.lastName||''}`.trim() || '(naam onbekend)';
  }

  function cryptoId(){
    return 'e_' + Math.random().toString(36).slice(2,10) + Date.now().toString(36);
  }

  function nowIso(){ return new Date().toISOString(); }

  function seedHistory(dateStr){
    return [{ ts: (dateStr ? dateStr+'T09:00:00' : nowIso()), user:'Systeem (voorbeelddata)', action:'created' }];
  }

  const seed = [
    {id:cryptoId(), firstName:'Lotte', lastName:'de Vries', birthDate:'1988-04-12', gender:'Vrouw', nationality:'Nederlandse',
     email:'lotte.devries@bedrijf.nl', phone:'06-12345678', street:'Kerkstraat 12', postalCode:'1017 GC', city:'Amsterdam', emergencyContactName:'Peter de Vries', emergencyContactPhone:'06-87654321',
     contractType:'Onbepaalde tijd', startDate:'2021-03-01', endDate:'', hoursPerWeek:40, probationEndDate:'',
     grossSalary:4800, salaryScale:'Schaal 9', lastRaiseDate:'2024-01-01',
     role:'HR Manager', dept:'HR', manager:'', status:'Actief', history:seedHistory('2021-03-01')},
    {id:cryptoId(), firstName:'Sem', lastName:'Bakker', birthDate:'1995-09-23', gender:'Man', nationality:'Nederlandse',
     email:'sem.bakker@bedrijf.nl', phone:'06-23456789', street:'Prinsengracht 45', postalCode:'1015 DX', city:'Amsterdam', emergencyContactName:'Marieke Bakker', emergencyContactPhone:'06-76543210',
     contractType:'Onbepaalde tijd', startDate:'2022-06-15', endDate:'', hoursPerWeek:40, probationEndDate:'',
     grossSalary:4200, salaryScale:'Schaal 8', lastRaiseDate:'',
     role:'Software Engineer', dept:'Engineering', manager:'Lotte de Vries', status:'Actief', history:seedHistory('2022-06-15')},
    {id:cryptoId(), firstName:'Anouk', lastName:'Jansen', birthDate:'1992-01-30', gender:'Vrouw', nationality:'Nederlandse',
     email:'anouk.jansen@bedrijf.nl', phone:'06-34567890', street:'Vondelstraat 8', postalCode:'1054 GD', city:'Amsterdam', emergencyContactName:'', emergencyContactPhone:'',
     contractType:'Bepaalde tijd', startDate:'2020-11-20', endDate:'2026-11-19', hoursPerWeek:32, probationEndDate:'',
     grossSalary:3600, salaryScale:'Schaal 6', lastRaiseDate:'2023-06-01',
     role:'Marketing Specialist', dept:'Marketing', manager:'', status:'Actief', history:seedHistory('2020-11-20')},
    {id:cryptoId(), firstName:'Daan', lastName:'Visser', birthDate:'1985-07-02', gender:'Man', nationality:'Nederlandse',
     email:'daan.visser@bedrijf.nl', phone:'06-45678901', street:'Zuidas 100', postalCode:'1082 MA', city:'Amsterdam', emergencyContactName:'', emergencyContactPhone:'',
     contractType:'Onbepaalde tijd', startDate:'2019-01-10', endDate:'', hoursPerWeek:40, probationEndDate:'',
     grossSalary:5200, salaryScale:'Schaal 10', lastRaiseDate:'2022-01-01',
     role:'Sales Manager', dept:'Sales', manager:'', status:'Inactief', history:seedHistory('2019-01-10')},
    {id:cryptoId(), firstName:'Fleur', lastName:'Smit', birthDate:'1997-12-05', gender:'Vrouw', nationality:'Nederlandse',
     email:'fleur.smit@bedrijf.nl', phone:'06-56789012', street:'Herengracht 200', postalCode:'1016 BS', city:'Amsterdam', emergencyContactName:'', emergencyContactPhone:'',
     contractType:'Bepaalde tijd', startDate:'2023-02-05', endDate:'2026-02-04', hoursPerWeek:40, probationEndDate:'2023-05-05',
     grossSalary:3400, salaryScale:'Schaal 5', lastRaiseDate:'',
     role:'Financieel Analist', dept:'Finance', manager:'', status:'Actief', history:seedHistory('2023-02-05')}
  ];

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

  const tableBody = document.getElementById('tableBody');
  const emptyState = document.getElementById('emptyState');
  const statsEl = document.getElementById('stats');
  const searchInput = document.getElementById('searchInput');
  const deptFilter = document.getElementById('deptFilter');
  const statusFilter = document.getElementById('statusFilter');
  const toast = document.getElementById('toast');

  function initials(e){
    return ((e.firstName||'')[0]||'').toUpperCase() + ((e.lastName||'')[0]||'').toUpperCase();
  }

  function formatDate(d){
    if(!d) return '—';
    const parts = d.split('T')[0].split('-');
    if(parts.length!==3) return d;
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }

  function formatDateTime(iso){
    if(!iso) return '';
    const d = new Date(iso);
    if(isNaN(d.getTime())) return iso;
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yyyy = d.getFullYear();
    const hh = String(d.getHours()).padStart(2,'0');
    const min = String(d.getMinutes()).padStart(2,'0');
    return `${dd}-${mm}-${yyyy} ${hh}:${min}`;
  }

  function formatFieldValue(field, value){
    if(value === '' || value === null || value === undefined) return '—';
    if(field.type === 'date') return formatDate(value);
    if(field.key === 'grossSalary') return '€ ' + Number(value).toLocaleString('nl-NL');
    if(field.key === 'hoursPerWeek') return value + ' u/week';
    return escapeHtml(String(value));
  }

  function showToast(msg){
    toast.textContent = msg;
    toast.classList.add('show');
    setTimeout(()=>toast.classList.remove('show'), 2200);
  }

  function updateDeptFilterOptions(){
    const depts = [...new Set(employees.map(e=>e.dept).filter(Boolean))].sort();
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
    const depts = new Set(employees.map(e=>e.dept).filter(Boolean)).size;
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
      const matchesQ = !q || fullName(e).toLowerCase().includes(q) || (e.email||'').toLowerCase().includes(q);
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
            <span class="avatar">${escapeHtml(initials(e))}</span>
            <span class="name-link" data-open="${e.id}">${escapeHtml(fullName(e))}</span>
          </div>
        </td>
        <td>${escapeHtml(e.role||'—')}</td>
        <td>${escapeHtml(e.dept||'—')}</td>
        <td><span class="badge badge-neutral">${escapeHtml(e.contractType||'—')}</span></td>
        <td><span class="badge ${e.status==='Actief'?'badge-active':'badge-inactive'}">${escapeHtml(e.status)}</span></td>
        <td>
          <div class="actions-cell">
            <button class="btn-secondary btn-small" data-open="${e.id}">Dossier</button>
            <button class="btn-danger btn-small" data-delete="${e.id}">Verwijderen</button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  /* ===================================================================
     Dossierscherm
     =================================================================== */
  const dossierOverlay = document.getElementById('dossierOverlay');
  const dossierTitle = document.getElementById('dossierTitle');
  const dossierSubtitle = document.getElementById('dossierSubtitle');
  const dossierAvatar = document.getElementById('dossierAvatar');
  const dossierEditBtn = document.getElementById('dossierEditBtn');
  const dossierCloseBtn = document.getElementById('dossierCloseBtn');
  const dossierCancelBtn = document.getElementById('dossierCancelBtn');
  const dossierSaveBtn = document.getElementById('dossierSaveBtn');
  const dossierFooter = document.getElementById('dossierFooter');
  const dossierTabs = document.getElementById('dossierTabs');
  const historyList = document.getElementById('historyList');

  let currentRecord = null;   // origineel record (of null bij nieuw)
  let draftRecord = null;     // werkkopie tijdens bewerken/aanmaken
  let dossierMode = 'view';   // 'view' | 'edit' | 'create'

  const CATS = ['personal','contact','contract','salary','function'];

  function fieldRowHtml(field, value, editable){
    if(!editable){
      return `<div class="field-row"><label>${escapeHtml(field.label)}</label><div class="field-view">${formatFieldValue(field, value)}</div></div>`;
    }
    const id = 'f_' + field.key;
    let input = '';
    if(field.type === 'select'){
      input = `<select id="${id}" data-field="${field.key}">` +
        field.options.map(o=>`<option value="${escapeHtml(o)}" ${o===value?'selected':''}>${o===''?'Kies...':escapeHtml(o)}</option>`).join('') +
        `</select>`;
    } else {
      const v = value===null||value===undefined?'':value;
      input = `<input type="${field.type}" id="${id}" data-field="${field.key}" value="${escapeHtml(String(v))}">`;
    }
    return `<div class="field-row"><label>${escapeHtml(field.label)}${field.required?' *':''}</label>${input}<div class="error-text" id="err_${field.key}">Dit veld is verplicht.</div></div>`;
  }

  function renderDossierPanels(){
    CATS.forEach(cat=>{
      const panel = document.getElementById('panel-'+cat);
      const fields = FIELDS.filter(f=>f.cat===cat);
      const editable = dossierMode !== 'view';
      panel.innerHTML = fields.map(f=>fieldRowHtml(f, draftRecord[f.key], editable)).join('');
    });
    renderHistory();
  }

  function renderHistory(){
    const rec = currentRecord || draftRecord;
    const history = (rec && rec.history) ? [...rec.history].sort((a,b)=> new Date(b.ts) - new Date(a.ts)) : [];
    if(history.length === 0){
      historyList.innerHTML = '<div class="history-empty">Nog geen wijzigingen geregistreerd.</div>';
      return;
    }
    historyList.innerHTML = history.map(h=>{
      let text = '';
      if(h.action === 'created'){
        text = `<b>${escapeHtml(h.user)}</b> heeft dit dossier aangemaakt.`;
      } else {
        const fieldLabel = FIELDS_BY_KEY[h.field] ? FIELDS_BY_KEY[h.field].label : h.field;
        text = `<b>${escapeHtml(h.user)}</b> wijzigde <b>${escapeHtml(fieldLabel)}</b>: "${escapeHtml(h.oldValue||'—')}" &rarr; "${escapeHtml(h.newValue||'—')}"`;
      }
      return `<div class="history-item"><div class="history-dot"></div><div><div class="history-text">${text}</div><div class="history-meta">${formatDateTime(h.ts)}</div></div></div>`;
    }).join('');
  }

  function switchTab(tabKey){
    dossierTabs.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.dataset.tab===tabKey));
    document.querySelectorAll('.tab-panel').forEach(p=>p.classList.toggle('active', p.dataset.panel===tabKey));
  }

  dossierTabs.addEventListener('click', (e)=>{
    const btn = e.target.closest('.tab-btn');
    if(btn) switchTab(btn.dataset.tab);
  });

  function updateDossierHeader(){
    const rec = draftRecord;
    dossierAvatar.textContent = initials(rec) || '?';
    dossierTitle.textContent = dossierMode==='create' ? 'Nieuwe werknemer' : fullName(rec);
    dossierSubtitle.textContent = dossierMode==='create' ? 'Vul het dossier in en klik op Opslaan' : `${rec.role||'—'} · ${rec.dept||'—'}`;
    dossierEditBtn.style.display = dossierMode==='create' ? 'none' : (dossierMode==='edit' ? 'none' : 'inline-block');
    dossierFooter.style.display = dossierMode==='view' ? 'none' : 'flex';
  }

  function openDossier(emp){
    currentRecord = emp;
    draftRecord = JSON.parse(JSON.stringify(emp));
    dossierMode = 'view';
    switchTab('personal');
    updateDossierHeader();
    renderDossierPanels();
    dossierOverlay.classList.add('open');
  }

  function openNewDossier(){
    currentRecord = null;
    draftRecord = emptyEmployee();
    dossierMode = 'create';
    switchTab('personal');
    updateDossierHeader();
    renderDossierPanels();
    dossierOverlay.classList.add('open');
  }

  function closeDossier(){
    dossierOverlay.classList.remove('open');
    currentRecord = null;
    draftRecord = null;
    dossierMode = 'view';
  }

  function enterEditMode(){
    dossierMode = 'edit';
    draftRecord = JSON.parse(JSON.stringify(currentRecord));
    updateDossierHeader();
    renderDossierPanels();
  }

  function collectDraftFromInputs(){
    FIELDS.forEach(f=>{
      const el = document.getElementById('f_'+f.key);
      if(el) draftRecord[f.key] = el.value;
    });
  }

  function validateDraft(){
    let valid = true;
    FIELDS.forEach(f=>{
      const errEl = document.getElementById('err_'+f.key);
      if(errEl) errEl.style.display = 'none';
      const el = document.getElementById('f_'+f.key);
      if(el) el.style.borderColor = 'var(--border)';
    });
    FIELDS.filter(f=>f.required).forEach(f=>{
      const value = draftRecord[f.key];
      let ok = !!(value && String(value).trim());
      if(f.key === 'email' && ok){
        ok = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
      }
      if(!ok){
        valid = false;
        const errEl = document.getElementById('err_'+f.key);
        const el = document.getElementById('f_'+f.key);
        if(errEl){ errEl.style.display='block'; errEl.textContent = f.key==='email' ? 'Voer een geldig e-mailadres in.' : 'Dit veld is verplicht.'; }
        if(el) el.style.borderColor = 'var(--danger)';
      }
    });
    return valid;
  }

  function saveDossier(){
    collectDraftFromInputs();
    if(!validateDraft()) return;

    if(dossierMode === 'create'){
      draftRecord.id = cryptoId();
      draftRecord.history = [{ ts: nowIso(), user: currentUserName, action:'created' }];
      employees.push(draftRecord);
      saveEmployees(employees);
      showToast('Werknemer en dossier aangemaakt.');
      closeDossier();
      render();
      return;
    }

    // edit mode: diff tegen currentRecord, wijzigingen loggen
    const changes = [];
    FIELDS.forEach(f=>{
      const oldVal = currentRecord[f.key] ?? '';
      const newVal = draftRecord[f.key] ?? '';
      if(String(oldVal) !== String(newVal)){
        changes.push({ ts: nowIso(), user: currentUserName, action:'updated', field:f.key, oldValue: oldVal, newValue: newVal });
      }
    });

    const idx = employees.findIndex(e=>e.id === currentRecord.id);
    if(idx > -1){
      const updated = {...currentRecord, ...draftRecord};
      updated.history = [...(currentRecord.history||[]), ...changes];
      employees[idx] = updated;
      saveEmployees(employees);
      currentRecord = updated;
      draftRecord = JSON.parse(JSON.stringify(updated));
    }
    showToast(changes.length ? `Dossier bijgewerkt (${changes.length} wijziging${changes.length>1?'en':''}).` : 'Geen wijzigingen gevonden.');
    dossierMode = 'view';
    updateDossierHeader();
    renderDossierPanels();
    render();
  }

  dossierEditBtn.addEventListener('click', enterEditMode);
  dossierCloseBtn.addEventListener('click', closeDossier);
  dossierCancelBtn.addEventListener('click', ()=>{
    if(dossierMode === 'create'){ closeDossier(); return; }
    dossierMode = 'view';
    draftRecord = JSON.parse(JSON.stringify(currentRecord));
    updateDossierHeader();
    renderDossierPanels();
  });
  dossierSaveBtn.addEventListener('click', saveDossier);
  dossierOverlay.addEventListener('click', (e)=>{ if(e.target === dossierOverlay) closeDossier(); });

  document.getElementById('addBtn').addEventListener('click', openNewDossier);

  tableBody.addEventListener('click', (e)=>{
    const openId = e.target.getAttribute('data-open');
    const delId = e.target.getAttribute('data-delete');
    if(openId){
      const emp = employees.find(x=>x.id === openId);
      if(emp) openDossier(emp);
    }
    if(delId){
      const emp = employees.find(x=>x.id === delId);
      if(emp && confirm(`Weet je zeker dat je ${fullName(emp)} wilt verwijderen? Dit verwijdert ook de volledige historie.`)){
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
    if(e.key === 'Escape' && dossierOverlay.classList.contains('open')) closeDossier();
  });

  render();
  } // einde startEmployeeApp
