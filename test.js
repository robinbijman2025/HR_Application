const fs = require('fs');
const { JSDOM } = require('jsdom');

let html = fs.readFileSync('hr-applicatie.html', 'utf8');
// Strip the remote MSAL script tag (no network in sandbox) and stub a minimal msal global instead.
html = html.replace(/<script\s+src="https:\/\/alcdn[^>]*><\/script>/, '<script>window.msal = { PublicClientApplication: function(cfg){ this.cfg=cfg; this.acc=null; this.initialize=async()=>{}; this.handleRedirectPromise=async()=>null; this.getAllAccounts=()=>[]; this.setActiveAccount=(a)=>{this.acc=a;}; this.loginPopup=async()=>({account:{name:"Test Gebruiker",username:"test@bedrijf.nl"}}); this.logoutPopup=async()=>{}; }; };</script>');

const dom = new JSDOM(html, { runScripts: 'dangerously', resources: 'usable', pretendToBeVisual: true, url: 'https://example.com/hr-applicatie.html' });
const { window } = dom;

function wait(ms){ return new Promise(r=>setTimeout(r, ms)); }

(async () => {
  await wait(200);
  const doc = window.document;

  // trigger login
  const loginBtn = doc.getElementById('msLoginBtn');
  loginBtn.dispatchEvent(new window.Event('click', {bubbles:true}));
  await wait(200);

  const appRoot = doc.getElementById('appRoot');
  console.log('appRoot visible after login:', appRoot.style.display === 'block');

  const rows = doc.querySelectorAll('#tableBody tr');
  console.log('employee rows rendered:', rows.length);

  // open first dossier
  const firstOpenBtn = doc.querySelector('[data-open]');
  firstOpenBtn.dispatchEvent(new window.Event('click', {bubbles:true}));
  await wait(50);
  console.log('dossier open:', doc.getElementById('dossierOverlay').classList.contains('open'));
  console.log('dossier title:', doc.getElementById('dossierTitle').textContent);

  // check history tab shows the created entry
  const historyHtml = doc.getElementById('historyList').innerHTML;
  console.log('history has content:', historyHtml.includes('history-item'));

  // enter edit mode, change role, save
  doc.getElementById('dossierEditBtn').dispatchEvent(new window.Event('click', {bubbles:true}));
  await wait(50);
  const roleInput = doc.getElementById('f_role');
  console.log('role input found in edit mode:', !!roleInput);
  roleInput.value = 'Senior HR Manager';
  doc.getElementById('dossierSaveBtn').dispatchEvent(new window.Event('click', {bubbles:true}));
  await wait(50);

  console.log('dossier title after save:', doc.getElementById('dossierTitle').textContent);
  console.log('dossier subtitle after save:', doc.getElementById('dossierSubtitle').textContent);
  const historyHtml2 = doc.getElementById('historyList').innerHTML;
  console.log('history mentions new role:', historyHtml2.includes('Senior HR Manager'));
  console.log('history mentions old role:', historyHtml2.includes('HR Manager'));

  // check table reflects updated role
  const roleCellText = doc.querySelector('#tableBody tr td:nth-child(2)').textContent;
  console.log('table role cell updated:', roleCellText.trim());

  // check localStorage persisted
  const stored = window.localStorage.getItem('hrapp_employees_v2');
  const parsed = JSON.parse(stored);
  console.log('stored employees count:', parsed.length);
  console.log('stored first employee role:', parsed[0].role);
  console.log('stored first employee history length:', parsed[0].history.length);

  // test add new employee flow
  doc.getElementById('addBtn').dispatchEvent(new window.Event('click', {bubbles:true}));
  await wait(50);
  console.log('create mode title:', doc.getElementById('dossierTitle').textContent);
  doc.getElementById('f_firstName').value = 'Nieuwe';
  doc.getElementById('f_lastName').value = 'Werknemer';
  doc.getElementById('f_email').value = 'nieuwe.werknemer@bedrijf.nl';
  doc.getElementById('f_role').value = 'Tester';
  doc.getElementById('f_dept').value = 'Engineering';
  doc.getElementById('dossierSaveBtn').dispatchEvent(new window.Event('click', {bubbles:true}));
  await wait(50);
  const rowsAfterAdd = doc.querySelectorAll('#tableBody tr');
  console.log('rows after add:', rowsAfterAdd.length);

  // test delete
  const delBtns = doc.querySelectorAll('[data-delete]');
  console.log('delete buttons found:', delBtns.length);

  window.close();
  console.log('TEST_DONE');
})().catch(e => { console.error('TEST_ERROR', e); process.exit(1); });
