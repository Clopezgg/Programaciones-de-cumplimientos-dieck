(() => {
  const seed = window.DIECK_DATA || {};
  const LS = {
    ui: 'dieck_sap_ui_v1',
    master: 'dieck_sap_master_v1',
    programming: 'dieck_sap_programming_v1',
    production: 'dieck_sap_production_v1',
    settings: 'dieck_sap_settings_v1',
    auth: 'dieck_sap_auth_v1'
  };

  const NAV = [
    { group: 'General', items: [
      { code: 'ZINI', name: 'Inicio / Dashboard', view: 'dashboard' },
      { code: 'ZCFG', name: 'Configuración SAP', view: 'config' }
    ] },
    { group: 'Programación', items: [
      { code: 'ZP001', name: 'Programación semanal', view: 'programming' },
      { code: 'ZP002', name: 'Importar programación', view: 'programming' }
    ] },
    { group: 'Producción', items: [
      { code: 'ZD001', name: 'Producción diaria', view: 'production' },
      { code: 'ZD002', name: 'Importar producción', view: 'production' }
    ] },
    { group: 'Control', items: [
      { code: 'ZC001', name: 'Cumplimiento semanal', view: 'compliance' },
      { code: 'ZL001', name: 'Lotificación mensual', view: 'lotification' },
      { code: 'ZM001', name: 'Maestro de materiales', view: 'master' },
      { code: 'ZDOC', name: 'Documentos / impresión', view: 'documents' }
    ] },
    { group: 'Seguridad', items: [
      { code: 'ZUSR', name: 'Usuarios y permisos', view: 'admin' },
      { code: 'ZROL', name: 'Roles / autorizaciones', view: 'admin' }
    ] }
  ];

  const dayNames = ['DOMINGO','LUNES','MARTES','MIÉRCOLES','JUEVES','VIERNES','SÁBADO'];
  const monthNames = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  const shortMonth = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  const WEEKDAY_INDEX = { 'LUNES':1, 'MARTES':2, 'MIÉRCOLES':3, 'MIERCOLES':3, 'JUEVES':4, 'VIERNES':5, 'SÁBADO':6, 'SABADO':6, 'DOMINGO':0 };

  const $ = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  const uid = () => 'id_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

  function defaultAuth() {
    return {
      users: [
        { id: 'u-super', username: 'admin', fullName: 'Administrador General', password: 'Admin123!', role: 'SUPER', active: true },
        { id: 'u-plan', username: 'planner', fullName: 'Planeación', password: 'Plan123!', role: 'PLAN', active: true },
        { id: 'u-prod', username: 'produccion', fullName: 'Producción', password: 'Prod123!', role: 'PROD', active: true },
        { id: 'u-read', username: 'consulta', fullName: 'Consulta', password: 'Read123!', role: 'READ', active: true }
      ],
      roles: [
        { code: 'SUPER', name: 'Superusuario', permissions: ['*'] },
        { code: 'ADMIN', name: 'Administración', permissions: ['dashboard','programming','production','compliance','lotification','master','config','documents','admin','ZINI','ZCFG','ZP001','ZP002','ZD001','ZD002','ZC001','ZL001','ZM001','ZDOC','ZUSR','ZROL'] },
        { code: 'PLAN', name: 'Planeación', permissions: ['dashboard','programming','documents','ZINI','ZP001','ZP002','ZDOC'] },
        { code: 'PROD', name: 'Operación', permissions: ['dashboard','production','compliance','lotification','documents','ZINI','ZD001','ZD002','ZC001','ZL001','ZDOC'] },
        { code: 'READ', name: 'Consulta', permissions: ['dashboard','documents','ZINI','ZDOC'] }
      ],
      session: null,
      audit: [],
      error: ''
    };
  }

  const esc = (value) => String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
  const selectOptions = (options, current) => options.map(opt => {
    const value = String(opt ?? '');
    const selected = String(current ?? '') === value ? ' selected' : '';
    return `<option value="${esc(value)}"${selected}>${esc(value)}</option>`;
  }).join('');

  const today = new Date();
  const defaultUI = {
    view: 'dashboard',
    date: toISO(today),
    week: isoWeek(today),
    theme: 'quartz',
    density: 'normal',
    lotThreshold: 3000,
    quickArea: 'PAQUETERIA',
    filters: { programming: '', production: '', master: '', documents: '' },
    filter: { query: '', area: '', from: '', to: '' }
  };

  const state = {
    ui: loadJSON(LS.ui, seed.ui ? {...defaultUI, ...seed.ui} : defaultUI),
    settings: loadJSON(LS.settings, seed.settings || { packRate: 235, millingRate: 100, defaultEnvas: 2, defaultTarima: 100 }),
    master: normalizeMaster(loadJSON(LS.master, seed.master || [])),
    programming: normalizeProgramming(loadJSON(LS.programming, seed.weeklyPlans || [])),
    production: normalizeProduction(loadJSON(LS.production, seed.production || [])),
    auth: loadJSON(LS.auth, seed.auth || defaultAuth()),
    importPreview: ''
  };

  let pickerState = { open: false, input: null, target: null, items: [], index: -1 };
  let saveTimer = null;

  const refs = {};

  function init() {
    cacheRefs();
    if (!['dashboard','programming','production','compliance','lotification','master','config','documents'].includes(state.ui.view)) {
      state.ui.view = 'dashboard';
    }
    renderNav();
    applyTheme(state.ui.theme);
    bindGlobalActions();
    renderAll();
    updateStatusBar();
  }

  function cacheRefs() {
    [
      'navTree','screenHost','screenTitle','screenSubtitle','bannerStats','statusWeek','statusDate',
      'statusView','statusSave','okCode','btnRunCode','btnLogout','sessionChip','importModal','importText','importKind',
      'importPreview','closeImportModal','btnPreviewImport','btnApplyImport','pickerPanel'
    ].forEach(id => refs[id] = $('#' + id));
  }

  function bindGlobalActions() {
    document.body.addEventListener('click', onBodyClick);
    document.body.addEventListener('input', onBodyInput);
    document.body.addEventListener('keydown', onBodyKeydown);
    window.addEventListener('resize', () => { if (pickerState.open) positionPicker(pickerState.input); });

    refs.okCode.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); runOkCode(refs.okCode.value); }
    });
    refs.btnRunCode.addEventListener('click', () => runOkCode(refs.okCode.value));
    if (refs.btnLogout) refs.btnLogout.addEventListener('click', () => handleAction('logout'));

    refs.closeImportModal.addEventListener('click', closeImportModal);
    refs.btnPreviewImport.addEventListener('click', previewImport);
    refs.btnApplyImport.addEventListener('click', applyImportFromModal);
    refs.importModal.addEventListener('click', e => {
      if (e.target === refs.importModal) closeImportModal();
    });
  }

  function onBodyClick(ev) {
    const navBtn = ev.target.closest('[data-nav]');
    if (navBtn) {
      const view = navBtn.dataset.nav;
      if (view) setView(view);
      return;
    }
    const actionBtn = ev.target.closest('[data-action]');
    if (actionBtn) {
      handleAction(actionBtn.dataset.action, actionBtn);
      return;
    }
    const pickerItem = ev.target.closest('.suggestion-item');
    if (pickerItem && pickerState.open) {
      chooseSuggestion(Number(pickerItem.dataset.index));
      return;
    }
    const transBtn = ev.target.closest('[data-transaction]');
    if (transBtn) {
      runOkCode(transBtn.dataset.transaction);
      return;
    }
    const rowAction = ev.target.closest('[data-row-action]');
    if (rowAction) {
      handleRowAction(rowAction);
      return;
    }
    const rowPlus = ev.target.closest('[data-plus]');
    if (rowPlus) {
      addRow(rowPlus.dataset.plus);
      return;
    }
    const pickerOpen = ev.target.closest('[data-picker]');
    if (pickerOpen) {
      const input = $('#' + pickerOpen.dataset.picker);
      if (input) openPicker(input);
      return;
    }
    if (!ev.target.closest('.suggestions')) closePicker();
  }

  function onBodyInput(ev) {
    const t = ev.target;
    if (t.matches('[data-bind]')) {
      const { bind, row, field, list } = t.dataset;
      updateRow(bind, row, field, t.value, list);
      if (field === 'descripcion' || field === 'codigo') {
        const rowObj = getRow(bind, row);
        autoFillFromMaster(rowObj, { sourceField: field, value: t.value });
        scheduleRender(false);
      }
      if (field === 'area' || field === 'programado' || field === 'cantidad' || field === 'total' || field === 'turno' || field === 'obs') {
        scheduleRender(false);
      }
      if (field === 'descripcion') {
        const q = t.value.trim();
        if (q.length >= 1) openPicker(t, q);
        else closePicker();
      }
      persistSoon();
    }
    if (t.matches('#importText, #importKind')) {
      persistSoon();
    }
    if (t.matches('[data-auth-bind]')) {
      updateAuthRow(t.dataset.authBind, t.dataset.row, t.dataset.field, t.type === 'checkbox' ? t.checked : t.value);
    }
    if (t.matches('[data-search-master]')) {
      const q = t.value.trim();
      const view = t.dataset.searchMaster;
      filterMasterTable(view, q);
    }
    if (t.matches('[data-search-table]')) {
      filterTableRows(t.dataset.searchTable, t.value.trim());
    }
  }

  function onBodyKeydown(ev) {
    const t = ev.target;
    if (!isAuthenticated() && t.matches('#authUser, #authPass') && ev.key === 'Enter') {
      ev.preventDefault();
      handleAction('login-submit');
      return;
    }

    if (pickerState.open && ['ArrowDown','ArrowUp','Enter','Escape'].includes(ev.key) && pickerState.input) {
      ev.preventDefault();
      if (ev.key === 'ArrowDown') movePicker(1);
      else if (ev.key === 'ArrowUp') movePicker(-1);
      else if (ev.key === 'Enter') {
        if (pickerState.index >= 0) chooseSuggestion(pickerState.index);
        else closePicker();
      } else if (ev.key === 'Escape') closePicker();
      return;
    }

    if (t.matches('input, select, textarea') && ev.key === 'Enter') {
      const list = focusablesIn(t.closest('tr') || t.closest('.panel') || document.body);
      const i = list.indexOf(t);
      if (i >= 0) {
        ev.preventDefault();
        const next = list[i + 1];
        if (next) next.focus();
        else if (t.closest('tr') && t.closest('tr').nextElementSibling) {
          const nextRow = t.closest('tr').nextElementSibling;
          const first = nextRow.querySelector('input,select,textarea,button');
          if (first) first.focus();
        } else {
          t.blur();
        }
      }
    }
  }

  function focusablesIn(root) {
    return $$('input, select, textarea, button', root).filter(el => !el.disabled && el.offsetParent !== null);
  }

  function handleAction(action, btn) {
    switch(action) {
      case 'quick-print': window.print(); break;
      case 'close-import-modal': closeImportModal(); break;
      case 'login-submit': submitLoginFromDom(); break;
      case 'demo-login': submitDemoLogin(); break;
      case 'logout': logout(); break;
      case 'new-production-row': addRow('production'); break;
      case 'new-programming-row': addRow('programming'); break;
      case 'save-all': persistAll(); break;
      case 'export-json': exportJSON(); break;
      case 'import-json': openImportModal('production'); break;
      case 'export-csv': exportCSV(); break;
      case 'clear-production': if (confirm('¿Vaciar producción?')) { state.production = []; persistAll(); renderAll(); } break;
      case 'clear-programming': if (confirm('¿Vaciar programación?')) { state.programming = []; persistAll(); renderAll(); } break;
      case 'open-import': openImportModal(btn.dataset.kind || 'production'); break;
      case 'theme': setTheme(btn.dataset.theme); break;
      case 'reset-ui': if (confirm('¿Restaurar configuración visual?')) { localStorage.removeItem(LS.ui); localStorage.removeItem(LS.settings); location.reload(); } break;
      case 'save-master': persistAll(true); break;
      case 'add-master': addMasterRow(); break;
      case 'restore-master': if (confirm('¿Restaurar maestro original?')) { state.master = normalizeMaster(seed.master || []); persistAll(); renderAll(); } break;
      case 'save-settings': persistAll(true); break;
      case 'print-report': window.print(); break;
      case 'toggle-density': state.ui.density = state.ui.density === 'compact' ? 'normal' : 'compact'; document.body.dataset.density = state.ui.density; persistAll(); renderAll(); break;
      case 'recalc-all': renderAll(); break;
      case 'go-history': setView('documents'); break;
      case 'open-modal-import': openImportModal(state.ui.view === 'programming' ? 'programming' : state.ui.view === 'master' ? 'master' : 'production'); break;
      case 'export-report': exportCSV('compliance'); break;
      case 'add-auth-user': addAuthUser(); break;
      case 'add-auth-role': addAuthRole(); break;
      case 'save-auth': persistAll(true); break;
      case 'seed-auth': if (confirm('¿Restaurar usuarios y roles demo?')) { state.auth = defaultAuth(); persistAll(); renderAll(); } break;
      case 'delete-auth-row': deleteAuthRow(btn.dataset.bind, btn.dataset.id); break;
    }
  }

  function handleRowAction(btn) {
    const bind = btn.dataset.bind;
    const id = btn.dataset.id;
    const action = btn.dataset.rowAction;
    const idx = state[bind].findIndex(r => r.id === id);
    if (idx < 0) return;
    if (action === 'delete') {
      state[bind].splice(idx, 1);
      persistAll();
      renderAll();
      return;
    }
    if (action === 'annul') {
      const row = state[bind][idx];
      row.annulled = !row.annulled;
      if (row.annulled && bind === 'production') {
        row.note = row.note ? row.note : 'ANULADO';
      }
      persistAll();
      renderAll();
      return;
    }
    if (action === 'copy') {
      const copy = {...state[bind][idx], id: uid()};
      state[bind].splice(idx + 1, 0, copy);
      persistAll();
      renderAll();
    }
  }

  function setView(view) {
    if (!isAuthenticated()) {
      state.ui.view = 'dashboard';
      persistAll(false);
      renderAll();
      return;
    }
    if (!canAccess(view)) {
      alert('No tienes permisos para abrir esa transacción.');
      return;
    }
    state.ui.view = view;
    persistAll(false);
    renderAll();
  }

  function runOkCode(raw) {
    const code = String(raw || '').trim().toUpperCase();
    if (!code) return;
    const normalized = code.replace(/^\/N/, '');
    const map = {
      ZINI: 'dashboard', ZCFG: 'config', ZP001: 'programming', ZP002: 'programming',
      ZD001: 'production', ZD002: 'production', ZC001: 'compliance', ZL001: 'lotification',
      ZM001: 'master', ZDOC: 'documents', ZUSR: 'admin', ZROL: 'admin'
    };
    if (map[normalized]) {
      if (!canAccess(map[normalized])) {
        alert('No tienes permiso para esa transacción.');
        return;
      }
      setView(map[normalized]);
      refs.okCode.value = '';
      return;
    }
    if (normalized === 'PRINT') window.print();
  }

  function renderNav() {
    const groups = NAV.map(group => ({
      ...group,
      items: group.items.filter(item => isAuthenticated() ? canAccess(item.view) : false)
    })).filter(group => group.items.length);
    refs.navTree.innerHTML = groups.length ? groups.map(group => `
      <div class="nav-group">
        <div class="nav-group-title">${group.group}</div>
        ${group.items.map(item => `
          <button class="nav-item ${state.ui.view === item.view ? 'active' : ''}" data-nav="${item.view}" type="button">
            <span><span class="code">${item.code}</span> · <span class="name">${item.name}</span></span>
            <span>▶</span>
          </button>
        `).join('')}
      </div>
    `).join('') : `
      <div class="nav-group">
        <div class="nav-group-title">Acceso</div>
        <div class="timeline-item">
          <strong>Inicia sesión</strong>
          <div class="small">Ingresa tus credenciales para ver las transacciones disponibles.</div>
        </div>
      </div>
    `;
  }

  function renderAll() {
    syncAuthSession();
    renderNav();
    updateStatusBar();
    const view = isAuthenticated() ? state.ui.view : 'auth';
    refs.bannerStats.innerHTML = isAuthenticated() ? renderBannerStats() : `
      <div class="stat-card"><span>Estado</span><strong>Bloqueado</strong></div>
      <div class="stat-card"><span>Modo</span><strong>Acceso</strong></div>
      <div class="stat-card"><span>Usuarios</span><strong>${fmtNumber(state.auth.users.length)}</strong></div>
      <div class="stat-card"><span>Roles</span><strong>${fmtNumber(state.auth.roles.length)}</strong></div>
    `;
    const titleMap = {
      dashboard: ['ERP Operativo', 'Accesos rápidos, indicadores y último movimiento.'],
      programming: ['Programación semanal', 'Captura por transacción con búsqueda rápida y navegación por Enter.'],
      production: ['Producción diaria', 'Captura de producción real por día, turno, producto y observación.'],
      compliance: ['Cumplimiento del plan semanal', 'Comparación automática entre lo programado y lo producido por día.'],
      lotification: ['Control de lotificación', 'Generación de lotes mensuales por familia y secuencia.'],
      master: ['Maestro de materiales', 'Mantenimiento de códigos, pesos, familia y parámetros de cálculo.'],
      config: ['Configuración del sistema', 'Tema, densidad, umbrales y parámetros operativos.'],
      documents: ['Documentos y reportes', 'Salida rápida de impresión y trazabilidad por módulo.'],
      admin: ['Usuarios, roles y permisos', 'Control de acceso, autorizaciones y auditoría.'],
      auth: ['Acceso al sistema', 'Ingresa para continuar.']
    };
    const screenMeta = titleMap[view] || titleMap.dashboard;
    refs.screenTitle.textContent = screenMeta[0];
    refs.screenSubtitle.textContent = screenMeta[1];
    refs.screenHost.innerHTML = renderView(view);
    bindView(view);
  }

  function renderBannerStats() {
    const week = currentWeekRange();
    const prod = state.production.filter(r => !r.annulled && isWithin(r.date, week.start, week.end)).reduce((a,b)=>a + num(b.total), 0);
    const plan = state.programming.filter(r => !r.annulled && isWithin(r.date, week.start, week.end)).reduce((a,b)=>a + num(b.programmed), 0);
    const comp = plan ? Math.min(100, (prod / plan) * 100) : 0;
    const lots = lotificationRows().length;
    return [
      ['Programado', fmtNumber(plan)],
      ['Producido', fmtNumber(prod)],
      ['Cumplimiento', fmtPct(comp)],
      ['Lotes', String(lots)]
    ].map(([l,v])=>`<div class="stat-card"><span>${l}</span><strong>${v}</strong></div>`).join('');
  }

  function renderView(view) {
    if (view === 'auth') return renderAuth();
    if (view === 'dashboard') return renderDashboard();
    if (view === 'programming') return renderProgramming();
    if (view === 'production') return renderProduction();
    if (view === 'compliance') return renderCompliance();
    if (view === 'lotification') return renderLotification();
    if (view === 'master') return renderMaster();
    if (view === 'config') return renderConfig();
    if (view === 'documents') return renderDocuments();
    if (view === 'admin') return renderAdmin();
    return renderDashboard();
  }

  function renderDashboard() {
    const range = currentWeekRange();
    const weekPlan = state.programming.filter(r => !r.annulled && isWithin(r.date, range.start, range.end));
    const weekProd = state.production.filter(r => !r.annulled && isWithin(r.date, range.start, range.end));
    const planSum = weekPlan.reduce((a,b)=>a + num(b.programmed), 0);
    const prodSum = weekProd.reduce((a,b)=>a + num(b.total), 0);
    const comp = planSum ? Math.min(100, (prodSum / planSum) * 100) : 0;
    const lastProd = [...weekProd].sort((a,b)=>String(b.date).localeCompare(String(a.date)))[0];
    const lastPlan = [...weekPlan].sort((a,b)=>String(b.date).localeCompare(String(a.date)))[0];
    return `
      ${dashboardCards(planSum, prodSum, comp, weekPlan.length, weekProd.length)}
      <div class="split-grid">
        <section class="panel panel-compact">
          <div class="section-head">
            <div>
              <h2>Transacciones SAP</h2>
              <p>Acceso directo a las pantallas funcionales.</p>
            </div>
          </div>
          <div class="toolbar">
            <button class="btn btn-primary" data-nav="programming" type="button">ZP001 Programación</button>
            <button class="btn btn-primary" data-nav="production" type="button">ZD001 Producción</button>
            <button class="btn btn-secondary" data-nav="compliance" type="button">ZC001 Cumplimiento</button>
            <button class="btn btn-secondary" data-nav="lotification" type="button">ZL001 Lotes</button>
            <button class="btn btn-secondary" data-nav="master" type="button">ZM001 Maestro</button>
          </div>
          <div style="margin-top:14px" class="note">
            Usa el campo <strong>OK Code</strong> para saltar de transacción igual que en SAP. La descripción del producto permite escribir libremente y también seleccionar sugerencias sin bloquearte.
          </div>
        </section>
        <section class="panel panel-compact">
          <div class="section-head">
            <div>
              <h2>Últimos movimientos</h2>
              <p>Lo más reciente de programación y producción.</p>
            </div>
          </div>
          <div class="timeline">
            <div class="timeline-item">
              <strong>Programación</strong>
              <div class="small">${lastPlan ? `${dateLabel(lastPlan.date)} · ${lastPlan.codigo || ''} · ${lastPlan.descripcion || ''}` : 'Sin registros esta semana'}</div>
            </div>
            <div class="timeline-item">
              <strong>Producción</strong>
              <div class="small">${lastProd ? `${dateLabel(lastProd.date)} · ${lastProd.codigo || ''} · ${lastProd.descripcion || ''}` : 'Sin registros esta semana'}</div>
            </div>
            <div class="timeline-item">
              <strong>Semana activa</strong>
              <div class="small">${longWeekLabel(range)}</div>
            </div>
          </div>
        </section>
      </div>
    `;
  }

  function dashboardCards(plan, prod, comp, nPlan, nProd) {
    const delta = prod - plan;
    return `
      <div class="card-grid">
        <div class="card-mini"><div class="label">Programado</div><div class="value">${fmtNumber(plan)}</div><div class="hint">${nPlan} filas activas</div></div>
        <div class="card-mini"><div class="label">Producido</div><div class="value">${fmtNumber(prod)}</div><div class="hint">${nProd} filas activas</div></div>
        <div class="card-mini"><div class="label">Diferencia</div><div class="value ${delta >= 0 ? 'kpi-good' : 'kpi-bad'}">${fmtSigned(delta)}</div><div class="hint">vs. programación</div></div>
        <div class="card-mini"><div class="label">Cumplimiento</div><div class="value ${comp >= 100 ? 'kpi-good' : comp >= 85 ? 'kpi-warn' : 'kpi-bad'}">${fmtPct(comp)}</div><div class="hint">semana activa</div></div>
      </div>
    `;
  }

  function renderProgramming() {
    const rows = state.programming;
    return `
      <div class="section-head">
        <div>
          <h2>Programación semanal</h2>
          <p>Captura por transacción. Escribes, presionas Enter y pasa al siguiente campo.</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="new-programming-row" type="button">Agregar fila</button>
          <button class="btn" data-action="open-modal-import" data-kind="programming" type="button">Pegar desde Excel</button>
          <button class="btn" data-action="clear-programming" type="button">Vaciar</button>
          <button class="btn btn-secondary" data-action="export-csv" type="button">Exportar CSV</button>
          <button class="btn btn-secondary" data-action="print-report" type="button">Imprimir</button>
        </div>
      </div>
      <div class="field-row no-print" style="max-width:420px">
        <label>Buscar en programación</label>
        <input class="input" data-filter-table="programming" type="text" placeholder="Código, descripción, área o tipo..." value="${esc(state.ui.filters?.programming || '')}">
      </div>
      <div class="note" style="margin-bottom:10px">La programación se guarda por fecha, código SAP y área. Si el producto no existe en el maestro, igualmente puedes escribirlo y luego corregirlo.</div>
      <div class="table-wrap no-print">
        <table data-table="programming">
          <thead><tr>
            <th>Fecha</th><th>Semana</th><th>Código SAP</th><th>Presentación</th><th>Área</th><th>Programado</th><th>Cantidad ME</th><th>Tipo</th><th></th>
          </tr></thead>
          <tbody>${rows.map(r => programmingRowHTML(r)).join('')}</tbody>
        </table>
      </div>
      <div class="report-frame print-only" style="display:none"></div>
    `;
  }

  function programmingRowHTML(r) {
    return `
      <tr class="${r.annulled ? 'row-annulled' : ''}" data-row="programming" data-id="${r.id}">
        <td><input class="input" data-bind="programming" data-row="${r.id}" data-field="date" type="date" value="${esc(r.date || '')}"></td>
        <td><input class="input" data-bind="programming" data-row="${r.id}" data-field="week" type="number" min="1" max="53" value="${esc(r.week ?? '')}"></td>
        <td><input class="input" data-bind="programming" data-row="${r.id}" data-field="codigo" type="text" value="${esc(r.codigo || '')}" data-picker="prog_desc_${r.id}"></td>
        <td><input class="input" id="prog_desc_${r.id}" data-bind="programming" data-row="${r.id}" data-field="descripcion" type="text" value="${esc(r.descripcion || '')}" autocomplete="off"></td>
        <td>
          <select class="input" data-bind="programming" data-row="${r.id}" data-field="area">
            ${selectOptions(['PAQUETERIA','M1','M2','SORTEX B'], r.area || '')}
          </select>
        </td>
        <td><input class="input" data-bind="programming" data-row="${r.id}" data-field="programmed" type="number" min="0" step="1" value="${esc(r.programmed ?? 0)}"></td>
        <td><input class="input" data-bind="programming" data-row="${r.id}" data-field="cantidad_me" type="number" min="0" step="1" value="${esc(r.cantidad_me ?? 0)}"></td>
        <td>
          <select class="input" data-bind="programming" data-row="${r.id}" data-field="type">
            ${selectOptions(['BOBINA IMPRESA','SACO QQ 100 LBS','SAQUITO 25 LBS','SUBPRODUCTO'], r.type || '')}
          </select>
        </td>
        <td class="table-actions">
          <button class="btn" data-row-action="copy" data-bind="programming" data-id="${r.id}" type="button">Copiar</button>
          <button class="btn btn-secondary" data-row-action="delete" data-bind="programming" data-id="${r.id}" type="button">Eliminar</button>
        </td>
      </tr>`;
  }

  function renderProduction() {
    const rows = state.production;
    return `
      <div class="section-head">
        <div>
          <h2>Producción diaria</h2>
          <p>Captura la producción real por día, turno, producto y observación.</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="new-production-row" type="button">Agregar fila</button>
          <button class="btn" data-action="open-modal-import" data-kind="production" type="button">Pegar desde Excel</button>
          <button class="btn" data-action="clear-production" type="button">Vaciar</button>
          <button class="btn btn-secondary" data-action="print-report" type="button">Imprimir</button>
        </div>
      </div>
      <div class="field-row no-print" style="max-width:420px">
        <label>Buscar en producción</label>
        <input class="input" data-filter-table="production" type="text" placeholder="Código, descripción, observación o turno..." value="${esc(state.ui.filters?.production || '')}">
      </div>
      <div class="note" style="margin-bottom:10px">Aquí ingresas lo producido realmente por día y turno. El sistema lo compara automáticamente contra la programación semanal.</div>
      <div class="table-wrap no-print">
        <table data-table="production">
          <thead><tr>
            <th>Fecha</th><th>Día</th><th>Turno</th><th>Código SAP</th><th>Presentación</th><th>Área</th><th>Total</th><th>Observación</th><th>Acción</th>
          </tr></thead>
          <tbody>${rows.map(r => productionRowHTML(r)).join('')}</tbody>
        </table>
      </div>
      <div class="report-frame print-only" style="display:none"></div>
    `;
  }

  function productionRowHTML(r) {
    return `
      <tr class="${r.annulled ? 'row-annulled' : ''}" data-row="production" data-id="${r.id}">
        <td><input class="input" data-bind="production" data-row="${r.id}" data-field="date" type="date" value="${esc(r.date || '')}"></td>
        <td>
          <select class="input" data-bind="production" data-row="${r.id}" data-field="dayName">
            ${selectOptions(dayNames, r.dayName || dayNameOf(r.date || state.ui.date))}
          </select>
        </td>
        <td>
          <select class="input" data-bind="production" data-row="${r.id}" data-field="shift">
            ${selectOptions(['A','B','C'], r.shift || 'A')}
          </select>
        </td>
        <td><input class="input" data-bind="production" data-row="${r.id}" data-field="codigo" type="text" value="${esc(r.codigo || '')}" data-picker="prod_desc_${r.id}"></td>
        <td><input class="input" id="prod_desc_${r.id}" data-bind="production" data-row="${r.id}" data-field="descripcion" type="text" value="${esc(r.descripcion || '')}" autocomplete="off"></td>
        <td>
          <select class="input" data-bind="production" data-row="${r.id}" data-field="area">
            ${selectOptions(['PAQUETERIA','M1','M2','SORTEX B'], r.area || '')}
          </select>
        </td>
        <td><input class="input" data-bind="production" data-row="${r.id}" data-field="total" type="number" min="0" step="1" value="${esc(r.total ?? 0)}"></td>
        <td><input class="input" data-bind="production" data-row="${r.id}" data-field="note" type="text" value="${esc(r.note || '')}" placeholder="Motivo, paro, evento..."></td>
        <td class="table-actions">
          <button class="btn" data-row-action="annul" data-bind="production" data-id="${r.id}" type="button">${r.annulled ? 'Reactivar' : 'Anular'}</button>
          <button class="btn btn-secondary" data-row-action="delete" data-bind="production" data-id="${r.id}" type="button">Eliminar</button>
        </td>
      </tr>`;
  }

  function renderCompliance() {
    const range = currentWeekRange();
    const data = complianceData(range);
    const totals = data.reduce((acc, r) => {
      acc.programmed += r.programmed;
      acc.produced += r.total;
      return acc;
    }, { programmed: 0, produced: 0 });
    const comp = totals.programmed ? Math.min(100, totals.produced / totals.programmed * 100) : 0;
    const days = rangeDays(range);
    return `
      <div class="section-head">
        <div>
          <h2>Control semanal SAP · Producción vs Programación</h2>
          <p>${longWeekLabel(range)}</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="print-report" type="button">Imprimir hoja</button>
          <button class="btn" data-action="export-report" type="button">Exportar CSV</button>
          <button class="btn btn-secondary" data-nav="programming" type="button">Ir a programación</button>
          <button class="btn btn-secondary" data-nav="production" type="button">Ir a producción</button>
        </div>
      </div>
      <div class="report-summary">
        <div class="mini"><span>Programado</span><strong>${fmtNumber(totals.programmed)}</strong></div>
        <div class="mini"><span>Producido</span><strong>${fmtNumber(totals.produced)}</strong></div>
        <div class="mini"><span>Diferencia</span><strong class="${totals.produced - totals.programmed >= 0 ? 'kpi-good' : 'kpi-bad'}">${fmtSigned(totals.produced - totals.programmed)}</strong></div>
        <div class="mini"><span>Cumplimiento</span><strong class="${comp >= 100 ? 'kpi-good' : comp >= 85 ? 'kpi-warn' : 'kpi-bad'}">${fmtPct(comp)}</strong></div>
      </div>
      <div class="report-frame">
        <div class="report-head">
          <h3>Programa de Cumplimiento del Plan Semanal · Semana ${state.ui.week}</h3>
          <div class="sub">Del ${spanishDate(range.start)} al ${spanishDate(range.end)}</div>
          <div class="sub">Actualizado por última vez · ${longNow()}</div>
          <div class="sub">Responsable · ${esc(state.ui.signature || 'Sin firma')}</div>
        </div>
        <div class="table-wrap" style="border:0;border-top:1px solid #b9c6d3">
          <table class="report-table">
            <thead>
              <tr>
                <th>Código SAP</th>
                <th>Área</th>
                <th>Presentación</th>
                <th>Programación</th>
                ${days.map(d => `<th>${d.label}<br><span class="small">${shortDate(d.date)}</span></th>`).join('')}
                <th>Total</th><th>Cumplimiento</th><th>Diferencia</th><th>SC / IC</th>
              </tr>
            </thead>
            <tbody>
              ${data.map(r => complianceRowHTML(r, days)).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function complianceRowHTML(r, days) {
    const total = r.total;
    const comp = r.programmed ? Math.min(100, total / r.programmed * 100) : 0;
    const ratio = r.programmed ? total / r.programmed * 100 : 0;
    return `
      <tr>
        <td><strong>${esc(r.codigo)}</strong></td>
        <td>${esc(r.area)}</td>
        <td>${esc(r.descripcion)}</td>
        <td>${fmtNumber(r.programmed)}</td>
        ${days.map(d => `<td>${r.byDay[d.key] ? fmtNumber(r.byDay[d.key]) : ''}</td>`).join('')}
        <td><strong>${fmtNumber(total)}</strong></td>
        <td class="${comp >= 100 ? 'kpi-good' : comp >= 85 ? 'kpi-warn' : 'kpi-bad'}">${fmtPct(comp)}</td>
        <td>${fmtSigned(total - r.programmed)}</td>
        <td>${fmtPct(ratio)}</td>
      </tr>
    `;
  }

  function renderLotification() {
    const rows = lotificationRows();
    return `
      <div class="section-head">
        <div>
          <h2>Control de lotificación</h2>
          <p>Mes actual, familia, secuencia y lote disponible según acumulado.</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="print-report" type="button">Imprimir</button>
          <button class="btn btn-secondary" data-nav="production" type="button">Ver producción</button>
          <button class="btn btn-secondary" data-nav="master" type="button">Ver maestro</button>
        </div>
      </div>
      <div class="field-row inline no-print">
        <div>
          <label>Umbral por lote</label>
          <input class="input" id="lotThreshold" type="number" min="1" value="${state.ui.lotThreshold}">
        </div>
        <div>
          <label>Mes activo</label>
          <input class="input" type="text" value="${monthNames[state.ui.date ? new Date(state.ui.date).getMonth() : new Date().getMonth()] || 'MARZO'}" readonly>
        </div>
        <div>
          <label>Año activo</label>
          <input class="input" type="text" value="${new Date(state.ui.date || today).getFullYear()}" readonly>
        </div>
      </div>
      <div class="report-frame">
        <div class="report-head">
          <h3>Control de Validación - Lotificaciones en Productos Terminados & Subproductos en Mes Actual</h3>
          <div class="sub">${new Date(state.ui.date || today).getFullYear()} · ${monthNames[new Date(state.ui.date || today).getMonth()]} · Semanas: ${rangeWeeksText()}</div>
          <div class="sub">Actualizado por última vez · ${longNow()}</div>
          <div class="sub">Responsable · ${esc(state.ui.signature || 'Sin firma')}</div>
        </div>
        <div class="table-wrap" style="border:0;border-top:1px solid #b9c6d3">
          <table class="report-table">
            <thead>
              <tr>
                <th>Código SAP</th><th>Presentación</th><th>Área</th><th>Familia</th><th>Prefijo</th><th>Producción mes</th><th>Secuencia</th><th>Lote actual</th><th>Disponible</th><th>Libre utilización</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map(r => `
                <tr>
                  <td><strong>${esc(r.codigo)}</strong></td>
                  <td>${esc(r.descripcion)}</td>
                  <td>${esc(r.area)}</td>
                  <td>${esc(r.family)}</td>
                  <td>${esc(r.prefix)}</td>
                  <td>${fmtNumber(r.monthQty)}</td>
                  <td>${r.sequence}</td>
                  <td>${r.lotCode}</td>
                  <td>${fmtNumber(r.available)}</td>
                  <td><strong>${r.freeUse}</strong></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderMaster() {
    return `
      <div class="section-head">
        <div>
          <h2>Maestro de materiales</h2>
          <p>Búsqueda libre, edición rápida y parámetros por producto.</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="add-master" type="button">Agregar material</button>
          <button class="btn btn-secondary" data-action="save-master" type="button">Guardar maestro</button>
          <button class="btn btn-secondary" data-action="restore-master" type="button">Restaurar maestro</button>
          <button class="btn" data-action="open-modal-import" data-kind="master" type="button">Importar</button>
        </div>
      </div>
      <div class="field-row no-print" style="max-width:420px">
        <label>Buscar material</label>
        <input class="input" data-search-master="master" type="text" placeholder="Código, descripción, familia, área..." value="${esc(state.ui.filters?.master || '')}">
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Código</th><th>Descripción</th><th>Peso</th><th>Unidad</th><th>Familia</th><th>Proceso</th><th>Ent.</th><th>Micrón/Max</th><th>Pulido</th><th>Área</th><th>Tipo</th><th>Tarima</th><th>Bobina U</th><th>Bobina R</th><th>Granza</th><th></th>
          </tr></thead>
          <tbody>
            ${state.master.map(r => masterRowHTML(r)).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function masterRowHTML(r) {
    return `
      <tr data-row="master" data-id="${r.id}">
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="codigo" value="${esc(r.codigo)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="descripcion" value="${esc(r.descripcion)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="peso" type="number" min="0" step="0.01" value="${esc(r.peso)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="unidad" value="${esc(r.unidad)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="family" value="${esc(r.family)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="proceso" value="${esc(r.proceso)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="entero_min" type="number" min="0" step="0.01" value="${esc(r.entero_min)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="mignon_max" type="text" value="${esc(r.mignon_max)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="pulido_min" type="text" value="${esc(r.pulido_min)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="area" value="${esc(r.area)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="tipo" value="${esc(r.tipo)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="tarima_div" type="number" min="0" step="1" value="${esc(r.tarima_div)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="bobina_unidades" type="number" min="0" step="1" value="${esc(r.bobina_unidades)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="bobina_rend" type="number" min="0" step="0.01" value="${esc(r.bobina_rend)}"></td>
        <td><input class="input" data-bind="master" data-row="${r.id}" data-field="granza_factor" type="number" min="0" step="0.0001" value="${esc(r.granza_factor)}"></td>
        <td class="table-actions">
          <button class="btn btn-secondary" data-row-action="delete" data-bind="master" data-id="${r.id}" type="button">Eliminar</button>
        </td>
      </tr>
    `;
  }

  function renderConfig() {
    return `
      <div class="section-head">
        <div>
          <h2>Configuración SAP</h2>
          <p>Tema visual, densidad, umbrales y tasas de cálculo.</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="save-settings" type="button">Guardar</button>
          <button class="btn btn-secondary" data-action="reset-ui" type="button">Restablecer</button>
        </div>
      </div>
      <div class="split-grid">
        <section class="panel panel-compact">
          <div class="panel-title">Tema visual</div>
          <div class="field-row">
            ${themeOption('quartz', 'Quartz Light')}
            ${themeOption('quartz-dark', 'Quartz Dark')}
            ${themeOption('belize', 'Belize Light')}
            ${themeOption('signature', 'Signature Classic')}
          </div>
        </section>
        <section class="panel panel-compact">
          <div class="panel-title">Parámetros operativos</div>
          <div class="field-row">
            <label>Umbral de lote</label>
            <input class="input" id="settingLot" type="number" min="1" value="${state.ui.lotThreshold}">
          </div>
          <div class="field-row inline">
            <div>
              <label>Tasa paquetería (u/hr)</label>
              <input class="input" id="settingPackRate" type="number" min="1" step="1" value="${state.settings.packRate}">
            </div>
            <div>
              <label>Tasa molinos (qq/hr)</label>
              <input class="input" id="settingMillRate" type="number" min="1" step="1" value="${state.settings.millingRate}">
            </div>
            <div>
              <label>Envas por defecto</label>
              <input class="input" id="settingEnvas" type="number" min="1" step="1" value="${state.settings.defaultEnvas}">
            </div>
          </div>
          <div class="field-row">
            <label>Tarima por defecto</label>
            <input class="input" id="settingTarima" type="number" min="1" step="1" value="${state.settings.defaultTarima}">
          </div>
          <div class="field-row">
            <label>Menú rápido</label>
            <select class="input" id="quickArea">
              ${selectOptions(['PAQUETERIA','M1','M2','SORTEX B'], state.ui.quickArea || 'PAQUETERIA')}
            </select>
          </div>
        </section>
      </div>
      <section class="panel panel-compact" style="margin-top:14px">
        <div class="panel-title">Documentos / accesos</div>
        <div class="toolbar">
          <button class="btn btn-primary" data-nav="compliance" type="button">Hoja semanal</button>
          <button class="btn btn-primary" data-nav="lotification" type="button">Lotes</button>
          <button class="btn btn-primary" data-nav="production" type="button">Producción</button>
          <button class="btn btn-primary" data-nav="programming" type="button">Programación</button>
        </div>
      </section>
    `;
  }

  function renderDocuments() {
    const range = currentWeekRange();
    return `
      <div class="section-head">
        <div>
          <h2>Documentos y reportes</h2>
          <p>Salida rápida de impresión y trazabilidad por módulo.</p>
        </div>
        <div class="toolbar no-print">
          <button class="btn btn-primary" data-action="print-report" type="button">Imprimir vista actual</button>
          <button class="btn" data-nav="compliance" type="button">Ir a cumplimiento</button>
          <button class="btn" data-nav="production" type="button">Ir a producción</button>
        </div>
      </div>
      <div class="split-grid">
        <section class="panel panel-compact">
          <div class="panel-title">Reportes disponibles</div>
          <div class="timeline">
            <div class="timeline-item"><strong>Control semanal</strong><div class="small">${longWeekLabel(range)}</div></div>
            <div class="timeline-item"><strong>Producción diaria</strong><div class="small">Registros de turnos y observaciones.</div></div>
            <div class="timeline-item"><strong>Lotificación</strong><div class="small">Lotes por familia y mes activo.</div></div>
          </div>
        </section>
        <section class="panel panel-compact">
          <div class="panel-title">Histórico reciente</div>
          <div class="timeline">${recentHistoryHTML()}</div>
        </section>
      </div>
    `;
  }

  function recentHistoryHTML() {
    const rows = [...state.production].filter(r => !r.annulled).sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0, 6);
    if (!rows.length) return '<div class="timeline-item"><strong>Sin producción todavía</strong><div class="small">Carga registros para ver el histórico.</div></div>';
    return rows.map(r => `<div class="timeline-item"><strong>${esc(dateLabel(r.date))} · ${esc(r.shift || '')}</strong><div class="small">${esc(r.codigo || '')} · ${esc(r.descripcion || '')} · ${fmtNumber(r.total)}</div></div>`).join('');
  }

  function themeOption(value, label) {
    return `
      <label class="timeline-item" style="display:flex;align-items:center;gap:10px;cursor:pointer">
        <input type="radio" name="theme" value="${value}" ${state.ui.theme === value ? 'checked' : ''}>
        <span><strong>${label}</strong><div class="small">Tema visual SAP</div></span>
      </label>
    `;
  }



function renderAuth() {
  return `
    <div class="auth-layout">
      <section class="panel auth-card">
        <div class="eyebrow">Acceso protegido</div>
        <h2>Iniciar sesión</h2>
        <p class="muted">Ingresa con tu usuario y contraseña para acceder a los módulos habilitados por rol.</p>
        <div class="field-row">
          <label for="authUser">Usuario</label>
          <input id="authUser" class="input" type="text" placeholder="admin" autocomplete="username">
        </div>
        <div class="field-row">
          <label for="authPass">Contraseña</label>
          <input id="authPass" class="input" type="password" placeholder="••••••••" autocomplete="current-password">
        </div>
        <div class="toolbar">
          <button class="btn btn-primary" data-action="login-submit" type="button">Entrar</button>
          <button class="btn btn-secondary" data-action="demo-login" type="button">Usar demo admin</button>
        </div>
        <div class="note auth-note">${esc(state.auth.error || 'La sesión se guarda localmente en este navegador.')}</div>
      </section>
      <section class="panel auth-side">
        <div class="panel-title">Credenciales demo</div>
        <div class="timeline">
          <div class="timeline-item"><strong>Admin</strong><div class="small">Usuario: admin · Clave: Admin123!</div></div>
          <div class="timeline-item"><strong>Planeación</strong><div class="small">Usuario: planner · Clave: Plan123!</div></div>
          <div class="timeline-item"><strong>Producción</strong><div class="small">Usuario: produccion · Clave: Prod123!</div></div>
          <div class="timeline-item"><strong>Consulta</strong><div class="small">Usuario: consulta · Clave: Read123!</div></div>
        </div>
        <div class="note" style="margin-top:12px">El sistema controla acceso por rol y transacción. Los módulos no autorizados no se muestran.</div>
      </section>
    </div>
  `;
}

function renderAdmin() {
  const userRows = state.auth.users.map(u => `
    <tr data-auth-id="${esc(u.id)}">
      <td><input class="input" data-auth-bind="users" data-row="${esc(u.id)}" data-field="username" value="${esc(u.username)}"></td>
      <td><input class="input" data-auth-bind="users" data-row="${esc(u.id)}" data-field="fullName" value="${esc(u.fullName)}"></td>
      <td>
        <select class="input" data-auth-bind="users" data-row="${esc(u.id)}" data-field="role">
          ${selectOptions(state.auth.roles.map(r => r.code), u.role)}
        </select>
      </td>
      <td><input class="input" data-auth-bind="users" data-row="${esc(u.id)}" data-field="password" value="${esc(u.password)}"></td>
      <td style="text-align:center"><input type="checkbox" data-auth-bind="users" data-row="${esc(u.id)}" data-field="active" ${u.active !== false ? 'checked' : ''}></td>
      <td>
        <div class="row-actions">
          <button class="btn" data-action="delete-auth-row" data-bind="users" data-id="${esc(u.id)}" type="button">Eliminar</button>
        </div>
      </td>
    </tr>
  `).join('');

  const roleRows = state.auth.roles.map(r => `
    <tr data-auth-role="${esc(r.code)}">
      <td><input class="input" data-auth-bind="roles" data-row="${esc(r.code)}" data-field="code" value="${esc(r.code)}"></td>
      <td><input class="input" data-auth-bind="roles" data-row="${esc(r.code)}" data-field="name" value="${esc(r.name)}"></td>
      <td><textarea class="input" data-auth-bind="roles" data-row="${esc(r.code)}" data-field="permissions" style="min-height:68px">${esc((r.permissions || []).join(', '))}</textarea></td>
      <td>
        <div class="row-actions">
          <button class="btn" data-action="delete-auth-row" data-bind="roles" data-id="${esc(r.code)}" type="button">Eliminar</button>
        </div>
      </td>
    </tr>
  `).join('');

  const auditRows = (state.auth.audit || []).slice(0, 12).map(a => `
    <div class="timeline-item">
      <strong>${esc(a.action)}</strong>
      <div class="small">${esc(a.user || 'sistema')} · ${esc(a.at || '')}</div>
      <div class="small">${esc(a.detail || '')}</div>
    </div>
  `).join('') || '<div class="timeline-item"><strong>Sin eventos</strong><div class="small">Los accesos y cambios aparecerán aquí.</div></div>';

  return `
    <div class="section-head">
      <div>
        <h2>Usuarios y autorizaciones</h2>
        <p>Controla acceso por usuario, rol y transacción.</p>
      </div>
      <div class="toolbar no-print">
        <button class="btn btn-primary" data-action="add-auth-user" type="button">Agregar usuario</button>
        <button class="btn btn-primary" data-action="add-auth-role" type="button">Agregar rol</button>
        <button class="btn btn-secondary" data-action="save-auth" type="button">Guardar seguridad</button>
        <button class="btn" data-action="seed-auth" type="button">Restaurar demo</button>
      </div>
    </div>

    <div class="split-grid">
      <section class="panel panel-compact">
        <div class="section-head">
          <div>
            <h2>Usuarios</h2>
            <p>Usuario, nombre, contraseña, rol y estado activo.</p>
          </div>
        </div>
        <div class="field-row no-print" style="max-width:360px">
          <label>Buscar usuarios</label>
          <input class="input" data-search-table="users" type="text" placeholder="Filtra por usuario, nombre o rol...">
        </div>
        <div class="table-wrap">
          <table data-table="users">
            <thead><tr><th>Usuario</th><th>Nombre</th><th>Rol</th><th>Contraseña</th><th>Activo</th><th></th></tr></thead>
            <tbody>${userRows}</tbody>
          </table>
        </div>
      </section>

      <section class="panel panel-compact">
        <div class="section-head">
          <div>
            <h2>Roles</h2>
            <p>Permisos separados por coma. Ejemplo: dashboard, programming, ZP001.</p>
          </div>
        </div>
        <div class="field-row no-print" style="max-width:360px">
          <label>Buscar roles</label>
          <input class="input" data-search-table="roles" type="text" placeholder="Filtra por código o nombre...">
        </div>
        <div class="table-wrap">
          <table data-table="roles">
            <thead><tr><th>Código</th><th>Nombre</th><th>Permisos</th><th></th></tr></thead>
            <tbody>${roleRows}</tbody>
          </table>
        </div>
      </section>
    </div>

    <div class="split-grid" style="margin-top:14px">
      <section class="panel panel-compact">
        <div class="section-head"><div><h2>Transacciones disponibles</h2><p>Catálogo visible según permisos.</p></div></div>
        <div class="permissions-grid">
          ${NAV.map(group => `
            <div class="permission-card">
              <strong>${esc(group.group)}</strong>
              <div class="small">${group.items.map(i => `${esc(i.code)} · ${esc(i.name)}`).join('<br>')}</div>
            </div>
          `).join('')}
        </div>
      </section>
      <section class="panel panel-compact">
        <div class="section-head"><div><h2>Auditoría reciente</h2><p>Accesos, cierres de sesión y cambios de seguridad.</p></div></div>
        <div class="timeline">${auditRows}</div>
      </section>
    </div>
  `;
}
  function bindView(view) {
    applyViewFilters(view);
    if (view === 'config') {
      $$('input[name="theme"]').forEach(r => r.addEventListener('change', () => setTheme(r.value)));
      $('#settingLot').addEventListener('input', e => state.ui.lotThreshold = num(e.target.value) || 3000);
      $('#settingPackRate').addEventListener('input', e => state.settings.packRate = num(e.target.value) || 235);
      $('#settingMillRate').addEventListener('input', e => state.settings.millingRate = num(e.target.value) || 100);
      $('#settingEnvas').addEventListener('input', e => state.settings.defaultEnvas = num(e.target.value) || 2);
      $('#settingTarima').addEventListener('input', e => state.settings.defaultTarima = num(e.target.value) || 100);
      $('#quickArea').addEventListener('change', e => state.ui.quickArea = e.target.value);
    }
    if (view === 'lotification') {
      $('#lotThreshold').addEventListener('input', e => { state.ui.lotThreshold = num(e.target.value) || 3000; persistAll(); renderAll(); });
    }
    if (view === 'admin') {
      filterTableRows('users', '');
      filterTableRows('roles', '');
    }
  }

  function applyViewFilters(view) {
    if (view === 'master') filterMasterTable('master', state.ui.filters?.master || '');
    if (view === 'programming') filterTableRows('programming', state.ui.filters?.programming || '');
    if (view === 'production') filterTableRows('production', state.ui.filters?.production || '');
  }

  function addRow(kind) {
    if (kind === 'programming') {
      state.programming.unshift({ id: uid(), date: state.ui.date, week: state.ui.week, codigo: '', descripcion: '', area: state.ui.quickArea || 'PAQUETERIA', programmed: 0, cantidad_me: 0, type: 'BOBINA IMPRESA', annulled: false });
    } else if (kind === 'production') {
      state.production.unshift({ id: uid(), date: state.ui.date, dayName: dayNameOf(state.ui.date), shift: 'A', codigo: '', descripcion: '', area: state.ui.quickArea || 'PAQUETERIA', total: 0, note: '', annulled: false });
    } else if (kind === 'master') {
      addMasterRow();
      return;
    }
    persistAll();
    renderAll();
    const row = $(`[data-row="${kind}"]`);
    if (row) {
      const first = row.querySelector('input,select,textarea');
      if (first) first.focus();
    }
  }

  function addMasterRow() {
    state.master.unshift({
      id: uid(), codigo: '', descripcion: '', peso: 0, unidad: '', family: '', proceso: '', entero_min: '', mignon_max: '', pulido_min: '', area: '', tipo: '', tarima_div: state.settings.defaultTarima, bobina_unidades: 5, bobina_rend: 1, granza_factor: 1.4286
    });
    persistAll();
    renderAll();
  }

  function updateRow(bind, rowId, field, value) {
    const row = getRow(bind, rowId);
    if (!row) return;
    if (bind === 'master' && ['peso','entero_min','tarima_div','bobina_unidades','bobina_rend','granza_factor'].includes(field)) row[field] = num(value);
    else if (bind === 'programming' && ['week','programmed','cantidad_me'].includes(field)) row[field] = num(value);
    else if (bind === 'production' && ['total'].includes(field)) row[field] = num(value);
    else row[field] = value;
    if (bind === 'production' && field === 'date') row.dayName = dayNameOf(value);
    if (bind === 'programming' && field === 'date') row.week = isoWeek(new Date(value));
    if ((field === 'codigo' || field === 'descripcion') && value) {
      autoFillFromMaster(row, { sourceField: field, value });
    }
    persistSoon();
    scheduleRender();
  }

  function autoFillFromMaster(row, { sourceField, value }) {
    const found = sourceField === 'codigo' ? findMasterByCode(value) : findMasterByText(value);
    if (!found) return;
    row.codigo = found.codigo;
    row.descripcion = found.descripcion;
    if ('area' in row && !row.area) row.area = found.area || row.area;
    if (row.hasOwnProperty('programmed') && !row.programmed) row.programmed = 0;
    if (row.hasOwnProperty('total') && !row.total) row.total = 0;
  }

  function getRow(bind, id) {
    return state[bind].find(r => r.id === id);
  }

  function findMasterByCode(code) {
    const norm = normalize(code);
    return state.master.find(m => normalize(m.codigo) === norm);
  }

  function findMasterByText(text) {
    const q = normalize(text);
    if (!q) return null;
    const exact = state.master.find(m => normalize(m.descripcion) === q);
    if (exact) return exact;
    return state.master.find(m => normalize(m.descripcion).includes(q) || normalize(m.codigo).includes(q)) || null;
  }

  function normalizeMaster(list) {
    return list.map(m => ({
      id: m.id || uid(),
      codigo: m.codigo || '',
      descripcion: m.descripcion || '',
      peso: num(m.peso),
      unidad: m.unidad || '',
      family: m.family || m.familia || inferFamily(m),
      proceso: m.proceso || '',
      entero_min: m.entero_min ?? '',
      mignon_max: m.mignon_max ?? '',
      pulido_min: m.pulido_min ?? '',
      area: m.area || inferArea(m),
      tipo: m.tipo || 'PT',
      tarima_div: num(m.tarima_div) || defaultTarimaFor(m),
      bobina_unidades: num(m.bobina_unidades) || defaultBobinaUnits(m),
      bobina_rend: num(m.bobina_rend) || defaultBobinaRend(m),
      granza_factor: num(m.granza_factor) || 1.4286,
      lot_prefix: m.lot_prefix || inferLotPrefix(m)
    }));
  }

  function normalizeProgramming(list) {
    return list.map(r => ({
      id: r.id || uid(),
      date: r.date || defaultUI.date,
      week: num(r.week) || isoWeek(new Date(r.date || defaultUI.date)),
      codigo: r.codigo || '',
      descripcion: r.descripcion || '',
      area: r.area || '',
      programmed: num(r.programado ?? r.programmed),
      cantidad_me: num(r.cantidad_me),
      type: r.tipo || r.type || 'BOBINA IMPRESA',
      annulled: !!r.annulled
    }));
  }

  function normalizeProduction(list) {
    return list.map(r => ({
      id: r.id || uid(),
      date: r.date || r.fecha || defaultUI.date,
      dayName: r.dayName || r.dia || dayNameOf(r.date || r.fecha || defaultUI.date),
      shift: r.shift || r.turno || 'A',
      codigo: r.codigo || r.material || '',
      descripcion: r.descripcion || r.presentacion || '',
      area: r.area || '',
      total: num(r.total ?? r.cantidad),
      note: r.note || r.observacion || '',
      annulled: !!r.annulled
    }));
  }

  function inferArea(m) {
    const d = normalize(m.descripcion || '');
    if (d.includes('PAQUETERIA') || d.includes('1 LBS') || d.includes('1.5 KGS') || d.includes('350 GRS') || d.includes('175 GRS') || d.includes('4 LBS')) return 'PAQUETERIA';
    if (d.includes('ESCALD')) return 'M2';
    if (d.includes('SEMOLINA')) return 'M1';
    return 'M1';
  }

  function inferFamily(m) {
    const d = normalize(m.descripcion || '');
    if (d.includes('SEMOLINA ESCALD')) return 'SE';
    if (d.includes('SEMOLINA')) return 'SB';
    if (d.includes('MIGA MEMBRET')) return 'PB';
    if (d.includes('MIGON ESCALD')) return 'ME';
    if (d.includes('MIGON BLANC')) return 'MB';
    if (d.includes('REPASO')) return 'RT';
    if (d.includes('INTEGRAL')) return 'AI';
    if (d.includes('ESCALD')) return 'AE';
    return 'AB';
  }

  function inferLotPrefix(m) { return inferFamily(m); }
  function defaultTarimaFor(m) {
    const d = normalize(m.descripcion || '');
    if (d.includes('1 LBS')) return 100;
    if (d.includes('175 GRS') || d.includes('350 GRS') || d.includes('1.5 KGS') || d.includes('1.75 KGS')) return 100;
    if (d.includes('25 LBS')) return 100;
    return 100;
  }
  function defaultBobinaUnits(m) {
    const d = normalize(m.descripcion || '');
    if (d.includes('PAQUETERIA') || d.includes('1 LBS') || d.includes('1.5 KGS') || d.includes('1.75 KGS') || d.includes('350 GRS') || d.includes('4 LBS') || d.includes('175 GRS')) return 5;
    return 0;
  }
  function defaultBobinaRend(m) {
    const d = normalize(m.descripcion || '');
    if (d.includes('PAQUETERIA') || d.includes('1 LBS') || d.includes('1.5 KGS') || d.includes('1.75 KGS') || d.includes('350 GRS') || d.includes('4 LBS') || d.includes('175 GRS')) return 1;
    return 0;
  }

  function currentWeekRange() {
    const d = new Date(state.ui.date || today);
    const start = mondayOf(d);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    return { start, end };
  }

  function rangeDays(range) {
    return Array.from({length: 7}, (_,i) => {
      const d = new Date(range.start);
      d.setDate(range.start.getDate() + i);
      return { date: d, key: toISO(d), label: dayNames[d.getDay()] };
    });
  }

  function longWeekLabel(range) {
    return `Semana ${state.ui.week} | Del ${spanishDate(range.start)} al ${spanishDate(range.end)}`;
  }

  function rangeWeeksText() {
    const w = state.ui.week || isoWeek(new Date(state.ui.date || today));
    return `${Math.max(1, w-3)} - ${Math.max(1, w+1)}`;
  }

  function complianceData(range) {
    const days = rangeDays(range);
    const plan = state.programming.filter(r => !r.annulled && isWithin(r.date, range.start, range.end));
    const prod = state.production.filter(r => !r.annulled && isWithin(r.date, range.start, range.end));
    const keys = new Map();
    for (const r of plan) {
      const key = compositeKey(r.codigo, r.descripcion, r.area);
      if (!keys.has(key)) keys.set(key, { codigo: r.codigo || '', descripcion: r.descripcion || '', area: r.area || '', programmed: 0, total: 0, byDay: Object.fromEntries(days.map(d=>[d.key,0])) });
      keys.get(key).programmed += num(r.programmed);
    }
    for (const r of prod) {
      const key = compositeKey(r.codigo, r.descripcion, r.area);
      if (!keys.has(key)) keys.set(key, { codigo: r.codigo || '', descripcion: r.descripcion || '', area: r.area || '', programmed: 0, total: 0, byDay: Object.fromEntries(days.map(d=>[d.key,0])) });
      const row = keys.get(key);
      row.total += num(r.total);
      if (row.byDay[r.date] == null) row.byDay[r.date] = 0;
      row.byDay[r.date] += num(r.total);
    }
    return [...keys.values()].sort((a,b) => (a.area + a.codigo).localeCompare(b.area + b.codigo));
  }

  function lotificationRows() {
    const d = new Date(state.ui.date || today);
    const year = String(d.getFullYear()).slice(-2);
    const month = pad2(d.getMonth() + 1);
    const monthKey = `${d.getFullYear()}-${month}`;
    const monthRows = state.production.filter(r => !r.annulled && String(r.date || '').startsWith(monthKey));
    const grouped = new Map();
    for (const row of monthRows) {
      const m = findMasterByCode(row.codigo) || findMasterByText(row.descripcion) || {};
      const prefix = m.lot_prefix || inferLotPrefix(m);
      const key = compositeKey(row.codigo || '', row.descripcion || '', row.area || '');
      if (!grouped.has(key)) grouped.set(key, { codigo: row.codigo || '', descripcion: row.descripcion || '', area: row.area || '', family: m.family || inferFamily(m), prefix, qty: 0, master: m });
      grouped.get(key).qty += num(row.total);
    }
    const threshold = num(state.ui.lotThreshold) || 3000;
    return [...grouped.values()].sort((a,b)=>a.codigo.localeCompare(b.codigo)).map(item => {
      const sequence = Math.max(1, Math.ceil(item.qty / threshold));
      return {
        ...item,
        monthQty: item.qty,
        sequence,
        lotCode: `${item.prefix}-${year}-${month}-${pad2(sequence)}`,
        available: Math.max(0, threshold - (item.qty % threshold || threshold)),
        freeUse: `${item.prefix}-${year}-${month}-${pad2(sequence)}`
      };
    });
  }

  function scheduleRender() {
    if (state._renderReq) return;
    state._renderReq = requestAnimationFrame(() => {
      state._renderReq = null;
      renderAll();
      updateStatusBar();
    });
  }

  function filterMasterTable(view, query) {
    const q = normalize(query || '');
    $$('#screenHost tbody tr').forEach(tr => {
      const txt = normalize(tr.textContent);
      tr.style.display = !q || txt.includes(q) ? '' : 'none';
    });
  }

  function filterTableRows(view, query) {
    const q = normalize(query || '');
    const root = $(`#screenHost table[data-table="${view}"] tbody`);
    if (!root) return;
    root.querySelectorAll('tr').forEach(tr => {
      const txt = normalize(tr.textContent);
      tr.style.display = !q || txt.includes(q) ? '' : 'none';
    });
  }

  function openPicker(input, query='') {
    const rect = input.getBoundingClientRect();
    const items = searchProducts(query || input.value || '');
    pickerState = { open: true, input, target: input, items, index: items.length ? 0 : -1 };
    renderPicker();
    positionPicker(input);
    refs.pickerPanel.hidden = false;
    refs.pickerPanel.style.display = 'block';
  }

  function renderPicker() {
    if (!pickerState.open) return;
    refs.pickerPanel.innerHTML = pickerState.items.length ? pickerState.items.slice(0, 10).map((item, i) => `
      <div class="suggestion-item ${i === pickerState.index ? 'active' : ''}" data-index="${i}">
        <div class="left">
          <div class="title">${esc(item.codigo)} · ${esc(item.descripcion)}</div>
          <div class="meta">${esc(item.area)} · ${esc(item.family)} · ${item.peso ? item.peso + ' lbs' : ''}</div>
        </div>
        <div class="pill">${esc(item.lot_prefix || item.family || '')}</div>
      </div>
    `).join('') : `<div class="suggestion-item"><div class="left"><div class="title">Sin coincidencias</div><div class="meta">Puedes seguir escribiendo libremente.</div></div></div>`;
  }

  function positionPicker(input) {
    if (!pickerState.open) return;
    const rect = input.getBoundingClientRect();
    refs.pickerPanel.style.left = `${Math.min(rect.left, window.innerWidth - 440)}px`;
    refs.pickerPanel.style.top = `${Math.min(rect.bottom + 6, window.innerHeight - 300)}px`;
    refs.pickerPanel.style.width = `${Math.max(320, rect.width)}px`;
  }

  function movePicker(step) {
    if (!pickerState.items.length) return;
    pickerState.index = (pickerState.index + step + pickerState.items.length) % pickerState.items.length;
    renderPicker();
  }

  function chooseSuggestion(i) {
    const item = pickerState.items[i];
    if (!item || !pickerState.input) return;
    const input = pickerState.input;
    const row = input.closest('tr');
    if (!row) return;
    const rowId = row.dataset.id;
    const bind = row.dataset.row;
    const rowObj = getRow(bind, rowId);
    if (!rowObj) return;
    rowObj.codigo = item.codigo;
    rowObj.descripcion = item.descripcion;
    if ('area' in rowObj && !rowObj.area) rowObj.area = item.area || rowObj.area;
    persistSoon();
    renderAll();
    closePicker();
    const nextInput = row.querySelector(`[data-field="${bind === 'master' ? 'descripcion' : 'area'}"]`) || row.querySelector('input, select');
    if (nextInput) nextInput.focus();
  }

  function closePicker() {
    pickerState.open = false;
    pickerState.input = null;
    pickerState.items = [];
    pickerState.index = -1;
    refs.pickerPanel.hidden = true;
    refs.pickerPanel.style.display = 'none';
    refs.pickerPanel.innerHTML = '';
  }

  function searchProducts(query) {
    const q = normalize(query);
    if (!q) return state.master.slice().sort((a,b)=>a.descripcion.localeCompare(b.descripcion));
    return state.master
      .map(m => {
        const hay = normalize(`${m.codigo} ${m.descripcion} ${m.family} ${m.area}`);
        let score = 0;
        if (normalize(m.codigo).startsWith(q)) score += 1000;
        if (normalize(m.descripcion).startsWith(q)) score += 900;
        if (hay.includes(q)) score += 200;
        score += Math.max(0, 100 - Math.abs(hay.length - q.length));
        return { ...m, _score: score };
      })
      .filter(m => m._score > 0)
      .sort((a,b) => b._score - a._score || a.descripcion.localeCompare(b.descripcion));
  }

  function openImportModal(kind) {
    refs.importKind.value = kind || 'production';
    refs.importText.value = '';
    refs.importPreview.textContent = '';
    refs.importModal.hidden = false;
    refs.importModal.style.display = 'grid';
    refs.importText.focus();
  }

  function closeImportModal() {
    refs.importModal.hidden = true;
    refs.importModal.style.display = 'none';
  }

  function previewImport() {
    const kind = refs.importKind.value;
    const text = refs.importText.value.trim();
    const parsed = parseImport(text, kind);
    refs.importPreview.textContent = parsed.preview || 'Sin datos para mostrar.';
    state.importPreview = text;
  }

  function applyImportFromModal() {
    const kind = refs.importKind.value;
    const text = refs.importText.value.trim();
    if (!text) return;
    const parsed = parseImport(text, kind);
    if (kind === 'production') state.production = parsed.rows.concat(state.production);
    else if (kind === 'programming') state.programming = parsed.rows.concat(state.programming);
    else if (kind === 'master') state.master = normalizeMaster(parsed.rows.concat(state.master));
    persistAll();
    closeImportModal();
    renderAll();
  }

  function parseImport(text, kind) {
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return { rows: [], preview: '' };
    const rows = [];
    const isHeader = looksLikeHeader(lines[0]);
    const header = isHeader ? splitRow(lines[0]) : [];
    const dataLines = isHeader ? lines.slice(1) : lines;
    for (const line of dataLines) {
      const cols = splitRow(line);
      if (!cols.length) continue;
      if (kind === 'production') rows.push(parseProductionImport(cols, header));
      else if (kind === 'programming') rows.push(parseProgrammingImport(cols, header));
      else rows.push(parseMasterImport(cols, header));
    }
    return { rows: rows.filter(Boolean), preview: rows.slice(0, 8).map(r => JSON.stringify(r, null, 2)).join('\n\n') };
  }

  function splitRow(line) {
    if (line.includes('\t')) return line.split('\t').map(s => s.trim());
    if (line.includes(';')) return line.split(';').map(s => s.trim());
    if (line.includes(',')) return splitCSVLine(line);
    return line.split(/\s{2,}/).map(s => s.trim());
  }

  function splitCSVLine(line) {
    const out = [];
    let cur = '';
    let quote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      const next = line[i + 1];
      if (ch === '"' && quote && next === '"') { cur += '"'; i++; continue; }
      if (ch === '"') { quote = !quote; continue; }
      if (ch === ',' && !quote) { out.push(cur.trim()); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur.trim());
    return out;
  }

  function normalizeHeader(s) { return normalize(s).replace(/\s+/g,''); }

  function looksLikeHeader(line) {
    const h = normalize(line);
    return ['CODIGO SAP','CODIGO','DESCRIPCION','PRESENTACION','PROGRAMADO','PRODUCCION','FECHA','SEMANA','TOTAL','TURNO','OBSERVACION','AREA','FAMILIA','UNIDAD','PESO'].some(token => h.includes(token));
  }

  function parseProductionImport(cols, header) {
    const map = colMap(header, cols);
    const positional = !header || !header.length;
    return {
      id: uid(),
      date: map.fecha || map.date || isoDateFromParts(map.anio, map.mes, map.dia) || (positional ? cols[0] : '') || state.ui.date,
      dayName: map.dia || (positional ? cols[1] : '') || dayNameOf(map.fecha || state.ui.date),
      shift: map.turno || (positional ? cols[2] : '') || 'A',
      codigo: map.material || map.codigo || (positional ? cols[3] : '') || '',
      descripcion: map.presentacion || map.descripcion || (positional ? cols[4] : '') || '',
      area: map.area || (positional ? cols[5] : '') || state.ui.quickArea || 'PAQUETERIA',
      total: num(map.total || map.cantidad || (positional ? cols[6] : '')),
      note: map.observacion || map.note || (positional ? cols[7] : '') || '',
      annulled: false
    };
  }

  function parseProgrammingImport(cols, header) {
    const map = colMap(header, cols);
    const positional = !header || !header.length;
    return {
      id: uid(),
      date: map.fecha || (positional ? cols[0] : '') || state.ui.date,
      week: num(map.semana) || isoWeek(new Date(map.fecha || (positional ? cols[0] : state.ui.date))),
      codigo: map['codigo sap'] || map.codigo || (positional ? cols[2] : '') || '',
      descripcion: map.presentacion || map.descripcion || (positional ? cols[3] : '') || '',
      area: map.area || (positional ? cols[4] : '') || state.ui.quickArea || 'PAQUETERIA',
      programmed: num(map.programado || (positional ? cols[5] : '')),
      cantidad_me: num(map['cantidad me'] || map.cantidadme || map.cantidad || (positional ? cols[6] : '')),
      type: map.tipo || (positional ? cols[7] : '') || 'BOBINA IMPRESA',
      annulled: false
    };
  }

  function parseMasterImport(cols, header) {
    const map = colMap(header, cols);
    const positional = !header || !header.length;
    return {
      id: uid(),
      codigo: map.codigo || map['codigo sku'] || (positional ? cols[0] : '') || '',
      descripcion: map.descripcion || map['descripcion del producto'] || (positional ? cols[1] : '') || '',
      peso: num(map.peso || map['peso/unidad'] || map['peso lbs'] || (positional ? cols[2] : '')),
      unidad: map.unidad || map['unidad de medida'] || (positional ? cols[3] : '') || '',
      family: map.familia || map.family || inferFamily({descripcion: map.descripcion || (positional ? cols[1] : '') || ''}),
      proceso: map.proceso || (positional ? cols[5] : '') || '',
      entero_min: map['entero min.'] || map['entero min'] || (positional ? cols[6] : '') || '',
      mignon_max: map['mignon max.'] || map['migon max.'] || map['pulido min.'] || (positional ? cols[7] : '') || '',
      pulido_min: map['pulido min.'] || map['pulido'] || (positional ? cols[8] : '') || '',
      area: map.area || (positional ? cols[9] : '') || inferArea({descripcion: map.descripcion || (positional ? cols[1] : '') || ''}),
      tipo: map.tipo || (positional ? cols[10] : '') || 'PT',
      tarima_div: num(map.tarima || map['tarima div'] || (positional ? cols[11] : '')) || state.settings.defaultTarima,
      bobina_unidades: num(map['bobina u'] || map['bobina unid.'] || (positional ? cols[12] : '')) || 5,
      bobina_rend: num(map['bobina r'] || map['bobina rend.'] || (positional ? cols[13] : '')) || 1,
      granza_factor: num(map.granza || map['granza factor'] || (positional ? cols[14] : '')) || 1.4286
    };
  }

  function colMap(header, cols) {
    const map = {};
    if (header && header.length) {
      header.forEach((h, i) => map[normalizeHeader(h)] = cols[i]);
      // also plain text keys
      header.forEach((h, i) => map[h.toLowerCase()] = cols[i]);
    }
    const alias = {
      'codigo sap': 'codigo sap', 'codigo': 'codigo', 'presentacion': 'presentacion', 'descripcion': 'descripcion',
      'programado': 'programado', 'cantidad me': 'cantidad me', 'cantidad': 'cantidad', 'tipo': 'tipo',
      'fecha': 'fecha', 'semana': 'semana', 'area': 'area', 'total': 'total', 'turno': 'turno', 'observacion': 'observacion', 'dia': 'dia', 'año': 'anio', 'ano': 'anio', 'mes': 'mes', 'material': 'material'
    };
    const out = {};
    for (const [k,v] of Object.entries(map)) out[normalize(k)] = v;
    return new Proxy(out, { get: (obj, prop) => obj[normalize(prop)] || obj[prop] });
  }

  function exportJSON() {
    const blob = new Blob([JSON.stringify({ ui: state.ui, settings: state.settings, master: state.master, programming: state.programming, production: state.production, auth: state.auth }, null, 2)], {type:'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'portal_sap_backup.json';
    a.click();
  }

  function exportCSV(kind='all') {
    let rows = [];
    if (kind === 'compliance') {
      const data = complianceData(currentWeekRange());
      const days = rangeDays(currentWeekRange());
      rows = [['CODIGO SAP','AREA','PRESENTACION','PROGRAMACION',...days.map(d=>d.label),'TOTAL','CUMPLIMIENTO','DIFERENCIA','SC/IC']];
      data.forEach(r => {
        const pct = r.programmed ? Math.min(100, (r.total / r.programmed) * 100) : 0;
        const ratio = r.programmed ? (r.total / r.programmed) * 100 : 0;
        rows.push([r.codigo, r.area, r.descripcion, r.programmed, ...days.map(d=>r.byDay[d.key]||0), r.total, safePct(pct, 2), r.total - r.programmed, safePct(ratio, 2)]);
      });
    } else {
      rows = [['TIPO','FECHA','CODIGO','DESCRIPCION','AREA','VALOR','TEXTO']];
      if (state.ui.view === 'production') state.production.forEach(r => rows.push(['PRODUCCION', r.date, r.codigo, r.descripcion, r.area, r.total, r.note]));
      else state.programming.forEach(r => rows.push(['PROGRAMACION', r.date, r.codigo, r.descripcion, r.area, r.programmed, '']));
    }
    const csv = rows.map(r => r.map(csvCell).join(',')).join('\n');
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `portal_sap_${state.ui.view}.csv`;
    a.click();
  }

  function persistSoon() {
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(() => persistAll(), 250);
  }

  function persistAll(refresh=true) {
    state.ui.week = isoWeek(new Date(state.ui.date || today));
    localStorage.setItem(LS.ui, JSON.stringify(state.ui));
    localStorage.setItem(LS.settings, JSON.stringify(state.settings));
    localStorage.setItem(LS.master, JSON.stringify(state.master));
    localStorage.setItem(LS.programming, JSON.stringify(state.programming));
    localStorage.setItem(LS.production, JSON.stringify(state.production));
    localStorage.setItem(LS.auth, JSON.stringify(state.auth));
    state.lastSaved = new Date();
    updateStatusBar();
    if (refresh) renderAll();
  }

  function saveJSON(key, value) { localStorage.setItem(key, JSON.stringify(value)); }
  function loadJSON(key, fallback) { try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : clone(fallback); } catch { return clone(fallback); } }

  function setTheme(theme) {
    state.ui.theme = theme;
    applyTheme(theme);
    persistAll();
  }

  function applyTheme(theme) {
    document.body.className = `theme-${['quartz','quartz-dark','belize','signature'].includes(theme) ? theme : 'quartz'}`;
    document.body.dataset.density = state.ui.density;
  }

  function updateStatusBar() {
    if (refs.statusWeek) refs.statusWeek.textContent = String(state.ui.week);
    if (refs.statusDate) refs.statusDate.textContent = spanishDate(new Date(state.ui.date || today));
    if (refs.statusView) refs.statusView.textContent = viewName(isAuthenticated() ? state.ui.view : 'auth');
    if (refs.statusSave) refs.statusSave.textContent = state.lastSaved ? shortTime(state.lastSaved) : 'sin guardar';
    const user = currentUser();
    if (refs.sessionChip) refs.sessionChip.textContent = user ? `${user.fullName || user.username} · ${user.role}` : 'Sin sesión';
    if (refs.btnLogout) refs.btnLogout.hidden = !user;
  }

  function viewName(view) {
    return ({ dashboard:'Inicio', programming:'Programación', production:'Producción', compliance:'Cumplimiento', lotification:'Lotificación', master:'Maestro', config:'Configuración', documents:'Documentos', admin:'Seguridad', auth:'Acceso' })[view] || view;
  }

  function num(v) { const n = Number(String(v).replace(/,/g,'').replace(/\s/g,'')); return Number.isFinite(n) ? n : 0; }
  function normalize(v) { return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim(); }
  function pad2(n) { return String(n).padStart(2, '0'); }
  function toISO(d) { const x = new Date(d); return `${x.getFullYear()}-${pad2(x.getMonth()+1)}-${pad2(x.getDate())}`; }
  function isoDateFromParts(year, month, day) {
    const y = num(year); const m = month ? monthToNumber(month) : 0; const d = num(day);
    if (!y || !m || !d) return '';
    return `${y}-${pad2(m)}-${pad2(d)}`;
  }
  function monthToNumber(m) {
    const s = normalize(m).slice(0,3);
    const map = { ENE:1, FEB:2, MAR:3, ABR:4, MAY:5, JUN:6, JUL:7, AGO:8, SEP:9, OCT:10, NOV:11, DIC:12 };
    return map[s] || num(m);
  }
  function clone(v) { return JSON.parse(JSON.stringify(v)); }
  function fmtNumber(n) { return new Intl.NumberFormat('es-HN').format(num(n)); }
  function fmtSigned(n) { const x = num(n); return `${x >= 0 ? '+' : ''}${fmtNumber(x)}`; }
  function fmtPct(n) { const x = Number(n); return `${(Number.isFinite(x) ? x : 0).toFixed(0)}%`; }
  function safePct(n, decimals = 0) {
    const x = Number(n);
    if (!Number.isFinite(x)) return `${(0).toFixed(decimals)}%`;
    return `${x.toFixed(decimals)}%`;
  }
  function dateLabel(d) { return spanishDate(new Date(d)) }
  function shortDate(d) { const x = new Date(d); return `${x.getDate()}/${x.getMonth()+1}`; }
  function spanishDate(d) { const x = new Date(d); return `${dayNames[x.getDay()]} ${x.getDate()} DE ${monthNames[x.getMonth()]} DE ${x.getFullYear()}`; }
  function longNow() { return new Intl.DateTimeFormat('es-HN', { dateStyle:'full', timeStyle:'medium' }).format(new Date()); }
  function shortTime(d) { return new Intl.DateTimeFormat('es-HN', { hour:'2-digit', minute:'2-digit' }).format(new Date(d)); }
  function dayNameOf(date) { return dayNames[new Date(date).getDay()]; }
  function monthKey(date) { const x = new Date(date); return `${x.getFullYear()}-${pad2(x.getMonth()+1)}`; }
  function isWithin(date, start, end) { const d = new Date(date); return d >= start && d <= end; }
  function mondayOf(date) { const d = new Date(date); const day = d.getDay() === 0 ? -6 : 1 - d.getDay(); d.setHours(0,0,0,0); d.setDate(d.getDate() + day); return d; }
  function isoWeek(date) { const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())); const dayNum = d.getUTCDay() || 7; d.setUTCDate(d.getUTCDate() + 4 - dayNum); const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1)); return Math.ceil((((d - yearStart) / 86400000) + 1) / 7); }
  function csvCell(v) { const s = String(v ?? ''); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
  function compositeKey(a,b,c) { return normalize([a,b,c].join('|')); }

  function handleInputExtras() {
    const lot = $('#lotThreshold'); if (lot) lot.addEventListener('input', e => { state.ui.lotThreshold = num(e.target.value) || 3000; persistSoon(); scheduleRender(); });
  }

  // extra direct bindings for inputs in re-rendered views
  document.addEventListener('change', e => {
    if (e.target.matches('#settingLot')) { state.ui.lotThreshold = num(e.target.value) || 3000; persistSoon(); }
    if (e.target.matches('#settingPackRate')) { state.settings.packRate = num(e.target.value) || 235; persistSoon(); }
    if (e.target.matches('#settingMillRate')) { state.settings.millingRate = num(e.target.value) || 100; persistSoon(); }
    if (e.target.matches('#settingEnvas')) { state.settings.defaultEnvas = num(e.target.value) || 2; persistSoon(); }
    if (e.target.matches('#settingTarima')) { state.settings.defaultTarima = num(e.target.value) || 100; persistSoon(); }
    if (e.target.matches('#quickArea')) { state.ui.quickArea = e.target.value; persistSoon(); }
    if (e.target.matches('input[name="theme"]')) { setTheme(e.target.value); }
    if (e.target.matches('[data-auth-bind]')) {
      updateAuthRow(e.target.dataset.authBind, e.target.dataset.row, e.target.dataset.field, e.target.type === 'checkbox' ? e.target.checked : e.target.value);
    }
  });



function currentUser() {
  const session = state.auth && state.auth.session;
  if (!session) return null;
  const user = state.auth.users.find(u => u.id === session.userId);
  if (!user || user.active === false) return null;
  return user;
}

function currentRole() {
  const user = currentUser();
  if (!user) return null;
  return state.auth.roles.find(r => normalize(r.code) === normalize(user.role)) || null;
}

function isAuthenticated() {
  return !!currentUser();
}

function canAccess(target) {
  const user = currentUser();
  if (!user) return false;
  const role = currentRole();
  if (!role) return false;
  const perms = Array.isArray(role.permissions) ? role.permissions : [];
  if (perms.some(p => normalize(p) === '*')) return true;
  const wanted = normalize(target);
  return perms.some(p => normalize(p) === wanted);
}

function persistAuth() {
  localStorage.setItem(LS.auth, JSON.stringify(state.auth));
}

function pushAudit(action, detail='') {
  if (!state.auth.audit) state.auth.audit = [];
  state.auth.audit.unshift({
    id: uid(),
    at: new Date().toLocaleString('es-HN'),
    user: currentUser() ? currentUser().username : 'sistema',
    action,
    detail
  });
  state.auth.audit = state.auth.audit.slice(0, 50);
}

function syncAuthSession() {
  if (!state.auth.session) return;
  const valid = state.auth.users.find(u => u.id === state.auth.session.userId && u.active !== false);
  if (!valid) {
    state.auth.session = null;
    persistAuth();
  }
}

function performLogin(username, password) {
  const user = state.auth.users.find(u => normalize(u.username) === normalize(username) && u.password === password && u.active !== false);
  if (!user) {
    state.auth.error = 'Credenciales inválidas o usuario inactivo.';
    persistAuth();
    return false;
  }
  state.auth.session = { userId: user.id, at: new Date().toISOString() };
  state.auth.error = '';
  state.ui.view = 'dashboard';
  pushAudit('LOGIN', `${user.username} inició sesión`);
  persistAll(false);
  return true;
}

function submitLoginFromDom() {
  const user = $('#authUser')?.value || '';
  const pass = $('#authPass')?.value || '';
  const ok = performLogin(user, pass);
  renderAll();
  if (!ok) {
    setTimeout(() => {
      const u = $('#authUser'); const p = $('#authPass');
      if (u) u.focus();
      if (p) p.value = '';
    }, 0);
  }
}

function submitDemoLogin() {
  const demo = state.auth.users.find(u => normalize(u.username) === 'ADMIN' && u.active !== false);
  if (!demo) return;
  performLogin(demo.username, demo.password);
  renderAll();
}

function logout() {
  const user = currentUser();
  if (user) pushAudit('LOGOUT', `${user.username} cerró sesión`);
  state.auth.session = null;
  persistAll(false);
  renderAll();
}

function updateAuthRow(listName, rowId, field, value) {
  if (!state.auth || !state.auth[listName]) return;
  const row = state.auth[listName].find(r => String(r.id ?? r.code) === String(rowId));
  if (!row) return;
  if (listName === 'users') {
    if (field === 'active') row.active = !!value;
    else row[field] = value;
  } else if (listName === 'roles') {
    if (field === 'permissions') {
      row.permissions = String(value || '')
        .split(',')
        .map(v => v.trim())
        .filter(Boolean);
    } else {
      row[field] = value;
    }
    if (field === 'code' && row.code) row.code = String(row.code).trim().toUpperCase();
  }
  persistAuth();
  scheduleRender();
}

function addAuthUser() {
  state.auth.users.unshift({
    id: uid(),
    username: '',
    fullName: '',
    password: '',
    role: 'READ',
    active: true
  });
  persistAuth();
  renderAll();
}

function addAuthRole() {
  state.auth.roles.unshift({
    code: 'NEW',
    name: 'Nuevo rol',
    permissions: ['dashboard']
  });
  persistAuth();
  renderAll();
}

function deleteAuthRow(bind, id) {
  if (!confirm('¿Eliminar este registro de seguridad?')) return;
  if (!state.auth[bind]) return;
  state.auth[bind] = state.auth[bind].filter(r => String(r.id ?? r.code) !== String(id));
  persistAuth();
  renderAll();
}

  document.addEventListener('DOMContentLoaded', init);
  if (document.readyState !== 'loading') init();
})();
