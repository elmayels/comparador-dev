const $ = (s, r=document) => r.querySelector(s);
const $$ = (s, r=document) => [...r.querySelectorAll(s)];

const state = {
  user: JSON.parse(localStorage.getItem('apu:user') || 'null'),
  token: localStorage.getItem('apu:token'),
  theme: localStorage.getItem('apu:theme') || 'dark',
  refs: null,
  baseUploadFiles: [],
  comparisonFiles: [],
  comparisonMode: 'MULTI',
  comparisonContractors: [
    { id: 1, name: 'PROV-A', conceptsFile: null, matrixFile: null },
    { id: 2, name: 'PROV-B', conceptsFile: null, matrixFile: null },
  ],
  lastComparisonRun: null,
  lastBaseBudgetRun: null,
  pendingBaseBudget: null,
  lastBaseError: null,
};

document.documentElement.dataset.theme = state.theme;

const routes = {
  '': landing,
  'login': login,
  'dashboard': dashboard,
  'base-budgets': baseBudgets,
  'base-budgets-new': baseBudgetsNew,
  'base-budgets-loading': baseBudgetLoading,
  'base-budgets-result': baseBudgetResult,
  'comparisons': comparisons,
  'comparison-new': comparisonNew,
  'comparison-processing': comparisonProcessing,
  'comparison-results': comparisonResults,
  'matrix-detail': matrixDetail,
  'my-runs': myRuns,
  'team-runs': teamRuns,
  'global-runs': globalRuns,
  'users': usersPage,
  'user-create': userCreate,
  'user-edit': userEdit,
  'references': referencesPage,
  'audit': auditPage,
  'profile': profile,
};

const demoUsers = [
  {name:'Super Administrador', email:'superadmin@demo.com', role:'SUPER_ADMIN', org:'Global', status:'Activo'},
  {name:'Administrador Demo', email:'admin@demo.com', role:'ADMIN', org:'Empresa Demo', status:'Activo'},
  {name:'Analista Demo', email:'analista@demo.com', role:'ANALYST', org:'Empresa Demo', status:'Activo'},
  {name:'Analista Obra Norte', email:'analista.norte@demo.com', role:'ANALYST', org:'Empresa Demo', status:'Inactivo'},
];


function money(n){return new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',maximumFractionDigits:0}).format(n)}
function go(hash){ location.hash = '#/' + hash; }
function currentRoute(){ return location.hash.replace(/^#\/?/, '').replace(/^\//,'') || ''; }
function can(view){
  const role = state.user?.role;
  if (!state.user) return false;
  if (role === 'SUPER_ADMIN') return true;
  if (view === 'audit' || view === 'global-runs') return false;
  if (view === 'users') return role === 'ADMIN';
  if (view === 'team-runs') return role === 'ADMIN';
  return true;
}

function render(){
  const r = currentRoute();
  const publicRoutes = ['', 'login'];
  if (!state.user && !publicRoutes.includes(r)) return go('login');
  if (state.user && r === 'login') return go('dashboard');
  (routes[r] || notFound)();
  bindCommon();
}

window.addEventListener('hashchange', render);

function bindCommon(){
  $$('[data-nav]').forEach(a => a.onclick = e => { e.preventDefault(); go(a.dataset.nav); });
  $$('[data-theme-toggle]').forEach(b => b.onclick = () => {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    localStorage.setItem('apu:theme', state.theme);
    document.documentElement.dataset.theme = state.theme;
    render();
  });
  $$('[data-logout]').forEach(b => b.onclick = () => {
    localStorage.removeItem('apu:user'); localStorage.removeItem('apu:token'); state.user=null; state.token=null; go('');
  });
}

function shell(content, title='Plataforma APU'){
  const menu = [
    ['dashboard','Dashboard',''],
    ['base-budgets','Presupuestos base',''],
    ['comparisons','Comparador',''],
    ['my-runs','Mis corridas',''],
    ['team-runs','Corridas equipo','ADMIN'],
    ['global-runs','Historial global','SUPER_ADMIN'],
    ['references','Referencias data',''],
    ['users','Usuarios','ADMIN_OR_SUPER'],
    ['audit','Auditoría','SUPER_ADMIN'],
    ['profile','Mi perfil',''],
  ].filter(([_,__,perm]) => {
    if (!perm) return true;
    if (perm === 'ADMIN') return state.user.role === 'ADMIN' || state.user.role === 'SUPER_ADMIN';
    if (perm === 'ADMIN_OR_SUPER') return ['ADMIN','SUPER_ADMIN'].includes(state.user.role);
    return state.user.role === perm;
  });
  const active = currentRoute();
  $('#app').innerHTML = `
    <div class="app-shell">
      <aside class="sidebar">
        <div class="side-head"><div class="logo">Q</div><div><strong>Quantia APU</strong><div class="muted small">V1 real</div></div></div>
        <nav class="menu">${menu.map(([h,l])=>`<a href="#/${h}" class="${active===h?'active':''}">${l}<span>›</span></a>`).join('')}</nav>
        <div class="side-foot">
          <div class="small muted">Usuario</div>
          <strong>${state.user.name}</strong>
          <div class="small muted">${state.user.role} · ${state.user.org}</div>
        </div>
      </aside>
      <main class="main">
        <header class="topbar"><div><strong>${title}</strong><div class="small muted">Presupuesto base · Comparador · Detalle APU · Datos reales</div></div><div class="actions"><button class="btn btn-secondary" data-theme-toggle>${state.theme==='dark'?'Tema claro':'Tema oscuro'}</button><button class="btn btn-secondary" data-logout>Salir</button></div></header>
        <section class="page">${content}</section>
      </main>
    </div>`;
}

function landing(){
  $('#app').innerHTML = `<div class="public marketing">
    <nav class="nav landing-nav">
      <div class="brand"><div class="logo">Q</div><div><h1>Quantia APU</h1><p>Presupuestos, mercado y APU con inteligencia artificial</p></div></div>
      <div class="landing-links"><a href="#market">Mercado</a><a href="#budgets">Presupuestos</a><a href="#workflow">Flujo</a><a href="#deliverables">Entregables</a></div>
      <div class="actions"><button class="btn btn-secondary" data-theme-toggle>${state.theme==='dark'?'Tema claro':'Tema oscuro'}</button><button class="btn btn-primary" data-nav="login">Iniciar sesión</button></div>
    </nav>

    <section class="landing-hero">
      <div class="hero-copy">
        <div class="eyebrow"><span></span> Comparativa de mercado con referencias Construdata e IA</div>
        <h2>Presupuesta, compara y negocia precios unitarios con evidencia de mercado.</h2>
        <p class="lead">Quantia APU convierte catálogos de conceptos, propuestas de proveedores y matrices APU en un análisis económico claro: presupuestos base, comparativas contra mercado, detalle técnico por proveedor y recomendaciones ejecutivas generadas con apoyo de IA.</p>
        <div class="hero-actions"><button class="btn btn-primary" data-nav="login">Explorar plataforma</button><button class="btn btn-secondary" onclick="document.getElementById('market').scrollIntoView({behavior:'smooth'})">Ver valor de mercado</button></div>
        <div class="trust-strip"><span>Referencias Construdata</span><span>Presupuesto desde conceptos base</span><span>Comparativa multi-proveedor</span><span>Detalle APU auditado</span></div>
      </div>
      <div class="product-showcase" aria-label="Vista previa del sistema">
        <div class="showcase-top"><span></span><span></span><span></span></div>
        <div class="score-card large"><label>Desviación contra mercado</label><strong>+8.7%</strong><small>La propuesta supera la referencia en partidas de alto impacto.</small></div>
        <div class="mini-dashboard">
          <div><label>Presupuesto base</label><strong>$56.5M</strong></div>
          <div><label>Cobertura ref.</label><strong>86%</strong></div>
          <div><label>Partidas críticas</label><strong>18</strong></div>
        </div>
        <div class="chart-card"><div class="bar b1"></div><div class="bar b2"></div><div class="bar b3"></div><div class="bar b4"></div></div>
        <div class="ai-note"><strong>IA sobre datos calculados</strong><p>Resume desviaciones, explica causas probables y prioriza partidas para revisión sin reemplazar la trazabilidad numérica.</p></div>
      </div>
    </section>

    <section class="landing-section problem-section">
      <div class="section-head"><span class="section-kicker">El reto</span><h3>Una oferta baja no siempre es competitiva; una oferta alta no siempre está mal justificada.</h3><p>La decisión correcta exige mirar cada concepto contra referencias de mercado, entender la composición de su APU y detectar qué insumos, rendimientos o porcentajes explican la diferencia.</p></div>
      <div class="pain-grid">
        <div class="pain-card"><strong>Mercado disperso</strong><p>Los precios unitarios cambian por alcance, ciudad, rendimiento, disponibilidad e insumos. El sistema ayuda a comparar contra referencias Construdata y archivos de mercado.</p></div>
        <div class="pain-card"><strong>Catálogos incompletos</strong><p>Antes de licitar, los equipos técnicos necesitan convertir conceptos base en un presupuesto defendible, con partidas en revisión y conceptos sin match.</p></div>
        <div class="pain-card"><strong>APU difíciles de auditar</strong><p>Cada proveedor puede presentar una matriz distinta. El detalle debe respetar su estructura y mostrar la comparación contra materiales, mano de obra y maquinaria de referencia.</p></div>
      </div>
    </section>

    <section id="market" class="landing-section ai-section">
      <div class="ai-gradient"><span class="section-kicker">Comparativa de mercado</span><h3>Evalúa propuestas contra valores de referencia Construdata y detecta desviaciones con impacto real.</h3><p>El análisis no se queda en el monto total. La plataforma cruza precios unitarios, importes, insumos y porcentajes contra referencias disponibles en data para identificar sobrecostos, partidas sin referencia, unidades dudosas y oportunidades de negociación.</p><div class="ai-pill-row"><span>Materiales de referencia</span><span>Mano de obra</span><span>Maquinaria y equipo</span><span>Porcentajes aplicados</span><span>Impacto económico</span></div></div>
    </section>

    <section id="budgets" class="landing-section">
      <div class="section-head"><span class="section-kicker">Presupuestos base</span><h3>Genera presupuestos desde un catálogo de conceptos base creado por ingeniería.</h3><p>Cuando aún no existe licitación, Quantia APU permite tomar un archivo de conceptos base, cruzarlo contra matrices Construdata y producir un presupuesto inicial para estimar costo, revisar alcance y preparar mejor la salida a mercado.</p></div>
      <div class="capability-grid">
        <article class="capability-card accent-blue"><div class="cap-icon">01</div><h4>Catálogo de conceptos base</h4><p>Parte del archivo elaborado por ingeniería con conceptos, unidades y cantidades del proyecto.</p><ul><li>Conceptos técnicos</li><li>Unidades y cantidades</li><li>Observaciones de alcance</li></ul></article>
        <article class="capability-card accent-green"><div class="cap-icon">02</div><h4>Matrices Construdata</h4><p>Busca equivalencias y referencias para estimar precios unitarios base de forma trazable.</p><ul><li>Matches por concepto</li><li>Partidas sin referencia</li><li>Confianza del cruce</li></ul></article>
        <article class="capability-card accent-purple"><div class="cap-icon">03</div><h4>Presupuesto estimado</h4><p>Genera una línea base económica para validar alcance, rangos y rubros principales antes de cotizar.</p><ul><li>Monto total base</li><li>Distribución por familias</li><li>Conceptos en revisión</li></ul></article>
        <article class="capability-card accent-orange"><div class="cap-icon">04</div><h4>Salida ejecutiva</h4><p>Entrega un Excel limpio y un resumen para revisar, ajustar y usar como soporte técnico.</p><ul><li>Resumen ejecutivo</li><li>Matriz presupuestada</li><li>Alertas de consistencia</li></ul></article>
      </div>
    </section>

    <section id="capabilities" class="landing-section">
      <div class="section-head"><span class="section-kicker">Capacidades</span><h3>Un sistema para presupuestar, comparar y explicar diferencias económicas.</h3><p>La solución separa el presupuesto base independiente de la comparación de propuestas, pero mantiene una misma lógica de trazabilidad, referencias y análisis ejecutivo.</p></div>
      <div class="capability-grid">
        <article class="capability-card accent-blue"><div class="cap-icon">A</div><h4>Presupuesto base independiente</h4><p>Conceptos de ingeniería + matrices Construdata para obtener una estimación previa.</p><ul><li>No depende de una licitación</li><li>Puede ejecutarse solo</li><li>Genera línea base</li></ul></article>
        <article class="capability-card accent-green"><div class="cap-icon">B</div><h4>Comparador de propuestas</h4><p>Analiza uno o varios proveedores, mostrando diferencias contra mercado y entre ofertas.</p><ul><li>Formato multi-proveedor</li><li>Ranking económico</li><li>Partidas críticas</li></ul></article>
        <article class="capability-card accent-purple"><div class="cap-icon">C</div><h4>Detalle APU por proveedor</h4><p>Cada matriz se analiza en su propia hoja porque no todos los proveedores declaran la misma estructura.</p><ul><li>Materiales</li><li>Mano de obra</li><li>Maquinaria y porcentajes</li></ul></article>
        <article class="capability-card accent-orange"><div class="cap-icon">D</div><h4>IA ejecutiva</h4><p>Transforma cálculos y validaciones en hallazgos, riesgos y preguntas de negociación.</p><ul><li>Causas probables</li><li>Recomendaciones</li><li>Prioridades</li></ul></article>
      </div>
    </section>

    <section id="workflow" class="landing-section split-section">
      <div>
        <span class="section-kicker">Flujo operativo</span>
        <h3>De archivos XLSX a una lectura económica clara, comparable y defendible.</h3>
        <p>Para cada proveedor se registra un nombre corto, su catálogo de conceptos y su matriz/APU. El sistema conserva el origen de cada archivo, valida la estructura y genera comparativos profesionales contra mercado.</p>
        <div class="flow-list"><div><b>1</b><span>Crear presupuesto base desde conceptos o iniciar comparación</span></div><div><b>2</b><span>Cargar por proveedor: nombre, conceptos XLSX y matriz/APU XLSX</span></div><div><b>3</b><span>Comparar precios, importes, insumos y porcentajes contra data de mercado</span></div><div><b>4</b><span>Generar Excel ejecutivo, detalle por proveedor y hallazgos IA</span></div></div>
      </div>
      <div class="workflow-panel">
        <div class="workflow-row active"><span>Conceptos base</span><strong>Presupuesto previo con Construdata</strong></div>
        <div class="workflow-row"><span>Proveedor 1</span><strong>Conceptos.xlsx + Matriz_APU.xlsx</strong></div>
        <div class="workflow-row"><span>Proveedor 2</span><strong>Conceptos.xlsx + Matriz_APU.xlsx</strong></div>
        <div class="workflow-row"><span>Referencias</span><strong>Materiales · MO · Maquinaria · %</strong></div>
        <div class="workflow-result"><strong>Resultado</strong><p>Comparativa horizontal + detalle independiente por proveedor + análisis de mercado.</p></div>
      </div>
    </section>

    <section class="landing-section ai-section">
      <div class="ai-gradient"><span class="section-kicker">Capa IA</span><h3>IA para explicar, no para inventar.</h3><p>La plataforma calcula con reglas determinísticas y usa IA para interpretar resultados: qué partidas concentran riesgo, qué insumos explican el sobrecosto, qué debe revisarse primero y qué preguntas conviene hacer al proveedor.</p><div class="ai-pill-row"><span>Explicación de sobrecostos</span><span>Priorización de negociación</span><span>Resumen ejecutivo</span><span>Alertas de consistencia</span></div></div>
    </section>

    <section id="deliverables" class="landing-section deliverables-section">
      <div class="section-head"><span class="section-kicker">Entregables</span><h3>Excel profesional y resultados diseñados para revisión técnica y decisión ejecutiva.</h3><p>El Excel se aprovecha como una interfaz de análisis: limpio, ordenado, filtrable y con la separación correcta entre comparativa general y detalle técnico.</p></div>
      <div class="deliverable-grid">
        <div class="deliverable-card"><strong>Resumen Ejecutivo</strong><p>KPIs, semáforos, ranking, cobertura de referencias y principales hallazgos.</p></div>
        <div class="deliverable-card"><strong>Comparativa</strong><p>Vista horizontal por proveedor para comparar P.U., importe, participación y mercado.</p></div>
        <div class="deliverable-card"><strong>Detalle por proveedor</strong><p>Una hoja independiente por matriz/APU, respetando su estructura y sus porcentajes declarados.</p></div>
        <div class="deliverable-card"><strong>Validaciones</strong><p>Insumos sin referencia, unidades dudosas, diferencias de cálculo y partidas para revisión.</p></div>
      </div>
    </section>

    <section class="landing-cta"><h3>Presupuestos más sólidos. Comparativos más claros. Negociaciones mejor sustentadas.</h3><p>Una plataforma para explotar referencias Construdata, catálogos base y matrices APU con una capa de IA que convierte datos técnicos en decisiones económicas.</p><button class="btn btn-primary" data-nav="login">Iniciar análisis</button></section>
  </div>`;
}

function login(){
  $('#app').innerHTML = `<div class="login-page"><div class="login-card"><div class="brand" style="margin-bottom:18px"><div class="logo">Q</div><div><h1>Quantia APU</h1><p>V1 real</p></div></div><h2>Iniciar sesión</h2><p>Usa uno de los usuarios demo para validar navegación, permisos e historial.</p><form id="loginForm" class="form-grid"><div><label class="label">Correo</label><input class="input" name="email" value="superadmin@demo.com"></div><div><label class="label">Contraseña</label><input class="input" name="password" type="password" value="demo123"></div><button class="btn btn-primary">Ingresar</button><div id="loginError" class="bad badge hidden"></div></form><div class="demo-box"><strong>Demo:</strong><br>superadmin@demo.com / demo123<br>admin@demo.com / demo123<br>analista@demo.com / demo123</div></div></div>`;
  $('#loginForm').onsubmit = async e => {
    e.preventDefault();
    const payload = Object.fromEntries(new FormData(e.target));
    try{
      const res = await fetch('/api/auth/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
      if(!res.ok) throw new Error('Credenciales inválidas');
      const data = await res.json(); state.user=data.user; state.token=data.token;
      localStorage.setItem('apu:user', JSON.stringify(data.user)); localStorage.setItem('apu:token', data.token); go('dashboard');
    }catch(err){ $('#loginError').textContent=err.message; $('#loginError').classList.remove('hidden'); }
  }
}

function pageHead(title, subtitle, action=''){
  return `<div class="page-head"><div><h2>${title}</h2><p>${subtitle}</p></div><div class="actions">${action}</div></div>`;
}
function kpis(items){return `<div class="grid cols-4">${items.map(i=>`<div class="card kpi"><label>${i.label}</label><strong>${i.value}</strong><span>${i.text||''}</span></div>`).join('')}</div>`}
function table(headers, rows){return `<div class="table-wrap"><table><thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`}

function fmtMoney(n){ const v = Number(n||0); return new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',maximumFractionDigits:0}).format(v); }
function fmtPct(n){ return `${Number(n||0).toFixed(1)}%`; }
function chartBar(label, amount, max, meta=''){
  const pct = max > 0 ? Math.max(2, Math.min(100, (Number(amount||0)/max)*100)) : 0;
  return `<div class="bar-row"><div class="bar-meta"><strong>${label}</strong><span>${fmtMoney(amount)}${meta?` · ${meta}`:''}</span></div><div class="bar-track"><span style="width:${pct}%"></span></div></div>`;
}
function donutLegend(items){
  const total = items.reduce((a,i)=>a+Number(i.amount||0),0);
  const max = Math.max(...items.map(i=>Number(i.amount||0)), 1);
  return `<div class="chart-list">${items.map(i=>chartBar(i.label, i.amount, max, total?`${((Number(i.amount||0)/total)*100).toFixed(1)}%`:'' )).join('')}</div>`;
}
function summaryFromRun(run){ return run?.summary || run || {}; }

function dashboard(){
  const role = state.user.role;
  const cards = role==='SUPER_ADMIN' ? [
    {label:'Usuarios activos',value:'45',text:'Visibilidad global'}, {label:'Corridas totales',value:'326',text:'Historial completo'}, {label:'Presupuestos base',value:'28',text:'Módulo independiente'}, {label:'Errores recientes',value:'5',text:'Auditoría disponible'}
  ] : role==='ADMIN' ? [
    {label:'Analistas activos',value:'12',text:'Equipo asignado'}, {label:'Corridas mes',value:'48',text:'Bajo su gestión'}, {label:'Con error',value:'3',text:'Requieren revisión'}, {label:'Presupuestos',value:'8',text:'Creados por equipo'}
  ] : [
    {label:'Mis corridas',value:'9',text:'Historial personal'}, {label:'Presupuestos base',value:'2',text:'Independientes'}, {label:'Pendientes',value:'1',text:'Con advertencias'}, {label:'Último riesgo',value:'Amarillo',text:'Comparativo reciente'}
  ];
  shell(`${pageHead('Dashboard', 'Entrada principal para generar, comparar y revisar resultados reales.', `<button class="btn btn-primary" data-nav="base-budgets-new">Nuevo presupuesto base real</button><button class="btn btn-secondary" data-nav="comparison-new">Nueva comparación</button>`)}${kpis(cards)}<div class="grid cols-2" style="margin-top:16px"><div class="card"><h3>Flujo recomendado</h3><p>Presupuesto base es opcional e independiente. El comparador puede ejecutarse sin presupuesto base.</p><div class="actions" style="margin-top:14px"><button class="btn btn-secondary" data-nav="base-budgets">Ver presupuestos</button><button class="btn btn-secondary" data-nav="comparisons">Ver comparador</button></div></div><div class="card"><h3>Regla canónica</h3><p><strong>El sistema calcula.</strong> La IA interpreta. El tablero comunica. La IA interpreta hallazgos sin modificar los cálculos.</p></div></div>`, 'Dashboard');
}

function baseBudgets(){
  const last = state.lastBaseBudgetRun;
  const lastBlock = last ? `<div style="margin-top:16px">${table(['Campo','Valor'], [['Ultima corrida real', last.id], ['Conceptos leidos', last.concepts || 0], ['Ejecutables', last.executableConcepts || 0], ['Filas Detalle Base', last.apuItems || 0], ['Excel', `<a href="${last.downloadUrl}">Descargar reporte real</a>`]])}</div>` : `<div class="callout" style="margin-top:16px"><strong>Sin corrida real en esta sesion.</strong> Usa el boton Nuevo presupuesto base, carga el archivo .xlsx y genera el reporte. Esta pantalla solo muestra corridas reales generadas en la sesión.</div>`;
  shell(`${pageHead('Presupuestos base reales', 'Modulo independiente: conceptos de ingenieria + data/construdata_matrices.xlsx.', `<button class="btn btn-primary" data-nav="base-budgets-new">Nuevo presupuesto base real</button>`)}${kpis([{label:'Modo',value:'REAL',text:'Datos reales'}, {label:'Entrada',value:'.xlsx',text:'Catalogo de conceptos'}, {label:'Fuente matriz',value:'Construdata',text:'data/construdata_matrices.xlsx'}, {label:'Excel',value:'Detalle Base',text:'Matriz real generada'}])}${lastBlock}`, 'Presupuestos base reales');
}

function filePicker(id, multiple=true){return `<div class="drop" data-drop="${id}"><strong>Arrastra o selecciona archivos .xlsx</strong><span class="muted small">Solo Excel .xlsx. No PDF, CSV ni ZIP.</span><input id="${id}" type="file" ${multiple?'multiple':''} accept=".xlsx" class="hidden"></div><div id="${id}List" class="file-list"></div>`}
function bindPicker(id, target){
  const drop = $(`[data-drop="${id}"]`), input = $('#'+id), list = $('#'+id+'List');
  drop.onclick = () => input.click();
  input.onchange = () => { state[target] = [...input.files]; renderFileList(list, state[target]); };
  renderFileList(list, state[target]);
}
function renderFileList(list, files){ list.innerHTML = files.length ? files.map(f=>`<div class="file-item"><div><strong>${f.name}</strong><div class="small muted">${(f.size/1024).toFixed(1)} KB</div></div><span class="badge ${f.name.toLowerCase().endsWith('.xlsx')?'ok':'bad'}">${f.name.toLowerCase().endsWith('.xlsx')?'Válido':'No permitido'}</span></div>`).join('') : ''; }


function ensureContractorCount(){
  const min = state.comparisonMode === 'SINGLE' ? 1 : 2;
  if (!state.comparisonContractors.length) state.comparisonContractors = [{ id: Date.now(), name: 'PROV-A', conceptsFile: null, matrixFile: null }];
  while (state.comparisonContractors.length < min) {
    const idx = state.comparisonContractors.length + 1;
    state.comparisonContractors.push({ id: Date.now() + idx, name: `PROV-${String.fromCharCode(64+idx)}`, conceptsFile: null, matrixFile: null });
  }
  if (state.comparisonMode === 'SINGLE' && state.comparisonContractors.length > 1) state.comparisonContractors = state.comparisonContractors.slice(0,1);
}

function xlsxStatus(file){
  if (!file) return '<span class="badge warn">Pendiente</span>';
  return `<span class="badge ${file.name.toLowerCase().endsWith('.xlsx')?'ok':'bad'}">${file.name.toLowerCase().endsWith('.xlsx')?'Válido':'No permitido'}</span>`;
}
function shortProviderName(name){ return String(name || '').trim().slice(0,10); }
function providerNamesParam(){
  const min = state.comparisonMode === 'SINGLE' ? 1 : 2;
  const active = state.comparisonContractors.slice(0, state.comparisonMode === 'SINGLE' ? 1 : state.comparisonContractors.length);
  return encodeURIComponent(active.map((c,i)=> shortProviderName(c.name) || `Prov ${i+1}`).join(','));
}
function reportUrl(){ return state.lastComparisonRun?.downloadUrl || `/api/reports/comparison?providers=${providerNamesParam()}`; }

function contractorCard(c, idx){
  const canRemove = state.comparisonMode === 'MULTI' && state.comparisonContractors.length > 2;
  return `<div class="contractor-card card" data-contractor="${c.id}">
    <div class="contractor-card-head">
      <div><h3>Proveedor ${idx + 1}</h3><p>Registra nombre corto, catálogo de conceptos y matriz/APU.</p></div>
      ${canRemove ? `<button class="btn btn-danger" data-remove-contractor="${c.id}">Eliminar</button>` : ''}
    </div>
    <div class="form-grid three">
      <div>
        <label class="label">Nombre corto proveedor <span class="muted small">máx. 10 caracteres</span></label>
        <input class="input" maxlength="10" data-contractor-name="${c.id}" value="${shortProviderName(c.name) || ''}" placeholder="Ej. PROV-A">
        <div class="small muted">${shortProviderName(c.name).length}/10 · Este nombre aparecerá en Excel.</div>
      </div>
      <div>
        <label class="label">Archivo de conceptos .xlsx</label>
        <div class="mini-upload" data-pick-concepts="${c.id}">
          <strong>${c.conceptsFile ? c.conceptsFile.name : 'Seleccionar conceptos'}</strong>
          <span>Catálogo/lista de conceptos ofertados</span>
        </div>
        <input id="concepts-${c.id}" class="hidden" type="file" accept=".xlsx" data-concepts-input="${c.id}">
        <div class="upload-status">${xlsxStatus(c.conceptsFile)}</div>
      </div>
      <div>
        <label class="label">Archivo matriz/APU .xlsx</label>
        <div class="mini-upload" data-pick-matrix="${c.id}">
          <strong>${c.matrixFile ? c.matrixFile.name : 'Seleccionar matriz/APU'}</strong>
          <span>Detalle de insumos, MO, maquinaria y porcentajes</span>
        </div>
        <input id="matrix-${c.id}" class="hidden" type="file" accept=".xlsx" data-matrix-input="${c.id}">
        <div class="upload-status">${xlsxStatus(c.matrixFile)}</div>
      </div>
    </div>
  </div>`;
}

function bindContractorInputs(){
  $$('[data-contractor-name]').forEach(input => input.oninput = () => {
    const c = state.comparisonContractors.find(x => String(x.id) === String(input.dataset.contractorName));
    if (c) { input.value = shortProviderName(input.value); c.name = input.value; }
  });
  $$('[data-pick-concepts]').forEach(el => el.onclick = () => $(`#concepts-${el.dataset.pickConcepts}`).click());
  $$('[data-pick-matrix]').forEach(el => el.onclick = () => $(`#matrix-${el.dataset.pickMatrix}`).click());
  $$('[data-concepts-input]').forEach(input => input.onchange = () => {
    const c = state.comparisonContractors.find(x => String(x.id) === String(input.dataset.conceptsInput));
    if (c) c.conceptsFile = input.files[0] || null;
    render();
  });
  $$('[data-matrix-input]').forEach(input => input.onchange = () => {
    const c = state.comparisonContractors.find(x => String(x.id) === String(input.dataset.matrixInput));
    if (c) c.matrixFile = input.files[0] || null;
    render();
  });
  $$('[data-remove-contractor]').forEach(btn => btn.onclick = () => {
    state.comparisonContractors = state.comparisonContractors.filter(c => String(c.id) !== String(btn.dataset.removeContractor));
    render();
  });
  const add = $('#addContractor');
  if (add) add.onclick = () => {
    const idx = state.comparisonContractors.length + 1;
    state.comparisonContractors.push({ id: Date.now(), name: `PROV-${String.fromCharCode(64+idx)}`, conceptsFile: null, matrixFile: null });
    render();
  };
}

function baseBudgetsNew(){
  state.lastBaseError = null;
  shell(`${pageHead('Nuevo presupuesto base', 'Crear presupuesto independiente desde archivo de conceptos de ingeniería y matrices Construdata en data.', `<button class="btn btn-secondary" data-nav="base-budgets">Volver</button>`)}<div class="stepper"><span class="step active">1 Datos</span><span class="step active">2 Carga .xlsx</span><span class="step">3 Procesamiento</span><span class="step">4 Resultado</span></div><div class="grid cols-2"><div class="card"><h3>Datos del presupuesto</h3><div class="form-grid"><div><label class="label">Nombre</label><input id="baseProjectName" class="input" value="Presupuesto base real"></div><div><label class="label">Cliente / Obra</label><input class="input" value="Proyecto con catálogo real"></div><div><label class="label">Fuente Construdata detectada</label><input class="input" value="data/construdata_matrices.xlsx" readonly></div></div></div><div class="card"><h3>Archivo de conceptos de ingeniería</h3><p>Este módulo no requiere proveedores ni licitación.</p>${filePicker('baseFiles', false)}<div class="actions" style="margin-top:16px"><button id="runBase" class="btn btn-primary">Generar presupuesto base</button></div></div></div><div class="callout" style="margin-top:16px"><strong>Flujo real:</strong> el archivo se envía a <span class="mono">/api/base-budgets/real-run</span>, se generan matriz, Comparativa, Detalle Base, Validaciones y Análisis IA. No usa endpoints demo.</div>`, 'Nuevo presupuesto base');
  bindPicker('baseFiles','baseUploadFiles');
  $('#runBase').onclick = () => {
    if (!state.baseUploadFiles.length) { alert('Carga un archivo .xlsx de conceptos de ingeniería.'); return; }
    if (state.baseUploadFiles.some(f=>!f.name.toLowerCase().endsWith('.xlsx'))) { alert('Solo se aceptan archivos .xlsx'); return; }
    const projectInput = $('#baseProjectName');
    state.pendingBaseBudget = {projectName: projectInput ? projectInput.value : 'Presupuesto base real', fileName: state.baseUploadFiles[0].name, startedAt: Date.now()};
    go('base-budgets-loading');
  };
}

function baseBudgetLoading(){
  const pending = state.pendingBaseBudget;
  if(!pending || !state.baseUploadFiles.length){ return go('base-budgets-new'); }
  shell(`${pageHead('Generando presupuesto base', 'Procesando archivo real y construyendo el tablero ejecutivo.')}<div class="processing-screen card"><div class="spinner-wrap"><div class="spinner"></div></div><h3>Estamos construyendo el análisis</h3><p>Archivo: <strong>${pending.fileName}</strong></p><div class="progress large"><span></span></div><div class="process-grid"><div class="process-step done">✓ Leyendo conceptos</div><div class="process-step done">✓ Buscando matrices Construdata</div><div class="process-step active">Calculando Detalle Base</div><div class="process-step">Preparando KPIs y gráficas</div><div class="process-step">Generando Excel profesional</div></div><div class="callout" style="margin-top:16px"><strong>Proceso real:</strong> al finalizar se mostrará monto total, costo directo, indirecto, cobertura Construdata, segmentos y descarga del Excel.</div></div>`, 'Generando presupuesto base');
  setTimeout(async () => {
    try{
      const fd = new FormData();
      fd.append('projectName', pending.projectName || 'Presupuesto base real');
      fd.append('concepts_file', state.baseUploadFiles[0]);
      const res = await fetch('/api/base-budgets/real-run', {method:'POST', body:fd});
      if(!res.ok){ const err = await res.json().catch(()=>({detail:'Error al generar presupuesto base'})); throw new Error(err.detail || 'Error al generar presupuesto base'); }
      const data = await res.json();
      state.lastBaseBudgetRun = data;
      state.pendingBaseBudget = null;
      state.lastBaseError = null;
      go('base-budgets-result');
    }catch(err){ state.lastBaseError = err.message; state.pendingBaseBudget = null; go('base-budgets-result'); }
  }, 120);
}

function baseBudgetResult(){
  const real = state.lastBaseBudgetRun;
  if(state.lastBaseError){
    shell(`${pageHead('Error al generar presupuesto base', 'El proceso devolvió una validación o error de procesamiento.', `<button class="btn btn-primary" data-nav="base-budgets-new">Intentar de nuevo</button>`)}<div class="callout"><strong>Detalle:</strong> ${state.lastBaseError}</div>`, 'Error presupuesto base');
    return;
  }
  if(real){
    const summary = summaryFromRun(real);
    const download = summary.downloadUrl || real.downloadUrl || '/api/real-runs/{run_id}/report';
    const coverage = summary.coverage || {};
    const breakdown = summary.costBreakdown || {};
    const breakdownItems = ['materials','labor','equipment','basics','indirect'].map(k=>breakdown[k]).filter(Boolean);
    const segs = summary.segments || [];
    const top = summary.topConcepts || [];
    const findings = summary.executiveFindings || [];
    const maxSeg = Math.max(...segs.map(s=>Number(s.amount||0)),1);
    const rows = [
      ['Estado', summary.status || real.status || 'COMPLETED_WITH_WARNINGS'],
      ['Archivo fuente', summary.sourceFile || real.sourceFile || '—'],
      ['Conceptos leídos', summary.conceptsRead ?? real.concepts ?? 0],
      ['Conceptos ejecutables', summary.conceptsExecutable ?? real.executableConcepts ?? 0],
      ['Filas Detalle Base', summary.detailRows ?? real.apuItems ?? 0],
      ['Cobertura Construdata', `${coverage.matched||0} con match / ${coverage.unmatched||0} en revisión`],
      ['Descarga', `<a href="${download}">${download}</a>`]
    ];
    const topRows = top.slice(0,8).map(c=>[c.code||'—', c.unit||'—', fmtMoney(c.unitPrice), fmtMoney(c.amount), `${Number(c.weightPct||0).toFixed(1)}%`, c.state||'—']);
    shell(`${pageHead('Resultado presupuesto base', 'Tablero ejecutivo generado con datos reales.', `<a class="btn btn-primary" href="${download}">Descargar Excel</a><button class="btn btn-secondary" data-nav="base-budgets-new">Nuevo presupuesto base</button>`)}
      ${kpis([
        {label:'Monto total',value:fmtMoney(summary.totalAmount),text:'Presupuesto base'},
        {label:'Costo directo',value:fmtMoney(summary.directCost),text:'Antes de indirecto'},
        {label:'Indirecto 25%',value:fmtMoney(summary.indirectCost),text:'Regla canónica'},
        {label:'Cobertura',value:fmtPct(coverage.coveragePct),text:`${coverage.matched||0} matches`}
      ])}
      <div class="grid cols-2" style="margin-top:16px">
        <div class="card"><h3>Composición del presupuesto</h3>${donutLegend(breakdownItems)}</div>
        <div class="card"><h3>Distribución por segmento</h3><div class="chart-list">${segs.length?segs.map(s=>chartBar(s.code, s.amount, maxSeg, `${Number(s.weightPct||0).toFixed(1)}%`)).join(''):'<p>Sin segmentos disponibles.</p>'}</div></div>
      </div>
      <div class="grid cols-2" style="margin-top:16px">
        <div class="card"><h3>Hallazgos ejecutivos</h3><ul class="finding-list">${findings.map(f=>`<li>${f}</li>`).join('') || '<li>Sin hallazgos calculados.</li>'}</ul></div>
        <div class="card"><h3>Cobertura de referencias</h3>${kpis([{label:'Con match',value:coverage.matched||0,text:'Construdata'}, {label:'En revisión',value:coverage.unmatched||0,text:'Sin matriz directa'}, {label:'Estimadas',value:coverage.estimated||0,text:'Revisión técnica'}, {label:'Validaciones',value:summary.validations||0,text:'Registros'}])}</div>
      </div>
      <div style="margin-top:16px"><h3>Top conceptos por impacto</h3>${table(['Código','Unidad','P.U.','Importe','% Part.','Estado'], topRows)}</div>
      <div style="margin-top:16px">${table(['Campo','Valor'], rows)}</div>`, 'Resultado presupuesto base');
    return;
  }
  shell(`${pageHead('Resultado presupuesto base', 'Aún no hay una corrida real cargada.', `<button class="btn btn-primary" data-nav="base-budgets-new">Cargar archivo base</button>`)}<div class="callout"><strong>Sin corrida real:</strong> carga un archivo .xlsx de conceptos para generar el presupuesto base con data real.</div>`, 'Resultado presupuesto base');
}

function comparisons(){
  const last = state.lastComparisonRun;
  const lastBlock = last ? `<div style="margin-top:16px">${table(['Campo','Valor'], [['Ultima corrida', last.id], ['Proveedores', (last.providers||[]).length], ['Excel', `<a href="${last.downloadUrl}">Descargar reporte</a>`]])}</div>` : `<div class="callout" style="margin-top:16px"><strong>Sin corrida en esta sesión.</strong> Carga archivos reales de proveedores para generar la comparativa.</div>`;
  shell(`${pageHead('Comparador', 'Comparar propuesta individual o múltiples proveedores. El presupuesto base es opcional.', `<button class="btn btn-primary" data-nav="comparison-new">Nueva comparación</button>`)}${kpis([{label:'Entrada',value:'.xlsx',text:'Conceptos + matriz'}, {label:'Mercado',value:'Construdata',text:'Referencias'}, {label:'Salida',value:'Excel',text:'Comparativa + Detalles'}, {label:'Análisis',value:'Resumen',text:'Hallazgos ejecutivos'}])}${lastBlock}`, 'Comparador');
}

function comparisonNew(){
  ensureContractorCount();
  const min = state.comparisonMode === 'MULTI' ? 2 : 1;
  const contractorCards = state.comparisonContractors.map((c,i)=>contractorCard(c,i)).join('');
  shell(`${pageHead('Nueva comparación', 'Por cada proveedor carga dos archivos .xlsx: conceptos y matriz/APU. El nombre ingresado será el nombre visible en web y Excel.', `<button class="btn btn-secondary" data-nav="comparisons">Volver</button>`)}
  <div class="grid cols-2">
    <div class="card"><h3>Tipo de comparación</h3>
      <div class="tabs"><button class="tab ${state.comparisonMode==='SINGLE'?'active':''}" data-mode="SINGLE">Un proveedor</button><button class="tab ${state.comparisonMode==='MULTI'?'active':''}" data-mode="MULTI">Múltiples proveedors</button></div>
      <p>${state.comparisonMode==='SINGLE'?'No se mostrará ranking económico; se analizará un proveedor contra referencias disponibles.':'Se mostrará ranking económico y comparativa horizontal por proveedor.'}</p>
      <div class="form-grid" style="margin-top:14px"><div><label class="label">Proyecto</label><input class="input" value="Comparativo Sucursal Norte"></div><div><label class="label">Presupuesto base opcional</label><select class="select"><option>No usar presupuesto base</option><option>BB-2026-0001 · Sucursal Norte</option></select></div></div>
    </div>
    <div class="card"><h3>Regla de carga</h3><p>Cada proveedor debe tener su <strong>catálogo de conceptos</strong> y su <strong>matriz/APU</strong>. La matriz/APU alimenta su propio tab de detalle en el Excel.</p><div class="callout" style="margin-top:12px"><strong>Importante:</strong> no todos los proveedors tienen la misma matriz. Por eso el Excel genera un tab de detalle separado por proveedor.</div></div>
  </div>
  <div class="contractor-list" style="margin-top:16px">${contractorCards}</div>
  ${state.comparisonMode==='MULTI'?'<div class="actions" style="margin-top:14px"><button id="addContractor" class="btn btn-secondary">+ Agregar proveedor</button></div>':''}
  <div class="actions" style="margin-top:18px"><button id="runComparison" class="btn btn-primary">Procesar con datos reales</button></div>
  <div class="callout" style="margin-top:16px"><strong>Detalle APU:</strong> por cada proveedor se generará un tab propio usando su matriz/APU declarada + referencias granulares desde data. Los porcentajes se aplican según la base declarada en su archivo: materiales, MO, maquinaria, directo o directo + indirecto.</div>`, 'Nueva comparación');
  $$('.tab').forEach(t=>t.onclick=()=>{state.comparisonMode=t.dataset.mode; ensureContractorCount(); render();});
  bindContractorInputs();
  $('#runComparison').onclick=async()=>{
    ensureContractorCount();
    const required = state.comparisonMode==='MULTI'?2:1;
    const contractors = state.comparisonContractors.slice(0, required === 1 ? 1 : state.comparisonContractors.length);
    if(contractors.length < required){alert(`Agrega mínimo ${required} proveedor(s).`); return;}
    const invalid = contractors.find(c => !shortProviderName(c.name) || shortProviderName(c.name).length > 10 || !c.conceptsFile || !c.matrixFile || !c.conceptsFile.name.toLowerCase().endsWith('.xlsx') || !c.matrixFile.name.toLowerCase().endsWith('.xlsx'));
    if(invalid){alert('Cada proveedor debe tener nombre corto de máximo 10 caracteres, archivo de conceptos .xlsx y archivo matriz/APU .xlsx.'); return;}
    state.comparisonContractors = contractors.map(c => ({...c, name: shortProviderName(c.name)}));
    const btn = $('#runComparison');
    btn.disabled = true; btn.textContent = 'Procesando archivos reales...';
    try{
      const fd = new FormData();
      fd.append('projectName', 'Comparativo real');
      state.comparisonContractors.forEach(c => {
        fd.append('provider_names', shortProviderName(c.name));
        fd.append('concept_files', c.conceptsFile);
        fd.append('matrix_files', c.matrixFile);
      });
      const res = await fetch('/api/comparisons/real-run', {method:'POST', body:fd});
      if(!res.ok){ const err = await res.json().catch(()=>({detail:'Error al procesar'})); throw new Error(err.detail || 'Error al procesar'); }
      state.lastComparisonRun = await res.json();
      go('comparison-results');
    }catch(err){ alert(err.message); btn.disabled = false; btn.textContent = 'Procesar con datos reales'; }
  };
}

function comparisonProcessing(){
  shell(`${pageHead('Generando análisis profesional', 'Procesamiento de archivos XLSX, cálculo APU y preparación del diagnóstico profesional.')}<div class="card"><div class="progress"><span></span></div><div class="grid cols-2" style="margin-top:18px"><div><h3>Etapas</h3><p>✓ Recepción de archivos .xlsx<br>✓ Parser de conceptos<br>✓ Parser de matriz/APU<br>✓ Modelo de análisis<br>✓ Comparación económica inicial<br>✓ Generación de Excel</p></div><div><h3>Mensaje actual</h3><p>Identificando componentes que explican diferencias: materiales, mano de obra, maquinaria, porcentajes e indirectos.</p><button class="btn btn-primary" data-nav="comparison-results">Ver diagnóstico profesional</button></div></div></div>`, 'Procesamiento');
}

function diagnosticReportUrl(url){
  if(!url || url === '#') return '#';
  const sep = url.includes('?') ? '&' : '?';
  return `${url}${sep}theme=${encodeURIComponent(state.theme)}`;
}

async function mountDiagnosticReport(url){
  const frame = $('#diagnostic-report-frame');
  const status = $('#diagnostic-report-status');
  if(!frame) return;
  if(status) status.innerHTML = '<strong>Preparando diagnóstico profesional...</strong><br><span class="muted">Cargando KPIs, tablas de evidencia y plan de revisión.</span>';
  try{
    const themedUrl = diagnosticReportUrl(url);
    const res = await fetch(themedUrl, {cache:'no-store'});
    if(!res.ok) throw new Error(`No fue posible cargar el diagnóstico (${res.status})`);
    const html = await res.text();
    frame.srcdoc = html;
    frame.classList.remove('hidden');
    if(status) status.classList.add('hidden');
  }catch(err){
    frame.classList.add('hidden');
    if(status){
      status.classList.remove('hidden');
      status.innerHTML = `<strong>No fue posible embeber el diagnóstico.</strong><br><span class="muted">${err.message}</span><div class="actions" style="margin-top:12px"><a class="btn btn-primary" href="${diagnosticReportUrl(url)}" target="_blank">Abrir diagnóstico profesional</a></div>`;
    }
  }
}

function comparisonResults(){
  const real = state.lastComparisonRun;
  const multi = state.comparisonMode !== 'SINGLE';
  if(real){
    const providers = real.providers || [];
    const summary = real.summary || {};
    const download = real.downloadUrl || reportUrl();
    const aiReport = summary.aiReportUrl || real.aiReportUrl || (real.id ? `/api/real-runs/${real.id}/ai-report` : '#');
    const themedReport = diagnosticReportUrl(aiReport);
    const single = (summary.providersCount || providers.length) <= 1;
    shell(`${pageHead('Diagnóstico profesional', single?'Lectura individual contra mercado, con evidencias y acciones de revisión.':'Comparativa contra mercado con ranking, evidencias y acciones de revisión.', `<a class="btn btn-primary" href="${download}">Descargar Excel</a><a class="btn btn-secondary" href="${themedReport}" target="_blank">Abrir en nueva pestaña</a>`)}
      <div class="callout" style="margin-bottom:16px"><strong>Resultado principal:</strong> este tablero concentra KPIs, tablas de evidencia, alertas contra mercado, top insumos y plan de revisión para el analista.</div>
      <div id="diagnostic-report-status" class="card" style="margin-bottom:16px"></div>
      <iframe id="diagnostic-report-frame" title="Diagnóstico profesional" class="hidden" style="width:100%;height:calc(100vh - 220px);min-height:820px;border:1px solid var(--line);border-radius:22px;background:var(--panel);box-shadow:var(--shadow);"></iframe>`, 'Diagnóstico profesional');
    mountDiagnosticReport(aiReport);
    return;
  }
  shell(`${pageHead('Diagnóstico profesional', multi?'Genera una comparativa para ver el diagnóstico profesional.':'Genera una comparativa individual para ver el diagnóstico profesional.', `<button class="btn btn-primary" data-nav="comparison-new">Nueva comparación</button>`)}<div class="callout"><strong>Sin corrida real:</strong> carga archivos .xlsx y ejecuta el análisis para generar el diagnóstico profesional.</div>`, 'Diagnóstico profesional');
}

function matrixDetail(){
  ensureContractorCount();
  const rows = state.comparisonContractors.map((c,i)=>[
    c.name || `Proveedor ${i+1}`,
    c.conceptsFile ? c.conceptsFile.name : 'conceptos_'+(i+1)+'.xlsx',
    c.matrixFile ? c.matrixFile.name : 'matriz_apu_'+(i+1)+'.xlsx',
    `Detalle - ${(c.name || `Proveedor ${i+1}`).substring(0,22)}`,
    '<span class="badge ok">Tab individual</span>'
  ]);
  shell(`${pageHead('Detalle de matriz/APU por proveedor', 'Cada proveedor genera su propio tab de detalle porque su matriz/APU puede tener estructura diferente.', `<a class="btn btn-primary" href="${reportUrl()}">Descargar Excel resultado</a>`)}${kpis([{label:'Materiales',value:'data',text:'construdata-materiales'}, {label:'Mano de obra',value:'3',text:'archivos detectados'}, {label:'Maquinaria',value:'data',text:'construdata-maquinaria'}, {label:'Detalle',value:'1 tab',text:'por proveedor'}])}<div class="callout" style="margin-bottom:16px"><strong>Regla:</strong> los nombres de proveedores provienen del textbox corto (máx. 10 caracteres) y el Excel no fuerza un detalle horizontal común. La hoja <span class="mono">Comparativa</span> sí puede ser horizontal por proveedor; el detalle se separa como <span class="mono">Detalle - Proveedor A</span>, <span class="mono">Detalle - Proveedor B</span>, etc.</div>${table(['Proveedor','Archivo conceptos','Archivo matriz/APU','Tab generado','Estado'], rows)}<div class="card" style="margin-top:16px"><h3>Subtotales y porcentajes</h3><p>En cada tab individual se calculan subtotales de Materiales, Mano de Obra y Maquinaria. Los insumos porcentuales se aplican sobre la sección declarada en la matriz del proveedor: % sobre materiales, % sobre MO, % sobre maquinaria, % sobre costo directo o % sobre directo + indirecto.</p></div>`, 'Detalle APU');
}

function myRuns(){
  const rows = [];
  if(state.lastBaseBudgetRun) rows.push(['Hoy', state.lastBaseBudgetRun.projectName || 'Presupuesto base', 'Presupuesto base', state.lastBaseBudgetRun.status || 'Completado', 'Revisión', `<a href="${state.lastBaseBudgetRun.downloadUrl}">Descargar</a>`]);
  if(state.lastComparisonRun) rows.push(['Hoy', state.lastComparisonRun.projectName || 'Comparativa', 'Comparativa', state.lastComparisonRun.status || 'Completado', 'Revisión', `<a href="${state.lastComparisonRun.downloadUrl}">Descargar</a>`]);
  const body = rows.length ? table(['Fecha','Proyecto','Tipo','Estado','Riesgo','Acciones'], rows) : '<div class="callout"><strong>Sin corridas en esta sesión.</strong> Ejecuta un presupuesto base o comparativa para ver historial aquí.</div>';
  shell(`${pageHead('Mis corridas', 'Historial personal de corridas reales y reportes generados.')}${body}`, 'Mis corridas');
}
function teamRuns(){ shell(`${pageHead('Corridas del equipo', 'Historial de equipo disponible para roles administradores.')}<div class="callout"><strong>Historial persistente pendiente.</strong> En esta versión se muestran las corridas de la sesión actual desde Mis corridas.</div>`, 'Corridas equipo'); }
function globalRuns(){ shell(`${pageHead('Historial global', 'Visible únicamente para Super Admin.')}<div class="callout"><strong>Auditoría global pendiente.</strong> La primera versión prioriza generación y descarga de reportes reales.</div>`, 'Historial global'); }

function usersPage(){
  shell(`${pageHead('Administración de usuarios', 'Crear, editar, activar o desactivar usuarios según rol.', `<button class="btn btn-primary" data-nav="user-create">Crear usuario</button>`)}${table(['Nombre','Correo','Rol','Organización','Estado','Acciones'], demoUsers.filter(u=>state.user.role==='SUPER_ADMIN'||u.role==='ANALYST').map(u=>[u.name,u.email,u.role,u.org,`<span class="badge ${u.status==='Activo'?'ok':'warn'}">${u.status}</span>`,`<button class="btn btn-secondary" data-nav="user-edit">Editar</button>`]))}`, 'Usuarios');
}
function userCreate(){ shell(`${pageHead('Crear usuario', 'El Admin solo crea Analistas. El Super Admin puede crear Admins y Analistas.')}<div class="card"><div class="form-grid two"><div><label class="label">Nombre</label><input class="input" value="Nuevo Analista"></div><div><label class="label">Correo</label><input class="input" value="nuevo@empresa.com"></div><div><label class="label">Rol</label><select class="select"><option>ANALYST</option>${state.user.role==='SUPER_ADMIN'?'<option>ADMIN</option>':''}</select></div><div><label class="label">Organización</label><input class="input" value="Empresa Demo"></div></div><div class="actions" style="margin-top:16px"><button class="btn btn-primary" data-nav="users">Crear usuario</button><button class="btn btn-secondary" data-nav="users">Cancelar</button></div></div>`, 'Crear usuario'); }
function userEdit(){ shell(`${pageHead('Editar usuario', 'Cambiar estado, rol permitido o restablecer contraseña.')}<div class="card"><div class="form-grid two"><div><label class="label">Nombre</label><input class="input" value="Analista Demo"></div><div><label class="label">Correo</label><input class="input" value="analista@demo.com"></div><div><label class="label">Rol</label><select class="select"><option>ANALYST</option></select></div><div><label class="label">Estado</label><select class="select"><option>Activo</option><option>Inactivo</option></select></div></div><div class="actions" style="margin-top:16px"><button class="btn btn-primary" data-nav="users">Guardar cambios</button><button class="btn btn-danger">Restablecer contraseña</button></div></div>`, 'Editar usuario'); }

async function ensureRefs(){ if(!state.refs){ const r = await fetch('/api/references'); state.refs = await r.json(); } return state.refs; }
async function referencesPage(){
  shell(`${pageHead('Referencias data', 'Archivos reales detectados desde la carpeta data del proyecto viejo.')}<div class="card"><p>Cargando referencias...</p></div>`, 'Referencias');
  const data = await ensureRefs();
  shell(`${pageHead('Referencias data', 'Archivos reales detectados desde la carpeta data. Se muestran con su uso canónico.')} ${kpis([{label:'Total archivos',value:data.summary.total,text:'data'}, {label:'Matrices base',value:data.summary.baseBudgetMatrix,text:'presupuesto base'}, {label:'Materiales',value:data.summary.materials,text:'detalle APU'}, {label:'MO / Equipo',value:data.summary.labor + data.summary.equipment,text:'detalle APU'}])}<div style="margin-top:16px">${table(['Archivo','Uso canónico','Tipo','Tamaño','Hojas'], data.references.map(r=>[r.name,r.canonicalUse,r.kind,`${(r.sizeBytes/1024/1024).toFixed(2)} MB`,(r.sheets||[]).map(s=>`${s.name} (${s.rows||'?' }x${s.columns||'?'})`).join('<br>')||'—']))}</div>`, 'Referencias');
}
function auditPage(){ shell(`${pageHead('Auditoría', 'Eventos relevantes para Super Admin.')} ${table(['Fecha','Usuario','Acción','Detalle'], [['17/06/2026','admin@demo.com','Creó usuario','analista.norte@demo.com'],['17/06/2026','analista@demo.com','Ejecutó corrida','RUN-2026-0145'],['16/06/2026','superadmin@demo.com','Consultó referencias','data/construdata_matrices.xlsx']])}`, 'Auditoría'); }
function profile(){ shell(`${pageHead('Mi perfil', 'Datos del usuario y cambio de contraseña.')}<div class="card"><div class="form-grid two"><div><label class="label">Nombre</label><input class="input" value="${state.user.name}"></div><div><label class="label">Correo</label><input class="input" value="${state.user.email}" readonly></div><div><label class="label">Rol</label><input class="input" value="${state.user.role}" readonly></div><div><label class="label">Último acceso</label><input class="input" value="17/06/2026 15:30" readonly></div></div><div class="actions" style="margin-top:16px"><button class="btn btn-primary">Guardar</button></div></div>`, 'Mi perfil'); }
function notFound(){ shell(pageHead('No encontrado','Ruta no disponible.'), 'No encontrado'); }

render();
