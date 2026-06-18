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
    { id: 1, name: 'Proveedor A', conceptsFile: null, matrixFile: null },
    { id: 2, name: 'Proveedor B', conceptsFile: null, matrixFile: null },
  ],
};

document.documentElement.dataset.theme = state.theme;

const routes = {
  '': landing,
  'login': login,
  'dashboard': dashboard,
  'base-budgets': baseBudgets,
  'base-budgets-new': baseBudgetsNew,
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

const mockBaseBudgets = [
  {id:'BB-2026-0001', name:'Presupuesto base Sucursal Norte', user:'Analista Demo', total:56520000, concepts:128, status:'Con advertencias', risk:'Medio', date:'17/06/2026'},
  {id:'BB-2026-0002', name:'Presupuesto ampliación bodega', user:'Administrador Demo', total:88400000, concepts:211, status:'Completado', risk:'Bajo', date:'14/06/2026'},
];
const mockRuns = [
  {id:'RUN-2026-0145', name:'Comparativo Sucursal Norte', user:'Analista Demo', type:'Múltiples proveedors', contractors:3, status:'Completa', risk:'Amarillo', date:'17/06/2026'},
  {id:'RUN-2026-0141', name:'Propuesta individual Planta A', user:'Analista Demo', type:'Un proveedor', contractors:1, status:'Completa', risk:'Rojo', date:'15/06/2026'},
  {id:'RUN-2026-0137', name:'Validación preliminar', user:'Administrador Demo', type:'Múltiples proveedors', contractors:4, status:'Fallida', risk:'N/A', date:'10/06/2026'},
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
        <div class="side-head"><div class="logo">Q</div><div><strong>Quantia APU</strong><div class="muted small">Canonical V0</div></div></div>
        <nav class="menu">${menu.map(([h,l])=>`<a href="#/${h}" class="${active===h?'active':''}">${l}<span>›</span></a>`).join('')}</nav>
        <div class="side-foot">
          <div class="small muted">Usuario</div>
          <strong>${state.user.name}</strong>
          <div class="small muted">${state.user.role} · ${state.user.org}</div>
        </div>
      </aside>
      <main class="main">
        <header class="topbar"><div><strong>${title}</strong><div class="small muted">Presupuesto base independiente · Comparador · Detalle APU</div></div><div class="actions"><button class="btn btn-secondary" data-theme-toggle>${state.theme==='dark'?'Tema claro':'Tema oscuro'}</button><button class="btn btn-secondary" data-logout>Salir</button></div></header>
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
  $('#app').innerHTML = `<div class="login-page"><div class="login-card"><div class="brand" style="margin-bottom:18px"><div class="logo">Q</div><div><h1>Quantia APU</h1><p>Canonical V0</p></div></div><h2>Iniciar sesión</h2><p>Usa uno de los usuarios demo para validar navegación, permisos e historial.</p><form id="loginForm" class="form-grid"><div><label class="label">Correo</label><input class="input" name="email" value="superadmin@demo.com"></div><div><label class="label">Contraseña</label><input class="input" name="password" type="password" value="demo123"></div><button class="btn btn-primary">Ingresar</button><div id="loginError" class="bad badge hidden"></div></form><div class="demo-box"><strong>Demo:</strong><br>superadmin@demo.com / demo123<br>admin@demo.com / demo123<br>analista@demo.com / demo123</div></div></div>`;
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

function dashboard(){
  const role = state.user.role;
  const cards = role==='SUPER_ADMIN' ? [
    {label:'Usuarios activos',value:'45',text:'Visibilidad global'}, {label:'Corridas totales',value:'326',text:'Historial completo'}, {label:'Presupuestos base',value:'28',text:'Módulo independiente'}, {label:'Errores recientes',value:'5',text:'Auditoría disponible'}
  ] : role==='ADMIN' ? [
    {label:'Analistas activos',value:'12',text:'Equipo asignado'}, {label:'Corridas mes',value:'48',text:'Bajo su gestión'}, {label:'Con error',value:'3',text:'Requieren revisión'}, {label:'Presupuestos',value:'8',text:'Creados por equipo'}
  ] : [
    {label:'Mis corridas',value:'9',text:'Historial personal'}, {label:'Presupuestos base',value:'2',text:'Independientes'}, {label:'Pendientes',value:'1',text:'Con advertencias'}, {label:'Último riesgo',value:'Amarillo',text:'Comparativo reciente'}
  ];
  shell(`${pageHead('Dashboard', 'Entrada principal según rol. La V0 separa presupuesto base, comparador, referencias e historial.', `<button class="btn btn-primary" data-nav="base-budgets-new">Nuevo presupuesto</button><button class="btn btn-secondary" data-nav="comparison-new">Nueva comparación</button>`)}${kpis(cards)}<div class="grid cols-2" style="margin-top:16px"><div class="card"><h3>Flujo recomendado</h3><p>Presupuesto base es opcional e independiente. El comparador puede ejecutarse sin presupuesto base.</p><div class="actions" style="margin-top:14px"><button class="btn btn-secondary" data-nav="base-budgets">Ver presupuestos</button><button class="btn btn-secondary" data-nav="comparisons">Ver comparador</button></div></div><div class="card"><h3>Regla canónica</h3><p><strong>El sistema calcula.</strong> La IA interpreta. El tablero comunica. En V0 la IA queda mockeada y separada.</p></div></div>`, 'Dashboard');
}

function baseBudgets(){
  shell(`${pageHead('Presupuestos base', 'Módulo independiente: conceptos de ingeniería + data/construdata_matrices.xlsx.', `<button class="btn btn-primary" data-nav="base-budgets-new">Nuevo presupuesto base</button>`)}${kpis([{label:'Total estimado',value:money(144920000),text:'Suma mock'}, {label:'Presupuestos',value:'2',text:'Histórico'}, {label:'Con revisión',value:'1',text:'Matches medios'}, {label:'Fuente matriz',value:'data',text:'Construdata matrices'}])}<div style="margin-top:16px">${table(['ID','Nombre','Usuario','Monto','Conceptos','Estado','Riesgo','Acciones'], mockBaseBudgets.map(b=>[b.id,b.name,b.user,money(b.total),b.concepts,`<span class="badge warn">${b.status}</span>`,b.risk,`<button class="btn btn-secondary" data-nav="base-budgets-result">Ver</button>`]))}</div>`, 'Presupuestos base');
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
function reportUrl(){ return `/api/reports/comparison?providers=${providerNamesParam()}`; }

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
  shell(`${pageHead('Nuevo presupuesto base', 'Crear presupuesto independiente desde archivo de conceptos de ingeniería y matrices Construdata en data.', `<button class="btn btn-secondary" data-nav="base-budgets">Volver</button>`)}<div class="stepper"><span class="step active">1 Datos</span><span class="step active">2 Carga .xlsx</span><span class="step">3 Match Construdata</span><span class="step">4 Resultado</span></div><div class="grid cols-2"><div class="card"><h3>Datos del presupuesto</h3><div class="form-grid"><div><label class="label">Nombre</label><input class="input" value="Presupuesto base Sucursal Norte"></div><div><label class="label">Cliente / Obra</label><input class="input" value="Cliente A / Obra Civil 2026"></div><div><label class="label">Fuente Construdata detectada</label><input class="input" value="data/construdata_matrices.xlsx" readonly></div></div></div><div class="card"><h3>Archivo de conceptos de ingeniería</h3><p>Este módulo no requiere proveedors ni licitación.</p>${filePicker('baseFiles', false)}<div class="actions" style="margin-top:16px"><button id="runBase" class="btn btn-primary">Generar presupuesto base mock</button></div></div></div><div class="callout" style="margin-top:16px"><strong>Separación clave:</strong> este flujo usa <span class="mono">construdata_matrices.xlsx</span>. El detalle APU de proveedors usa materiales/MO/maquinaria desde data y no busca matrices Construdata.</div>`, 'Nuevo presupuesto base');
  bindPicker('baseFiles','baseUploadFiles');
  $('#runBase').onclick = async () => {
    if (!state.baseUploadFiles.length) { alert('Carga un archivo .xlsx de conceptos de ingeniería.'); return; }
    if (state.baseUploadFiles.some(f=>!f.name.toLowerCase().endsWith('.xlsx'))) { alert('Solo se aceptan archivos .xlsx'); return; }
    go('base-budgets-result');
  };
}

function baseBudgetResult(){
  shell(`${pageHead('Resultado presupuesto base', 'Resultado mock: matriz presupuestada independiente lista para descarga y revisión.', `<a class="btn btn-primary" href="/api/reports/base">Descargar Excel base</a>`)}${kpis([{label:'Monto estimado',value:money(56520000),text:'Presupuesto base'}, {label:'Conceptos',value:'128',text:'Archivo ingeniería'}, {label:'Con match',value:'104',text:'Contra matrices Construdata'}, {label:'En revisión',value:'17',text:'Match medio'}])}<div class="grid cols-2" style="margin-top:16px"><div class="card"><h3>Hallazgos IA mock</h3><p>7 conceptos no tienen match directo. 17 conceptos requieren revisión técnica por unidad o descripción ambigua.</p></div><div class="card"><h3>Distribución por familia</h3>${table(['Familia','Importe','Participación'], [['Movimiento de tierras',money(10560000),'18.7%'],['Concreto',money(16800000),'29.7%'],['Acero',money(11160000),'19.7%'],['Acabados',money(18000000),'31.9%']])}</div></div><div style="margin-top:16px">${table(['Código','Concepto ingeniería','Unidad','Cantidad','Concepto Construdata','PU ref.','Importe','Estado'], [['001','Excavación manual','m3','120','Excavación manual material común',money(88000),money(10560000),'<span class="badge ok">Con precio</span>'],['002','Concreto f\'c 250','m3','40','Concreto 250 kg/cm2',money(420000),money(16800000),'<span class="badge ok">Con precio</span>'],['003','Acero de refuerzo','kg','1800','Acero fy 4200',money(6200),money(11160000),'<span class="badge warn">Revisar</span>'],['004','Partida especial','gl','1','Sin referencia','—','—','<span class="badge bad">Sin match</span>']])}</div>`, 'Resultado presupuesto base');
}

function comparisons(){
  shell(`${pageHead('Comparador', 'Comparar propuesta individual o múltiples proveedors. El presupuesto base es opcional.', `<button class="btn btn-primary" data-nav="comparison-new">Nueva comparación</button>`)}${kpis([{label:'Corridas',value:'3',text:'Historial demo'}, {label:'Mejor oferta',value:money(1180000),text:'Última corrida'}, {label:'Partidas críticas',value:'18',text:'Último análisis'}, {label:'Riesgo global',value:'Amarillo',text:'Última corrida'}])}<div style="margin-top:16px">${table(['ID','Proyecto','Tipo','Proveedors','Estado','Riesgo','Acciones'], mockRuns.map(r=>[r.id,r.name,r.type,r.contractors,`<span class="badge ${r.status==='Fallida'?'bad':'ok'}">${r.status}</span>`,r.risk,`<button class="btn btn-secondary" data-nav="comparison-results">Ver</button>`]))}</div>`, 'Comparador');
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
  <div class="actions" style="margin-top:18px"><button id="runComparison" class="btn btn-primary">Validar y procesar mock</button></div>
  <div class="callout" style="margin-top:16px"><strong>Detalle APU:</strong> por cada proveedor se generará un tab propio usando su matriz/APU declarada + referencias granulares desde data. Los porcentajes se aplican según la base declarada en su archivo: materiales, MO, maquinaria, directo o directo + indirecto.</div>`, 'Nueva comparación');
  $$('.tab').forEach(t=>t.onclick=()=>{state.comparisonMode=t.dataset.mode; ensureContractorCount(); render();});
  bindContractorInputs();
  $('#runComparison').onclick=()=>{
    ensureContractorCount();
    const required = state.comparisonMode==='MULTI'?2:1;
    const contractors = state.comparisonContractors.slice(0, required === 1 ? 1 : state.comparisonContractors.length);
    if(contractors.length < required){alert(`Agrega mínimo ${required} proveedor(s).`); return;}
    const invalid = contractors.find(c => !shortProviderName(c.name) || shortProviderName(c.name).length > 10 || !c.conceptsFile || !c.matrixFile || !c.conceptsFile.name.toLowerCase().endsWith('.xlsx') || !c.matrixFile.name.toLowerCase().endsWith('.xlsx'));
    if(invalid){alert('Cada proveedor debe tener nombre corto de máximo 10 caracteres, archivo de conceptos .xlsx y archivo matriz/APU .xlsx.'); return;}
    state.comparisonContractors = state.comparisonContractors.map(c => ({...c, name: shortProviderName(c.name)}));
    go('comparison-processing');
  };
}

function comparisonProcessing(){
  shell(`${pageHead('Procesamiento IA / análisis', 'Procesamiento simulado para validar UX y etapas canónicas.')}<div class="card"><div class="progress"><span></span></div><div class="grid cols-2" style="margin-top:18px"><div><h3>Etapas</h3><p>✓ Lectura de .xlsx<br>✓ Normalización de columnas<br>✓ Homologación de conceptos<br>✓ Comparación económica<br>→ Detalle APU con referencias data<br>→ IA mock de hallazgos<br>○ Generación de Excel</p></div><div><h3>Mensaje actual</h3><p>Identificando componentes que explican diferencias: materiales, mano de obra, maquinaria, porcentajes e indirectos.</p><button class="btn btn-primary" data-nav="comparison-results">Ver resultados mock</button></div></div></div>`, 'Procesamiento');
}

function comparisonResults(){
  const multi = state.comparisonMode !== 'SINGLE';
  const names = state.comparisonContractors.map((c,i)=>shortProviderName(c.name)||`Prov ${i+1}`);
  const n1 = names[0] || 'PROV-A', n2 = names[1] || 'PROV-B', n3 = names[2] || 'PROV-C';
  shell(`${pageHead('Resultados de comparación', multi?'Ranking económico y tablero ejecutivo mock.':'Análisis individual sin ranking económico.', `<a class="btn btn-primary" href="${reportUrl()}">Descargar Excel resultado</a><button class="btn btn-secondary" data-nav="matrix-detail">Ver detalle APU</button>`)}${kpis(multi?[{label:'Mejor oferta',value:money(1180000),text:(state.comparisonContractors[1]?.name || state.comparisonContractors[0]?.name || 'Proveedor')}, {label:'Riesgo global',value:'Amarillo',text:'Con advertencias'}, {label:'Críticas',value:'18',text:'A revisar'}, {label:'Sin referencia',value:'14',text:'Catálogos granulares'}]:[{label:'Monto ofertado',value:money(1250000),text:(state.comparisonContractors[0]?.name || 'Proveedor')}, {label:'Desv. referencia',value:'+12.3%',text:'Sin ranking'}, {label:'Críticas',value:'18',text:'A revisar'}, {label:'Semáforo',value:'Amarillo',text:'Riesgo medio'}])}<div class="grid cols-2" style="margin-top:16px"><div class="card"><h3>Hallazgos IA mock</h3><p>El sobrecosto se concentra en concreto, acero e instalaciones. Se recomienda revisar primero partidas con mayor impacto monetario, no solo las de mayor desviación porcentual.</p></div><div class="card"><h3>Recomendaciones</h3><p>Solicitar desglose APU, validar rendimientos, revisar precios de maquinaria fuera de referencia y negociar porcentajes superiores al rango esperado.</p></div></div><div style="margin-top:16px">${multi?table(['Ranking','Proveedor','Monto','Dif vs menor','Desv. promedio','Semáforo','Acciones'], [['1',n2,money(1180000),'0.0%','-7.8%','<span class="badge ok">Verde</span>','Ver detalle'],['2',n1,money(1250000),'+5.9%','+2.4%','<span class="badge warn">Amarillo</span>','Ver detalle'],['3',n3,money(1410000),'+19.5%','+14.3%','<span class="badge bad">Rojo</span>','Ver detalle']]):table(['Proveedor','Monto','Desv. referencia','Partidas críticas','Semáforo','Observación'], [[n1,money(1250000),'+12.3%','18','<span class="badge warn">Amarillo</span>','Sin ranking por ser análisis individual']])}</div><div style="margin-top:16px">${table(['Partida crítica','Proveedor','PU ofertado','PU ref. data','Dif %','Impacto','Prioridad'], [['Concreto f\'c 250','Proveedor C',money(2850000),money(2300000),'+24%',money(95000),'Alta'],['Acero refuerzo',n1,money(42000),money(35000),'+20%',money(62000),'Alta'],['Luminarias LED',n3,money(1200000),money(980000),'+22%',money(48000),'Media']])}</div>`, 'Resultados');
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

function myRuns(){ shell(`${pageHead('Mis corridas', 'Historial personal del usuario autenticado.')} ${table(['Fecha','Proyecto','Tipo','Estado','Riesgo','Acciones'], mockRuns.filter(r=>state.user.role!=='ANALYST'||r.user==='Analista Demo').map(r=>[r.date,r.name,r.type,`<span class="badge ${r.status==='Fallida'?'bad':'ok'}">${r.status}</span>`,r.risk,'Ver · Descargar · Duplicar']))}`, 'Mis corridas'); }
function teamRuns(){ shell(`${pageHead('Corridas del equipo', 'Visible para Admin y Super Admin.')} ${table(['Fecha','Usuario','Proyecto','Estado','Riesgo','Acciones'], mockRuns.map(r=>[r.date,r.user,r.name,`<span class="badge ${r.status==='Fallida'?'bad':'ok'}">${r.status}</span>`,r.risk,'Ver · Descargar']))}`, 'Corridas equipo'); }
function globalRuns(){ shell(`${pageHead('Historial global', 'Visible únicamente para Super Admin.')} ${table(['Fecha','Usuario','Rol','Proyecto','Estado','Riesgo','Auditoría'], mockRuns.map((r,i)=>[r.date,r.user,i===2?'ADMIN':'ANALYST',r.name,`<span class="badge ${r.status==='Fallida'?'bad':'ok'}">${r.status}</span>`,r.risk,'Ver evento']))}`, 'Historial global'); }

function usersPage(){
  shell(`${pageHead('Administración de usuarios', 'Crear, editar, activar o desactivar usuarios según rol.', `<button class="btn btn-primary" data-nav="user-create">Crear usuario</button>`)}${table(['Nombre','Correo','Rol','Organización','Estado','Acciones'], demoUsers.filter(u=>state.user.role==='SUPER_ADMIN'||u.role==='ANALYST').map(u=>[u.name,u.email,u.role,u.org,`<span class="badge ${u.status==='Activo'?'ok':'warn'}">${u.status}</span>`,`<button class="btn btn-secondary" data-nav="user-edit">Editar</button>`]))}`, 'Usuarios');
}
function userCreate(){ shell(`${pageHead('Crear usuario', 'El Admin solo crea Analistas. El Super Admin puede crear Admins y Analistas.')}<div class="card"><div class="form-grid two"><div><label class="label">Nombre</label><input class="input" value="Nuevo Analista"></div><div><label class="label">Correo</label><input class="input" value="nuevo@empresa.com"></div><div><label class="label">Rol</label><select class="select"><option>ANALYST</option>${state.user.role==='SUPER_ADMIN'?'<option>ADMIN</option>':''}</select></div><div><label class="label">Organización</label><input class="input" value="Empresa Demo"></div></div><div class="actions" style="margin-top:16px"><button class="btn btn-primary" data-nav="users">Crear usuario mock</button><button class="btn btn-secondary" data-nav="users">Cancelar</button></div></div>`, 'Crear usuario'); }
function userEdit(){ shell(`${pageHead('Editar usuario', 'Cambiar estado, rol permitido o restablecer contraseña.')}<div class="card"><div class="form-grid two"><div><label class="label">Nombre</label><input class="input" value="Analista Demo"></div><div><label class="label">Correo</label><input class="input" value="analista@demo.com"></div><div><label class="label">Rol</label><select class="select"><option>ANALYST</option></select></div><div><label class="label">Estado</label><select class="select"><option>Activo</option><option>Inactivo</option></select></div></div><div class="actions" style="margin-top:16px"><button class="btn btn-primary" data-nav="users">Guardar cambios</button><button class="btn btn-danger">Restablecer contraseña</button></div></div>`, 'Editar usuario'); }

async function ensureRefs(){ if(!state.refs){ const r = await fetch('/api/references'); state.refs = await r.json(); } return state.refs; }
async function referencesPage(){
  shell(`${pageHead('Referencias data', 'Archivos reales detectados desde la carpeta data del proyecto viejo.')}<div class="card"><p>Cargando referencias...</p></div>`, 'Referencias');
  const data = await ensureRefs();
  shell(`${pageHead('Referencias data', 'Archivos reales detectados desde la carpeta data. Se muestran con su uso canónico.')} ${kpis([{label:'Total archivos',value:data.summary.total,text:'data'}, {label:'Matrices base',value:data.summary.baseBudgetMatrix,text:'presupuesto base'}, {label:'Materiales',value:data.summary.materials,text:'detalle APU'}, {label:'MO / Equipo',value:data.summary.labor + data.summary.equipment,text:'detalle APU'}])}<div style="margin-top:16px">${table(['Archivo','Uso canónico','Tipo','Tamaño','Hojas'], data.references.map(r=>[r.name,r.canonicalUse,r.kind,`${(r.sizeBytes/1024/1024).toFixed(2)} MB`,(r.sheets||[]).map(s=>`${s.name} (${s.rows||'?' }x${s.columns||'?'})`).join('<br>')||'—']))}</div>`, 'Referencias');
}
function auditPage(){ shell(`${pageHead('Auditoría', 'Eventos relevantes mock para Super Admin.')} ${table(['Fecha','Usuario','Acción','Detalle'], [['17/06/2026','admin@demo.com','Creó usuario','analista.norte@demo.com'],['17/06/2026','analista@demo.com','Ejecutó corrida','RUN-2026-0145'],['16/06/2026','superadmin@demo.com','Consultó referencias','data/construdata_matrices.xlsx']])}`, 'Auditoría'); }
function profile(){ shell(`${pageHead('Mi perfil', 'Datos del usuario y cambio de contraseña mock.')}<div class="card"><div class="form-grid two"><div><label class="label">Nombre</label><input class="input" value="${state.user.name}"></div><div><label class="label">Correo</label><input class="input" value="${state.user.email}" readonly></div><div><label class="label">Rol</label><input class="input" value="${state.user.role}" readonly></div><div><label class="label">Último acceso</label><input class="input" value="17/06/2026 15:30" readonly></div></div><div class="actions" style="margin-top:16px"><button class="btn btn-primary">Guardar mock</button></div></div>`, 'Mi perfil'); }
function notFound(){ shell(pageHead('No encontrado','Ruta no disponible en V0.'), 'No encontrado'); }

render();
