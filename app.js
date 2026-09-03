import { storage } from './db.js';
import { normalizeKey, parseInventoryFiles } from './parser.js';
import { downloadBlob, exportBackup, exportExcel, readBackup, reportMarkup } from './reports.js';

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
const app = $('#app');
const modalRoot = $('#modal-root');
const toastRoot = $('#toast-root');

const MAGIC_PROFILES = [
  ['^MZ01', 'CPU', 'LENOVO', 'THINK CENTRE M75s GEN 5', 'Arrendamiento'],
  ['^VR00', 'MONITOR', 'LENOVO', 'S22I-30', 'Arrendamiento'],
  ['^8SSD51', 'TECLADO', 'LENOVO', 'KU1601', 'Arrendamiento'],
  ['^8SSM51', 'MOUSE', 'LENOVO', 'MOJUUO', 'Arrendamiento'],
  ['^PF[A-Z0-9]{6}', 'LAPTOP', 'LENOVO', 'THINKPAD', 'Arrendamiento'],
  ['^12240', 'REGULADOR DE VOLTAJE', 'SMARTBITT', 'SBNB500', 'Arrendamiento'],
  ['^22WZ', 'TELÉFONO', 'AVAYA', 'VANTAGE 12', 'Cámara'],
  ['^17WZ[A-Z0-9]{8,}', 'TELÉFONO', 'AVAYA', '9611G', 'Cámara']
].map(([regex, descripcion, marca, modelo, possession]) => ({ id: crypto.randomUUID(), regex, descripcion, marca, modelo, possession }));

const SPACE_TYPES = ['ARCHIVO', 'AREA SECRETARIAL', 'BODEGA', 'CUBICULO', 'FOTOCOPIADO', 'PAPELERIA', 'MODULO', 'OFICINA', 'PASILLO', 'RECEPCION', 'SALA DE ESPERA', 'SALA DE JUNTAS', 'COCINA', 'COMEDOR', 'SITE', '(OTRO MANUAL)'];
const BUILDINGS = ['EDIF. A', 'EDIF. B', 'EDIF. C', 'EDIF. E', 'EDIF. F', 'EDIF. G', 'EDIF. H', 'EDIF. I', 'EDIF. CENDI', 'EDIF. TALLERES GRAFICOS', 'EDIF. RESGUARDO Y SEGURIDAD E1', 'EDIF. RESGUARDO Y SEGURIDAD P1', 'ESTACIONAMIENTO 1', 'ESTACIONAMIENTO 2', 'ESTACIONAMIENTO 3', 'ESTACIONAMIENTO 4', 'ESTACIONAMIENTO HELIPUERTO', '(OTRO MANUAL)'];
const FLOORS = ['BASAMENTO', 'PLANTA BAJA', 'PISO 1', 'PISO 2', 'PISO 3', 'PISO 4', 'AZOTEA', 'SOTANO', '(OTRO MANUAL)'];

const iconPaths = {
  inventory: '<path d="M4 7.5 12 3l8 4.5v9L12 21l-8-4.5z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/>',
  users: '<circle cx="12" cy="8" r="3.3"/><path d="M5.2 20c.5-3.5 3-5.4 6.8-5.4s6.3 1.9 6.8 5.4"/>',
  add: '<path d="M12 5v14M5 12h14"/>', notes: '<path d="M6 3h9l3 3v15H6z"/><path d="M9 11h6M9 15h6M9 7h2"/>',
  reports: '<path d="M7 3h8l3 3v15H7z"/><path d="M15 3v4h4M10 12h5M10 16h5"/>',
  scale: '<path d="M12 3v18M5 7h14M6 7l-3 7h6l-3-7m12 0-3 7h6l-3-7"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1-2.2 2.2-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.5v.2h-3.2v-.2a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.9.3l-.1.1-2.2-2.2.1-.1a1.7 1.7 0 0 0 .3-1.9 1.7 1.7 0 0 0-1.5-1H5V11h.2a1.7 1.7 0 0 0 1.5-1 1.7 1.7 0 0 0-.3-1.9l-.1-.1 2.2-2.2.1.1a1.7 1.7 0 0 0 1.9.3 1.7 1.7 0 0 0 1-1.5V4.5h3.2v.2a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.9-.3l.1-.1 2.2 2.2-.1.1a1.7 1.7 0 0 0-.3 1.9 1.7 1.7 0 0 0 1.5 1h.2v3.2h-.2a1.7 1.7 0 0 0-1.5 1Z"/>',
  search: '<circle cx="10.5" cy="10.5" r="5.5"/><path d="m15 15 4 4"/>', upload: '<path d="M12 16V3m0 0L7.5 7.5M12 3l4.5 4.5M5 14.5V20h14v-5.5"/>',
  download: '<path d="M12 3v13m0 0 4.5-4.5M12 16l-4.5-4.5M5 20h14"/>', logout: '<path d="M10 5H5v14h5M14 8l4 4-4 4M18 12H9"/>',
  undo: '<path d="M9 7 5 11l4 4"/><path d="M5 11h8a5 5 0 1 1 0 10h-1"/>', qr: '<path d="M4 4h6v6H4zM14 4h6v6h-6zM4 14h6v6H4zM15 15h2v2h-2zm3 0h2v5h-2zm-3 3h2v2h-2z"/>',
  close: '<path d="m6 6 12 12M18 6 6 18"/>', check: '<path d="m5 12 4 4L19 6"/>', camera: '<path d="M4 8h4l1.4-2h5.2L16 8h4v11H4z"/><circle cx="12" cy="13" r="3.2"/>',
  filter: '<path d="M4 5h16l-6 7v5l-4 2v-7z"/>', pencil: '<path d="m4 16-.5 4.5L8 20l10-10-4-4zM15 7l2 2"/>', warning: '<path d="M12 3 2.8 20h18.4z"/><path d="M12 9v4m0 3h.01"/>',
  map: '<path d="m9 19-5 2.5V5.8L9 3l6 2.8L20 3v15.7l-5 2.5z"/><path d="M9 3v16m6-13.2v15.7"/>', pin: '<path d="M20 10.5c0 5.1-8 10.5-8 10.5S4 15.6 4 10.5a8 8 0 1 1 16 0Z"/><circle cx="12" cy="10.5" r="2.3"/>', trash: '<path d="M4 7h16M10 11v5m4-5v5M9 7l.8-3h4.4l.8 3m-9 0 .8 13h10.4L18 7"/>'
};
function icon(name, className = '') { return `<svg class="${className}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${iconPaths[name] || ''}</svg>`; }
const escapeHtml = (value) => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
const shortName = (name) => name.split(' ').filter(Boolean).slice(0, 2).map((word) => word[0]).join('').slice(0, 2);
const localDate = (value) => new Date(value).toLocaleString('es-MX', { dateStyle: 'medium', timeStyle: 'short' });
const uuid = () => crypto.randomUUID();

let state;
let activeModule = 'inventory';
let selected = new Set();
let history = [];
let inventoryFilters = { query: '', area: '', status: '', page: 1 };
let conciliation = null;
let scannerStream = null;
let layoutUserId = '';
let layoutLocationName = '';
let layoutPlanUrl = '';
let installPrompt = null;
let modalObjectUrls = [];

function emptyState() {
  return {
    schemaVersion: 5, auditor: null, activeUserId: '', inventory: [], users: [], additionalItems: [], notes: [],
    magicProfiles: structuredClone(MAGIC_PROFILES), activity: [], areaNames: {}, layouts: {}, updatedAt: Date.now()
  };
}

function inferSpaceType(value) {
  const label = String(value || '').trim().toLocaleUpperCase('es-MX');
  return SPACE_TYPES.find((type) => label.startsWith(type)) || label.replace(/\s+\d+$/, '') || 'OFICINA';
}

function normalizeLocation(location, details = {}) {
  const raw = typeof location === 'string' ? { name: location, ...details } : { ...(location || {}) };
  const name = String(raw.name || raw.nombre || raw.location || '').trim();
  return {
    ...raw,
    id: raw.id || uuid(),
    name,
    type: String(raw.type || raw.tipoEspacio || inferSpaceType(name)).trim(),
    building: String(raw.building || raw.edificio || '').trim(),
    floor: String(raw.floor || raw.piso || '').trim(),
    photoKey: String(raw.photoKey || '')
  };
}

function normalizeBackup(source = {}) {
  const rawUsers = Array.isArray(source.users) && source.users.length ? source.users : (source.resguardantes || []);
  const users = rawUsers.map((user) => {
    const locations = (user.locations || []).map((location) => normalizeLocation(location, user.locationDetails?.[location] || {})).filter((location) => location.name);
    const savedActiveId = String(user.activeLocationId || '');
    return {
      ...user,
      id: user.id || uuid(),
      name: String(user.name || user.nombre || '').trim(),
      area: String(user.area || '').trim(),
      photoKey: String(user.photoKey || ''),
      locations,
      activeLocationId: locations.some((location) => location.id === savedActiveId) ? savedActiveId : (locations.at(-1)?.id || '')
    };
  });
  const inventoryRows = (source.inventory || []).map((item) => ({
    id: item.id || uuid(), clave: normalizeKey(item.clave ?? item['CLAVE UNICA'] ?? ''),
    descripcion: item.descripcion ?? item.DESCRIPCION ?? item.DESCRripcion ?? '', marca: item.marca ?? item.MARCA ?? '',
    modelo: item.modelo ?? item.MODELO ?? '', serie: item.serie ?? item.SERIE ?? '', area: item.area ?? item.areaOriginal ?? '',
    areaName: item.areaName ?? '', bookType: item.bookType ?? item.listadoOriginal ?? '', source: item.source ?? '',
    status: item.status || (item.UBICADO === 'SI' ? 'ubicado' : 'pendiente'), retag: item.retag ?? item.RE_ETIQUETADO === 'SI',
    userId: item.userId || users.find((user) => user.name === item['NOMBRE DE USUARIO'])?.id || '', location: item.location ?? item.ubicacionEspecifica ?? '',
    note: item.note || '', photoKey: item.photoKey || '', createdAt: item.createdAt || Date.now(), updatedAt: item.updatedAt || Date.now()
  })).filter((item) => item.clave);
  const inventory = [...inventoryRows.reduce((byKey, item) => {
    const previous = byKey.get(item.clave);
    const score = (entry) => Number(Boolean(entry.userId)) * 8 + Number(Boolean(entry.location)) * 4 + Number(entry.status === 'ubicado') * 2 + Number(Boolean(entry.photoKey));
    if (!previous || score(item) > score(previous)) byKey.set(item.clave, item);
    return byKey;
  }, new Map()).values()];
  const additionalItems = (source.additionalItems || []).map((item) => ({
    id: item.id || uuid(), clave: normalizeKey(item.clave ?? item.claveAsignada ?? ''), descripcion: item.descripcion ?? '', marca: item.marca ?? '', modelo: item.modelo ?? '', serie: item.serie ?? '',
    possession: item.possession ?? item.posesion ?? 'Cámara', personal: item.personal === true || item.personal === 'Si', detail: item.detail ?? item.areaProcedencia ?? item.numContrato ?? item.grupoParlamentario ?? '',
    userId: item.userId || users.find((user) => user.name === item.usuario)?.id || '', location: item.location ?? item.ubicacionEspecifica ?? '', createdAt: item.createdAt || Date.now(), updatedAt: item.updatedAt || Date.now(), photoKey: item.photoKey || ''
  }));
  const layouts = Object.fromEntries(Object.entries(source.layouts || {}).map(([userId, layout]) => [userId, { planKey: String(layout?.planKey || ''), pins: Object.fromEntries(Object.entries(layout?.pins || {}).map(([name, pin]) => [name, { x: Number(pin?.x), y: Number(pin?.y) }])) }]));
  return { ...emptyState(), ...source, schemaVersion: 5, users, inventory, additionalItems, notes: source.notes || [], layouts, magicProfiles: source.magicProfiles?.length ? source.magicProfiles : structuredClone(MAGIC_PROFILES), activity: source.activity || [] };
}

async function persist() { state.updatedAt = Date.now(); await storage.save(state); }
function snapshot() { history.push(structuredClone(state)); if (history.length > 10) history.shift(); }
async function mutate(action, message) { snapshot(); action(); addActivity(message); await persist(); render(); }
function addActivity(message) { state.activity.unshift({ id: uuid(), message, at: Date.now() }); state.activity = state.activity.slice(0, 12); }
function notify(message, type = 'success') {
  const el = document.createElement('div');
  el.className = `toast ${type === 'error' ? 'error' : type === 'warning' ? 'warning' : ''}`;
  el.innerHTML = `${icon(type === 'error' || type === 'warning' ? 'warning' : 'check')}<span>${escapeHtml(message)}</span>`;
  toastRoot.append(el); setTimeout(() => el.remove(), 4600);
}

function getActiveUser() { return state.users.find((user) => user.id === state.activeUserId); }
function activeLocationFor(userId) {
  const user = state.users.find((entry) => entry.id === userId);
  const locations = user?.locations || [];
  return locations.find((location) => location.id === user?.activeLocationId) || locations.at(-1) || null;
}
function activateUserLocation(userId, locationId) {
  const user = state.users.find((entry) => entry.id === userId);
  const location = user?.locations?.find((entry) => entry.id === locationId) || activeLocationFor(userId);
  if (!user || !location) return;
  state.activeUserId = userId;
  user.activeLocationId = location.id;
  layoutUserId = userId;
  layoutLocationName = location.name;
}
function personName(userId) { return state.users.find((user) => user.id === userId)?.name || 'Sin asignar'; }
function statuses() { return { total: state.inventory.length, located: state.inventory.filter((item) => item.status === 'ubicado').length, pending: state.inventory.filter((item) => item.status !== 'ubicado').length }; }
function areas() { return [...new Set([...state.inventory.map((item) => item.area), ...Object.keys(state.areaNames || {})].map((area) => String(area || '').trim()).filter(Boolean))].sort(); }
function catalogValues(defaults, property) {
  const values = state.users.flatMap((user) => (user.locations || []).map((location) => location[property]).filter(Boolean));
  return [...new Set([...defaults, ...values].map((value) => String(value).trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es-MX'));
}
function spaceTypes() { return catalogValues(SPACE_TYPES, 'type'); }
function buildings() { return catalogValues(BUILDINGS, 'building'); }
function floors() { return catalogValues(FLOORS, 'floor'); }
function comparableName(value) { return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLocaleUpperCase('es-MX'); }
function userAreaOptions(selected = '') {
  const values = [...new Set([...areas(), ...state.users.map((user) => user.area), selected].map((area) => String(area || '').trim()).filter(Boolean))].sort();
  return `<option value="">Selecciona un área cargada</option>${values.map((area) => `<option value="${escapeHtml(area)}" ${area === selected ? 'selected' : ''}>${escapeHtml(area)} ${escapeHtml(state.areaNames[area] || '')}</option>`).join('')}`;
}
function datalist(id, values) { return `<datalist id="${id}">${values.map((value) => `<option value="${escapeHtml(value)}"></option>`).join('')}</datalist>`; }
function filteredInventory() {
  const query = inventoryFilters.query.trim().toLocaleLowerCase('es-MX');
  return state.inventory.filter((item) => (!inventoryFilters.area || item.area === inventoryFilters.area) && (!inventoryFilters.status || item.status === inventoryFilters.status) && (!query || [item.clave, item.descripcion, item.marca, item.modelo, item.serie, item.areaName, personName(item.userId)].join(' ').toLocaleLowerCase('es-MX').includes(query)));
}
function locationsFor(userId) { return state.users.find((user) => user.id === userId)?.locations || []; }
function nextLocation(label, ignoreLocationId = '', userId = '') {
  const used = locationsFor(userId).filter((location) => location.id !== ignoreLocationId).map((location) => location.name);
  const base = label.trim().toUpperCase() || 'OFICINA';
  const nums = used.map((value) => value.match(new RegExp(`^${base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s+(\\d+)$`))?.[1]).filter(Boolean).map(Number);
  return `${base} ${String((nums.length ? Math.max(...nums) : 0) + 1).padStart(2, '0')}`;
}
function locationValues(form, ignoreLocationId = '', userId = '') {
  const type = String(form.get('spaceType') || '').trim().toLocaleUpperCase('es-MX');
  const customName = String(form.get('locationName') || '').trim().toLocaleUpperCase('es-MX');
  return {
    id: ignoreLocationId || uuid(),
    type: type || 'OFICINA',
    name: customName || nextLocation(type, ignoreLocationId, userId),
    building: String(form.get('building') || '').trim().toLocaleUpperCase('es-MX'),
    floor: String(form.get('floor') || '').trim().toLocaleUpperCase('es-MX'),
    photoKey: ''
  };
}
function userHasLocation(user, locationName, ignoreLocationId = '') { return (user.locations || []).some((location) => location.id !== ignoreLocationId && comparableName(location.name) === comparableName(locationName)); }
function relatedItems(userId, locationName = '') { return [...state.inventory, ...state.additionalItems].filter((item) => item.userId === userId && (!locationName || item.location === locationName)); }
function suggestedSimilarName(name) {
  const base = String(name).trim(); let count = 2; let candidate = `${base} (${count})`;
  const existing = new Set(state.users.map((user) => comparableName(user.name)));
  while (existing.has(comparableName(candidate))) candidate = `${base} (${++count})`;
  return candidate;
}

function loginView() {
  app.innerHTML = `<main class="login-page"><section class="login-card"><img src="./assets/camara-logo.png" alt="Cámara de Diputados, LXVI Legislatura" /><h1>Inventario Legislativo</h1><p>Control de campo, resguardo documental y conciliación de bienes, aun sin conexión.</p><form id="login-form" class="field"><label for="auditor-name">Nombre para esta sesión</label><input id="auditor-name" autocomplete="name" minlength="3" placeholder="Escribe tu nombre" required /><button class="btn" type="submit">Continuar ${icon('logout')}</button></form><p class="login-tip">Este identificador se guarda únicamente en este navegador. La versión pública no incorpora directorios de personal ni autentica usuarios.</p></section></main>`;
  $('#login-form').onsubmit = async (event) => {
    event.preventDefault(); const name = $('#auditor-name').value.trim();
    if (name.length < 3) { notify('Escribe un nombre de al menos tres caracteres.', 'warning'); return; }
    state.auditor = { id: uuid(), name }; addActivity(`Sesión local iniciada por ${name}.`); await persist(); render();
  };
}

const moduleNames = { inventory: 'Inventario', users: 'Resguardantes', layout: 'Croquis de ubicaciones', additional: 'Adicionales', notes: 'Notas', reports: 'Reportes', conciliation: 'Conciliación', settings: 'Configuración' };
function navItem(module, label, glyph) { return `<button class="nav-item ${activeModule === module ? 'active' : ''}" data-action="nav" data-module="${module}" aria-current="${activeModule === module ? 'page' : 'false'}">${icon(glyph, 'nav-icon')}<span>${label}</span></button>`; }
function connectionLabel() { return navigator.onLine ? 'Guardado local · en línea' : 'Guardado local · sin conexión'; }
function shell() {
  const activeUser = getActiveUser();
  const content = ({ inventory: inventoryModule, users: usersModule, layout: layoutModule, additional: additionalModule, notes: notesModule, reports: reportsModule, conciliation: conciliationModule, settings: settingsModule })[activeModule]();
  app.innerHTML = `<div class="app-shell"><aside class="sidebar"><div class="brand"><img src="./assets/camara-logo.png" alt="Escudo de la Cámara de Diputados" /><div class="brand-caption"><strong>Cámara de<br>Diputados</strong><span>LXVI Legislatura</span></div></div><nav class="main-nav" aria-label="Módulos principales">${navItem('inventory', 'Inventario', 'inventory')}${navItem('users', 'Resguardantes', 'users')}${navItem('layout', 'Croquis', 'map')}${navItem('additional', 'Adicionales', 'add')}${navItem('notes', 'Notas', 'notes')}${navItem('reports', 'Reportes', 'reports')}${navItem('conciliation', 'Conciliación', 'scale')}${navItem('settings', 'Configuración', 'settings')}</nav><div class="sidebar-bottom"><button data-action="logout">${icon('logout', 'nav-icon')}<span>Cerrar sesión</span></button></div></aside><main class="workspace"><header class="topbar"><div><h1 class="page-title">Inventario Legislativo</h1><p class="page-subtitle">${escapeHtml(moduleNames[activeModule])}${activeUser ? ` · Resguardante activo: ${escapeHtml(activeUser.name)}${activeLocationFor(activeUser.id) ? ` · Ubicación: ${escapeHtml(activeLocationFor(activeUser.id).name)}` : ''}` : ' · Sin resguardante activo'}</p></div><div class="topbar-actions"><span class="sync-status ${navigator.onLine ? '' : 'offline'}" role="status"><i class="sync-dot"></i>${connectionLabel()}</span>${installPrompt ? `<button class="icon-button install-button" data-action="install-app" title="Instalar aplicación" aria-label="Instalar aplicación">${icon('download')}</button>` : ''}<button class="icon-button" data-action="undo" title="Deshacer" ${history.length ? '' : 'disabled'}>${icon('undo')}</button><div class="auditor"><span class="auditor-avatar">${escapeHtml(shortName(state.auditor.name))}</span><div><strong>${escapeHtml(state.auditor.name)}</strong><span>Auditor en sesión</span></div></div></div></header>${content}</main></div>`;
  app.onclick = handleClick; app.onchange = handleChange; app.oninput = handleInput;
  if (activeModule === 'layout') void hydrateLayoutPlan();
}

function inventoryModule() {
  const counts = statuses(); const all = filteredInventory(); const perPage = 15; const pages = Math.max(1, Math.ceil(all.length / perPage)); inventoryFilters.page = Math.min(inventoryFilters.page, pages);
  const items = all.slice((inventoryFilters.page - 1) * perPage, inventoryFilters.page * perPage);
  const selectAll = items.length && items.every((item) => selected.has(item.id));
  const areaOptions = areas().map((area) => `<option value="${escapeHtml(area)}" ${inventoryFilters.area === area ? 'selected' : ''}>${escapeHtml(area)} ${escapeHtml(state.areaNames[area] || '')}</option>`).join('');
  if (!state.inventory.length) return `<section class="module"><div class="module-panel empty-state">${icon('inventory')}<div><h2>El inventario está listo para recibir su primera fuente</h2><p>Importa los libros Oracle HTML, XLSX/XLSM o CSV. El lector identifica el área, el tipo de libro y conserva las claves con formato decimal.</p><button class="btn" data-action="choose-import">${icon('upload')} Cargar libros de inventario</button><button class="btn ghost" data-action="load-demo">Ver recorrido con datos de ejemplo</button><input id="inventory-file" class="hidden" data-file="inventory" type="file" multiple accept=".xls,.xlsx,.xlsm,.csv,text/html" /></div></div></section>`;
  const rows = items.map((item) => `<tr><td><input aria-label="Seleccionar ${escapeHtml(item.clave)}" class="check" type="checkbox" data-action="toggle-selection" data-id="${item.id}" ${selected.has(item.id) ? 'checked' : ''}></td><td class="mono">${escapeHtml(item.clave)}</td><td><button class="item-title btn ghost" data-action="detail" data-id="${item.id}">${escapeHtml(item.descripcion || 'Sin descripción')}</button><span class="item-meta">${escapeHtml([item.marca, item.modelo, item.serie].filter(Boolean).join(' · ') || 'Sin datos complementarios')}</span></td><td>${escapeHtml(item.area || '—')}</td><td>${escapeHtml(item.location || 'Pendiente')}</td><td><span class="status-pill ${item.status === 'ubicado' ? 'located' : 'pending'}">${item.status === 'ubicado' ? 'Ubicado' : 'Pendiente'}</span></td><td>${escapeHtml(personName(item.userId))}</td><td><button class="icon-button ellipsis" data-action="detail" data-id="${item.id}" aria-label="Abrir detalle">•••</button></td></tr>`).join('');
  return `<section class="module"><div class="section-tabs"><button class="subtab active">Inventario</button><button class="subtab" data-action="set-status" data-status="pendiente">Pendientes por ubicar <span class="tag">${counts.pending}</span></button><button class="subtab" data-action="nav" data-module="notes">Historial</button></div><div class="content-grid"><section class="module-panel"><div class="filters"><div class="field search-field"><label for="inventory-search">Búsqueda global</label>${icon('search')}<input id="inventory-search" value="${escapeHtml(inventoryFilters.query)}" placeholder="Buscar clave, serie o descripción" /></div><div class="field"><label for="area-filter">Área</label><select id="area-filter"><option value="">Todas las áreas</option>${areaOptions}</select></div><div class="field"><label for="status-filter">Estatus</label><select id="status-filter"><option value="">Todos los estatus</option><option value="ubicado" ${inventoryFilters.status === 'ubicado' ? 'selected' : ''}>Ubicado</option><option value="pendiente" ${inventoryFilters.status === 'pendiente' ? 'selected' : ''}>Pendiente</option></select></div><button class="filter-clear" data-action="clear-filters">Limpiar</button></div><div class="bulkbar"><label class="selected-count"><input class="check" type="checkbox" data-action="toggle-page" ${selectAll ? 'checked' : ''}> <b>${selected.size}</b> seleccionados</label><button class="btn secondary small" data-action="assign">${icon('inventory')} Asignar ubicación</button><button class="btn secondary small" data-action="retag">Re-etiquetar</button><button class="btn secondary small" data-action="bulk-note">${icon('notes')} Nota</button><button class="btn secondary small" data-action="scan">${icon('qr')} Escanear QR</button><button class="btn small" data-action="choose-import">${icon('upload')} Cargar libros</button><input id="inventory-file" class="hidden" data-file="inventory" type="file" multiple accept=".xls,.xlsx,.xlsm,.csv,text/html" /></div><div class="table-wrap"><table class="inventory-table"><thead><tr><th></th><th>Clave</th><th>Descripción</th><th>Área</th><th>Ubicación</th><th>Estatus</th><th>Resguardante</th><th></th></tr></thead><tbody>${rows || `<tr><td colspan="8"><div class="empty-state"><div><h2>Sin coincidencias</h2><p>Ajusta la búsqueda o limpia los filtros para ver los bienes cargados.</p></div></div></td></tr>`}</tbody></table></div><div class="table-footer"><span>Mostrando ${items.length ? (inventoryFilters.page - 1) * perPage + 1 : 0}–${Math.min(inventoryFilters.page * perPage, all.length)} de ${all.length} resultados</span><nav class="pager"><button class="page-button" data-action="page" data-page="${inventoryFilters.page - 1}" ${inventoryFilters.page === 1 ? 'disabled' : ''}>‹</button><button class="page-button active">${inventoryFilters.page}</button><button class="page-button" data-action="page" data-page="${inventoryFilters.page + 1}" ${inventoryFilters.page === pages ? 'disabled' : ''}>›</button></nav></div></section>${summaryAside(counts)}</div></section>`;
}

function summaryAside(counts) {
  const activity = state.activity.slice(0, 4).map((entry) => `<div class="activity-item"><span class="activity-icon">${icon('check')}</span><div><strong>${escapeHtml(entry.message)}</strong><time>${localDate(entry.at)}</time></div></div>`).join('') || '<p class="page-subtitle">Aún no hay actividad registrada.</p>';
  return `<aside class="stat-aside"><section class="side-card"><h2>Resumen general</h2><div class="summary-number"><strong>${counts.total.toLocaleString('es-MX')}</strong><span>Total de bienes</span></div><div class="status-list"><div class="status-row"><span><i class="status-dot ok"></i>Ubicados</span><b>${counts.located}</b></div><div class="status-row"><span><i class="status-dot pending"></i>Pendientes</span><b>${counts.pending}</b></div><div class="status-row"><span><i class="status-dot alert"></i>Por re-etiquetar</span><b>${state.inventory.filter((item) => item.retag).length}</b></div></div></section><section class="side-card"><h2>Actividad reciente</h2><div class="activity">${activity}</div></section></aside>`;
}

function locationFieldsMarkup(prefix, location = {}) {
  const type = location.type || inferSpaceType(location.name || 'OFICINA');
  return `${datalist(`${prefix}-space-types`, spaceTypes())}${datalist(`${prefix}-buildings`, buildings())}${datalist(`${prefix}-floors`, floors())}
    <div class="field"><label for="${prefix}-space-type">Tipo de espacio</label><input id="${prefix}-space-type" name="spaceType" list="${prefix}-space-types" required value="${escapeHtml(type)}" placeholder="OFICINA" /></div>
    <div class="field"><label for="${prefix}-location-name">Nombre visible</label><input id="${prefix}-location-name" name="locationName" value="${escapeHtml(location.name || '')}" placeholder="Automático: OFICINA 01" /></div>
    <div class="field"><label for="${prefix}-building">Edificio</label><input id="${prefix}-building" name="building" list="${prefix}-buildings" value="${escapeHtml(location.building || '')}" placeholder="EDIF. A" /></div>
    <div class="field"><label for="${prefix}-floor">Piso</label><input id="${prefix}-floor" name="floor" list="${prefix}-floors" value="${escapeHtml(location.floor || '')}" placeholder="PLANTA BAJA" /></div>`;
}

function usersModule() {
  const rows = state.users.map((user) => {
    const locations = user.locations || []; const latest = locations.at(-1); const activeLocation = activeLocationFor(user.id);
    const chips = locations.map((location) => `<span class="location-chip">${icon('pin')} ${escapeHtml(location.name)}</span>`).join('') || '<span class="location-chip muted">Sin ubicaciones</span>';
    const selector = locations.length ? `<label class="user-location-selector">Ubicación para registrar bienes<select data-user-location-select data-user-id="${user.id}">${locations.map((location) => `<option value="${location.id}" ${location.id === activeLocation?.id ? 'selected' : ''}>${escapeHtml(location.name)}${location.building ? ` · ${escapeHtml(location.building)}` : ''}</option>`).join('')}</select></label>` : '';
    return `<article class="user-row user-card ${user.id === state.activeUserId ? 'active-user' : ''}"><div class="user-identity"><span class="user-avatar">${escapeHtml(shortName(user.name || '?'))}</span><div><strong>${escapeHtml(user.name)}</strong><span>Área ${escapeHtml(user.area || 'sin asignar')} · ${locations.length} ${locations.length === 1 ? 'ubicación' : 'ubicaciones'}${latest ? ` · Última: ${escapeHtml(latest.name)}` : ''}</span>${selector}<div class="location-chips">${chips}</div></div></div><div class="user-actions">${user.id === state.activeUserId ? '<span class="tag">Activo</span>' : ''}<button class="btn secondary small" data-action="edit-user" data-id="${user.id}">${icon('pencil')} Gestionar</button><button class="btn secondary small" data-action="activate-user" data-id="${user.id}">${user.id === state.activeUserId ? 'Quitar' : 'Activar'}</button></div></article>`;
  }).join('') || `<div class="empty-state"><div><h2>Aún no hay resguardantes</h2><p>Crea el primer resguardante y una ubicación para asignar bienes desde el inventario.</p></div></div>`;
  const areaReady = areas().length > 0;
  const importHelp = areaReady ? '' : `<div class="form-tip span-2">Aún no hay áreas en esta página. Carga aquí los listados Oracle o restaura un respaldo: los datos de otra ventana, navegador o dirección no se transfieren automáticamente.<div class="form-actions"><button class="btn secondary small" type="button" data-action="choose-import">${icon('upload')} Cargar listados ahora</button><input id="inventory-file" class="hidden" data-file="inventory" type="file" multiple accept=".xls,.xlsx,.xlsm,.csv,text/html" /></div></div>`;
  return `<section class="module"><div class="two-column"><section class="form-panel"><h2>Nuevo resguardante</h2><p>El área se toma de los listados cargados. Los catálogos de ubicación conservan valores editables y aprenden los nuevos registros.</p><form id="user-form" class="form-grid"><div class="field span-2"><label for="user-name">Nombre completo</label><input id="user-name" name="name" required autocomplete="name" placeholder="Nombre de la persona responsable" /></div><div class="field span-2"><label for="user-area">Área</label><select id="user-area" name="area" required ${areaReady ? '' : 'disabled'}>${userAreaOptions()}</select>${areaReady ? '' : '<small class="field-hint">Carga un listado para habilitar las áreas detectadas.</small>'}</div>${importHelp}${locationFieldsMarkup('user')}<div class="field"><label for="user-photo-new">Foto del resguardante</label><input id="user-photo-new" name="userPhoto" type="file" accept="image/*" capture="user" /></div><div class="field"><label for="location-photo-new">Foto de la primera ubicación</label><input id="location-photo-new" name="locationPhoto" type="file" accept="image/*" capture="environment" /></div><div class="form-actions span-2"><button class="btn" type="submit" ${areaReady ? '' : 'disabled'}>${icon('add')} Crear resguardante</button></div></form></section><section class="module-panel"><div class="panel-heading"><div><h2>Resguardantes registrados</h2><p>${state.users.length} personas · ${state.users.reduce((count, user) => count + (user.locations || []).length, 0)} ubicaciones</p></div></div><div class="users-list">${rows}</div></section></div></section>`;
}

function additionalModule() {
  const rows = state.additionalItems.map((item) => `<article class="additional-card"><span class="additional-mark">+</span><div><strong>${escapeHtml(item.descripcion || 'Bien adicional sin descripción')}</strong><span>${escapeHtml(item.clave || 'Clave pendiente')} · ${escapeHtml(item.possession)} · ${escapeHtml(personName(item.userId))}</span></div><button class="btn secondary small" data-action="additional-detail" data-id="${item.id}">Ver</button></article>`).join('') || `<div class="empty-state"><div><h2>Sin bienes adicionales</h2><p>Registra los bienes de Cámara, arrendamiento, personales o del grupo parlamentario para conciliarlos después.</p></div></div>`;
  const userOptions = `<option value="">Sin asignar</option>${state.users.map((user) => `<option value="${user.id}" ${user.id === state.activeUserId ? 'selected' : ''}>${escapeHtml(user.name)}</option>`).join('')}`;
  return `<section class="module"><div class="adicional-layout"><section class="form-panel"><h2>Registrar bien adicional</h2><p>La serie se valida en tiempo real contra inventario y perfiles de autollenado.</p><form id="additional-form" class="form-grid"><div class="field"><label>Posesión</label><select name="possession" id="ad-possession"><option>Cámara</option><option>Arrendamiento</option><option>Propiedad del Grupo</option></select></div><div class="field"><label>Usuario</label><select name="userId">${userOptions}</select></div><div class="field"><label>Clave única</label><input name="clave" id="ad-clave" placeholder="Automática para Cámara" /></div><div class="field"><label>Serie</label><input name="serie" id="ad-serie" placeholder="Escribe o escanea serie" /></div><div class="field span-2"><label>Descripción</label><input name="descripcion" id="ad-descripcion" required placeholder="Descripción del bien" /></div><div class="field"><label>Marca</label><input name="marca" id="ad-marca" placeholder="Marca" /></div><div class="field"><label>Modelo</label><input name="modelo" id="ad-modelo" placeholder="Modelo" /></div><div class="field span-2"><label>Dato complementario</label><input name="detail" id="ad-detail" placeholder="Área de procedencia, contrato o grupo parlamentario" /></div><label class="selected-count span-2"><input class="check" name="personal" type="checkbox"> Es un bien personal</label><div class="form-actions span-2"><button class="btn secondary" type="button" data-action="magic-fill">Autollenar por serie</button><button class="btn" type="submit">${icon('add')} Registrar adicional</button></div></form></section><section class="module-panel"><div class="panel-heading"><div><h2>Bienes adicionales</h2><p>${state.additionalItems.length} registros independientes del libro Oracle.</p></div><button class="btn secondary small" data-action="bulk-magic">Aplicar perfiles</button></div><div>${rows}</div></section></div></section>`;
}

function notesModule() {
  const notes = [...state.notes].sort((a, b) => b.createdAt - a.createdAt).map((note) => `<article class="note-row"><div><strong>${escapeHtml(note.clave)} · ${escapeHtml(note.title || 'Nota de inventario')}</strong><span>${escapeHtml(note.text)}</span><span>${localDate(note.createdAt)} · ${escapeHtml(note.author)}</span></div><button class="btn secondary small" data-action="detail" data-id="${note.itemId}">Ver bien</button></article>`).join('') || `<div class="empty-state"><div><h2>Sin notas de inventario</h2><p>Agrega observaciones individuales o masivas desde la tabla de inventario.</p></div></div>`;
  return `<section class="module"><section class="module-panel"><div class="panel-heading"><div><h2>Notas y observaciones</h2><p>Bitácora inmutable en la sesión local. Puedes exportar el inventario completo desde Configuración.</p></div><button class="btn" data-action="new-note">${icon('add')} Nueva nota</button></div><div class="notes-list">${notes}</div></section></section>`;
}

function reportsModule() {
  const cards = [
    ['resguardo', 'Resguardo individual', 'Relación de bienes asignados a una persona y sus ubicaciones.', 'users'],
    ['pendientes', 'Bienes pendientes por ubicar', 'Control por área de los bienes que aún requieren validación de campo.', 'warning'],
    ['adicionales', 'Mobiliario y equipo adicional', 'Relación para regularización de bienes fuera del libro fuente.', 'inventory'],
    ['album', 'Álbum fotográfico', 'Prepara una relación imprimible de evidencias vinculadas a bienes y ubicaciones.', 'camera']
  ].map(([type, title, text, glyph]) => `<article class="report-card">${icon(glyph)}<h3>${title}</h3><p>${text}</p><button class="btn secondary small" data-action="make-report" data-report="${type}">Generar reporte</button></article>`).join('');
  return `<section class="module"><div class="report-form"><div class="field"><label>Área del reporte</label><select id="report-area"><option value="">Todas las áreas</option>${areas().map((area) => `<option value="${escapeHtml(area)}">${escapeHtml(area)}</option>`).join('')}</select></div><div class="field"><label>Resguardante</label><select id="report-user"><option value="">Todos los resguardantes</option>${state.users.map((user) => `<option value="${user.id}">${escapeHtml(user.name)}</option>`).join('')}</select></div><button class="btn" data-action="make-report" data-report="resguardo">${icon('reports')} Generar selección</button></div><div class="report-grid">${cards}</div></section>`;
}

function conciliationModule() {
  const diff = conciliation ? `<div class="concil-summary"><div class="concil-stat green"><span>Altas</span><strong>${conciliation.additions.length}</strong></div><div class="concil-stat red"><span>Bajas</span><strong>${conciliation.removals.length}</strong></div><div class="concil-stat copper"><span>Modificados</span><strong>${conciliation.modified.length}</strong></div></div><div class="diff-list">${[...conciliation.additions.map((item) => `<div class="diff-row alta"><div><b>ALTA · ${escapeHtml(item.clave)}</b><small>${escapeHtml(item.descripcion)}</small></div><span class="tag">Nuevo</span></div>`), ...conciliation.removals.map((item) => `<div class="diff-row baja"><div><b>BAJA · ${escapeHtml(item.clave)}</b><small>${escapeHtml(item.descripcion)}</small></div><span class="tag">No encontrado</span></div>`), ...conciliation.modified.map(({ old, next, fields }) => `<div class="diff-row mod"><div><b>CAMBIO · ${escapeHtml(old.clave)}</b><small>${escapeHtml(fields.join(', '))} · ${escapeHtml(next.descripcion)}</small></div><span class="tag">Revisar</span></div>`)].join('') || '<p class="page-subtitle">No se detectaron diferencias.</p>'}</div><div class="form-actions"><button class="btn" data-action="apply-conciliation">Aplicar actualización</button><button class="btn secondary" data-action="export-conciliation">Exportar resumen</button></div>` : `<div class="empty-state"><div>${icon('scale')}<h2>Compara un nuevo corte contra tu inventario</h2><p>Sube uno o varios listados Oracle. Antes de tocar la sesión, la app separa altas, bajas y diferencias de datos para que puedas revisarlas.</p><button class="btn" data-action="choose-conciliation">${icon('upload')} Seleccionar libros para comparar</button><input id="conciliation-file" class="hidden" data-file="conciliation" type="file" multiple accept=".xls,.xlsx,.xlsm,.csv,text/html" /></div></div>`;
  return `<section class="module"><section class="module-panel"><div class="panel-heading"><div><h2>Conciliador de listados</h2><p>Vista previa de cambios antes de agregarlos, modificarlos o darles de baja.</p></div>${conciliation ? '<button class="btn secondary small" data-action="choose-conciliation">Comparar otro corte</button><input id="conciliation-file" class="hidden" data-file="conciliation" type="file" multiple accept=".xls,.xlsx,.xlsm,.csv,text/html" />' : ''}</div>${diff}</section></section>`;
}

function safeCoordinate(value, fallback) { const number = Number(value); return Number.isFinite(number) ? Math.max(5, Math.min(95, number)) : fallback; }
function layoutModule() {
  if (!state.users.length) return `<section class="module"><div class="module-panel empty-state">${icon('map')}<div><h2>Primero registra un resguardante</h2><p>El croquis organiza las ubicaciones declaradas para cada resguardante. Crea el primer registro y vuelve aquí para posicionarlas.</p><button class="btn" data-action="nav" data-module="users">Ir a resguardantes</button></div></div></section>`;
  const user = state.users.find((entry) => entry.id === layoutUserId) || getActiveUser() || state.users[0];
  layoutUserId = user.id;
  const locations = user.locations || [];
  if (!locations.some((location) => location.name === layoutLocationName)) layoutLocationName = activeLocationFor(user.id)?.name || locations[0]?.name || '';
  const layout = state.layouts?.[user.id] || { planKey: '', pins: {} };
  const allItems = [...state.inventory, ...state.additionalItems].filter((item) => item.userId === user.id);
  const userOptions = state.users.map((entry) => `<option value="${entry.id}" ${entry.id === user.id ? 'selected' : ''}>${escapeHtml(entry.name)} · ${escapeHtml(entry.area || 'Sin área')}</option>`).join('');
  const locationOptions = locations.map((location) => `<option value="${escapeHtml(location.name)}" ${location.name === layoutLocationName ? 'selected' : ''}>${escapeHtml(location.name)}${location.building ? ` · ${escapeHtml(location.building)}` : ''}</option>`).join('');
  const pins = locations.map((location, index) => {
    const fallbackX = 22 + ((index * 27) % 57); const fallbackY = 30 + ((index * 19) % 43); const point = layout.pins?.[location.name] || {};
    const x = safeCoordinate(point.x, fallbackX); const y = safeCoordinate(point.y, fallbackY); const count = allItems.filter((item) => item.location === location.name).length;
    return `<button class="layout-pin ${location.name === layoutLocationName ? 'selected' : ''}" style="left:${x}%;top:${y}%" data-action="layout-pin" data-location="${escapeHtml(location.name)}" aria-pressed="${location.name === layoutLocationName}" title="${escapeHtml(location.name)} · ${count} bien${count === 1 ? '' : 'es'}">${icon('pin')}<span>${escapeHtml(location.name)}</span></button>`;
  }).join('');
  const selectedCount = allItems.filter((item) => item.location === layoutLocationName).length;
  const selectedLocation = locations.find((location) => location.name === layoutLocationName);
  return `<section class="module layout-module"><div class="panel-heading layout-heading"><div><h2>Croquis de ubicaciones</h2><p>Selecciona una ubicación y haz clic sobre el plano para colocar o mover su pin. La distribución queda guardada en este equipo.</p></div><div class="layout-kpis"><span><b>${locations.length}</b> ubicaciones</span><span><b>${allItems.filter((item) => item.location).length}</b> bienes ubicados</span><span><b>${Object.keys(layout.pins || {}).length}</b> pines ajustados</span></div></div><div class="layout-toolbar"><div class="field"><label for="layout-user-select">Resguardante</label><select id="layout-user-select">${userOptions}</select></div><div class="field"><label for="layout-location-select">Ubicación a posicionar</label><select id="layout-location-select" ${locations.length ? '' : 'disabled'}><option value="">${locations.length ? 'Selecciona una ubicación' : 'Sin ubicaciones registradas'}</option>${locationOptions}</select></div><div class="layout-toolbar-help">${icon('pin')} Haz clic sobre el plano para mover el pin seleccionado.</div></div><div class="layout-workbench"><section class="layout-board"><div id="layout-canvas" class="layout-canvas" data-action="layout-canvas" data-plan-key="${escapeHtml(layout.planKey || '')}" aria-label="Plano editable de ${escapeHtml(user.name)}"><div class="layout-zones" aria-hidden="true"><span class="zone zone-entry">Entrada</span><span class="zone zone-hall">Pasillo</span><span class="zone zone-office">Área de trabajo</span><span class="zone zone-file">Archivo</span></div>${pins || '<span class="layout-canvas-empty">Registra ubicaciones para comenzar el croquis.</span>'}</div><div class="layout-legend"><span><i></i> Ubicación registrada</span><span><i class="copper"></i> Ubicación seleccionada</span><span>Plano base: ${layout.planKey ? 'cargado' : 'referencia esquemática'}</span></div></section><aside class="layout-inspector"><div class="inspector-title">${icon('pin')} <div><span>Ubicación seleccionada</span><strong>${escapeHtml(layoutLocationName || 'Sin ubicación')}</strong></div></div><dl class="layout-details"><div><dt>Edificio</dt><dd>${escapeHtml(selectedLocation?.building || 'Sin especificar')}</dd></div><div><dt>Planta / piso</dt><dd>${escapeHtml(selectedLocation?.floor || 'Sin especificar')}</dd></div><div><dt>Bienes asignados</dt><dd>${selectedCount}</dd></div></dl><div class="layout-inspector-actions"><button class="btn secondary" data-action="choose-layout-plan">${icon('upload')} Subir plano</button><button class="btn secondary" data-action="remove-layout-plan" ${layout.planKey ? '' : 'disabled'}>${icon('trash')} Quitar plano</button><input id="layout-plan-input" class="hidden" data-file="layout-plan" type="file" accept="image/png,image/jpeg,image/webp" /><button class="btn" data-action="clear-layout-pin" ${layout.pins?.[layoutLocationName] ? '' : 'disabled'}>${icon('pencil')} Restablecer posición</button></div><p class="layout-note">El plano es opcional y se conserva solo en el almacenamiento local del navegador. No se envía a ningún servicio.</p></aside></div></section>`;
}
async function hydrateLayoutPlan() {
  const canvas = $('#layout-canvas'); const key = canvas?.dataset.planKey;
  if (!canvas || !key) return;
  try {
    const blob = await storage.photo(key); if (!blob || !canvas.isConnected) return;
    if (layoutPlanUrl) URL.revokeObjectURL(layoutPlanUrl); layoutPlanUrl = URL.createObjectURL(blob);
    const image = document.createElement('img'); image.className = 'layout-plan-image'; image.src = layoutPlanUrl; image.alt = 'Plano de referencia cargado'; canvas.prepend(image); canvas.classList.add('has-plan');
  } catch (error) { console.warn('No fue posible mostrar el plano local.', error); }
}

function settingsModule() {
  const profiles = state.magicProfiles.map((profile) => `<article class="profile-row"><div><strong>${escapeHtml(profile.descripcion)} · ${escapeHtml(profile.marca)} ${escapeHtml(profile.modelo)}</strong><span><code>${escapeHtml(profile.regex)}</code> · ${escapeHtml(profile.possession)}</span></div><button class="btn secondary small" data-action="remove-profile" data-id="${profile.id}">Eliminar</button></article>`).join('');
  return `<section class="module"><div class="settings-grid"><section class="settings-card"><h2>Respaldo y recuperación</h2><p>El respaldo ZIP contiene el estado de trabajo y permite restaurar o fusionar otra tableta sin enviar datos a un servidor.</p><div class="settings-actions"><button class="btn" data-action="export-backup">${icon('download')} Exportar ZIP</button><button class="btn secondary" data-action="choose-restore">Restaurar ZIP</button><button class="btn secondary" data-action="choose-merge">Fusionar ZIP</button><input id="restore-file" class="hidden" data-file="restore" type="file" accept=".zip" /><input id="merge-file" class="hidden" data-file="merge" type="file" accept=".zip" /></div></section><section class="settings-card"><h2>Exportación de datos</h2><p>Genera un libro Excel compatible, con hojas separadas para inventario y adicionales, listo para auditoría o archivo.</p><div class="settings-actions"><button class="btn" data-action="export-excel">${icon('download')} Exportar Excel</button><button class="btn secondary" data-action="export-reetag">Re-etiquetado</button></div></section><section class="settings-card"><h2>Perfiles de autollenado</h2><p>Reutiliza reglas de serie para prellenar marca, modelo, descripción y posesión de bienes adicionales.</p><form id="profile-form" class="form-grid"><div class="field"><label>Regex de serie</label><input name="regex" required placeholder="^MZ01" /></div><div class="field"><label>Posesión</label><select name="possession"><option>Cámara</option><option>Arrendamiento</option><option>Propiedad del Grupo</option></select></div><div class="field"><label>Descripción</label><input name="descripcion" required /></div><div class="field"><label>Marca / modelo</label><input name="marcaModelo" required placeholder="LENOVO | THINKPAD" /></div><div class="form-actions"><button class="btn" type="submit">Agregar perfil</button></div></form><div class="profile-list">${profiles}</div></section><section class="settings-card danger-card"><h2>Zona de seguridad</h2><p>Borra la sesión local, incluyendo los datos almacenados en este navegador. Exporta un ZIP antes de continuar.</p><div class="settings-actions"><button class="btn danger" data-action="clear-data">Borrar toda la sesión</button></div></section></div></section>`;
}

function render() { if (!state.auditor) loginView(); else shell(); }

function clearModalObjectUrls() { modalObjectUrls.forEach((url) => URL.revokeObjectURL(url)); modalObjectUrls = []; }
function modalPhotoUrl(blob) { if (!blob) return ''; const url = URL.createObjectURL(blob); modalObjectUrls.push(url); return url; }
function openModal(title, body, { wide = false, footer = '<button class="btn secondary" data-action="close-modal">Cerrar</button>' } = {}) {
  modalRoot.innerHTML = `<div class="modal-backdrop" data-action="close-modal"><section class="modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}"><header class="modal-header"><div><h2>${escapeHtml(title)}</h2></div><button class="icon-button" data-action="close-modal" aria-label="Cerrar">${icon('close')}</button></header><div class="modal-body">${body}</div><footer class="modal-footer">${footer}</footer></section></div>`;
}
function closeModal() { stopScanner(); clearModalObjectUrls(); modalRoot.innerHTML = ''; }
function confirmModal(title, text, callback, confirmLabel = 'Confirmar', danger = false) { openModal(title, `<p style="margin:0;color:var(--muted);line-height:1.55">${escapeHtml(text)}</p>`, { footer: `<button class="btn secondary" data-action="close-modal">Cancelar</button><button class="btn ${danger ? 'danger' : ''}" data-action="modal-confirm">${escapeHtml(confirmLabel)}</button>` }); $('#modal-root [data-action="modal-confirm"]').onclick = async () => { closeModal(); await callback(); }; }

async function openUserEditor(userId) {
  const user = state.users.find((entry) => entry.id === userId); if (!user) { notify('El resguardante ya no existe.', 'warning'); return; }
  clearModalObjectUrls();
  const photos = await Promise.all([user.photoKey ? storage.photo(user.photoKey) : null, ...(user.locations || []).map((location) => location.photoKey ? storage.photo(location.photoKey) : null)]);
  const profileUrl = modalPhotoUrl(photos[0]);
  const locations = (user.locations || []).map((location, index) => {
    const photoUrl = modalPhotoUrl(photos[index + 1]); const linked = relatedItems(user.id, location.name).length;
    return `<article class="location-card"><div class="location-thumb">${photoUrl ? `<img src="${photoUrl}" alt="Foto de ${escapeHtml(location.name)}" />` : icon('camera')}</div><div><strong>${escapeHtml(location.name)}</strong><span>${escapeHtml(location.type || 'Espacio')} · ${escapeHtml(location.building || 'Edificio sin especificar')} · ${escapeHtml(location.floor || 'Piso sin especificar')}</span><span>${linked} bien${linked === 1 ? '' : 'es'} asignado${linked === 1 ? '' : 's'}</span></div><div class="location-actions"><button class="btn secondary small" data-action="edit-location" data-user-id="${user.id}" data-location-id="${location.id}">${icon('pencil')} Editar</button><button class="btn secondary small" data-action="choose-location-photo" data-user-id="${user.id}" data-location-id="${location.id}">${icon('camera')} ${photoUrl ? 'Reemplazar' : 'Foto'}</button>${photoUrl ? `<button class="btn secondary small" data-action="remove-location-photo" data-user-id="${user.id}" data-location-id="${location.id}">Quitar foto</button>` : ''}<button class="btn secondary small" data-action="remove-location" data-user-id="${user.id}" data-location-id="${location.id}">Eliminar</button></div></article>`;
  }).join('') || '<p class="page-subtitle">Este resguardante aún no tiene ubicaciones.</p>';
  openModal(`Gestionar resguardante`, `<form id="user-edit-form" class="form-grid" data-user-id="${user.id}"><div class="field span-2"><label for="edit-user-name">Nombre completo</label><input id="edit-user-name" name="name" required value="${escapeHtml(user.name)}" /></div><div class="field span-2"><label for="edit-user-area">Área</label><select id="edit-user-area" name="area" required>${userAreaOptions(user.area)}</select></div><div class="photo-manager span-2"><div class="profile-photo">${profileUrl ? `<img src="${profileUrl}" alt="Foto de ${escapeHtml(user.name)}" />` : `<span>${escapeHtml(shortName(user.name || '?'))}</span>`}</div><div><strong>Foto del resguardante</strong><p>Captura o selecciona una imagen. Se conserva únicamente en este navegador.</p><button class="btn secondary small" type="button" data-action="choose-user-photo" data-user-id="${user.id}">${icon('camera')} ${profileUrl ? 'Reemplazar foto' : 'Agregar foto'}</button>${profileUrl ? `<button class="btn secondary small" type="button" data-action="remove-user-photo" data-user-id="${user.id}">Quitar foto</button>` : ''}</div></div><input id="user-photo-input" class="hidden" type="file" accept="image/*" capture="user" data-user-id="${user.id}" /></form><section class="managed-locations"><div class="managed-locations-heading"><div><h3>Ubicaciones</h3><p>Una persona puede tener todas las ubicaciones que necesite.</p></div><button class="btn small" data-action="add-location" data-user-id="${user.id}">${icon('add')} Agregar ubicación</button></div><div class="managed-locations-list">${locations}</div><input id="location-photo-input" class="hidden" type="file" accept="image/*" capture="environment" /></section>`, { wide: true, footer: `<button class="btn secondary" data-action="close-modal">Cerrar</button><button class="btn danger" data-action="delete-user" data-user-id="${user.id}">Eliminar resguardante</button><button class="btn" type="submit" form="user-edit-form">Guardar cambios</button>` });
}

function openLocationEditor(userId, locationId = '') {
  const user = state.users.find((entry) => entry.id === userId); if (!user) return;
  clearModalObjectUrls();
  const location = (user.locations || []).find((entry) => entry.id === locationId);
  const editing = Boolean(location);
  openModal(editing ? `Editar ${location.name}` : `Agregar ubicación a ${user.name}`, `<form id="location-form" class="form-grid" data-user-id="${user.id}" data-location-id="${location?.id || ''}">${locationFieldsMarkup('location-editor', location || {})}<p class="form-tip span-2">Si dejas vacío el nombre visible, se asignará automáticamente el siguiente consecutivo del tipo de espacio.</p></form>`, { footer: `<button class="btn secondary" data-action="edit-user" data-id="${user.id}">Cancelar</button><button class="btn" type="submit" form="location-form">${editing ? 'Guardar ubicación' : 'Agregar ubicación'}</button>` });
}

function openLocationReassignment(sourceUserId, sourceLocationId) {
  const sourceUser = state.users.find((user) => user.id === sourceUserId);
  const sourceLocation = sourceUser?.locations?.find((location) => location.id === sourceLocationId);
  if (!sourceUser || !sourceLocation) return;
  const linked = relatedItems(sourceUserId, sourceLocation.name);
  const candidates = state.users.filter((user) => (user.locations || []).some((location) => user.id !== sourceUserId || location.id !== sourceLocationId));
  const preferred = candidates.find((user) => user.id === sourceUserId) || candidates[0];
  const userOptions = candidates.map((user) => `<option value="${user.id}" ${user.id === preferred?.id ? 'selected' : ''}>${escapeHtml(user.name)} · ${escapeHtml(user.area || 'Sin área')}</option>`).join('');
  openModal('Reasignar bienes antes de eliminar ubicación', `<p class="page-subtitle">${linked.length} bien${linked.length === 1 ? '' : 'es'} está${linked.length === 1 ? '' : 'n'} asignado${linked.length === 1 ? '' : 's'} a <strong>${escapeHtml(sourceLocation.name)}</strong>. Selecciona otra ubicación del mismo resguardante o de otro; después podrás eliminar esta ubicación.</p><div class="form-grid"><div class="field span-2"><label for="relocate-user">Nuevo resguardante</label><select id="relocate-user">${userOptions || '<option value="">No hay otra ubicación disponible</option>'}</select></div><div class="field span-2"><label for="relocate-location">Nueva ubicación</label><select id="relocate-location"></select></div></div>`, { footer: `<button class="btn secondary" data-action="edit-user" data-id="${sourceUserId}">Cancelar</button><button class="btn" data-action="commit-location-reassignment" data-source-user-id="${sourceUserId}" data-source-location-id="${sourceLocationId}">Reasignar bienes</button>` });
  updateLocationReassignmentOptions();
}

function updateLocationReassignmentOptions() {
  const userId = $('#relocate-user')?.value;
  const select = $('#relocate-location');
  const commit = $('[data-action="commit-location-reassignment"]');
  if (!select || !commit) return;
  const options = locationsFor(userId).filter((location) => !(userId === commit.dataset.sourceUserId && location.id === commit.dataset.sourceLocationId));
  const preferred = activeLocationFor(userId);
  select.innerHTML = options.length ? options.map((location) => `<option value="${location.id}" ${location.id === preferred?.id ? 'selected' : ''}>${escapeHtml(location.name)}${location.building ? ` · ${escapeHtml(location.building)}` : ''}</option>`).join('') : '<option value="">No hay otra ubicación disponible</option>';
  commit.disabled = !options.length;
}

function openDuplicateUserModal(user, proposedName) {
  const suggestion = suggestedSimilarName(proposedName);
  openModal('Resguardante ya registrado', `<p class="page-subtitle">${escapeHtml(user.name)} ya está registrado en el área ${escapeHtml(user.area || 'sin asignar')}. Puedes añadirle otra ubicación o crear un registro similar con un nombre diferenciado.</p>`, { footer: `<button class="btn secondary" data-action="close-modal">Cancelar</button><button class="btn secondary" data-action="duplicate-add-location" data-user-id="${user.id}">${icon('add')} Agregar ubicación</button><button class="btn" data-action="duplicate-similar" data-suggestion="${escapeHtml(suggestion)}">Usar “${escapeHtml(suggestion)}”</button>` });
}

async function showDetail(id, additional = false) {
  const collection = additional ? state.additionalItems : state.inventory; const item = collection.find((entry) => entry.id === id); if (!item) { notify('El registro ya no existe.', 'warning'); return; }
  clearModalObjectUrls();
  const user = state.users.find((entry) => entry.id === item.userId); const photo = item.photoKey ? await storage.photo(item.photoKey) : null; const photoUrl = modalPhotoUrl(photo);
  const fields = additional ? [['Clave', item.clave], ['Descripción', item.descripcion], ['Marca', item.marca], ['Modelo', item.modelo], ['Serie', item.serie], ['Posesión', item.possession], ['Personal', item.personal ? 'Sí' : 'No'], ['Resguardante', user?.name || 'Sin asignar'], ['Ubicación', item.location || 'Sin asignar'], ['Dato complementario', item.detail || '—']] : [['Clave única', item.clave], ['Descripción', item.descripcion], ['Marca', item.marca], ['Modelo', item.modelo], ['Serie', item.serie || '—'], ['Área', `${item.area} ${item.areaName}`], ['Tipo de libro', item.bookType || '—'], ['Estatus', item.status], ['Resguardante', user?.name || 'Sin asignar'], ['Ubicación', item.location || 'Sin asignar'], ['Re-etiquetado', item.retag ? 'Sí' : 'No'], ['Actualizado', localDate(item.updatedAt)]];
  openModal(additional ? 'Detalle de bien adicional' : 'Detalle de bien', `<div class="detail-grid">${fields.map(([label, value]) => `<div><small>${escapeHtml(label)}</small><strong>${escapeHtml(value)}</strong></div>`).join('')}</div><div class="modal-photo">${photoUrl ? `<img src="${photoUrl}" alt="Evidencia del bien" />` : `<span>Sin evidencia fotográfica</span>`}</div><input id="photo-input" class="hidden" type="file" accept="image/*" capture="environment" data-item="${id}" data-additional="${additional}">`, { wide: true, footer: `<button class="btn secondary" data-action="close-modal">Cerrar</button><button class="btn secondary" data-action="choose-photo">${icon('camera')} ${photo ? 'Reemplazar foto' : 'Agregar foto'}</button>${additional ? '' : `<button class="btn" data-action="open-assign" data-id="${id}">Asignar ubicación</button>`}` });
}

function openAssign(ids = [...selected]) {
  const targets = ids.filter((id) => state.inventory.some((item) => item.id === id));
  if (!targets.length) { notify('Selecciona al menos un bien para asignar.', 'warning'); return; }
  const options = state.users.map((user) => `<option value="${user.id}" ${user.id === state.activeUserId ? 'selected' : ''}>${escapeHtml(user.name)} · ${escapeHtml(user.area || 'Sin área')}</option>`).join('');
  openModal('Asignar ubicación', `<p class="page-subtitle">${targets.length} bien${targets.length === 1 ? '' : 'es'} será${targets.length === 1 ? '' : 'n'} marcado${targets.length === 1 ? '' : 's'} como ubicado${targets.length === 1 ? '' : 's'}.</p><div class="form-grid"><div class="field span-2"><label>Resguardante</label><select id="assign-user"><option value="">Selecciona una persona</option>${options}</select></div><div class="field span-2"><label>Ubicación</label><select id="assign-location"><option value="">Primero selecciona un resguardante</option></select></div></div>`, { footer: `<button class="btn secondary" data-action="close-modal">Cancelar</button><button class="btn" data-action="commit-assign" data-ids="${targets.join(',')}">Asignar bienes</button>` });
  $('#assign-user').onchange = () => updateAssignLocations(); updateAssignLocations();
}
function updateAssignLocations() { const userId = $('#assign-user')?.value; const locations = locationsFor(userId); const activeLocation = activeLocationFor(userId); const select = $('#assign-location'); if (select) select.innerHTML = `<option value="">Selecciona una ubicación</option>${locations.map((location) => `<option value="${escapeHtml(location.name)}" ${location.id === activeLocation?.id ? 'selected' : ''}>${escapeHtml(location.name)}${location.building ? ` · ${escapeHtml(location.building)}` : ''}</option>`).join('')}`; }

function openNote(ids = [...selected]) {
  const itemIds = ids.filter((id) => state.inventory.some((item) => item.id === id));
  if (!itemIds.length) { notify('Selecciona al menos un bien para agregar una nota.', 'warning'); return; }
  openModal('Agregar observación', `<div class="field"><label for="note-title">Asunto</label><input id="note-title" placeholder="Ej. Dato por verificar" /></div><div class="field" style="margin-top:14px"><label for="note-text">Nota</label><textarea id="note-text" rows="5" required placeholder="Describe la observación de campo..."></textarea></div>`, { footer: `<button class="btn secondary" data-action="close-modal">Cancelar</button><button class="btn" data-action="commit-note" data-ids="${itemIds.join(',')}">Guardar nota</button>` });
}

function openStandaloneNote() {
  if (!state.inventory.length) { notify('Carga al menos un bien antes de crear una nota.', 'warning'); return; }
  openModal('Nueva nota de inventario', `<div class="field"><label for="note-item">Bien relacionado</label><select id="note-item">${state.inventory.map((item) => `<option value="${item.id}">${escapeHtml(item.clave)} · ${escapeHtml(item.descripcion)}</option>`).join('')}</select></div><div class="field" style="margin-top:14px"><label for="note-title">Asunto</label><input id="note-title" placeholder="Ej. Dato por verificar" /></div><div class="field" style="margin-top:14px"><label for="note-text">Nota</label><textarea id="note-text" rows="5" required placeholder="Describe la observación de campo..."></textarea></div>`, { footer: `<button class="btn secondary" data-action="close-modal">Cancelar</button><button class="btn" data-action="commit-note-picker">Guardar nota</button>` });
}

function demoData() {
  const office = { id: uuid(), name: 'OFICINA 01', type: 'OFICINA', building: 'EDIFICIO DEMO', floor: 'PLANTA BAJA', photoKey: '' };
  const warehouse = { id: uuid(), name: 'BODEGA 01', type: 'BODEGA', building: 'EDIFICIO DEMO', floor: 'PLANTA BAJA', photoKey: '' };
  const users = [{ id: uuid(), name: 'RESPONSABLE DEMO 01', area: '0604500', activeLocationId: office.id, locations: [office] }, { id: uuid(), name: 'RESPONSABLE DEMO 02', area: '0604500', activeLocationId: warehouse.id, locations: [warehouse] }];
  const items = [['22632', 'ARCHIVERO DE MADERA CON 3 GAVETAS', 'SELTA', 'SIN MODELO', 'SIN SERIE 22632'], ['25468', 'ARCHIVERO DE MADERA CON 4 GAVETAS', 'SELTA', 'SIN MODELO A', 'SIN SERIE 25468'], ['.26463', 'CESTO PARA BASURA DE MADERA', 'SIN MARCA', 'LINEA EVOLUTION', ''], ['100976', 'ARCHIVERO DE MADERA CHAPA ENCINO', 'SIN MARCA', 'SIN MODELO', '']].map(([clave, descripcion, marca, modelo, serie], index) => ({ id: uuid(), clave, descripcion, marca, modelo, serie, area: '0604500', areaName: 'SUBDIRECCIÓN DEL CENTRO DE DESARROLLO INFANTIL', bookType: index < 2 ? 'CÁMARA' : 'BIENES MENORES', source: 'Demostración', status: index < 2 ? 'ubicado' : 'pendiente', retag: index === 1, userId: index < 2 ? users[0].id : '', location: index < 2 ? 'OFICINA 01' : '', createdAt: Date.now(), updatedAt: Date.now() }));
  return { users, items };
}

async function handleClick(event) {
  const target = event.target.closest('[data-action]'); if (!target) return; const action = target.dataset.action;
  if (action === 'close-modal') { closeModal(); return; }
  if (action === 'nav') { activeModule = target.dataset.module; selected.clear(); render(); return; }
  if (action === 'logout') { state.auditor = null; await persist(); render(); return; }
  if (action === 'install-app') { if (!installPrompt) { notify('La instalación se habilita cuando el navegador la ofrece.', 'warning'); return; } installPrompt.prompt(); const { outcome } = await installPrompt.userChoice; installPrompt = null; render(); notify(outcome === 'accepted' ? 'La instalación fue solicitada al navegador.' : 'La instalación se dejó pendiente.', outcome === 'accepted' ? 'success' : 'warning'); return; }
  if (action === 'undo') { const previous = history.pop(); if (!previous) return; state = previous; await persist(); notify('Se revirtió la última operación.'); render(); return; }
  if (action === 'choose-import') { $('#inventory-file')?.click(); return; }
  if (action === 'choose-conciliation') { $('#conciliation-file')?.click(); return; }
  if (action === 'choose-restore') { $('#restore-file')?.click(); return; }
  if (action === 'choose-merge') { $('#merge-file')?.click(); return; }
  if (action === 'choose-layout-plan') { $('#layout-plan-input')?.click(); return; }
  if (action === 'load-demo') { const demo = demoData(); await mutate(() => { state.users = demo.users; state.inventory = demo.items; state.activeUserId = demo.users[0].id; state.areaNames['0604500'] = demo.items[0].areaName; }, 'Se cargó el recorrido de demostración.'); notify('Datos de demostración cargados.'); return; }
  if (action === 'toggle-selection') { target.checked ? selected.add(target.dataset.id) : selected.delete(target.dataset.id); render(); return; }
  if (action === 'toggle-page') { filteredInventory().slice((inventoryFilters.page - 1) * 15, inventoryFilters.page * 15).forEach((item) => target.checked ? selected.add(item.id) : selected.delete(item.id)); render(); return; }
  if (action === 'clear-filters') { inventoryFilters = { query: '', area: '', status: '', page: 1 }; render(); return; }
  if (action === 'set-status') { inventoryFilters.status = target.dataset.status; inventoryFilters.page = 1; render(); return; }
  if (action === 'page') { inventoryFilters.page = Math.max(1, Number(target.dataset.page)); render(); return; }
  if (action === 'detail') { await showDetail(target.dataset.id); return; }
  if (action === 'additional-detail') { await showDetail(target.dataset.id, true); return; }
  if (action === 'assign') { openAssign(); return; }
  if (action === 'open-assign') { openAssign([target.dataset.id]); return; }
  if (action === 'commit-assign') { const userId = $('#assign-user').value; const location = $('#assign-location').value; const ids = target.dataset.ids.split(','); if (!userId || !location) { notify('Selecciona resguardante y ubicación.', 'warning'); return; } await mutate(() => { const selectedLocation = locationsFor(userId).find((entry) => entry.name === location); ids.forEach((id) => { const item = state.inventory.find((entry) => entry.id === id); if (item) Object.assign(item, { userId, location, status: 'ubicado', retag: true, updatedAt: Date.now() }); }); activateUserLocation(userId, selectedLocation?.id); selected.clear(); }, `Se asignaron ${ids.length} bien${ids.length === 1 ? '' : 'es'} a ${personName(userId)}.`); closeModal(); return; }
  if (action === 'retag') { const ids = [...selected]; if (!ids.length) { notify('Selecciona bienes para marcarlos para re-etiquetado.', 'warning'); return; } await mutate(() => ids.forEach((id) => { const item = state.inventory.find((entry) => entry.id === id); if (item) item.retag = true; }), `Se marcaron ${ids.length} bien${ids.length === 1 ? '' : 'es'} para re-etiquetado.`); return; }
  if (action === 'bulk-note') { openNote(); return; }
  if (action === 'new-note') { openStandaloneNote(); return; }
  if (action === 'commit-note') { const text = $('#note-text').value.trim(); const title = $('#note-title').value.trim(); const ids = target.dataset.ids.split(','); if (!text) { notify('Escribe una observación.', 'warning'); return; } await mutate(() => { ids.forEach((id) => { const item = state.inventory.find((entry) => entry.id === id); if (item) state.notes.push({ id: uuid(), itemId: id, clave: item.clave, title, text, author: state.auditor.name, createdAt: Date.now() }); }); selected.clear(); }, `Se agregaron ${ids.length} observación${ids.length === 1 ? '' : 'es'} de inventario.`); closeModal(); return; }
  if (action === 'commit-note-picker') { const text = $('#note-text').value.trim(); const title = $('#note-title').value.trim(); const id = $('#note-item').value; const item = state.inventory.find((entry) => entry.id === id); if (!text || !item) { notify('Selecciona un bien y escribe una observación.', 'warning'); return; } await mutate(() => state.notes.push({ id: uuid(), itemId: id, clave: item.clave, title, text, author: state.auditor.name, createdAt: Date.now() }), 'Se agregó una observación de inventario.'); closeModal(); return; }
  if (action === 'edit-user') { await openUserEditor(target.dataset.id); return; }
  if (action === 'add-location') { openLocationEditor(target.dataset.userId); return; }
  if (action === 'edit-location') { openLocationEditor(target.dataset.userId, target.dataset.locationId); return; }
  if (action === 'duplicate-add-location') { openLocationEditor(target.dataset.userId); return; }
  if (action === 'duplicate-similar') { closeModal(); const input = $('#user-name'); if (input) { input.value = target.dataset.suggestion; input.focus(); input.select(); } return; }
  if (action === 'choose-user-photo') { $('#user-photo-input')?.click(); return; }
  if (action === 'choose-location-photo') { const input = $('#location-photo-input'); if (input) { input.dataset.userId = target.dataset.userId; input.dataset.locationId = target.dataset.locationId; input.click(); } return; }
  if (action === 'remove-user-photo') { const user = state.users.find((entry) => entry.id === target.dataset.userId); if (!user?.photoKey) return; const key = user.photoKey; confirmModal('Quitar foto del resguardante', 'La fotografía se eliminará de este dispositivo. Esta acción no modifica los bienes asignados.', async () => { await mutate(() => { const current = state.users.find((entry) => entry.id === user.id); if (current) current.photoKey = ''; }, 'Se quitó la foto del resguardante.'); await storage.deletePhoto(key); await openUserEditor(user.id); }, 'Quitar foto', true); return; }
  if (action === 'remove-location-photo') { const user = state.users.find((entry) => entry.id === target.dataset.userId); const location = user?.locations?.find((entry) => entry.id === target.dataset.locationId); if (!user || !location?.photoKey) return; const key = location.photoKey; confirmModal('Quitar foto de ubicación', `La fotografía de ${location.name} se eliminará de este dispositivo.`, async () => { await mutate(() => { const current = state.users.find((entry) => entry.id === user.id); const targetLocation = current?.locations?.find((entry) => entry.id === location.id); if (targetLocation) targetLocation.photoKey = ''; }, `Se quitó la foto de ${location.name}.`); await storage.deletePhoto(key); await openUserEditor(user.id); }, 'Quitar foto', true); return; }
  if (action === 'remove-location') { const user = state.users.find((entry) => entry.id === target.dataset.userId); const location = user?.locations?.find((entry) => entry.id === target.dataset.locationId); if (!user || !location) return; const linked = relatedItems(user.id, location.name); if (linked.length) { openLocationReassignment(user.id, location.id); return; } const photoKey = location.photoKey; confirmModal('Eliminar ubicación', `Se eliminará ${location.name} del resguardante. No hay bienes vinculados a esta ubicación.`, async () => { await mutate(() => { const current = state.users.find((entry) => entry.id === user.id); if (!current) return; current.locations = current.locations.filter((entry) => entry.id !== location.id); current.activeLocationId = current.locations.at(-1)?.id || ''; const layout = state.layouts?.[current.id]; if (layout?.pins?.[location.name]) { const pins = { ...layout.pins }; delete pins[location.name]; state.layouts[current.id] = { ...layout, pins }; } layoutUserId = current.id; layoutLocationName = activeLocationFor(current.id)?.name || ''; }, `Se eliminó la ubicación ${location.name}.`); if (photoKey) await storage.deletePhoto(photoKey); await openUserEditor(user.id); }, 'Eliminar ubicación', true); return; }
  if (action === 'commit-location-reassignment') { const sourceUserId = target.dataset.sourceUserId; const sourceLocationId = target.dataset.sourceLocationId; const sourceUser = state.users.find((user) => user.id === sourceUserId); const sourceLocation = sourceUser?.locations?.find((location) => location.id === sourceLocationId); const targetUserId = $('#relocate-user')?.value; const targetLocationId = $('#relocate-location')?.value; const targetUser = state.users.find((user) => user.id === targetUserId); const targetLocation = targetUser?.locations?.find((location) => location.id === targetLocationId); if (!sourceUser || !sourceLocation || !targetUser || !targetLocation) { notify('Selecciona una ubicación de destino válida.', 'warning'); return; } let moved = 0; await mutate(() => { [...state.inventory, ...state.additionalItems].forEach((item) => { if (item.userId === sourceUserId && item.location === sourceLocation.name) { Object.assign(item, { userId: targetUser.id, location: targetLocation.name, status: item.status === 'pendiente' ? 'ubicado' : item.status, updatedAt: Date.now() }); moved++; } }); activateUserLocation(targetUser.id, targetLocation.id); }, `Se reasignaron ${moved} bien${moved === 1 ? '' : 'es'} a ${targetUser.name} · ${targetLocation.name}.`); closeModal(); await openUserEditor(sourceUserId); return; }
  if (action === 'delete-user') { const user = state.users.find((entry) => entry.id === target.dataset.userId); if (!user) return; const linked = relatedItems(user.id); const photoKeys = [user.photoKey, ...(user.locations || []).map((location) => location.photoKey), state.layouts?.[user.id]?.planKey].filter(Boolean); confirmModal('Eliminar resguardante', `${user.name} se eliminará. Sus ${linked.length} bien${linked.length === 1 ? '' : 'es'} quedarán sin resguardante y, en inventario, volverán a pendiente. También se eliminarán sus fotografías locales.`, async () => { await mutate(() => { state.inventory.forEach((item) => { if (item.userId === user.id) Object.assign(item, { userId: '', location: '', status: 'pendiente', updatedAt: Date.now() }); }); state.additionalItems.forEach((item) => { if (item.userId === user.id) Object.assign(item, { userId: '', location: '', updatedAt: Date.now() }); }); state.users = state.users.filter((entry) => entry.id !== user.id); if (state.activeUserId === user.id) state.activeUserId = ''; if (layoutUserId === user.id) { layoutUserId = ''; layoutLocationName = ''; } delete state.layouts[user.id]; }, `Se eliminó a ${user.name} y se desvincularon ${linked.length} bienes.`); await Promise.all(photoKeys.map((key) => storage.deletePhoto(key))); }, 'Eliminar resguardante', true); return; }
  if (action === 'activate-user') { const id = target.dataset.id; const willDeactivate = state.activeUserId === id; await mutate(() => { if (willDeactivate) state.activeUserId = ''; else activateUserLocation(id, activeLocationFor(id)?.id); }, willDeactivate ? 'Se desactivó el resguardante de contexto.' : `Se activó a ${personName(id)} y su ubicación seleccionada.`); return; }
  if (action === 'magic-fill') { magicFillForm(); return; }
  if (action === 'bulk-magic') { let changes = 0; await mutate(() => { state.additionalItems.forEach((item) => { const profile = matchProfile(item.serie); if (profile) { Object.assign(item, { descripcion: profile.descripcion, marca: profile.marca, modelo: profile.modelo, possession: profile.possession, updatedAt: Date.now() }); changes++; } }); }, 'Se aplicaron perfiles de autollenado.'); notify(changes ? `${changes} adicional${changes === 1 ? '' : 'es'} actualizado${changes === 1 ? '' : 's'}.` : 'No hubo series que coincidieran con un perfil.', changes ? 'success' : 'warning'); return; }
  if (action === 'make-report') { const filters = { area: $('#report-area')?.value || '', userId: $('#report-user')?.value || '' }; $('#print-root').innerHTML = reportMarkup(state, target.dataset.report, filters); window.print(); return; }
  if (action === 'apply-conciliation') { if (!conciliation) return; confirmModal('Aplicar conciliación', `Se agregarán ${conciliation.additions.length} altas, se actualizarán ${conciliation.modified.length} registros y se darán de baja ${conciliation.removals.length} bienes. Esta acción puede deshacerse solo durante esta sesión.`, async () => { await mutate(() => { const removeIds = new Set(conciliation.removals.map((item) => item.id)); state.inventory = state.inventory.filter((item) => !removeIds.has(item.id)); conciliation.modified.forEach(({ old, next }) => Object.assign(old, { ...next, id: old.id, status: old.status, userId: old.userId, location: old.location, retag: old.retag, updatedAt: Date.now() })); state.inventory.push(...conciliation.additions); conciliation = null; }, 'Se aplicó la conciliación de listados.'); }); return; }
  if (action === 'export-conciliation') { exportConciliation(); return; }
  if (action === 'export-backup') { try { await exportBackup(state); notify('Respaldo ZIP generado.'); } catch (error) { notify(error.message, 'error'); } return; }
  if (action === 'export-excel') { exportExcel(state); notify('Libro Excel compatible generado.'); return; }
  if (action === 'export-reetag') { exportRetag(); return; }
  if (action === 'remove-profile') { const id = target.dataset.id; await mutate(() => { state.magicProfiles = state.magicProfiles.filter((profile) => profile.id !== id); }, 'Se eliminó un perfil de autollenado.'); return; }
  if (action === 'clear-data') { confirmModal('Borrar toda la sesión', 'Esta operación elimina datos y fotografías guardados localmente en este navegador. Exporta un respaldo ZIP antes de continuar.', async () => { await storage.clear(); state = emptyState(); history = []; selected.clear(); conciliation = null; notify('La sesión local fue eliminada.'); render(); }, 'Borrar sesión', true); return; }
  if (action === 'choose-photo') { $('#photo-input')?.click(); return; }
  if (action === 'layout-pin') { layoutLocationName = target.dataset.location; render(); return; }
  if (action === 'layout-canvas') { const user = state.users.find((entry) => entry.id === layoutUserId); if (!user || !layoutLocationName) { notify('Selecciona una ubicación antes de posicionarla.', 'warning'); return; } const rect = target.getBoundingClientRect(); const x = Math.max(5, Math.min(95, ((event.clientX - rect.left) / rect.width) * 100)); const y = Math.max(5, Math.min(95, ((event.clientY - rect.top) / rect.height) * 100)); await mutate(() => { state.layouts ||= {}; const current = state.layouts[user.id] || { planKey: '', pins: {} }; state.layouts[user.id] = { ...current, pins: { ...(current.pins || {}), [layoutLocationName]: { x, y } } }; }, `Se actualizó el pin de ${layoutLocationName} en el croquis.`); return; }
  if (action === 'clear-layout-pin') { const user = state.users.find((entry) => entry.id === layoutUserId); if (!user || !layoutLocationName) return; await mutate(() => { const current = state.layouts?.[user.id]; if (!current) return; const pins = { ...(current.pins || {}) }; delete pins[layoutLocationName]; state.layouts[user.id] = { ...current, pins }; }, `Se restableció la posición de ${layoutLocationName}.`); return; }
  if (action === 'remove-layout-plan') { const user = state.users.find((entry) => entry.id === layoutUserId); const planKey = user && state.layouts?.[user.id]?.planKey; if (!user || !planKey) return; confirmModal('Quitar plano de referencia', 'Se eliminará la imagen del plano guardada en este navegador. Los pines y las ubicaciones no se modificarán.', async () => { await storage.deletePhoto(planKey); await mutate(() => { const current = state.layouts?.[user.id] || { pins: {} }; state.layouts[user.id] = { ...current, planKey: '' }; }, 'Se eliminó el plano local de referencia.'); }, 'Quitar plano', true); return; }
  if (action === 'scan') { await openScanner(); return; }
}

async function handleChange(event) {
  if (event.target.id === 'area-filter') { inventoryFilters.area = event.target.value; inventoryFilters.page = 1; render(); return; }
  if (event.target.id === 'status-filter') { inventoryFilters.status = event.target.value; inventoryFilters.page = 1; render(); return; }
  if (event.target.id === 'assign-user') { updateAssignLocations(); return; }
  if (event.target.id === 'relocate-user') { updateLocationReassignmentOptions(); return; }
  if (event.target.matches('[data-user-location-select]')) { const userId = event.target.dataset.userId; const locationId = event.target.value; const location = locationsFor(userId).find((entry) => entry.id === locationId); if (!location) return; await mutate(() => activateUserLocation(userId, locationId), `Se dejó ${location.name} como ubicación activa de ${personName(userId)}.`); return; }
  if (event.target.id === 'layout-user-select') { layoutUserId = event.target.value; layoutLocationName = activeLocationFor(layoutUserId)?.name || locationsFor(layoutUserId)[0]?.name || ''; render(); return; }
  if (event.target.id === 'layout-location-select') { layoutLocationName = event.target.value; render(); return; }
  if (event.target.matches('[data-file]')) { await handleFile(event.target.dataset.file, event.target.files); event.target.value = ''; return; }
  if (event.target.id === 'user-photo-input') { const userId = event.target.dataset.userId; const saved = await saveUserPhotoFile(userId, event.target.files?.[0]); if (saved) await openUserEditor(userId); event.target.value = ''; return; }
  if (event.target.id === 'location-photo-input') { const userId = event.target.dataset.userId; const saved = await saveLocationPhotoFile(userId, event.target.dataset.locationId, event.target.files?.[0]); if (saved) await openUserEditor(userId); event.target.value = ''; return; }
  if (event.target.id === 'photo-input') { await savePhoto(event.target); return; }
}
function handleInput(event) {
  if (event.target.id === 'inventory-search') {
    const caret = event.target.selectionStart;
    inventoryFilters.query = event.target.value; inventoryFilters.page = 1; render();
    requestAnimationFrame(() => { const input = $('#inventory-search'); if (input) { input.focus(); input.setSelectionRange(caret, caret); } });
  }
}

function selectedPhoto(form, name) { const file = form.get(name); return file && typeof file === 'object' && file.size ? file : null; }
function validPhoto(file) {
  if (!file) return true;
  if (!file.type?.startsWith('image/')) { notify('Selecciona una imagen válida.', 'warning'); return false; }
  if (file.size > 8 * 1024 * 1024) { notify('La fotografía supera el límite de 8 MB.', 'warning'); return false; }
  return true;
}
async function saveUserPhotoFile(userId, file) {
  if (!validPhoto(file)) return false;
  const user = state.users.find((entry) => entry.id === userId); if (!user || !file) return Boolean(user);
  const key = user.photoKey || `user-${user.id}`;
  await storage.savePhoto(key, file);
  await mutate(() => { const current = state.users.find((entry) => entry.id === userId); if (current) current.photoKey = key; }, `Se actualizó la foto de ${user.name}.`);
  return true;
}
async function saveLocationPhotoFile(userId, locationId, file) {
  if (!validPhoto(file)) return false;
  const user = state.users.find((entry) => entry.id === userId); const location = user?.locations?.find((entry) => entry.id === locationId); if (!user || !location || !file) return Boolean(location);
  const key = location.photoKey || `location-${user.id}-${location.id}`;
  await storage.savePhoto(key, file);
  await mutate(() => { const current = state.users.find((entry) => entry.id === userId); const target = current?.locations?.find((entry) => entry.id === locationId); if (target) target.photoKey = key; }, `Se actualizó la foto de ${location.name}.`);
  return true;
}
function renameLocationReferences(userId, previousName, nextName) {
  if (previousName === nextName) return 0;
  let changed = 0;
  [...state.inventory, ...state.additionalItems].forEach((item) => {
    if (item.userId === userId && item.location === previousName) { item.location = nextName; item.updatedAt = Date.now(); changed++; }
  });
  const layout = state.layouts?.[userId];
  if (layout?.pins?.[previousName]) {
    const pins = { ...layout.pins, [nextName]: layout.pins[previousName] }; delete pins[previousName]; state.layouts[userId] = { ...layout, pins };
  }
  if (layoutUserId === userId && layoutLocationName === previousName) layoutLocationName = nextName;
  return changed;
}

document.addEventListener('submit', async (event) => {
  if (event.target.id === 'user-form') {
    event.preventDefault(); const form = new FormData(event.target); const name = String(form.get('name')).replace(/\s+/g, ' ').trim(); const area = String(form.get('area')).trim(); if (!name || !area) return;
    const existing = state.users.find((user) => comparableName(user.name) === comparableName(name)); if (existing) { openDuplicateUserModal(existing, name); return; }
    const user = { id: uuid(), name, area, photoKey: '', activeLocationId: '', locations: [] }; const location = locationValues(form, '', user.id); user.locations = [location]; user.activeLocationId = location.id;
    const userPhoto = selectedPhoto(form, 'userPhoto'); const locationPhoto = selectedPhoto(form, 'locationPhoto'); if (!validPhoto(userPhoto) || !validPhoto(locationPhoto)) return;
    await mutate(() => { state.users.push(user); activateUserLocation(user.id, location.id); }, `Se registró a ${name} con ${location.name}.`);
    if (userPhoto) await saveUserPhotoFile(user.id, userPhoto);
    if (locationPhoto) await saveLocationPhotoFile(user.id, location.id, locationPhoto);
    return;
  }
  if (event.target.id === 'user-edit-form') {
    event.preventDefault(); const form = new FormData(event.target); const user = state.users.find((entry) => entry.id === event.target.dataset.userId); if (!user) return;
    const name = String(form.get('name')).replace(/\s+/g, ' ').trim(); const area = String(form.get('area')).trim(); if (!name || !area) return;
    const duplicate = state.users.find((entry) => entry.id !== user.id && comparableName(entry.name) === comparableName(name)); if (duplicate) { notify(`Ya existe ${duplicate.name}. Usa un nombre diferenciado o agrega una ubicación a ese registro.`, 'warning'); return; }
    await mutate(() => { const current = state.users.find((entry) => entry.id === user.id); if (current) Object.assign(current, { name, area }); }, `Se actualizaron los datos de ${name}.`);
    await openUserEditor(user.id); return;
  }
  if (event.target.id === 'location-form') {
    event.preventDefault(); const form = new FormData(event.target); const user = state.users.find((entry) => entry.id === event.target.dataset.userId); if (!user) return;
    const locationId = event.target.dataset.locationId; const previous = user.locations.find((entry) => entry.id === locationId); const location = locationValues(form, locationId, user.id); if (userHasLocation(user, location.name, locationId)) { notify(`Ya existe una ubicación llamada ${location.name} para ${user.name}. Usa otro nombre visible.`, 'warning'); return; }
    if (previous) {
      const previousName = previous.name; location.photoKey = previous.photoKey || ''; let changed = 0;
      await mutate(() => { const current = state.users.find((entry) => entry.id === user.id); const target = current?.locations?.find((entry) => entry.id === previous.id); if (!target) return; Object.assign(target, location); changed = renameLocationReferences(user.id, previousName, location.name); activateUserLocation(user.id, location.id); }, `Se actualizó ${location.name}${changed ? ` y ${changed} bien${changed === 1 ? '' : 'es'} vinculado${changed === 1 ? '' : 's'}` : ''}.`);
    } else {
      await mutate(() => { const current = state.users.find((entry) => entry.id === user.id); if (!current) return; current.locations.push(location); activateUserLocation(user.id, location.id); }, `Se agregó ${location.name} a ${user.name}.`);
    }
    await openUserEditor(user.id); return;
  }
  if (event.target.id === 'additional-form') { event.preventDefault(); const form = new FormData(event.target); const series = String(form.get('serie')).trim().toUpperCase(); const userId = String(form.get('userId')); const possession = String(form.get('possession')); const personal = form.get('personal') === 'on'; let key = String(form.get('clave')).trim(); if (!key && possession === 'Cámara' && !personal) key = nextAdditionalKey(userId); if (!key) { notify('La clave única es obligatoria para este tipo de posesión.', 'warning'); return; } if (!String(form.get('descripcion')).trim()) { notify('Agrega una descripción para el bien adicional.', 'warning'); return; } const duplicate = [...state.inventory, ...state.additionalItems].find((item) => item.serie && item.serie.toUpperCase() === series); if (duplicate) { notify(`La serie ya existe en ${duplicate.clave}. Revisa el registro antes de continuar.`, 'warning'); } await mutate(() => state.additionalItems.push({ id: uuid(), clave: key, descripcion: String(form.get('descripcion')).trim(), marca: String(form.get('marca')).trim(), modelo: String(form.get('modelo')).trim(), serie: series, possession, personal, detail: String(form.get('detail')).trim() || (possession === 'Arrendamiento' ? 'LXVIDG AJ- 070/2024' : ''), userId, location: activeLocationFor(userId)?.name || '', createdAt: Date.now(), updatedAt: Date.now() }), `Se registró el adicional ${key || 'sin clave'}.`); return; }
  if (event.target.id === 'profile-form') { event.preventDefault(); const form = new FormData(event.target); const [marca, modelo] = String(form.get('marcaModelo')).split('|').map((part) => part.trim()); try { new RegExp(String(form.get('regex'))); } catch { notify('La expresión regular no es válida.', 'error'); return; } await mutate(() => state.magicProfiles.push({ id: uuid(), regex: String(form.get('regex')).trim(), descripcion: String(form.get('descripcion')).trim(), marca, modelo, possession: String(form.get('possession')) }), 'Se agregó un perfil de autollenado.'); return; }
});

function matchProfile(serie) { return state.magicProfiles.find((profile) => { try { return new RegExp(profile.regex, 'i').test(serie || ''); } catch { return false; } }); }
function magicFillForm() { const serie = $('#ad-serie')?.value.trim(); const profile = matchProfile(serie); if (!profile) { notify('No existe un perfil que coincida con esa serie.', 'warning'); return; } $('#ad-descripcion').value = profile.descripcion; $('#ad-marca').value = profile.marca; $('#ad-modelo').value = profile.modelo; $('#ad-possession').value = profile.possession; notify(`Perfil aplicado: ${profile.descripcion}.`); }
function nextAdditionalKey(userId) { const area = state.users.find((user) => user.id === userId)?.area || '0000000'; const count = state.additionalItems.filter((item) => item.clave.startsWith(`CD-${area}-`)).length + 1; return `CD-${area}-${String(count).padStart(3, '0')}`; }

async function handleFile(kind, files) {
  if (!files?.length) return;
  try {
    if (kind === 'layout-plan') { await saveLayoutPlan(files[0]); return; }
    if (kind === 'inventory' || kind === 'conciliation') {
      notify(`Leyendo ${files.length} archivo${files.length === 1 ? '' : 's'}…`); const results = await parseInventoryFiles(files); const incoming = results.flatMap(({ result }) => result.items); const metadata = results.map(({ result }) => result).filter((result) => result.area);
      if (kind === 'inventory') { const existing = new Set(state.inventory.map((item) => item.clave)); const fresh = incoming.filter((item) => !existing.has(item.clave)); await mutate(() => { state.inventory.push(...fresh); metadata.forEach((result) => { state.areaNames[result.area] = result.areaName || state.areaNames[result.area] || ''; }); }, `Se importaron ${fresh.length} bienes desde ${files.length} libro${files.length === 1 ? '' : 's'}.`); notify(`${fresh.length} bienes cargados; ${incoming.length - fresh.length} duplicados omitidos.`); }
      else { conciliation = compareInventory(incoming); render(); notify(`Conciliación preparada con ${incoming.length} registros fuente.`); }
    }
    if (kind === 'restore' || kind === 'merge') { const incoming = normalizeBackup(await readBackup(files[0])); if (kind === 'restore') { confirmModal('Restaurar respaldo', 'La restauración reemplazará los datos locales actuales. Esta operación no se puede deshacer después de cerrar esta sesión.', async () => { snapshot(); state = incoming; await persist(); activeModule = 'inventory'; render(); notify('Respaldo restaurado.'); }, 'Restaurar', true); } else { confirmModal('Fusionar respaldo', 'Se agregarán bienes, adicionales y ubicaciones que no existan. Los conflictos de una misma clave se mantendrán en la sesión actual para su revisión.', async () => { await mutate(() => mergeBackup(incoming), 'Se fusionó un respaldo externo con la sesión local.'); notify('Fusión completada.'); }); } }
  } catch (error) { console.error(error); notify(`No fue posible procesar el archivo: ${error.message}`, 'error'); }
}

function compareInventory(incoming) {
  const current = new Map(state.inventory.map((item) => [item.clave, item])); const next = new Map(incoming.map((item) => [item.clave, item]));
  const additions = incoming.filter((item) => !current.has(item.clave)); const removals = state.inventory.filter((item) => !next.has(item.clave)); const modified = [];
  for (const item of incoming) { const old = current.get(item.clave); if (!old) continue; const fields = [['descripcion', 'Descripción'], ['marca', 'Marca'], ['modelo', 'Modelo'], ['serie', 'Serie'], ['area', 'Área'], ['bookType', 'Tipo de libro']].filter(([key]) => String(old[key] || '').trim() !== String(item[key] || '').trim()).map(([, label]) => label); if (fields.length) modified.push({ old, next: item, fields }); }
  return { additions, removals, modified, at: Date.now() };
}
function mergeBackup(incoming) {
  const keys = new Set(state.inventory.map((item) => item.clave)); state.inventory.push(...incoming.inventory.filter((item) => !keys.has(item.clave)));
  const userNames = new Map(state.users.map((user) => [user.name.toUpperCase(), user])); incoming.users.forEach((user) => { const existing = userNames.get(user.name.toUpperCase()); if (!existing) state.users.push(user); else { const locations = new Set(existing.locations.map((location) => location.name)); existing.locations.push(...user.locations.filter((location) => !locations.has(location.name))); } });
  const additionalIds = new Set(state.additionalItems.map((item) => item.id)); state.additionalItems.push(...incoming.additionalItems.filter((item) => !additionalIds.has(item.id)));
  state.notes.push(...incoming.notes.filter((note) => !state.notes.some((existing) => existing.id === note.id))); state.areaNames = { ...state.areaNames, ...incoming.areaNames }; state.layouts ||= {};
  incoming.users.forEach((sourceUser) => { const targetUser = state.users.find((user) => user.name.toUpperCase() === sourceUser.name.toUpperCase()); const incomingLayout = incoming.layouts?.[sourceUser.id]; if (!targetUser || !incomingLayout) return; const current = state.layouts[targetUser.id] || { planKey: '', pins: {} }; state.layouts[targetUser.id] = { ...incomingLayout, ...current, pins: { ...(incomingLayout.pins || {}), ...(current.pins || {}) }, planKey: current.planKey || '' }; });
}
function exportConciliation() { if (!conciliation) return; const rows = [['Tipo', 'Clave', 'Descripción', 'Detalle'], ...conciliation.additions.map((item) => ['ALTA', item.clave, item.descripcion, 'Nuevo registro']), ...conciliation.removals.map((item) => ['BAJA', item.clave, item.descripcion, 'No apareció en el nuevo corte']), ...conciliation.modified.map(({ old, next, fields }) => ['MODIFICADO', old.clave, next.descripcion, fields.join(', ')])]; const csv = rows.map((row) => row.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\n'); downloadBlob(new Blob([`\ufeff${csv}`], { type: 'text/csv;charset=utf-8' }), `Conciliacion_${new Date().toISOString().slice(0, 10)}.csv`); notify('Resumen de conciliación exportado.'); }
function exportRetag() { const rows = state.inventory.filter((item) => item.retag); const csv = [['Clave única', 'Descripción', 'Usuario', 'Ubicación'], ...rows.map((item) => [item.clave, item.descripcion, personName(item.userId), item.location])].map((row) => row.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\n'); downloadBlob(new Blob([`\ufeff${csv}`], { type: 'text/csv;charset=utf-8' }), `Reetiquetado_${new Date().toISOString().slice(0, 10)}.csv`); notify('Lista de re-etiquetado exportada.'); }

async function savePhoto(input) { const file = input.files?.[0]; if (!file) return; if (file.size > 8 * 1024 * 1024) { notify('La fotografía supera el límite de 8 MB.', 'warning'); return; } const additional = input.dataset.additional === 'true'; const items = additional ? state.additionalItems : state.inventory; const item = items.find((entry) => entry.id === input.dataset.item); if (!item) return; const key = item.photoKey || `${additional ? 'additional' : 'inventory'}-${item.id}`; await storage.savePhoto(key, file); await mutate(() => { item.photoKey = key; item.updatedAt = Date.now(); }, `Se agregó evidencia fotográfica a ${item.clave}.`); closeModal(); await showDetail(item.id, additional); }
async function saveLayoutPlan(file) { const user = state.users.find((entry) => entry.id === layoutUserId); if (!user || !file) return; if (file.size > 8 * 1024 * 1024) { notify('El plano supera el límite de 8 MB.', 'warning'); return; } const key = `layout-${user.id}`; await storage.savePhoto(key, file); await mutate(() => { state.layouts ||= {}; const current = state.layouts[user.id] || { pins: {} }; state.layouts[user.id] = { ...current, planKey: key }; }, `Se cargó el plano de referencia de ${user.name}.`); }

async function openScanner() {
  if (!navigator.mediaDevices?.getUserMedia) { notify('Este navegador no permite abrir la cámara.', 'error'); return; }
  openModal('Escanear código QR', `<p class="page-subtitle">La cámara se abrirá únicamente en este dispositivo. Al leer una clave, se buscará dentro del inventario cargado.</p><video id="scanner-video" class="scanner-video" autoplay playsinline muted></video><p id="scanner-help" class="page-subtitle">Preparando cámara…</p>`);
  try { scannerStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } }, audio: false }); const video = $('#scanner-video'); video.srcObject = scannerStream; await video.play(); if (!('BarcodeDetector' in window)) { $('#scanner-help').textContent = 'Este navegador no incluye lector QR nativo. Usa la búsqueda global con la clave visible en la etiqueta.'; return; } const detector = new BarcodeDetector({ formats: ['qr_code'] }); const read = async () => { if (!scannerStream || !video.videoWidth) return requestAnimationFrame(read); const codes = await detector.detect(video); if (codes[0]?.rawValue) { const result = codes[0].rawValue.trim(); stopScanner(); closeModal(); const item = state.inventory.find((entry) => entry.clave === result) || state.additionalItems.find((entry) => entry.clave === result); if (item) { await showDetail(item.id, state.additionalItems.includes(item)); } else { inventoryFilters.query = result; activeModule = 'inventory'; render(); notify('Código leído; no hay coincidencia exacta, se aplicó a la búsqueda.', 'warning'); } return; } requestAnimationFrame(read); }; requestAnimationFrame(read); } catch (error) { $('#scanner-help').textContent = 'No se pudo abrir la cámara. Verifica el permiso del navegador e inténtalo de nuevo.'; console.error(error); }
}
function stopScanner() { if (scannerStream) scannerStream.getTracks().forEach((track) => track.stop()); scannerStream = null; }

async function init() {
  modalRoot.onclick = (event) => {
    const actionTarget = event.target.closest('[data-action]');
    if (!actionTarget) return;
    if (actionTarget.classList.contains('modal-backdrop') && event.target.closest('.modal')) return;
    handleClick(event);
  };
  modalRoot.onchange = handleChange;
  state = normalizeBackup(await storage.load() || emptyState());
  window.addEventListener('beforeinstallprompt', (event) => { event.preventDefault(); installPrompt = event; if (state?.auditor) render(); });
  window.addEventListener('appinstalled', () => { installPrompt = null; if (state?.auditor) { render(); notify('La aplicación se instaló correctamente.'); } });
  window.addEventListener('online', () => { if (state?.auditor) { render(); notify('Conexión recuperada. Tus datos siguen guardados solo en este dispositivo.'); } });
  window.addEventListener('offline', () => { if (state?.auditor) { render(); notify('Sin conexión: puedes seguir trabajando con los datos locales.', 'warning'); } });
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js?v=13').catch(() => {});
  render();
}
init().catch((error) => { console.error(error); app.innerHTML = `<main class="login-page"><section class="login-card"><h1>No se pudo abrir la sesión local</h1><p>${escapeHtml(error.message)}</p></section></main>`; });
