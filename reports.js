const esc = (value) => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const today = () => new Date().toISOString().slice(0, 10);

export function downloadBlob(blob, fileName) {
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 500);
}

function xmlRows(headers, rows) {
  const cell = (value) => `<Cell><Data ss:Type="String">${esc(value)}</Data></Cell>`;
  return `<Row>${headers.map(cell).join('')}</Row>${rows.map((row) => `<Row>${headers.map((header) => cell(row[header])).join('')}</Row>`).join('')}`;
}

export function exportExcel(state) {
  const headers = ['Clave única', 'Descripción', 'Marca', 'Modelo', 'Serie', 'Área', 'Usuario', 'Ubicación', 'Estatus', 'Re-etiquetado', 'Actualizado'];
  const people = new Map(state.users.map((user) => [user.id, user.name]));
  const inventory = state.inventory.map((item) => ({
    'Clave única': item.clave, 'Descripción': item.descripcion, Marca: item.marca, Modelo: item.modelo,
    Serie: item.serie, Área: `${item.area} ${item.areaName}`.trim(), Usuario: people.get(item.userId) || '',
    Ubicación: item.location, Estatus: item.status, 'Re-etiquetado': item.retag ? 'Sí' : 'No',
    Actualizado: new Date(item.updatedAt).toLocaleString('es-MX')
  }));
  const additionalHeaders = ['Clave única', 'Descripción', 'Marca', 'Modelo', 'Serie', 'Posesión', 'Personal', 'Usuario', 'Ubicación'];
  const additions = state.additionalItems.map((item) => ({
    'Clave única': item.clave, 'Descripción': item.descripcion, Marca: item.marca, Modelo: item.modelo,
    Serie: item.serie, Posesión: item.possession, Personal: item.personal ? 'Sí' : 'No',
    Usuario: people.get(item.userId) || '', Ubicación: item.location || ''
  }));
  const sheet = (name, heads, rows) => `<Worksheet ss:Name="${esc(name)}"><Table>${xmlRows(heads, rows)}</Table></Worksheet>`;
  const xml = `<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">${sheet('Inventario', headers, inventory)}${sheet('Adicionales', additionalHeaders, additions)}</Workbook>`;
  downloadBlob(new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8' }), `Inventario_Legislativo_${today()}.xls`);
}

export async function exportBackup(state) {
  if (!window.JSZip) throw new Error('El motor de respaldo no está disponible.');
  const zip = new window.JSZip();
  zip.file('inventario-legislativo.json', JSON.stringify({ ...state, exportedAt: new Date().toISOString() }, null, 2));
  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
  downloadBlob(blob, `Respaldo_Inventario_${today()}.zip`);
}

export async function readBackup(file) {
  if (!window.JSZip) throw new Error('El motor de respaldo no está disponible.');
  const zip = await window.JSZip.loadAsync(file);
  const entry = zip.file('inventario-legislativo.json') || zip.file('session.json');
  if (!entry) throw new Error('El respaldo no contiene inventario-legislativo.json ni session.json.');
  return JSON.parse(await entry.async('text'));
}

export function reportMarkup(state, type, filters = {}) {
  const people = new Map(state.users.map((user) => [user.id, user]));
  let items = state.inventory;
  if (filters.area) items = items.filter((item) => item.area === filters.area);
  if (filters.userId) items = items.filter((item) => item.userId === filters.userId);
  if (type === 'pendientes') items = items.filter((item) => item.status !== 'ubicado');
  const title = { resguardo: 'Resguardo individual', pendientes: 'Bienes pendientes por ubicar', adicionales: 'Mobiliario y equipo adicional', album: 'Álbum fotográfico' }[type] || 'Reporte de inventario';
  if (type === 'adicionales') {
    items = state.additionalItems;
  }
  const rows = items.map((item, index) => `<tr><td>${index + 1}</td><td>${esc(item.clave)}</td><td>${esc(item.descripcion)}</td><td>${esc(item.marca)}</td><td>${esc(item.modelo)}</td><td>${esc(item.serie)}</td><td>${esc(people.get(item.userId)?.name || '')}</td><td>${esc(item.location || '')}</td></tr>`).join('');
  return `<article class="print-document"><header><img src="./assets/camara-logo.png" alt="Cámara de Diputados" /><div><p>CÁMARA DE DIPUTADOS · LXVI LEGISLATURA</p><h1>${esc(title)}</h1><p>Dirección de Almacén e Inventarios · ${new Date().toLocaleDateString('es-MX', { dateStyle: 'long' })}</p></div></header><p class="legal-copy">Documento generado para control administrativo y resguardo de bienes. Verifique los datos antes de recabar firmas.</p><table><thead><tr><th>#</th><th>Clave</th><th>Descripción</th><th>Marca</th><th>Modelo</th><th>Serie</th><th>Resguardante</th><th>Ubicación</th></tr></thead><tbody>${rows || '<tr><td colspan="8">Sin registros que correspondan al filtro.</td></tr>'}</tbody></table><footer><div>Elaboró<br/><strong>${esc(state.auditor?.name || 'Sin sesión')}</strong></div><div>Recibió<br/><strong>_______________________________</strong></div></footer></article>`;
}
