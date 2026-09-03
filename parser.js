const CELL_REF = /^([A-Z]+)(\d+)$/;

const clean = (value) => String(value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
function canonicalDecimalKey(raw) {
  const decimal = raw.startsWith('.') ? `0${raw}` : raw;
  if (!/^0\.\d+$/.test(decimal)) return '';
  const fraction = decimal.split('.')[1];
  const numeric = Number(decimal);
  // Excel stores some 5–6 digit keys as binary float artifacts (for example,
  // 0.26462999999999998). Only round when it is already within a tiny distance
  // of a six-decimal value, so legitimate longer keys remain untouched.
  const rounded = Number(numeric.toFixed(6));
  if (fraction.length > 6 && Math.abs(numeric - rounded) < 1e-10) return rounded.toFixed(6).replace(/0+$/, '').replace(/^0/, '');
  return decimal.replace(/^0/, '');
}
export const normalizeKey = (value) => {
  const raw = clean(value).replace(/,/g, '');
  const decimal = canonicalDecimalKey(raw);
  if (decimal) return decimal;
  if (/^\d{5,6}$/.test(raw)) return raw;
  return raw;
};

const isInventoryKey = (value) => /^(?:\d{5,6}|\.\d+)$/.test(normalizeKey(value));
const getFirst = (rows, matcher) => rows.flat().map(clean).find((value) => matcher.test(value)) || '';

function metadata(rows, documentText = []) {
  const all = [...rows.flat(), ...documentText].map(clean).filter(Boolean);
  const areaCandidates = [...all, ...all.slice(0, -1).map((value, index) => `${value} ${all[index + 1]}`)];
  const areaText = areaCandidates.find((value) => /(?:ÁREA|AREA)\s*:?\s*\d+/i.test(value)) || '';
  const areaMatch = areaText.match(/(?:ÁREA|AREA)\s*:?\s*(\d+)(?:\s*[-:–—]?\s*)(.*)/i);
  const bookText = all.find((value) => /LIBRO\s*:/i.test(value)) || '';
  const bookMatch = bookText.match(/LIBRO\s*:\s*(.+)/i);
  return {
    area: areaMatch?.[1] || '',
    areaName: clean(areaMatch?.[2] || ''),
    bookType: clean(bookMatch?.[1] || ''),
    sourceTitle: getFirst(rows, /LISTADO DE BIENES/i)
  };
}

function buildItems(rows, sourceName, documentText = []) {
  const info = metadata(rows, documentText);
  const items = [];
  for (const row of rows) {
    const key = normalizeKey(row[0]);
    if (!isInventoryKey(key)) continue;
    items.push({
      id: crypto.randomUUID(),
      clave: key,
      descripcion: clean(row[1]),
      marca: clean(row[4]),
      modelo: clean(row[5]),
      serie: clean(row[6]),
      area: info.area,
      areaName: info.areaName,
      bookType: info.bookType,
      source: sourceName,
      status: 'pendiente',
      retag: false,
      userId: '',
      location: '',
      createdAt: Date.now(),
      updatedAt: Date.now()
    });
  }
  return { ...info, items };
}

function parseHtmlBook(text, sourceName) {
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const rows = [...doc.querySelectorAll('tr')].map((tr) => [...tr.querySelectorAll('th,td')].map((cell) => clean(cell.textContent)));
  if (!rows.length) throw new Error('No se encontraron filas de una exportación Oracle/HTML.');
  const documentText = [...doc.querySelectorAll('p')].map((paragraph) => clean(paragraph.textContent));
  return buildItems(rows, sourceName, documentText);
}

function getText(element) {
  return [...element.getElementsByTagName('*')].filter((node) => node.localName === 't').map((node) => node.textContent).join('') || element.textContent || '';
}

function columnFromRef(ref) { return ref.match(CELL_REF)?.[1] || ''; }

async function parseOoxml(buffer, sourceName) {
  if (!window.JSZip) throw new Error('El lector local de Excel no se cargó. Recargue la aplicación.');
  const zip = await window.JSZip.loadAsync(buffer);
  const sharedXml = await zip.file('xl/sharedStrings.xml')?.async('text');
  const parser = new DOMParser();
  const shared = sharedXml ? [...parser.parseFromString(sharedXml, 'application/xml').getElementsByTagName('*')]
    .filter((node) => node.localName === 'si').map(getText) : [];
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
  let sheetPath = 'xl/worksheets/sheet1.xml';
  if (workbookXml && relsXml) {
    const workbook = parser.parseFromString(workbookXml, 'application/xml');
    const firstSheet = [...workbook.getElementsByTagName('*')].find((node) => node.localName === 'sheet');
    const relationId = firstSheet?.getAttribute('r:id') || firstSheet?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const relationships = parser.parseFromString(relsXml, 'application/xml');
    const relation = [...relationships.getElementsByTagName('*')].find((node) => node.localName === 'Relationship' && node.getAttribute('Id') === relationId);
    if (relation?.getAttribute('Target')) sheetPath = `xl/${relation.getAttribute('Target').replace(/^\/+/, '').replace(/^xl\//, '')}`;
  }
  const sheetXml = await zip.file(sheetPath)?.async('text');
  if (!sheetXml) throw new Error('No se encontró la primera hoja del libro.');
  const sheet = parser.parseFromString(sheetXml, 'application/xml');
  const rowNodes = [...sheet.getElementsByTagName('*')].filter((node) => node.localName === 'row');
  const rows = rowNodes.map((row) => {
    const values = [];
    [...row.getElementsByTagName('*')].filter((node) => node.localName === 'c').forEach((cell) => {
      const col = columnFromRef(cell.getAttribute('r') || '');
      const index = col.split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
      const valueNode = [...cell.getElementsByTagName('*')].find((node) => node.localName === 'v');
      const inlineNode = [...cell.getElementsByTagName('*')].find((node) => node.localName === 'is');
      let value = valueNode?.textContent || (inlineNode ? getText(inlineNode) : '');
      if (cell.getAttribute('t') === 's') value = shared[Number(value)] || '';
      values[index] = clean(value);
    });
    return values;
  });
  return buildItems(rows, sourceName);
}

function parseCsv(text, sourceName) {
  const delimiter = text.includes(';') ? ';' : ',';
  const rows = text.split(/\r?\n/).map((line) => line.split(delimiter).map(clean));
  return buildItems(rows, sourceName);
}

export async function parseInventoryFile(file) {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer.slice(0, 8));
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (isZip) return parseOoxml(buffer, file.name);
  const text = new TextDecoder('utf-8').decode(buffer);
  if (/<(?:html|table|tr)[\s>]/i.test(text)) return parseHtmlBook(text, file.name);
  return parseCsv(text, file.name);
}

export async function parseInventoryFiles(files) {
  const results = await Promise.all([...files].map(async (file) => ({ file: file.name, result: await parseInventoryFile(file) })));
  return results;
}
