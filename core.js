export const FIELDS = [
  ['sku','SKU cliente *','text'], ['quantity','Cantidad de piezas','number'],
  ['name','Nombre *','text'], ['description','Descripción *','textarea'],
  ['condition','Condición *','text'], ['status','Estado *','text'],
  ['basePrice','Precio base (pesos) *','number'], ['currentPrice','Precio actual (pesos) *','number'],
  ['currency','Moneda *','text'], ['liquidation','Es liquidación','text'],
  ['receivedAt','Fecha de recepción','text'], ['liquidationAt','Inicio de liquidación','text'],
  ['weight','Peso (gramos)','number'], ['length','Largo (mm)','number'],
  ['width','Ancho (mm)','number'], ['height','Alto (mm)','number'],
  ['category','Categoría','text'], ['tag1','Etiqueta 1','text'], ['tag2','Etiqueta 2','text'],
  ['tag3','Etiqueta 3','text'], ['url','URL del producto','text'],
  ['variationGroup','Grupo de variaciones','text'], ['option1','Opción 1','text'],
  ['value1','Valor 1','text'], ['option2','Opción 2','text'], ['value2','Valor 2','text'],
  ['option3','Opción 3','text'], ['value3','Valor 3','text'], ['option4','Opción 4','text'],
  ['value4','Valor 4','text'], ['option5','Opción 5','text'], ['value5','Valor 5','text'],
  ['warehouse','Código de almacén','text'], ['location','Código de ubicación','text']
].map(([key,label,type])=>({key,label,type}));

export const DEFAULT_RULES = {
  ivaMode:'add', iva:16, condition:'nuevo', status:'draft', currency:'MXN',
  warehouse:'A-Almacen', categorySpaces:true, categoryDe:true,
  categoryLower:true, categoryAccents:false, locationDashes:true
};

export const DEFAULT_MAPPING_KEYS = ['sku','name','description','quantity','basePrice','currentPrice','category','location'];
const FIELD_MAP = Object.fromEntries(FIELDS.map(f=>[f.key,f]));
export function normalize(value) {
  return String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
}
export function xmlText(value) {
  return String(value??'').replace(/_x([\da-f]{4})_/gi,'_x005F_x$1_').replace(/\r/g,'_x000D_')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
export function decodeXml(value) {
  return String(value??'').replace(/&#(x[\da-f]+|\d+);|&(amp|lt|gt|quot|apos);/gi, (m,n,named)=>{
    if(n){const number=n[0].toLowerCase()==='x'?parseInt(n.slice(1),16):parseInt(n,10);return number<=0x10ffff?String.fromCodePoint(number):m;}
    return ({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"})[named.toLowerCase()];
  }).replace(/_x([\da-f]{4})_/gi,(_,code)=>String.fromCharCode(parseInt(code,16)));
}
function attrs(text) {
  const out={};
  for(const m of text.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) out[m[1]]=decodeXml(m[2]??m[3]);
  return out;
}
function tags(xml,tag) {
  const result=[];
  const re=new RegExp('<(?:[\\w.-]+:)?'+tag+'\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?'+tag+'>','g');
  for(const m of xml.matchAll(re))result.push(m[1]);
  return result;
}
function plainText(xml) {return tags(xml,'t').map(decodeXml).join('');}
function zipPath(base,target) {
  const p=target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target;
  const parts=[];for(const part of p.split('/')){if(part==='..')parts.pop();else if(part&&part!=='.')parts.push(part);}return parts.join('/');
}
export function columnName(index) {
  let result='';for(let n=index+1;n>0;n=Math.floor((n-1)/26))result=String.fromCharCode(65+(n-1)%26)+result;return result;
}
function columnIndex(text) {let n=0;for(const c of text.toUpperCase())n=n*26+c.charCodeAt(0)-64;return n-1;}

export async function readWorkbook(bytes,Zip) {
  const zip=await Zip.loadAsync(bytes);
  const bookFile=zip.file('xl/workbook.xml');
  if(!bookFile)throw new Error('El archivo no es un libro Excel .xlsx válido.');
  const [book,rels,stringsXml]=await Promise.all([
    bookFile.async('string'), zip.file('xl/_rels/workbook.xml.rels').async('string'),
    zip.file('xl/sharedStrings.xml')?.async('string')??''
  ]);
  const relMap={};for(const m of rels.matchAll(/<(?:[\w.-]+:)?Relationship\b([^>]*?)\/?\s*>/g)){const a=attrs(m[1]);relMap[a.Id]=a.Target;}
  const sheets=[];
  for(const m of book.matchAll(/<(?:[\w.-]+:)?sheet\b([^>]*?)\/?\s*>/g)){
    const a=attrs(m[1]),rid=a['r:id']??Object.entries(a).find(([k])=>k.endsWith(':id'))?.[1];
    if(relMap[rid])sheets.push({name:a.name,path:zipPath('xl/workbook.xml',relMap[rid]),hidden:a.state==='hidden'||a.state==='veryHidden'});
  }
  if(!sheets.length)throw new Error('No se encontraron hojas en el archivo.');
  const strings=tags(stringsXml,'si').map(plainText);
  return {zip,sheets,strings,book};
}

export async function readSheet(book,index=0) {
  const sheet=book.sheets[index];
  if(!sheet)throw new Error('Selecciona una hoja disponible.');
  const file=book.zip.file(sheet.path);
  if(!file)throw new Error('No se pudo leer la hoja seleccionada.');
  const xml=await file.async('string');
  const byRow=new Map();
  // Read actual cells, rather than iterating Excel's often inflated used range.
  const cellRe=/<(?:[\w.-]+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?c>)/g;
  let meaningfulCells=0;
  for(const m of xml.matchAll(cellRe)){
    const a=attrs(m[1]),address=a.r?.match(/^([A-Za-z]+)(\d+)$/);if(!address)continue;
    const inner=m[2]??'';
    const raw=tags(inner,'v')[0];
    let value=null;
    if(a.t==='inlineStr')value=plainText(inner);
    else if(raw!==undefined){
      if(a.t==='s')value=book.strings[Number(raw)]??'';
      else if(a.t==='b')value=raw==='1';
      else if(a.t==='str'||a.t==='e')value=decodeXml(raw);
      else value=raw.trim()===''?null:(Number.isFinite(Number(raw))?Number(raw):decodeXml(raw));
    }
    if(value===null||value==='')continue;
    if(++meaningfulCells>1200000)throw new Error('La hoja tiene demasiadas celdas para procesarla aquí. Divide el archivo en lotes.');
    const rowNum=Number(address[2]),col=columnIndex(address[1]);
    if(!byRow.has(rowNum))byRow.set(rowNum,new Map());
    byRow.get(rowNum).set(col,value);
  }
  const rows=[...byRow].sort((a,b)=>a[0]-b[0]).map(([number,cells])=>({number,cells:Object.fromEntries(cells)}));
  return {name:sheet.name,path:sheet.path,xml,rows};
}

export function readCsv(text) {
  const separator=(text.split(/\r?\n/,1)[0].match(/;/g)?.length??0)>(text.split(/\r?\n/,1)[0].match(/,/g)?.length??0)?';':',';
  const records=[];let record=[],cell='',quoted=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(c==='"'){if(quoted&&text[i+1]==='"'){cell+='"';i++;}else quoted=!quoted;}
    else if(c===separator&&!quoted){record.push(cell);cell='';}
    else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&text[i+1]==='\n')i++;record.push(cell);records.push(record);record=[];cell='';}
    else cell+=c;
  }
  if(quoted)throw new Error('Hay comillas sin cerrar en el archivo CSV.');
  if(cell!==''||record.length){record.push(cell);records.push(record);}
  return {name:'CSV',rows:records.map((row,i)=>({number:i+1,cells:Object.fromEntries(row.map((v,j)=>[j,v.replace(/^\uFEFF/,'')]))})).filter(r=>Object.values(r.cells).some(v=>v!==''))};
}

export function sourceTable(sheet,headerRow) {
  if(!sheet.rows.length)throw new Error('La hoja está vacía.');
  const selected=headerRow?sheet.rows.find(r=>r.number===Number(headerRow)):
    sheet.rows.slice(0,20).reduce((best,r)=>Object.values(r.cells).length>Object.values(best.cells).length?r:best,sheet.rows[0]);
  if(!selected)throw new Error('La fila de encabezados no existe.');
  const maxCol=Math.max(...Object.keys(selected.cells).map(Number));
  const headers=Array.from({length:maxCol+1},(_,i)=>String(selected.cells[i]??`Columna ${columnName(i)}`));
  const records=sheet.rows.filter(r=>r.number>selected.number&&Object.values(r.cells).some(v=>String(v).trim()!==''))
    .map(r=>({id:`source-${r.number}`,sourceRow:r.number,values:headers.map((_,i)=>r.cells[i]??null)}));
  if(records.length>60000)throw new Error('Esta versión admite hasta 60,000 filas. Divide el archivo en lotes.');
  return {headers,headerRow:selected.number,records,name:sheet.name};
}

const ALIASES={
  sku:['skucliente','sku','codigoproducto','codigodeproducto','clave1','clave','codigo'],
  name:['nombre','descripcion','producto'],
  description:['caracteristicas','descripciondetallada','descripcionlarga'],
  quantity:['cantidaddepiezas'],
  basePrice:['preciobasepesos','preciobase','precio','precio1'],
  currentPrice:['precioactualpesos','precioactual','precio','precio1'],
  category:['categoria','categoryslug'], location:['codigodeubicacion','codigoubicacion','ubicacion'],
  warehouse:['codigodealmacen','codigoalmacen','almacen'],condition:['condicion'],status:['estado'],currency:['moneda']
};
export function autoMap(headers) {
  const keys=headers.map(normalize),mapping={};
  for(const field of FIELDS){
    const possibilities=[...(ALIASES[field.key]??[]),normalize(field.label)];
    if(field.key==='description'&&keys.includes('nombre'))possibilities.unshift('descripcion');
    mapping[field.key]=possibilities.map(a=>keys.indexOf(a)).find(i=>i>=0)??-1;
  }
  return mapping;
}

export function numberValue(value) {
  if(typeof value==='number')return Number.isFinite(value)?value:null;
  let text=String(value??'').trim();if(!text)return null;
  text=text.replace(/\b(?:MXN|USD)\b/gi,'').replace(/[$\s]/g,'');
  if(text.includes(',')&&text.includes('.')){
    if(text.lastIndexOf(',')>text.lastIndexOf('.'))text=text.replace(/\./g,'').replace(',','.');
    else text=text.replace(/,/g,'');
  }else if(text.includes(',')){
    text=/^-?\d{1,3}(,\d{3})+$/.test(text)?text.replace(/,/g,''):text.replace(',','.');
  }
  return /^-?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)&&Number.isFinite(Number(text))?Number(text):null;
}
export function priceWithIva(value,rules) {
  const n=numberValue(value);if(n===null)return null;
  const rate=numberValue(rules.iva)??0;
  const factor=rules.ivaMode==='add'?1+rate/100:rules.ivaMode==='only'?rate/100:1;
  return Math.round((n*factor+Number.EPSILON)*100)/100;
}
export function categoryValue(value,rules) {
  let text=String(value??'').trim();
  if(rules.categoryAccents)text=text.normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  if(rules.categoryLower)text=text.toLowerCase();
  if(rules.categoryDe)text=text.replace(/(^|[\s-])de(?=$|[\s-])/gi,'$1-');
  if(rules.categorySpaces)text=text.replace(/\s+/g,'-');
  return text.replace(/-+/g,'-').replace(/^-|-$/g,'');
}
export function locationValue(value,rules) {
  const text=String(value??'').trim();if(!rules.locationDashes)return text;
  return text.replace(/([A-Za-zÁÉÍÓÚÜÑáéíóúüñ])(\d)/g,'$1-$2')
    .replace(/(\d)([A-Za-zÁÉÍÓÚÜÑáéíóúüñ])/g,'$1-$2')
    .replace(/[\s,]+/g,'-').replace(/-+/g,'-').replace(/^-|-$/g,'');
}
export function convertRecords(table,mapping,rules,categoryMap={},overrides={},manualRows=[]) {
  return [...table.records.map(record=>{
    const values={};for(const f of FIELDS){const col=mapping[f.key];const val=col>=0?record.values[col]:null;values[f.key]=f.type==='number'?numberValue(val):String(val??'');}
    const rawCategory=values.category,rawLocation=values.location;
    values.sku=values.sku.trim();
    values.basePrice=priceWithIva(mapping.basePrice>=0?record.values[mapping.basePrice]:null,rules);
    values.currentPrice=priceWithIva(mapping.currentPrice>=0?record.values[mapping.currentPrice]:null,rules);
    for(const key of ['condition','status','currency','warehouse'])if(rules[key]!=='')values[key]=rules[key];
    values.category=categoryMap[rawCategory]||categoryValue(rawCategory,rules);
    values.location=locationValue(rawLocation,rules);
    return {id:record.id,sourceRow:record.sourceRow,rawCategory,rawLocation,locationReviewed:Object.hasOwn(overrides[record.id]??{},'location'),values:{...values,...(overrides[record.id]??{})}};
  }),...manualRows.map(row=>({...row,values:{...row.values,...(overrides[row.id]??{})}}))];
}

export async function templateInfo(book) {
  let index=book.sheets.findIndex(s=>normalize(s.name)==='productos');if(index<0)index=0;
  const sheet=await readSheet(book,index);
  const table=sourceTable(sheet);
  const mapping=autoMap(table.headers);
  if(mapping.sku<0||mapping.name<0||mapping.basePrice<0)throw new Error('La plantilla debe tener columnas de SKU, Nombre y Precio base. Usa la plantilla base o un formato compatible.');
  const catalogs={categories:[],warehouses:[],conditions:['nuevo','usado'],statuses:['borrador','draft'],currencies:['MXN']};
  for(const [kind,name] of [['categories','categorias'],['warehouses','almacenes']]){
    const i=book.sheets.findIndex(s=>normalize(s.name)===name);if(i<0)continue;
    const aux=await readSheet(book,i);
    catalogs[kind]=aux.rows.slice(1).map(r=>({value:String(r.cells[0]??''),label:String(r.cells[1]??r.cells[0]??''),code:String(r.cells[1]??'')})).filter(r=>r.value);
    catalogs[kind]=catalogs[kind].filter((x,i,a)=>a.findIndex(y=>y.value===x.value)===i);
  }
  const vi=book.sheets.findIndex(s=>normalize(s.name)==='valores');
  if(vi>=0){const v=await readSheet(book,vi);for(const [key,col] of [['conditions',0],['statuses',1],['currencies',2]]){
    const list=[...new Set(v.rows.slice(1).map(r=>String(r.cells[col]??'')).filter(Boolean))];if(list.length)catalogs[key]=list;
  }}
  return {sheet,index,headers:table.headers,headerRow:table.headerRow,mapping,catalogs};
}

export function validateRows(rows,catalogs,excluded=new Set()) {
  const duplicates=new Map();
  for(const row of rows)if(!excluded.has(row.id)&&row.values.sku.trim()){
    const key=row.values.sku.trim();if(!duplicates.has(key))duplicates.set(key,[]);duplicates.get(key).push(row.id);
  }
  const cats=new Set(catalogs?.categories?.map(x=>x.value)??[]);
  const warehouses=new Set((catalogs?.warehouses??[]).flatMap(x=>[x.value,x.code]));
  return rows.map(row=>{
    const v=row.values,issues=[];
    function add(level,field,message){issues.push({level,field,message});}
    for(const key of ['sku','name','description','condition','status','currency'])if(!String(v[key]??'').trim())add('error',key,`${FIELD_MAP[key].label.replace(' *','')}: falta el dato.`);
    for(const key of ['basePrice','currentPrice'])if(typeof v[key]!=='number'||!Number.isFinite(v[key])||v[key]<0)add('error',key,`${FIELD_MAP[key].label.replace(' (pesos) *','')}: escribe un precio válido.`);
    if(v.quantity!==null&&v.quantity!==''&&(!Number.isInteger(v.quantity)||v.quantity<1))add('error','quantity','Cantidad: debe ser un entero mayor que cero.');
    if(v.sku&&duplicates.get(v.sku.trim())?.length>1)add('warning','sku',`SKU repetido en ${duplicates.get(v.sku.trim()).length} filas.`);
    if(v.category&&cats.size&&!cats.has(v.category))add('error','category','La categoría no existe en el catálogo de la plantilla.');
    if(!v.category)add('warning','category','Sin categoría: el producto quedará incompleto.');
    if(v.warehouse&&warehouses.size&&!warehouses.has(v.warehouse))add('error','warehouse','El almacén no coincide con el catálogo de la plantilla.');
    if(Boolean(v.warehouse)!==Boolean(v.location))add('error','location','Completa tanto el almacén como la ubicación.');
    for(const [key,list] of [['condition',catalogs?.conditions],['status',catalogs?.statuses],['currency',catalogs?.currencies]])if(v[key]&&list?.length&&!list.includes(v[key]))add('error',key,`El valor de ${FIELD_MAP[key].label.replace(' *','').toLowerCase()} no está permitido en la plantilla.`);
    if(/^(?:no encontrado|0)$/i.test(String(v.name).trim()))add('warning','name','El nombre de origen parece estar incompleto.');
    if(/^(?:no encontrado|0)$/i.test(String(v.description).trim()))add('warning','description','La descripción de origen parece estar incompleta.');
    if(row.rawLocation?.includes(',')&&!row.locationReviewed)add('warning','location','La ubicación original contiene una coma. Revisa los niveles.');
    return {...row,issues,excluded:excluded.has(row.id),errorCount:issues.filter(x=>x.level==='error').length,warningCount:issues.filter(x=>x.level==='warning').length};
  });
}

export function filterRows(rows,{search='',category='',status='all'}={}) {
  const term=normalize(search);
  return rows.filter(row=>{
    if(status==='excluded'?!row.excluded:row.excluded)return false;
    if(category&&row.values.category!==category)return false;
    if(status==='errors'&&!row.errorCount)return false;
    if(status==='warnings'&&!row.warningCount)return false;
    if(status==='clean'&&(row.errorCount||row.warningCount))return false;
    if(status==='noerrors'&&row.errorCount)return false;
    if(status==='duplicates'&&!row.issues.some(x=>x.field==='sku'&&x.level==='warning'))return false;
    if(term&&!normalize([row.values.sku,row.values.name,row.values.category,row.values.location].join(' ')).includes(term))return false;
    return true;
  });
}

export function bulkChanges(rows,{fields,operation,value,find=''}) {
  if(!rows.length)throw new Error('No hay filas para modificar.');
  if(!fields.length)throw new Error('Selecciona un campo.');
  const changes=[];
  for(const row of rows)for(const field of fields){
    const meta=FIELD_MAP[field];if(!meta)throw new Error('El campo no existe.');
    let next=value;
    if(operation==='replace'){
      if(!find)throw new Error('Escribe el texto que quieres buscar.');
      next=String(row.values[field]??'').split(find).join(String(value));
    }else if(operation==='percent'){
      const p=numberValue(value),current=numberValue(row.values[field]);
      if(p===null)throw new Error('Escribe un porcentaje válido.');
      if(current===null)continue;
      next=Math.round((current*(1+p/100)+Number.EPSILON)*100)/100;
    }
    if(meta.type==='number'){
      next=String(next).trim()===''?null:numberValue(next);
      if(next===null&&String(value).trim()!=='')throw new Error('Escribe un número válido.');
    }else next=String(next??'');
    if(row.values[field]!==next)changes.push({id:row.id,field,before:row.values[field],after:next});
  }
  return changes;
}

function cellXml(address,value,style) {
  const s=style?` s="${style}"`:'';
  if(value===null||value===undefined||value==='')return `<c r="${address}"${s}/>`;
  if(typeof value==='number'&&Number.isFinite(value))return `<c r="${address}"${s} t="n"><v>${value}</v></c>`;
  return `<c r="${address}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlText(value)}</t></is></c>`;
}
export async function exportWorkbook(templateBytes,info,rows,Zip) {
  if(!rows.length)throw new Error('No hay productos para exportar.');
  const zip=await Zip.loadAsync(templateBytes);
  let xml=await zip.file(info.sheet.path).async('string');
  const sheetData=xml.match(/<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/);
  if(!sheetData)throw new Error('La plantilla no contiene una tabla de productos compatible.');
  const before=[];let rowStyle='';const columnStyles={};
  for(const match of sheetData[1].matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)){
    const a=attrs(match[1]),n=Number(a.r);
    if(n<=info.headerRow)before.push(match[0]);
    if(n===info.headerRow+1){
      rowStyle=match[1].replace(/\br="[^"]*"/,'').replace(/\bspans="[^"]*"/,'').trim();
      for(const c of (match[2]??'').matchAll(/<c\b([^>]*?)(?:\/>|>)/g)){const ca=attrs(c[1]);if(ca.r&&ca.s)columnStyles[ca.r.replace(/\d/g,'')]=ca.s;}
    }
  }
  const inverse=Object.fromEntries(Object.entries(info.mapping).filter(([,i])=>i>=0).map(([key,i])=>[i,key]));
  const generated=rows.map((row,i)=>{
    const r=info.headerRow+1+i;
    const cells=info.headers.map((_,j)=>cellXml(columnName(j)+r,inverse[j]?row.values[inverse[j]]:null,columnStyles[columnName(j)])).join('');
    return `<row r="${r}"${rowStyle?' '+rowStyle:''}>${cells}</row>`;
  });
  const last=info.headerRow+rows.length,columns=columnName(info.headers.length-1);
  xml=xml.replace(sheetData[0],'<sheetData>'+before.join('')+generated.join('')+'</sheetData>');
  xml=xml.replace(/<dimension\b[^>]*\/>/,`<dimension ref="A1:${columns}${last}"/>`);
  xml=xml.replace(/\bsqref="([A-Z]+)(\d+):([A-Z]+)\d+"/g,(all,c,r,c2)=>Number(r)>info.headerRow?`sqref="${c}${r}:${c2}${Math.max(last,Number(r))}"`:all);
  xml=xml.replace(/<autoFilter\b([^>]*)ref="([A-Z]+\d+):[A-Z]+\d+"([^>]*)\/>/g,(_,a,start,b)=>`<autoFilter${a}ref="${start}:${columns}${last}"${b}/>`);
  zip.file(info.sheet.path,xml);
  // Resize any native Excel table linked to the product sheet as well.
  const relPart=info.sheet.path.replace(/([^/]+)$/, '_rels/$1.rels');
  if(zip.file(relPart)){
    const rels=await zip.file(relPart).async('string');
    for(const m of rels.matchAll(/<Relationship\b([^>]*?)\/?\s*>/g)){
      const a=attrs(m[1]);if(!a.Type?.endsWith('/table'))continue;
      const path=zipPath(info.sheet.path,a.Target);if(!zip.file(path))continue;
      let tableXml=await zip.file(path).async('string');
      tableXml=tableXml.replace(/\bref="([A-Z]+\d+):[A-Z]+\d+"/g,(_,start)=>`ref="${start}:${columns}${last}"`);
      zip.file(path,tableXml);
    }
  }
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
}

export function reportCsv(rows) {
  const records=[['Fila de origen','SKU','Nivel','Campo','Detalle']];
  for(const row of rows)for(const issue of row.issues)records.push([row.sourceRow??'Nuevo',row.values.sku,issue.level==='error'?'Error':'Aviso',FIELD_MAP[issue.field]?.label??issue.field,issue.message]);
  return '\uFEFF'+records.map(record=>record.map(x=>{let s=String(x??'');if(/^[=+@-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}).join(',')).join('\r\n');
}
