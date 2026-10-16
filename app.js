import {
  FIELDS, DEFAULT_RULES, DEFAULT_MAPPING_KEYS, autoMap, sourceTable, readWorkbook,
  readSheet, readCsv, templateInfo, convertRecords, validateRows, filterRows,
  numberValue, priceWithIva, categoryValue, bulkChanges, exportWorkbook, reportCsv, columnName
} from './core.js';

const $=id=>document.getElementById(id);
const h=value=>String(value??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
const icon=(name,cls='')=>`<svg class="icon ${cls}" aria-hidden="true"><use href="#i-${name}"/></svg>`;
const fmt=value=>new Intl.NumberFormat('es-MX').format(value);
const clone=value=>structuredClone(value);
const fieldMeta=Object.fromEntries(FIELDS.map(f=>[f.key,f]));
const MAX_BYTES=30*1024*1024;
const state={
  sourceName:'',sourceBytes:null,sourceBook:null,sourceSheet:null,sheetIndex:0,table:null,
  templateBytes:null,templateName:'Plantilla de carga Noventia',templateType:'builtin',templateInfo:null,
  rules:clone(DEFAULT_RULES),appliedRules:clone(DEFAULT_RULES),mapping:{},appliedMapping:{},
  categoryMap:{},appliedCategoryMap:{},overrides:{},manualRows:[],excluded:new Set(),
  rows:[],filtered:[],selected:new Set(),history:[],page:1,pageSize:20,search:'',categoryFilter:'',statusFilter:'all',
  dirtyRules:false,showAllMapping:false,isDemo:false,editId:null,pendingNewId:null
};
let toastTimer,saveTimer,dbPromise,saveVersion=0,storageFailed=false;

function toast(message,error=false){clearTimeout(toastTimer);$('toast').textContent=message;$('toast').classList.toggle('error-toast',error);$('toast').hidden=false;toastTimer=setTimeout(()=>$('toast').hidden=true,error?7000:4500);}
function loading(title,note='Esto puede tardar unos segundos.'){$('loading-title').textContent=title;$('loading-note').textContent=note;$('loading-overlay').hidden=false;}
function stopLoading(){$('loading-overlay').hidden=true;}
function breathe(){return new Promise(resolve=>requestAnimationFrame(()=>setTimeout(resolve,0)));}
function errorMessage(error){return error?.message??'No se pudo completar la operación.';}
function hasEdits(){return Object.keys(state.overrides).length||state.manualRows.length||state.excluded.size;}
function addHistory(entry){state.history.push(entry);if(state.history.length>20)state.history.shift();}
function sourceLabel(){return state.sourceName||'Catálogo manual';}

function db(){
  if(!dbPromise)dbPromise=new Promise((resolve,reject)=>{
    const request=indexedDB.open('noventia-catalogo-v1',1);
    request.onupgradeneeded=()=>request.result.createObjectStore('drafts');
    request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);
  });return dbPromise;
}
async function readDraft(){try{const store=(await db()).transaction('drafts','readonly').objectStore('drafts');return await new Promise((res,rej)=>{const r=store.get('current');r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);});}catch{return null;}}
async function writeDraft(value){const database=await db();await new Promise((res,rej)=>{const tx=database.transaction('drafts','readwrite');tx.objectStore('drafts').put(value,'current');tx.oncomplete=res;tx.onerror=()=>rej(tx.error);tx.onabort=()=>rej(tx.error);});}
async function deleteDraft(){try{const database=await db();await new Promise((res,rej)=>{const tx=database.transaction('drafts','readwrite');tx.objectStore('drafts').delete('current');tx.oncomplete=res;tx.onerror=()=>rej(tx.error);});}catch{}}
function serialize(){return {
  version:1,savedAt:Date.now(),sourceName:state.sourceName,sourceBytes:state.sourceBytes,
  sheetIndex:state.sheetIndex,table:state.table,templateBytes:state.templateBytes,templateName:state.templateName,
  templateType:state.templateType,rules:state.rules,appliedRules:state.appliedRules,mapping:state.mapping,
  appliedMapping:state.appliedMapping,categoryMap:state.categoryMap,appliedCategoryMap:state.appliedCategoryMap,
  overrides:state.overrides,manualRows:state.manualRows.filter(r=>r.id!==state.pendingNewId),excluded:[...state.excluded],dirtyRules:state.dirtyRules,isDemo:state.isDemo
};}
function scheduleSave(){
  clearTimeout(saveTimer);saveVersion++;if(storageFailed)return;
  $('save-status').innerHTML='<span class="status-dot"></span>Guardando avance…';
  saveTimer=setTimeout(async()=>{
    const version=saveVersion;
    try{await writeDraft(serialize());if(version===saveVersion)$('save-status').innerHTML=icon('check')+'Guardado en este navegador';}
    catch{storageFailed=true;$('save-status').textContent='Guardado local no disponible';toast('No se pudo guardar el avance en este navegador. Descarga tu Excel para conservar los cambios.',true);}
  },700);
}

function syncRuleInputs(){
  document.querySelectorAll('[data-rule]').forEach(input=>{const value=state.rules[input.dataset.rule];if(input.type==='checkbox')input.checked=!!value;else input.value=value??'';});
  updatePriceExample();$('pending-rules').hidden=!state.dirtyRules;
}
function updatePriceExample(){
  $('iva-example').textContent=new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN'}).format(priceWithIva(100,state.rules));
  $('iva-example-tag').textContent=state.rules.ivaMode==='only'?'Solo impuesto':state.rules.ivaMode==='keep'?'Sin ajuste':'IVA incluido';
  $('iva-rate').disabled=state.rules.ivaMode==='keep';
}
function markPending(){state.dirtyRules=true;$('pending-rules').hidden=false;updatePriceExample();scheduleSave();}
function activateTab(name){
  for(const tab of document.querySelectorAll('[data-tab]')){const active=tab.dataset.tab===name;tab.classList.toggle('active',active);tab.setAttribute('aria-selected',String(active));$('pane-'+tab.dataset.tab).hidden=!active;}
  if(name==='categories')renderCategoryMap();
}
function renderMapping(){
  const headers=state.table?.headers??[];
  const fields=state.showAllMapping?FIELDS:DEFAULT_MAPPING_KEYS.map(key=>fieldMeta[key]);
  $('mapping-grid').innerHTML=fields.map(field=>{
    const empty=field.key==='quantity'?'Sin asignar (la plantilla usa 1)':'Sin columna';
    return `<label class="field">${h(field.label)}<select data-mapping="${field.key}" ${!state.table?'disabled':''}><option value="-1">${empty}</option>${headers.map((header,i)=>`<option value="${i}" ${Number(state.mapping[field.key])===i?'selected':''}>${columnName(i)} · ${h(header)}</option>`).join('')}</select></label>`;
  }).join('');
  $('show-mapping-btn').textContent=state.showAllMapping?'Mostrar solo los campos principales':'Mostrar las 34 columnas';
  $('automap-btn').disabled=!state.table;
}
function renderCategoryMap(){
  if(!state.table){$('category-equivalences').innerHTML='<div class="small-empty">Sube un inventario para revisar sus categorías.</div>';return;}
  const col=state.mapping.category;
  if(!(col>=0)){$('category-equivalences').innerHTML='<div class="small-empty">Asigna primero la columna de categoría en “Columnas del archivo”.</div>';return;}
  const counts=new Map();for(const r of state.table.records){const cat=String(r.values[col]??'').trim();if(cat)counts.set(cat,(counts.get(cat)??0)+1);}
  const catalog=state.templateInfo?.catalogs.categories??[];
  const nodes=[...counts].sort((a,b)=>b[1]-a[1]);
  if(!nodes.length){$('category-equivalences').innerHTML='<div class="small-empty">El archivo no contiene categorías.</div>';return;}
  $('category-equivalences').innerHTML=nodes.map(([cat,count])=>`<div class="equivalence-row"><div>${h(cat)} <small>(${fmt(count)})</small></div>${icon('arrow')}<select aria-label="Equivalencia de ${h(cat)}" data-category-map="${h(cat)}"><option value="">${h(categoryValue(cat,state.rules))} · conservar</option>${catalog.map(c=>`<option value="${h(c.value)}" ${state.categoryMap[cat]===c.value?'selected':''}>${h(c.value)}</option>`).join('')}</select></div>`).join('');
}
function renderCatalogOptions(){
  const cats=state.templateInfo?.catalogs;
  $('category-options').innerHTML=(cats?.categories??[]).map(c=>`<option value="${h(c.value)}">${h(c.label)}</option>`).join('');
  $('warehouse-options').innerHTML=(cats?.warehouses??[]).flatMap(c=>[c.value,c.code]).filter((v,i,a)=>v&&a.indexOf(v)===i).map(v=>`<option value="${h(v)}"></option>`).join('');
  $('currency-options').innerHTML=(cats?.currencies??['MXN']).map(v=>`<option value="${h(v)}"></option>`).join('');
}
function updateFiles(){
  const count=state.table?.records.length??0;
  $('source-title').textContent=state.table?sourceLabel():'Arrastra aquí tu inventario';
  $('source-subtitle').innerHTML=state.table?'Haz clic para cargar otro archivo.':'o <b>elige un archivo</b> de tu computadora';
  $('source-tag').textContent=state.table?'Cargado':'Sin archivo';$('source-tag').className='tag '+(state.table?'tag-teal':'tag-muted');
  $('source-meta').textContent=state.table?`${fmt(count)} filas · ${state.table.headers.length} columnas`:'Archivos .xlsx o .csv · Hasta 30 MB';
  $('sheet-control').hidden=!state.sourceBook;
  if(state.sourceBook){$('source-sheet').innerHTML=state.sourceBook.sheets.map((s,i)=>`<option value="${i}" ${i===state.sheetIndex?'selected':''}>${h(s.name)}</option>`).join('');$('header-row').value=state.table.headerRow;}
  $('template-title').textContent=state.templateName;
  $('template-subtitle').textContent=state.templateInfo?'Haz clic para cambiar tu plantilla.':'Elige una plantilla .xlsx para continuar.';
  $('template-tag').textContent=state.templateType==='builtin'?'Incluida':'Personalizada';
  $('template-meta').textContent=state.templateInfo?`${state.templateInfo.headers.length} columnas · ${state.templateInfo.catalogs.categories.length} categorías`:'Sin plantilla disponible';
  $('demo-notice').hidden=!state.isDemo;
  $('step-1').classList.toggle('complete',!!state.table);$('step-2').classList.toggle('current',!!state.table);$('step-3').classList.toggle('current',!!state.rows.length);
}
function rebuild({resetPage=false,save=true}={}){
  if(resetPage)state.page=1;
  const table=state.table??{records:[]};
  const converted=convertRecords(table,state.appliedMapping,state.appliedRules,state.appliedCategoryMap,state.overrides,state.manualRows);
  state.rows=validateRows(converted,state.templateInfo?.catalogs,state.excluded);
  const available=new Set(state.rows.map(r=>r.id));state.selected=new Set([...state.selected].filter(id=>available.has(id)));
  renderAll();if(save)scheduleSave();
}
function renderAll(){
  const active=state.rows.filter(r=>!r.excluded),hasRows=!!state.rows.length;
  const duplicates=new Map();for(const r of active)if(r.values.sku)duplicates.set(r.values.sku,(duplicates.get(r.values.sku)??0)+1);
  const groupCount=[...duplicates.values()].filter(n=>n>1).length;
  $('metric-total').textContent=hasRows?fmt(active.length):'—';
  $('metric-total-note').textContent=hasRows?`${fmt(new Set(active.map(r=>r.values.sku).filter(Boolean)).size)} SKU únicos${state.excluded.size?' · '+fmt(state.excluded.size)+' excluidos':''}`:'Esperando tu archivo';
  $('metric-clean').textContent=hasRows?fmt(active.filter(r=>r.errorCount===0).length):'—';
  $('metric-errors').textContent=hasRows?fmt(active.filter(r=>r.errorCount>0).length):'—';
  $('metric-duplicates').textContent=hasRows?fmt(groupCount):'—';
  $('apply-btn').disabled=!state.table;$('export-btn').disabled=!active.length||!state.templateInfo;
  $('bulk-btn').disabled=!active.length;$('undo-btn').disabled=!state.history.length;
  $('report-btn').disabled=!active.some(r=>r.issues.length);
  $('reset-overrides-btn').disabled=!Object.keys(state.overrides).length;
  const cats=[...new Set(state.rows.map(r=>r.values.category).filter(Boolean))].sort();
  if(state.categoryFilter&&!cats.includes(state.categoryFilter))state.categoryFilter='';
  $('category-filter').innerHTML='<option value="">Todas las categorías</option>'+cats.map(c=>`<option value="${h(c)}" ${state.categoryFilter===c?'selected':''}>${h(c)}</option>`).join('');
  $('status-filter').value=state.statusFilter;
  const mismatch=active.some(r=>r.issues.some(i=>i.field==='category'||i.field==='warehouse'));
  $('review-banner').hidden=!mismatch;
  $('review-banner-text').textContent='Hay categorías o almacenes que no coinciden con tu plantilla.';
  updateFiles();renderTable();
}
function currentFiltered(){return filterRows(state.rows,{search:state.search,category:state.categoryFilter,status:state.statusFilter});}
function renderTable(){
  state.filtered=currentFiltered();
  const total=state.filtered.length,pages=Math.max(1,Math.ceil(total/state.pageSize));state.page=Math.min(state.page,pages);
  const offset=(state.page-1)*state.pageSize,visible=state.filtered.slice(offset,offset+state.pageSize);
  $('products-table').hidden=!visible.length;$('table-empty').hidden=!!visible.length;
  if(!visible.length){
    $('table-empty').innerHTML=state.rows.length?`<div class="empty-grid">${icon('search')}</div><h3>No hay productos en esta vista</h3><p>Prueba otro filtro o busca un término diferente.</p><button type="button" class="button button-outline" id="clear-filters-btn">Limpiar filtros</button>`:
    `<div class="empty-grid">${icon('grid')}<span class="empty-mini-check">${icon('check')}</span></div><h3>Tu próximo catálogo empieza aquí</h3><p>Sube tu inventario y revisa el resultado antes de descargarlo.<br>Puedes cambiar una fila o cientos a la vez.</p><button type="button" class="button button-outline" id="empty-upload-btn">${icon('upload')}Elegir mi inventario</button>`;
  }
  function input(row,key,cls='',list=''){
    const field=fieldMeta[key],numeric=field.type==='number',invalid=row.issues.some(i=>i.field===key&&i.level==='error');
    const value=row.values[key];
    return `<input class="cell-input ${cls} ${invalid?'invalid':''}" data-row="${h(row.id)}" data-cell="${key}" aria-label="${h(field.label)} de ${h(row.values.sku||'fila '+row.sourceRow)}" ${numeric?'type="number" step="'+(key==='quantity'?'1':'0.01')+'"':'type="text"'} value="${h(value??'')}" ${list?'list="'+list+'"':''}>`;
  }
  $('products-body').innerHTML=visible.map(row=>{
    const v=row.values;
    const review=row.excluded?'<span class="tag tag-muted">Excluido</span>':row.errorCount?`<button type="button" class="issue-button error" data-open="${h(row.id)}">${icon('alert')}${row.errorCount} ${row.errorCount===1?'error':'errores'}</button>`:row.warningCount?`<button type="button" class="issue-button warning" data-open="${h(row.id)}">${icon('info')}${row.warningCount} ${row.warningCount===1?'aviso':'avisos'}</button>`:`<button type="button" class="issue-button ok" data-open="${h(row.id)}">${icon('check')}Sin errores</button>`;
    return `<tr class="${row.excluded?'excluded-row':''} ${state.selected.has(row.id)?'selected-row':''}"><td class="select-col"><input type="checkbox" data-select="${h(row.id)}" aria-label="Seleccionar ${h(v.sku||'producto')}" ${state.selected.has(row.id)?'checked':''}></td><td class="sku-col">${input(row,'sku','sku-input')}<span class="cell-meta">${row.sourceRow?'Origen · fila '+row.sourceRow:'Producto nuevo'}</span></td><td>${input(row,'name','name-input')}<span class="cell-meta">${h(v.condition||'Sin condición')} · ${h(v.status||'Sin estado')} · ${h(v.currency)}</span></td><td>${input(row,'category','category-input','category-options')}</td><td>${input(row,'quantity','quantity-input')}</td><td>${input(row,'basePrice','number-input')}</td><td>${input(row,'currentPrice','number-input')}</td><td>${input(row,'location','location-input')}</td><td>${review}</td><td><button type="button" class="button button-icon row-edit-button" data-open="${h(row.id)}" aria-label="Editar ${h(v.name||v.sku)}">${icon('edit')}</button></td></tr>`;
  }).join('');
  const selected=visible.filter(r=>state.selected.has(r.id)).length;
  $('select-page').checked=visible.length>0&&selected===visible.length;$('select-page').indeterminate=selected>0&&selected<visible.length;
  $('table-range').textContent=total?`${fmt(offset+1)}–${fmt(Math.min(offset+state.pageSize,total))} de ${fmt(total)} productos`:state.rows.length?'0 productos con estos filtros':'Sin productos todavía';
  $('page-number').textContent=`${state.page} / ${fmt(pages)}`;$('prev-page').disabled=state.page<=1;$('next-page').disabled=state.page>=pages;
  $('selection-bar').hidden=!state.selected.size;$('selection-count').textContent=`${fmt(state.selected.size)} ${state.selected.size===1?'producto seleccionado':'productos seleccionados'}`;
}

function applySettings(showMessage=true){
  const iva=numberValue(state.rules.iva);
  if(iva===null||iva<0||iva>100){toast('Escribe un IVA entre 0 y 100.',true);return;}
  if(!state.table)return;
  state.appliedRules=clone(state.rules);state.appliedMapping=clone(state.mapping);state.appliedCategoryMap=clone(state.categoryMap);
  state.dirtyRules=false;$('pending-rules').hidden=true;
  rebuild({resetPage:true});if(showMessage)toast(`Reglas aplicadas a ${fmt(state.table.records.length)} filas. Se conservaron los ajustes manuales.`);
}
function resetEdits(){state.overrides={};state.manualRows=[];state.excluded=new Set();state.selected=new Set();state.history=[];state.categoryMap={};state.appliedCategoryMap={};state.pendingNewId=null;}
function confirmReplace(){return !hasEdits()||confirm('Cargar otro archivo iniciará una carga nueva. Descarga primero el Excel actual si quieres conservarlo. ¿Continuar?');}
async function importSource(file){
  if(!file)return;if(file.size>MAX_BYTES){toast('El archivo supera los 30 MB. Divide el inventario en lotes.',true);return;}
  if(!/\.(xlsx|csv)$/i.test(file.name)){toast('Selecciona un archivo .xlsx o .csv. Para .xls, guárdalo primero como .xlsx.',true);return;}
  if(!confirmReplace())return;
  loading('Leyendo tu inventario…','Solo se leen las filas con datos, aunque el archivo tenga filas vacías al final.');await breathe();
  try{
    let book=null,sheet,bytes=null,index=0;
    if(/\.csv$/i.test(file.name))sheet=readCsv(await file.text());
    else{bytes=await file.arrayBuffer();book=await readWorkbook(bytes,window.JSZip);index=Math.max(0,book.sheets.findIndex(s=>!s.hidden));sheet=await readSheet(book,index);}
    const table=sheet.rows.length?sourceTable(sheet):{headers:[],records:[],headerRow:1,name:sheet.name};delete sheet.xml;
    resetEdits();Object.assign(state,{sourceName:file.name,sourceBytes:bytes,sourceBook:book,sourceSheet:sheet,sheetIndex:index,table,isDemo:false,mapping:autoMap(table.headers),page:1,search:'',categoryFilter:'',statusFilter:'all'});
    $('search-input').value='';renderMapping();activateTab('mapping');applySettings(false);
    toast(table.records.length?`${fmt(table.records.length)} filas cargadas. Revisa las columnas y los avisos.`:'Esta hoja no contiene productos. Elige otra hoja del archivo.');
  }catch(e){toast('No se pudo cargar el inventario. '+errorMessage(e),true);}finally{stopLoading();$('source-input').value='';}
}
async function changeSheet(){
  if(!state.sourceBook)return;
  if(hasEdits()&&!confirm('Cambiar de hoja o de encabezados reiniciará los ajustes de esta carga. ¿Continuar?')){updateFiles();return;}
  const index=Number($('source-sheet').value),requestedHeader=Number($('header-row').value);
  loading('Leyendo la hoja seleccionada…');await breathe();
  try{
    const sheet=index===state.sheetIndex&&state.sourceSheet?state.sourceSheet:await readSheet(state.sourceBook,index);
    const table=sheet.rows.length?sourceTable(sheet,index===state.sheetIndex?requestedHeader:undefined):{headers:[],records:[],headerRow:1,name:sheet.name};delete sheet.xml;
    resetEdits();state.sheetIndex=index;state.sourceSheet=sheet;state.table=table;state.mapping=autoMap(table.headers);renderMapping();applySettings(false);toast('Hoja actualizada.');
  }catch(e){toast(errorMessage(e),true);updateFiles();}finally{stopLoading();}
}
async function loadTemplate(bytes,name,type='custom'){
  const book=await readWorkbook(bytes,window.JSZip),info=await templateInfo(book);
  Object.assign(state,{templateBytes:bytes,templateName:name,templateType:type,templateInfo:info});renderCatalogOptions();renderCategoryMap();rebuild({save:false});
}
async function importTemplate(file){
  if(!file)return;if(file.size>MAX_BYTES||!/\.xlsx$/i.test(file.name)){toast('La plantilla debe ser un archivo .xlsx de hasta 30 MB.',true);return;}
  loading('Leyendo tu plantilla…');await breathe();
  try{await loadTemplate(await file.arrayBuffer(),file.name);scheduleSave();toast('Plantilla actualizada. Las categorías y los almacenes ya se comparan con este archivo.');}
  catch(e){toast(errorMessage(e),true);}finally{stopLoading();$('template-input').value='';}
}

function saveManualChanges(changes,label='edit'){
  if(!changes.length){toast('No había valores distintos para modificar.');return;}
  const history=[];
  for(const c of changes){
    const existing=state.overrides[c.id]??{};
    history.push({...c,hadOverride:Object.hasOwn(existing,c.field),previousOverride:existing[c.field]});
    state.overrides[c.id]={...existing,[c.field]:c.after};
  }
  addHistory({type:'edit',changes:history,label});rebuild();
}
function undo(){
  const entry=state.history.pop();if(!entry)return;
  if(entry.type==='edit')for(const c of [...entry.changes].reverse()){
    if(c.hadOverride)state.overrides[c.id][c.field]=c.previousOverride;
    else{delete state.overrides[c.id]?.[c.field];if(state.overrides[c.id]&&!Object.keys(state.overrides[c.id]).length)delete state.overrides[c.id];}
  }
  if(entry.type==='exclude')for(const c of entry.changes){if(c.before)state.excluded.add(c.id);else state.excluded.delete(c.id);}
  if(entry.type==='add'){state.manualRows=state.manualRows.filter(r=>r.id!==entry.id);delete state.overrides[entry.id];state.excluded.delete(entry.id);}
  if(entry.type==='resetOverrides')state.overrides=entry.before;
  rebuild();toast('Se deshizo el último cambio.');
}
function excludeRows(ids,restore=false){
  const changes=[];for(const id of ids){changes.push({id,before:state.excluded.has(id)});restore?state.excluded.delete(id):state.excluded.add(id);}
  if(changes.length)addHistory({type:'exclude',changes});state.selected.clear();rebuild();toast(`${fmt(changes.length)} ${restore?'productos restaurados.':'productos excluidos del Excel.'}`);
}
function editFieldHtml(field,row,extra=false){
  const issue=row.issues?.find(i=>i.field===field.key&&i.level==='error');
  const value=row.values[field.key]??'',list=field.key==='category'?'category-options':field.key==='warehouse'?'warehouse-options':field.key==='currency'?'currency-options':'';
  const wide=field.key==='description'?'wide':field.key==='name'?'half-wide':'';
  const control=field.type==='textarea'?`<textarea name="${field.key}" id="edit-${field.key}" class="${issue?'invalid':''}" rows="4">${h(value)}</textarea>`:`<input id="edit-${field.key}" name="${field.key}" class="${issue?'invalid':''}" value="${h(value)}" ${field.type==='number'?`type="number" step="${field.key==='quantity'?'1':'0.01'}"`:'type="text"'} ${list?'list="'+list+'"':''}>`;
  return `<label class="field ${!extra?wide:''}">${h(field.label)}${control}${issue?`<span class="field-error">${h(issue.message)}</span>`:''}</label>`;
}
function openEdit(id,focusField){
  const row=state.rows.find(r=>r.id===id);if(!row)return;state.editId=id;
  $('edit-origin').textContent=row.sourceRow?`FILA ${row.sourceRow} DEL ARCHIVO ORIGINAL`:'PRODUCTO NUEVO';
  $('edit-heading').textContent=row.values.sku?'Editar '+row.values.sku:'Editar producto';
  $('edit-issues').innerHTML=row.issues?.length?'<ul class="issue-list">'+row.issues.map(i=>`<li>${h(i.message)}</li>`).join('')+'</ul>':'';
  const mainKeys=['sku','name','description','quantity','basePrice','currentPrice','condition','status','currency','category','warehouse','location'];
  $('edit-main-fields').innerHTML=mainKeys.map(key=>editFieldHtml(fieldMeta[key],row)).join('');
  $('edit-extra-fields').innerHTML=FIELDS.filter(f=>!mainKeys.includes(f.key)).map(f=>editFieldHtml(f,row,true)).join('');
  $('exclude-row-btn').textContent=row.excluded?'Restaurar al Excel':'Excluir del Excel';
  $('edit-dialog').showModal();setTimeout(()=>$(focusField?'edit-'+focusField:'edit-name')?.focus(),40);
}
function newRow(){
  const id='manual-'+crypto.randomUUID();const values=Object.fromEntries(FIELDS.map(f=>[f.key,f.type==='number'?null:'']));
  Object.assign(values,{condition:state.appliedRules.condition||'nuevo',status:state.appliedRules.status||'draft',currency:state.appliedRules.currency||'MXN',warehouse:state.appliedRules.warehouse||''});
  state.pendingNewId=id;state.manualRows.push({id,sourceRow:null,values});rebuild({save:false});openEdit(id,'sku');
}
function bulkTarget(){const scope=$('bulk-scope').value,filteredIds=new Set(state.filtered.map(r=>r.id));return state.rows.filter(r=>!r.excluded&&(scope==='all'||scope==='selected'&&state.selected.has(r.id)||scope==='filtered'&&filteredIds.has(r.id)));}
function updateBulk(){
  const field=$('bulk-field').value,op=$('bulk-operation').value;
  $('bulk-find-wrap').hidden=op!=='replace';$('bulk-value-label').textContent=op==='percent'?'Porcentaje de ajuste (ej. 10 o -10)':op==='replace'?'Reemplazar por':'Nuevo valor';
  const value=$('bulk-value');value.setAttribute('list',field==='category'?'category-options':field==='warehouse'?'warehouse-options':'');
  const count=bulkTarget().length;$('bulk-summary').textContent=`Se modificarán ${fmt(count)} productos${field==='prices'?' en ambos precios':''}.`;
  $('bulk-submit').disabled=!count;
}
function openBulk(){
  $('bulk-scope').value=state.selected.size?'selected':'filtered';$('bulk-value').value='';$('bulk-find').value='';$('bulk-error').hidden=true;updateBulk();$('bulk-dialog').showModal();
}
function exportTarget(){
  const mode=document.querySelector('input[name="export-mode"]:checked')?.value??'all';
  return state.rows.filter(r=>!r.excluded&&(mode==='all'||mode==='clean'&&!r.errorCount||mode==='selected'&&state.selected.has(r.id)));
}
function updateExportCount(){const count=exportTarget().length;$('export-count').textContent=`${fmt(count)} productos`;$('confirm-export').disabled=!count;}
function openExport(){
  if(state.dirtyRules&&state.table){applySettings(false);if(state.dirtyRules)return;toast('Se aplicaron las reglas pendientes antes de preparar la descarga.');}
  const active=state.rows.filter(r=>!r.excluded),clean=active.filter(r=>!r.errorCount),selected=active.filter(r=>state.selected.has(r.id));
  const errors=active.length-clean.length;
  $('export-warning').hidden=!errors;$('export-warning').textContent=`Hay ${fmt(errors)} filas con errores que podrían ser rechazadas al importar. Puedes corregirlas o descargar solo las filas sin errores.`;
  $('export-choices').innerHTML=[['all','Todo el catálogo','Incluye las filas con incidencias para seguir trabajando.',active.length],['clean','Solo filas sin errores','Los avisos y los SKU repetidos se mantienen.',clean.length],['selected','Solo productos seleccionados','Respeta la selección que hiciste en la tabla.',selected.length]].map(([key,title,note,count])=>`<label class="export-choice"><input type="radio" name="export-mode" value="${key}" ${key==='all'?'checked':''} ${!count?'disabled':''}><div>${h(title)}<small>${h(note)}</small></div><strong>${fmt(count)}</strong></label>`).join('');
  const base=sourceLabel().replace(/\.(xlsx|csv)$/i,'').replace(/[^\p{L}\p{N}_-]+/gu,'_').replace(/_+$/,'')||'Catalogo_Noventia';
  $('export-filename').value=base+'_plantilla.xlsx';updateExportCount();$('export-dialog').showModal();
}
function download(data,filename,type){const blob=data instanceof Blob?data:new Blob([data],{type});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
async function exportFile(){
  if(!state.templateInfo){toast('Carga una plantilla antes de exportar.',true);return;}
  const rows=exportTarget();if(!rows.length)return;
  let filename=$('export-filename').value.trim().replace(/[<>:"/\\|?*\x00-\x1f]/g,'_')||'Catalogo_Noventia_plantilla.xlsx';if(!/\.xlsx$/i.test(filename))filename+='.xlsx';
  $('export-dialog').close();loading('Preparando tu Excel…',`Conservando la plantilla y escribiendo ${fmt(rows.length)} productos.`);await breathe();
  try{const bytes=await exportWorkbook(state.templateBytes,state.templateInfo,rows,window.JSZip);download(bytes,filename,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');toast(`Excel descargado con ${fmt(rows.length)} productos.`);}
  catch(e){toast('No se pudo exportar. '+errorMessage(e),true);}finally{stopLoading();}
}

function demo(){
  if(!confirmReplace())return;
  resetEdits();const headers=['Código Producto','Descripción','Caracteristicas','Categoria','Precio','Existencia','Ubicación'];
  const samples=[
    ['01062026001','Bolsa de mano','Color: beige. Material: sintético.','Accesorios de mujer',350,2,'A9-4'],
    ['01062026002','Reloj de pulsera','Correa ajustable. Color: azul.','Relojería',220,1,'B7-3'],
    ['01062026003','Mochila urbana','Dos compartimientos. Color: negro.','Accesorios unisex',180,4,'C1-3'],
    ['01062026003','Mochila urbana','Dos compartimientos. Color: verde.','Accesorios unisex',180,1,'C1-3'],
    ['01062026005','Organizador de viaje','Set de tres piezas.','Accesorios unisex','',1,'A4-1']
  ];
  const rows=[{number:1,cells:Object.fromEntries(headers.map((x,i)=>[i,x]))},...samples.map((r,i)=>({number:i+2,cells:Object.fromEntries(r.map((x,j)=>[j,x]))}))];
  Object.assign(state,{sourceName:'Ejemplo_inventario.xlsx',sourceBytes:null,sourceBook:null,sourceSheet:{name:'Ejemplo',rows},sheetIndex:0,table:sourceTable({name:'Ejemplo',rows}),isDemo:true,page:1,search:'',categoryFilter:'',statusFilter:'all'});
  state.mapping=autoMap(headers);$('search-input').value='';renderMapping();applySettings(false);toast('Ejemplo cargado. Puedes editarlo y probar la descarga.');
}
async function resetAll(){
  if((state.rows.length||state.table)&&!confirm('¿Iniciar una carga nueva? Descarga tu Excel antes si necesitas conservar los cambios.'))return;
  clearTimeout(saveTimer);saveVersion++;resetEdits();Object.assign(state,{sourceName:'',sourceBytes:null,sourceBook:null,sourceSheet:null,sheetIndex:0,table:null,mapping:{},appliedMapping:{},rows:[],filtered:[],isDemo:false,page:1,search:'',categoryFilter:'',statusFilter:'all',dirtyRules:false});
  $('search-input').value='';$('pending-rules').hidden=true;renderMapping();renderCategoryMap();rebuild({save:false});await deleteDraft();toast('Lista para una nueva carga.');
}

function bindEvents(){
  $('source-input').addEventListener('change',e=>importSource(e.target.files[0]));$('template-input').addEventListener('change',e=>importTemplate(e.target.files[0]));
  for(const [id,handler] of [['source-drop',importSource],['template-drop',importTemplate]]){
    const drop=$(id);for(const name of ['dragover','dragenter'])drop.addEventListener(name,e=>{e.preventDefault();drop.classList.add('drag-over');});
    for(const name of ['dragleave','drop'])drop.addEventListener(name,()=>drop.classList.remove('drag-over'));
    drop.addEventListener('drop',e=>{e.preventDefault();handler(e.dataTransfer.files[0]);});
  }
  $('source-sheet').addEventListener('change',changeSheet);$('header-row').addEventListener('change',changeSheet);
  $('demo-btn').addEventListener('click',demo);$('reset-btn').addEventListener('click',resetAll);$('apply-btn').addEventListener('click',()=>applySettings());
  document.querySelectorAll('[data-tab]').forEach(tab=>tab.addEventListener('click',()=>activateTab(tab.dataset.tab)));
  document.querySelectorAll('[data-rule]').forEach(input=>input.addEventListener('change',()=>{state.rules[input.dataset.rule]=input.type==='checkbox'?input.checked:input.type==='number'?input.value:input.value.trim();markPending();}));
  $('mapping-grid').addEventListener('change',e=>{if(e.target.dataset.mapping){state.mapping[e.target.dataset.mapping]=Number(e.target.value);markPending();}});
  $('show-mapping-btn').addEventListener('click',()=>{state.showAllMapping=!state.showAllMapping;renderMapping();});
  $('automap-btn').addEventListener('click',()=>{if(state.table){state.mapping=autoMap(state.table.headers);renderMapping();markPending();toast('Columnas detectadas. Aplica las reglas para actualizar la tabla.');}});
  $('category-equivalences').addEventListener('change',e=>{if(e.target.dataset.categoryMap!==undefined){state.categoryMap[e.target.dataset.categoryMap]=e.target.value;markPending();}});
  $('use-warehouse-btn').addEventListener('click',()=>{const first=state.templateInfo?.catalogs.warehouses[0];if(!first){toast('La plantilla no incluye un catálogo de almacenes.',true);return;}state.rules.warehouse=first.value;$('rule-warehouse').value=first.value;markPending();toast('Almacén seleccionado. Aplica las reglas para actualizar todos los productos.');});
  $('resolve-categories-btn').addEventListener('click',()=>{activateTab('categories');$('rules-section').scrollIntoView({behavior:'smooth'});});
  $('search-input').addEventListener('input',e=>{state.search=e.target.value;state.page=1;renderTable();});
  $('category-filter').addEventListener('change',e=>{state.categoryFilter=e.target.value;state.page=1;renderTable();});
  $('status-filter').addEventListener('change',e=>{state.statusFilter=e.target.value;state.page=1;renderTable();});
  document.querySelectorAll('[data-metric]').forEach(metric=>metric.addEventListener('click',()=>{
    if(!state.rows.length)return;state.statusFilter=metric.dataset.metric==='clean'?'noerrors':metric.dataset.metric;state.page=1;$('status-filter').value=state.statusFilter;renderTable();$('table-container').scrollIntoView({behavior:'smooth',block:'center'});
  }));
  $('prev-page').addEventListener('click',()=>{state.page--;renderTable();});$('next-page').addEventListener('click',()=>{state.page++;renderTable();});
  $('page-size').addEventListener('change',e=>{state.pageSize=Number(e.target.value);state.page=1;renderTable();});
  $('select-page').addEventListener('change',e=>{const visible=state.filtered.slice((state.page-1)*state.pageSize,state.page*state.pageSize);for(const row of visible)e.target.checked?state.selected.add(row.id):state.selected.delete(row.id);renderTable();});
  $('products-body').addEventListener('change',e=>{
    if(e.target.dataset.select){e.target.checked?state.selected.add(e.target.dataset.select):state.selected.delete(e.target.dataset.select);renderTable();return;}
    const {row:id,cell:field}=e.target.dataset;if(!id||!field)return;
    const row=state.rows.find(r=>r.id===id),after=fieldMeta[field].type==='number'?numberValue(e.target.value):e.target.value;
    saveManualChanges([{id,field,before:row.values[field],after}]);
  });
  $('products-body').addEventListener('click',e=>{const open=e.target.closest('[data-open]');if(open)openEdit(open.dataset.open);});
  $('table-container').addEventListener('click',e=>{if(e.target.closest('#empty-upload-btn'))$('source-input').click();if(e.target.closest('#clear-filters-btn')){state.search='';state.categoryFilter='';state.statusFilter='all';state.page=1;$('search-input').value='';renderAll();}});
  $('select-filtered-btn').addEventListener('click',()=>{state.filtered.forEach(r=>state.selected.add(r.id));renderTable();});
  $('clear-selection-btn').addEventListener('click',()=>{state.selected.clear();renderTable();});
  $('exclude-selected-btn').addEventListener('click',()=>{if(!state.selected.size)return;excludeRows([...state.selected]);});
  $('undo-btn').addEventListener('click',undo);$('bulk-btn').addEventListener('click',openBulk);$('add-row-btn').addEventListener('click',newRow);
  $('reset-overrides-btn').addEventListener('click',()=>{if(confirm('¿Quitar todos los ajustes manuales y volver a los valores de las reglas?')){addHistory({type:'resetOverrides',before:clone(state.overrides)});state.overrides={};rebuild();toast('Se quitaron los ajustes manuales. Puedes deshacerlo.');}});
  $('edit-form').addEventListener('submit',e=>{
    e.preventDefault();const row=state.rows.find(r=>r.id===state.editId);if(!row)return;
    const data=new FormData(e.currentTarget),changes=[];
    for(const f of FIELDS){const val=data.get(f.key);if(val===null)continue;const after=f.type==='number'?numberValue(val):String(val);if(after!==row.values[f.key])changes.push({id:row.id,field:f.key,before:row.values[f.key],after});}
    if(state.pendingNewId===row.id){
      const saved=state.manualRows.find(r=>r.id===row.id);for(const change of changes)saved.values[change.field]=change.after;
      state.pendingNewId=null;addHistory({type:'add',id:row.id});rebuild();toast('Producto agregado.');
    }else if(changes.length){saveManualChanges(changes);toast('Producto guardado.');}$('edit-dialog').close();
  });
  $('edit-dialog').addEventListener('close',()=>{if(state.pendingNewId){state.manualRows=state.manualRows.filter(r=>r.id!==state.pendingNewId);delete state.overrides[state.pendingNewId];state.pendingNewId=null;rebuild({save:false});}});
  $('exclude-row-btn').addEventListener('click',()=>{const row=state.rows.find(r=>r.id===state.editId);if(row){if(state.pendingNewId!==row.id)excludeRows([row.id],row.excluded);$('edit-dialog').close();}});
  for(const id of ['bulk-scope','bulk-field','bulk-operation'])$(id).addEventListener('change',updateBulk);
  $('bulk-form').addEventListener('submit',e=>{
    e.preventDefault();try{
      const target=bulkTarget(),field=$('bulk-field').value,fields=field==='prices'?['basePrice','currentPrice']:[field];
      if($('bulk-operation').value==='percent'&&fields.some(k=>fieldMeta[k].type!=='number'))throw new Error('El ajuste porcentual solo se aplica a campos numéricos.');
      const changes=bulkChanges(target,{fields,operation:$('bulk-operation').value,value:$('bulk-value').value,find:$('bulk-find').value});
      saveManualChanges(changes,'bulk');$('bulk-dialog').close();if(changes.length)toast(`Se actualizaron ${fmt(new Set(changes.map(c=>c.id)).size)} productos.`);
    }catch(error){$('bulk-error').textContent=errorMessage(error);$('bulk-error').hidden=false;}
  });
  $('export-btn').addEventListener('click',openExport);$('export-choices').addEventListener('change',updateExportCount);$('confirm-export').addEventListener('click',exportFile);
  $('report-btn').addEventListener('click',()=>download(reportCsv(state.rows.filter(r=>!r.excluded)),'Incidencias_catalogo_Noventia.csv','text/csv;charset=utf-8'));
  document.querySelectorAll('[data-close]').forEach(button=>button.addEventListener('click',()=>$(button.dataset.close).close()));
  $('bulk-field').innerHTML='<option value="prices">Ambos precios</option>'+FIELDS.map(f=>`<option value="${f.key}">${h(f.label.replace(' *',''))}</option>`).join('');
}

async function init(){
  bindEvents();renderMapping();syncRuleInputs();
  const noErrors=document.createElement('option');noErrors.value='noerrors';noErrors.textContent='Sin errores';$('status-filter').append(noErrors);
  try{const response=await fetch('./plantilla-base.xlsx');if(!response.ok)throw new Error('No se pudo cargar la plantilla base.');await loadTemplate(await response.arrayBuffer(),'Plantilla de carga Noventia','builtin');}
  catch(error){toast(errorMessage(error)+' Puedes cargar tu propia plantilla.',true);updateFiles();}
  const draft=await readDraft();
  if(draft?.version===1&&(draft.table||draft.manualRows?.length)){
    loading('Recuperando tu avance…');await breathe();
    try{
      if(draft.templateBytes)await loadTemplate(draft.templateBytes,draft.templateName,draft.templateType);
      for(const key of ['sourceName','sourceBytes','sheetIndex','table','rules','appliedRules','mapping','appliedMapping','categoryMap','appliedCategoryMap','overrides','manualRows','dirtyRules','isDemo'])if(draft[key]!==undefined)state[key]=draft[key];
      state.excluded=new Set(draft.excluded??[]);if(state.sourceBytes)state.sourceBook=await readWorkbook(state.sourceBytes,window.JSZip);
      syncRuleInputs();renderMapping();renderCategoryMap();rebuild({save:false});toast('Recuperamos la última carga guardada en este navegador.');
    }catch{toast('No se pudo recuperar la sesión anterior. Puedes cargar tus archivos de nuevo.',true);}finally{stopLoading();}
  }
  renderAll();
}
init();
