/**
 * E33 · TABLERO DE ARTISTA (co-manager)
 * Script dentro de un Google Sheet (Extensiones → Apps Script).
 *
 * TODO lo específico del artista vive en la hoja "Config" del Sheet.
 * Para cambiar de artista NO se edita código: se editan las celdas de "Config".
 *
 * Puesta en marcha:
 *   1) Menú 🔄 Tablero → Configurar / restaurar hojas  (crea Config, Recursos, Info)
 *   2) Llena la hoja Config (artista, cliente, ids de las 2 bases de Notion, equipo)
 *   3) Propiedades del script: NOTION_TOKEN = secret_xxx
 *   4) Menú 🔄 Tablero → Sincronizar ahora
 *
 * PERMISOS de la integración de Notion: Read + Update + Insert content,
 * conectada a las DOS bases (central y la del artista).
 */

var NV = '2022-06-28';
var T_PEND='Pendientes', T_COB='Cobros', T_SPL='Splits', T_LAN='Lanzamientos';
var T_LOG='Registro', T_RES='Recursos', T_INFO='Info', T_CFG='Config';

// Estados de la base central (pendientes)
var S = { DONE:'✅ Hecho', PROG:'En Ejecución', BLOCK:'🚫 Trabado', TODO:'📥 Pendientes por asignar' };

/* ===================== CONFIG ===================== */
var CFG_DEFAULTS = [
  ['Artista', 'Reelian'],
  ['Cliente (valor en base central)', 'REELIAN'],
  ['ID base central (pendientes)', '8bdbde52a4334fe0b157767fb73c2c5e'],
  ['ID base del artista', '35faac7f7cb380daa35ff61fcf9508ba'],
  ['Propiedad % del artista en Notion', '% Reelian'],
  ['Equipo (Nombre:código:estado en Notion)', 'Manfred:1234:🧑 Tareas de Manfred, Pepe:1111:👤 Tareas de Pepe, Isis:5678:👩 Tareas de Isis'],
  ['Enlace de agenda (botón flotante)', 'https://calendar.app.google/w5ABFfnJfz547JJN6'],
  ['Texto del botón de agenda', 'Reunión con Pepe'],
  ['Zona horaria', 'America/Mexico_City']
];

function cfg_(){
  var ss=SpreadsheetApp.getActive();
  var sh=ss.getSheetByName(T_CFG);
  var map={};
  if(sh && sh.getLastRow()>1){
    sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues().forEach(function(r){
      if(r[0]) map[String(r[0]).trim()]=String(r[1]||'').trim();
    });
  }
  var g=function(k,d){ return map[k]||d; };
  var equipo=[];
  g('Equipo (Nombre:código:estado en Notion)','').split(',').forEach(function(p){
    var b=p.split(':'); var n=(b[0]||'').trim();
    if(n) equipo.push({nombre:n, pin:(b[1]||'').trim(), estado:(b[2]||'').trim()||('👤 Tareas de '+n)});
  });
  return {
    artista:  g('Artista','Artista'),
    cliente:  g('Cliente (valor en base central)',''),
    dbCentral:g('ID base central (pendientes)','').replace(/-/g,''),
    dbArtista:g('ID base del artista','').replace(/-/g,''),
    propPct:  g('Propiedad % del artista en Notion','% Artista'),
    equipo:   equipo,
    agendaUrl:g('Enlace de agenda (botón flotante)',''),
    agendaTxt:g('Texto del botón de agenda','Agendar reunión'),
    tz:       g('Zona horaria', Session.getScriptTimeZone()||'America/Mexico_City')
  };
}

function onOpen(){
  SpreadsheetApp.getUi().createMenu('🔄 Tablero')
    .addItem('Sincronizar ahora','syncAll')
    .addItem('Programar sync (15 min)','setupTrigger')
    .addItem('Configurar / restaurar hojas','seedRecursos')
    .addItem('Ver estado','showStatus')
    .addToUi();
}
function token_(){ var t=PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN'); if(!t) throw new Error('Falta NOTION_TOKEN en Propiedades del script.'); return t; }
function H_(){ return {'Authorization':'Bearer '+token_(),'Notion-Version':NV}; }

/* ---------- lectores de propiedades ---------- */
function pSel(pr,n){ return (pr[n]&&pr[n].select)?pr[n].select.name:''; }
function pNum(pr,n){ return (pr[n]&&typeof pr[n].number==='number')?pr[n].number:''; }
function pChk(pr,n){ return (pr[n]&&pr[n].checkbox===true); }
function pTit(pr,n){ var a=(pr[n]&&pr[n].title)||[]; return a.map(function(x){return x.plain_text;}).join(''); }
function pTxt(pr,n){ var a=(pr[n]&&pr[n].rich_text)||[]; return a.map(function(x){return x.plain_text;}).join(''); }
function pDate(pr,n){ return (pr[n]&&pr[n].date)?pr[n].date.start:''; }

function queryDB_(db){
  if(!db) throw new Error('Falta un ID de base en la hoja Config.');
  var out=[], cursor=null, more=true, g=0;
  while(more && g<30){ g++;
    var payload={page_size:100}; if(cursor) payload.start_cursor=cursor;
    var res=UrlFetchApp.fetch('https://api.notion.com/v1/databases/'+db+'/query',{method:'post',contentType:'application/json',headers:H_(),payload:JSON.stringify(payload),muteHttpExceptions:true});
    var j=JSON.parse(res.getContentText());
    if(j.object==='error') throw new Error('Notion: '+(j.message||res.getContentText()));
    out=out.concat(j.results||[]); more=j.has_more; cursor=j.next_cursor;
  }
  return out;
}
function writeTab_(name, header, rows){
  var ss=SpreadsheetApp.getActive();
  var sh=ss.getSheetByName(name)||ss.insertSheet(name);
  sh.clearContents();
  var data=[header].concat(rows);
  var r=sh.getRange(1,1,data.length,header.length);
  r.setNumberFormat('@'); r.setValues(data);
  sh.getRange(1,1,1,header.length).setFontWeight('bold');
}
function colMap_(tab,col){
  var m={};
  try{
    var sh=SpreadsheetApp.getActive().getSheetByName(tab);
    if(sh&&sh.getLastRow()>1){
      var head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
      var ii=head.indexOf('Id'), ci=head.indexOf(col);
      if(ii>=0&&ci>=0){
        sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues().forEach(function(r){ if(r[ii]) m[r[ii]]=r[ci]; });
      }
    }
  }catch(e){}
  return m;
}

/* ===================== SYNC ===================== */
function syncAll(){
  var C=cfg_();

  // 1) Pendientes desde la base central (filtrado por Cliente)
  var pend=[];
  queryDB_(C.dbCentral).forEach(function(p){
    var pr=p.properties||{};
    if(C.cliente && pSel(pr,'Cliente')!==C.cliente) return;
    var tarea=pTit(pr,'Tarea'); if(!tarea) return;
    pend.push([pSel(pr,'Estado'), tarea, pDate(pr,'Vencimiento'), p.id, p.url||'', p.created_time||'']);
  });
  writeTab_(T_PEND, ['Estado','Tarea','Vencimiento','Id','Url','Creado'], pend);

  // 2) Base del artista → Cobros / Splits / Lanzamientos
  var cob=[], spl=[], lan=[];
  var prevEtapa=colMap_(T_SPL,'Etapa'), prevDocSpl=colMap_(T_SPL,'Doc'), prevDocLan=colMap_(T_LAN,'Doc');
  var IGNORAR_ESTADOS=['Archivado','Cancelado','Plan de trabajo','Lanzamientos 2026','Seguimiento'];
  queryDB_(C.dbArtista).forEach(function(p){
    var pr=p.properties||{};
    var nombre=pTit(pr,'Nombre'); if(!nombre) return;
    if(nombre.indexOf('TEMPLATE')>=0) return;
    var tipo=pSel(pr,'Tipo de entrada'), estado=pSel(pr,'Estado');
    if(IGNORAR_ESTADOS.indexOf(estado)>=0) return;
    if(estado===C.artista) return; // estado heredado de Trello con el nombre del artista

    var fee=pNum(pr,'Fee USD'), anti=pNum(pr,'Anticipo USD'), saldo=pNum(pr,'Saldo USD');
    var aRec=pChk(pr,'Anticipo recibido'), sRec=pChk(pr,'Saldo recibido');
    if(fee!=='' && fee>0 && !(aRec && sRec) && estado!=='Hecho' && estado!==S.DONE){
      cob.push([nombre, pTxt(pr,'Artista / Contraparte'), fee, anti, saldo, aRec?'sí':'no', sRec?'sí':'no', estado, p.id]);
    }
    if(!pChk(pr,'Split firmado') && (tipo==='Track'||tipo==='Split')){
      spl.push([nombre, pTxt(pr,'Colaboradores'), pNum(pr,C.propPct), 'no', estado, p.id, prevEtapa[p.id]||'Solicitado', prevDocSpl[p.id]||'']);
    }
    if(tipo==='Lanzamiento'){
      lan.push([nombre, pDate(pr,'Fecha clave'), estado, pTxt(pr,'Sello / Distribuidor'), p.id, prevDocLan[p.id]||'']);
    }
  });
  writeTab_(T_COB, ['Nombre','Artista','Fee','Anticipo','Saldo','AnticipoRec','SaldoRec','Estado','Id'], cob);
  writeTab_(T_SPL, ['Track','Colaboradores','Pct','Firmado','Estado','Id','Etapa','Doc'], spl);
  writeTab_(T_LAN, ['Titulo','Fecha','Estado','Sello','Id','Doc'], lan);

  PropertiesService.getScriptProperties().setProperty('LAST_SYNC', new Date().toISOString());
}

/* ===================== WEB APP ===================== */
function doGet(){
  var C=cfg_();
  var t=HtmlService.createTemplateFromFile('Index');
  t.pendJson=JSON.stringify(readTab_(T_PEND));
  t.cobJson =JSON.stringify(readTab_(T_COB));
  t.splJson =JSON.stringify(readTab_(T_SPL));
  t.lanJson =JSON.stringify(readTab_(T_LAN));
  t.resJson =JSON.stringify(readTab_(T_RES));
  t.infoJson=JSON.stringify(readTab_(T_INFO));
  t.cfgJson =JSON.stringify(C);
  t.updatedAt=PropertiesService.getScriptProperties().getProperty('LAST_SYNC')||new Date().toISOString();
  var out=t.evaluate();
  out.setTitle(C.artista+' — E33');
  out.addMetaTag('viewport','width=device-width, initial-scale=1');
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}
function readTab_(name){
  var sh=SpreadsheetApp.getActive().getSheetByName(name);
  if(!sh||sh.getLastRow()<2) return [];
  var v=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues();
  var head=v[0], out=[];
  for(var i=1;i<v.length;i++){ var o={}; for(var c=0;c<head.length;c++) o[head[c]]=v[i][c]; out.push(o); }
  return out;
}

/* ---------- ACCIONES: pendientes → base central ---------- */
function setState(pageId, tarea, estado, note, who){
  if(!estado) throw new Error('Estado vacío.');
  var props={'Estado':{'select':{'name':estado}}};
  if(note&&note.trim()) props['Notas']=nota_(pageId,note,who);
  patch_(pageId,props);
  setCell_(T_PEND,'Id',pageId,'Estado',estado);
  log_(tarea, estado===S.DONE?'Hecho':(estado===S.PROG?'En progreso':(estado===S.BLOCK?'Trabado':'Estado→'+estado)), note||'', who);
  return {ok:true};
}
function postpone(pageId, tarea, iso, note, who){
  var props={'Vencimiento':{'date':{'start':iso}}};
  if(note&&note.trim()) props['Notas']=nota_(pageId,note,who);
  patch_(pageId,props); setCell_(T_PEND,'Id',pageId,'Vencimiento',iso);
  log_(tarea,'Pospuesto→'+iso, note||'', who); return {ok:true,date:iso};
}
function createTask(tarea, iso, estado, note, who){
  if(!tarea||!tarea.trim()) throw new Error('Escribe la tarea.');
  var C=cfg_();
  var props={'Tarea':{'title':[{'text':{'content':tarea.trim().substring(0,1900)}}]},'Estado':{'select':{'name':estado||S.TODO}}};
  if(C.cliente) props['Cliente']={'select':{'name':C.cliente}};
  if(iso) props['Vencimiento']={'date':{'start':iso}};
  if(note&&note.trim()) props['Notas']={'rich_text':[{'text':{'content':('✍️ '+(who||'')+': '+note.trim()).substring(0,1900)}}]};
  var res=UrlFetchApp.fetch('https://api.notion.com/v1/pages',{method:'post',contentType:'application/json',headers:H_(),payload:JSON.stringify({parent:{database_id:C.dbCentral},properties:props}),muteHttpExceptions:true});
  var j=JSON.parse(res.getContentText()); if(j.object==='error') throw new Error(j.message||'Error al crear');
  try{ SpreadsheetApp.getActive().getSheetByName(T_PEND).appendRow([estado||S.TODO,tarea.trim(),iso||'',j.id,j.url||'',j.created_time||'']); }catch(e){}
  log_(tarea.trim(),'Creada', note||'', who);
  return {ok:true, row:{Estado:estado||S.TODO, Tarea:tarea.trim(), Vencimiento:iso||'', Id:j.id, Url:j.url||'', Creado:j.created_time||''}};
}

/* ---------- ACCIONES: base del artista ---------- */
function markCobro(pageId, tarea, which, who){ // 'anticipo' | 'saldo' | 'ambos'
  var props={};
  if(which==='anticipo'||which==='ambos') props['Anticipo recibido']={'checkbox':true};
  if(which==='saldo'||which==='ambos')    props['Saldo recibido']={'checkbox':true};
  patch_(pageId,props);
  if(props['Anticipo recibido']) setCell_(T_COB,'Id',pageId,'AnticipoRec','sí');
  if(props['Saldo recibido'])    setCell_(T_COB,'Id',pageId,'SaldoRec','sí');
  log_(tarea,'Cobro '+which,'',who); return {ok:true};
}
function markSplit(pageId, tarea, who){
  patch_(pageId,{'Split firmado':{'checkbox':true}});
  setCell_(T_SPL,'Id',pageId,'Firmado','sí');
  log_(tarea,'Split firmado','',who); return {ok:true};
}
function setSplitStage(pageId, tarea, etapa, who){
  patch_(pageId,{'Split firmado':{'checkbox':etapa==='Firmado'}});
  setCell_(T_SPL,'Id',pageId,'Firmado',etapa==='Firmado'?'sí':'no');
  setCell_(T_SPL,'Id',pageId,'Etapa',etapa);
  log_(tarea,'Split '+etapa,'',who);
  return {ok:true};
}
function setRelease(pageId, tarea, estado, who){
  patch_(pageId,{'Estado':{'select':{'name':estado}}});
  setCell_(T_LAN,'Id',pageId,'Estado',estado);
  log_(tarea,'Lanzamiento→'+estado,'',who); return {ok:true};
}

/* ---------- ALTA de items en base del artista ---------- */
function createArtistPage_(props){
  var C=cfg_();
  var res=UrlFetchApp.fetch('https://api.notion.com/v1/pages',{method:'post',contentType:'application/json',headers:H_(),payload:JSON.stringify({parent:{database_id:C.dbArtista},properties:props}),muteHttpExceptions:true});
  var j=JSON.parse(res.getContentText());
  if(j.object==='error') throw new Error(j.message||'Notion rechazó (¿permiso Insert content y base del artista conectada?)');
  return j;
}
function createCobro(nombre, artista, fee, saldo, who){
  if(!nombre||!nombre.trim()) throw new Error('Falta el concepto.');
  var props={'Nombre':{'title':[{'text':{'content':nombre.trim().substring(0,1900)}}]},'Tipo de entrada':{'select':{'name':'Track'}},'Estado':{'select':{'name':'Pendiente'}}};
  var f=parseFloat(fee), s=parseFloat(saldo);
  if(!isNaN(f)) props['Fee USD']={'number':f};
  if(!isNaN(s)) props['Saldo USD']={'number':s};
  if(artista&&artista.trim()) props['Artista / Contraparte']={'rich_text':[{'text':{'content':artista.trim()}}]};
  var j=createArtistPage_(props);
  try{ SpreadsheetApp.getActive().getSheetByName(T_COB).appendRow([nombre.trim(),artista||'',isNaN(f)?'':f,'',isNaN(s)?'':s,'no','no','Pendiente',j.id]); }catch(e){}
  log_(nombre.trim(),'Cobro nuevo','',who);
  return {ok:true,row:{Nombre:nombre.trim(),Artista:artista||'',Fee:isNaN(f)?'':String(f),Anticipo:'',Saldo:isNaN(s)?'':String(s),AnticipoRec:'no',SaldoRec:'no',Estado:'Pendiente',Id:j.id}};
}
function createSplit(track, colaboradores, pct, who){
  if(!track||!track.trim()) throw new Error('Falta el track.');
  var C=cfg_();
  var props={'Nombre':{'title':[{'text':{'content':track.trim().substring(0,1900)}}]},'Tipo de entrada':{'select':{'name':'Track'}},'Split firmado':{'checkbox':false},'Estado':{'select':{'name':'Pendiente'}}};
  var p2=parseFloat(pct); if(!isNaN(p2)) props[C.propPct]={'number':p2};
  if(colaboradores&&colaboradores.trim()) props['Colaboradores']={'rich_text':[{'text':{'content':colaboradores.trim()}}]};
  var j=createArtistPage_(props);
  try{ SpreadsheetApp.getActive().getSheetByName(T_SPL).appendRow([track.trim(),colaboradores||'',isNaN(p2)?'':p2,'no','Pendiente',j.id,'Solicitado','']); }catch(e){}
  log_(track.trim(),'Split nuevo','',who);
  return {ok:true,row:{Track:track.trim(),Colaboradores:colaboradores||'',Pct:isNaN(p2)?'':String(p2),Firmado:'no',Estado:'Pendiente',Id:j.id,Etapa:'Solicitado',Doc:''}};
}
function createRelease(titulo, fecha, estado, who){
  if(!titulo||!titulo.trim()) throw new Error('Falta el título.');
  var props={'Nombre':{'title':[{'text':{'content':titulo.trim().substring(0,1900)}}]},'Tipo de entrada':{'select':{'name':'Lanzamiento'}},'Estado':{'select':{'name':estado||'🎛 Producción'}}};
  if(fecha) props['Fecha clave']={'date':{'start':fecha}};
  var j=createArtistPage_(props);
  try{ SpreadsheetApp.getActive().getSheetByName(T_LAN).appendRow([titulo.trim(),fecha||'',estado||'🎛 Producción','',j.id,'']); }catch(e){}
  log_(titulo.trim(),'Lanzamiento nuevo','',who);
  return {ok:true,row:{Titulo:titulo.trim(),Fecha:fecha||'',Estado:estado||'🎛 Producción',Sello:'',Id:j.id,Doc:''}};
}
function setDoc(section, pageId, tarea, link, who){
  var tab = section==='lan' ? T_LAN : T_SPL;
  setCell_(tab,'Id',pageId,'Doc',link||'');
  log_(tarea, 'Documento '+((link&&link.trim())?'agregado':'quitado'), link||'', who);
  return {ok:true};
}

/* ---------- helpers ---------- */
function patch_(pageId, props){
  var res=UrlFetchApp.fetch('https://api.notion.com/v1/pages/'+pageId,{method:'patch',contentType:'application/json',headers:H_(),payload:JSON.stringify({properties:props}),muteHttpExceptions:true});
  var j=JSON.parse(res.getContentText());
  if(j.object==='error') throw new Error(j.message||'Notion rechazó (¿permiso Update content y base conectada?)');
  return j;
}
function nota_(pageId, note, who){
  var cur='';
  try{
    var g=UrlFetchApp.fetch('https://api.notion.com/v1/pages/'+pageId,{headers:H_(),muteHttpExceptions:true});
    var gj=JSON.parse(g.getContentText());
    var na=(gj.properties&&gj.properties['Notas']&&gj.properties['Notas'].rich_text)||[];
    cur=na.map(function(x){return x.plain_text;}).join('');
  }catch(e){}
  var st=Utilities.formatDate(new Date(), cfg_().tz, 'dd/MM');
  var add=(cur?cur+'  -  ':'')+'✍️ '+(who||'')+' '+st+': '+note.trim();
  return {'rich_text':[{'text':{'content':add.substring(0,1900)}}]};
}
function setCell_(tab, keyCol, keyVal, setCol, value){
  try{
    var sh=SpreadsheetApp.getActive().getSheetByName(tab); if(!sh||sh.getLastRow()<2) return;
    var head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    var ki=head.indexOf(keyCol)+1, si=head.indexOf(setCol)+1; if(ki<1||si<1) return;
    var ids=sh.getRange(2,ki,sh.getLastRow()-1,1).getDisplayValues();
    for(var i=0;i<ids.length;i++){ if(ids[i][0]===keyVal){ sh.getRange(i+2,si).setValue(value); return; } }
  }catch(e){}
}
function log_(tarea, accion, detalle, who){
  try{
    var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(T_LOG);
    if(!sh){ sh=ss.insertSheet(T_LOG); sh.appendRow(['Fecha/hora','Ítem','Acción','Detalle','Quién']); sh.getRange(1,1,1,5).setFontWeight('bold'); }
    sh.appendRow([new Date(), tarea||'', accion, detalle||'', who||'—']);
  }catch(e){}
}

/* ===================== UTILIDADES ===================== */
function setupTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='syncAll') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(15).create();
  syncAll(); SpreadsheetApp.getUi().alert('Sync cada 15 min activado. Ya sincronicé ahora.');
}
function seedRecursos(){
  var ss=SpreadsheetApp.getActive();

  if(!ss.getSheetByName(T_CFG)){
    var sc=ss.insertSheet(T_CFG);
    sc.getRange(1,1,1,2).setValues([['Campo','Valor']]).setFontWeight('bold');
    sc.getRange(2,1,CFG_DEFAULTS.length,2).setValues(CFG_DEFAULTS);
    sc.setColumnWidth(1,330); sc.setColumnWidth(2,420);
  }else{
    var sc2=ss.getSheetByName(T_CFG);
    var have = sc2.getLastRow()>1 ? sc2.getRange(2,1,sc2.getLastRow()-1,1).getDisplayValues().map(function(r){return r[0];}) : [];
    CFG_DEFAULTS.forEach(function(kv){ if(have.indexOf(kv[0])<0) sc2.appendRow(kv); });
  }

  if(!ss.getSheetByName(T_RES)){
    var sh=ss.insertSheet(T_RES);
    sh.getRange(1,1,1,2).setValues([['Seccion','Link']]).setFontWeight('bold');
    var rows=[
      ['📋 Info General',''],['💰 Adelantos / Cobros',''],['🚀 Lanzamientos',''],
      ['📄 Splits / Contratos',''],['🎵 Catálogo',''],['💸 Gastos del artista',''],
      ['🎤 Shows',''],['🏕️ Campamentos','']
    ];
    sh.getRange(2,1,rows.length,2).setValues(rows);
    sh.setColumnWidth(1,190); sh.setColumnWidth(2,520);
  }

  if(!ss.getSheetByName(T_INFO)){
    var si=ss.insertSheet(T_INFO);
    si.getRange(1,1,1,2).setValues([['Campo','Valor']]).setFontWeight('bold');
    si.setColumnWidth(1,200); si.setColumnWidth(2,360);
  }
  var si2=ss.getSheetByName(T_INFO);
  var must=['Nombre artístico','Nombre legal','Rol principal','Género musical','País de origen','Ciudad base','Manager','Booking Agent','PRO principal','Publisher','IPI','Artistas principales','Email','Teléfono / WhatsApp','Instagram'];
  var haveI = si2.getLastRow()>1 ? si2.getRange(2,1,si2.getLastRow()-1,1).getDisplayValues().map(function(r){return r[0];}) : [];
  must.forEach(function(k){ if(haveI.indexOf(k)<0) si2.appendRow([k,'']); });

  SpreadsheetApp.getUi().alert('Listo. Llena la hoja "Config" (artista, cliente e ids de Notion) y luego "Info" y "Recursos". El tablero los toma al recargar.');
}
function showStatus(){
  var C=cfg_(), p=PropertiesService.getScriptProperties();
  SpreadsheetApp.getUi().alert(
    'Artista: '+C.artista+
    '\nCliente (central): '+(C.cliente||'(vacío)')+
    '\nBase central: '+(C.dbCentral||'(falta)')+
    '\nBase artista: '+(C.dbArtista||'(falta)')+
    '\nEquipo: '+C.equipo.map(function(e){return e.nombre;}).join(', ')+
    '\nToken Notion: '+(p.getProperty('NOTION_TOKEN')?'ok':'FALTA')+
    '\nÚltima sync: '+(p.getProperty('LAST_SYNC')||'(nunca)')
  );
}
