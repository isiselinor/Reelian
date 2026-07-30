/**
 * E33 · TABLERO DE ARTISTA (co-manager) — con módulo de CATÁLOGO
 * Script dentro de un Google Sheet (Extensiones → Apps Script).
 *
 * TODO lo específico del artista vive en la hoja "Config" del Sheet.
 * Para cambiar de artista NO se edita código: se editan las celdas de "Config".
 *
 * PERMISOS de la integración de Notion: Read + Update + Insert content,
 * conectada a las TRES bases (central, artista y catálogo).
 */

var NV = '2022-06-28';
var T_PEND='Pendientes', T_COB='Cobros', T_SPL='Splits', T_LAN='Lanzamientos', T_CAT='Catálogo';
var T_LOG='Registro', T_RES='Recursos', T_INFO='Info', T_CFG='Config', T_GAS='Gastos', T_CRE='Creditos';

// Estados de la base central (pendientes)
var S = { DONE:'✅ Hecho', PROG:'En Ejecución', BLOCK:'🚫 Trabado', TODO:'📥 Pendientes por asignar' };

/* ===================== CONFIG ===================== */
var CFG_DEFAULTS = [
  ['Artista', 'Reelian'],
  ['Cliente (valor en base central)', 'REELIAN'],
  ['ID base central (pendientes)', '8bdbde52a4334fe0b157767fb73c2c5e'],
  ['ID base del artista', '35faac7f7cb380daa35ff61fcf9508ba'],
  ['ID base catálogo (producción)', '64cf1324bdf34533bba69066a0a21e9f'],
  ['ID base gastos', '845ee5fa8bfa4d46b186f8ec79c809e7'],
  ['ID base créditos', 'cc1f8d9f540c41018b109a1fdc130aa7'],
  ['Carpeta Drive del catálogo (ID o link)', ''],
  ['Token de Samply (co-manager)', ''],
  ['Project ID de Samply (solo ese proyecto)', ''],
  ['Secreto del webhook Samply', 'reelian33'],
  ['Propiedad % del artista en Notion', '% Reelian'],
  ['Equipo (Nombre:código:estado en Notion)', 'Manfred:1234:🧑 Tareas de Manfred, Pepe:1111:👤 Tareas de Pepe, Isis:5678:👩 Tareas de Isis'],
  ['Enlace de agenda (botón flotante)', 'https://calendar.app.google/Wq95PbNQJ229vcLV8'],
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
    dbCatalogo:g('ID base catálogo (producción)','').replace(/-/g,''),
    dbGastos: g('ID base gastos','').replace(/-/g,''),
    dbCreditos: g('ID base créditos','').replace(/-/g,''),
    driveCat: g('Carpeta Drive del catálogo (ID o link)',''),
    samplyToken:  g('Token de Samply (co-manager)',''),
    samplyProject:g('Project ID de Samply (solo ese proyecto)',''),
    samplySecret: g('Secreto del webhook Samply','reelian33'),
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
    .addItem('Probar carpeta (Drive)','probarCarpeta')
    .addItem('Crear carpetas faltantes (Drive)','crearCarpetasFaltantes')
    .addItem('Vaciar catálogo (empezar de cero)','vaciarCatalogo')
    .addSeparator()
    .addItem('Samply · registrar webhook','registrarWebhookSamply')
    .addItem('Samply · ver proyectos','verProyectosSamply')
    .addItem('Samply · ver último aviso','verUltimoSamply')
    .addSeparator()
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
function pUrl(pr,n){ return (pr[n]&&pr[n].url)?pr[n].url:''; }

function queryDB_(db){
  if(!db) throw new Error('Falta un ID de base en la hoja Config.');
  var out=[], cursor=null, more=true, g=0;
  while(more && g<40){ g++;
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
    if(estado===C.artista) return;

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

  // 3) Catálogo de producción (base propia) → pestaña Catálogo (las 12 etapas)
  if(C.dbCatalogo){
    var cat=[];
    queryDB_(C.dbCatalogo).forEach(function(p){
      var pr=p.properties||{};
      var cancion=pTit(pr,'Canción'); if(!cancion) return;
      cat.push([
        cancion, pTxt(pr,'Artista'), pTxt(pr,'Productor'), pSel(pr,'Estatus'), pSel(pr,'Fase'),
        pSel(pr,'Grabación'), pSel(pr,'Prod'),
        pTxt(pr,'Mezcla'), pUrl(pr,'Doc Mezcla'), pTxt(pr,'Feedback'),
        pSel(pr,'Master'), pUrl(pr,'Doc Master'),
        pChk(pr,'Versiones Alt')?'sí':'no',
        pChk(pr,'Stems Dropbox')?'sí':'no', pUrl(pr,'Doc Stems'),
        pSel(pr,'Label Copy'), pUrl(pr,'Doc Label Copy'),
        pSel(pr,'Contratos Prod'), pUrl(pr,'Doc Contratos'),
        pTxt(pr,'Split Autoral'), pUrl(pr,'Doc Split'),
        pDate(pr,'Release Date'), pUrl(pr,'Doc Release'),
        pSel(pr,'Pagado'), pUrl(pr,'Doc Pagado'),
        pSel(pr,'Registrado'), pUrl(pr,'Doc Registrado'),
        pChk(pr,'Cover Art')?'sí':'no', pUrl(pr,'Doc Cover'),
        pChk(pr,'Canva')?'sí':'no', pUrl(pr,'Doc Canva'),
        pChk(pr,'Videos')?'sí':'no', pUrl(pr,'Doc Videos'),
        pChk(pr,'Fotos')?'sí':'no', pUrl(pr,'Doc Fotos'),
        pChk(pr,'Redes')?'sí':'no', pUrl(pr,'Doc Redes'),
        pNum(pr,'Fee USD'), pNum(pr,'Saldo USD'),
        pTxt(pr,'Notas'), pUrl(pr,'Links'), pUrl(pr,'Carpeta'), p.id
      ]);
    });
    writeTab_(T_CAT, ['Cancion','Artista','Productor','Estatus','Fase','Grabacion','Prod','Mezcla','DocMezcla','Feedback','Master','DocMaster','VersAlt','Stems','DocStems','LabelCopy','DocLabelCopy','Contratos','DocContratos','Split','DocSplit','Release','DocRelease','Pagado','DocPagado','Registrado','DocRegistrado','CoverArt','DocCover','Canva','DocCanva','Videos','DocVideos','Fotos','DocFotos','Redes','DocRedes','Fee','Saldo','Notas','Link','Carpeta','Id'], cat);
  }

  // 4) Gastos por tema → pestaña Gastos
  if(C.dbGastos){
    var gas=[];
    queryDB_(C.dbGastos).forEach(function(p){
      var pr=p.properties||{};
      gas.push([ pTit(pr,'Concepto'), pTxt(pr,'Canción'), pTxt(pr,'SongId'), pSel(pr,'Fase'), pNum(pr,'Monto USD'), pDate(pr,'Fecha'), pTxt(pr,'Notas'), p.id ]);
    });
    writeTab_(T_GAS, ['Concepto','Cancion','SongId','Fase','Monto','Fecha','Notas','Id'], gas);
  }

  // 5) Créditos / ficha técnica por tema → pestaña Creditos
  if(C.dbCreditos){
    var cre=[];
    queryDB_(C.dbCreditos).forEach(function(p){
      var pr=p.properties||{};
      cre.push([ pTit(pr,'Nombre'), pSel(pr,'Rol'), pChk(pr,'Acreditado')?'sí':'no', pTxt(pr,'Observaciones'), pTxt(pr,'Canción'), pTxt(pr,'SongId'), p.id ]);
    });
    writeTab_(T_CRE, ['Nombre','Rol','Acreditado','Obs','Cancion','SongId','Id'], cre);
  }

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
  t.catJson =JSON.stringify(readTab_(T_CAT));
  t.gasJson =JSON.stringify(readTab_(T_GAS));
  t.credJson=JSON.stringify(readTab_(T_CRE));
  t.resJson =JSON.stringify(readTab_(T_RES));
  t.infoJson=JSON.stringify(readTab_(T_INFO));
  var Csafe={}; for(var k in C){ if(k!=='samplyToken'&&k!=='samplySecret') Csafe[k]=C[k]; }
  t.cfgJson =JSON.stringify(Csafe);
  t.updatedAt=PropertiesService.getScriptProperties().getProperty('LAST_SYNC')||new Date().toISOString();
  var out=t.evaluate();
  out.setTitle(C.artista+' - E33');
  out.addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
function markCobro(pageId, tarea, which, who){
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

/* ---------- ACCIONES: catálogo de producción ---------- */
var CAT_COL={'Grabación':'Grabacion','Prod':'Prod','Master':'Master','Label Copy':'LabelCopy','Contratos Prod':'Contratos','Pagado':'Pagado','Registrado':'Registrado','Estatus':'Estatus','Fase':'Fase','Mezcla':'Mezcla','Feedback':'Feedback','Split Autoral':'Split','Fee USD':'Fee','Saldo USD':'Saldo','Carpeta':'Carpeta','Notas':'Notas','Links':'Link','Versiones Alt':'VersAlt','Stems Dropbox':'Stems','Release Date':'Release','Doc Mezcla':'DocMezcla','Doc Master':'DocMaster','Doc Stems':'DocStems','Doc Label Copy':'DocLabelCopy','Doc Contratos':'DocContratos','Doc Split':'DocSplit','Doc Release':'DocRelease','Doc Pagado':'DocPagado','Doc Registrado':'DocRegistrado','Cover Art':'CoverArt','Doc Cover':'DocCover','Canva':'Canva','Doc Canva':'DocCanva','Videos':'Videos','Doc Videos':'DocVideos','Fotos':'Fotos','Doc Fotos':'DocFotos','Redes':'Redes','Doc Redes':'DocRedes'};
function setCatField(pageId, cancion, prop, kind, value, who){
  var props={}; var num=NaN;
  if(kind==='select'){ props[prop]= value ? {'select':{'name':value}} : {'select':null}; }
  else if(kind==='checkbox'){ props[prop]={'checkbox': (value===true||value==='sí'||value==='true')}; }
  else if(kind==='date'){ props[prop]= value ? {'date':{'start':value}} : {'date':null}; }
  else if(kind==='url'){ props[prop]= {'url': value||null}; }
  else if(kind==='number'){ num=parseFloat(String(value).replace(/[^0-9.\-]/g,'')); props[prop]= isNaN(num) ? {'number':null} : {'number':num}; }
  else { props[prop]={'rich_text':[{'text':{'content':String(value||'').substring(0,1900)}}]}; }
  patch_(pageId,props);
  var col=CAT_COL[prop]||prop;
  var sv = kind==='checkbox' ? ((value===true||value==='sí'||value==='true')?'sí':'no') : (kind==='number' ? (isNaN(num)?'':num) : (value||''));
  setCell_(T_CAT,'Id',pageId,col,sv);
  log_(cancion, prop+'→'+sv, '', who);
  return {ok:true};
}
/* ---------- Carpetas de Drive por canción ---------- */
var SONG_SUBS=['01 Mezcla','02 Master','03 Stems','04 Label Copy','05 Contratos','06 Split','07 Administración','08 Cover Art','09 Canva','10 Videos','11 Fotos','12 Redes','13 Release'];
function folderId_(s){ if(!s) return ''; var m=String(s).match(/[-\w]{25,}/); return m?m[0]:''; }
function driveCatParent_(C){
  var id=folderId_(C.driveCat);
  if(!id){
    try{ var rows=readTab_(T_RES); for(var i=0;i<rows.length;i++){ var s=(rows[i].Seccion||'').toLowerCase(); if(s.indexOf('catálogo')>=0||s.indexOf('catalogo')>=0){ id=folderId_(rows[i].Link); if(id) break; } } }catch(e){}
  }
  if(!id) return null;
  try{ return DriveApp.getFolderById(id); }catch(e){ return null; }
}
function makeSongFolder_(C, cancion){
  var parent=driveCatParent_(C); if(!parent) return '';
  var name=String(cancion||'').trim().replace(/[\\/]/g,'-'); if(!name) return '';
  var f=parent.createFolder(name);
  SONG_SUBS.forEach(function(n){ try{ f.createFolder(n); }catch(e){} });
  return f.getUrl();
}
function createSong(cancion, artista, productor, split, who){
  if(!cancion||!cancion.trim()) throw new Error('Falta el nombre de la canción.');
  var C=cfg_();
  var prod=(productor&&productor.trim())||C.artista;
  var props={'Canción':{'title':[{'text':{'content':cancion.trim().substring(0,1900)}}]},'Estatus':{'select':{'name':'⚫️ Pipeline'}},'Fase':{'select':{'name':'Producción'}},'Productor':{'rich_text':[{'text':{'content':prod}}]}};
  if(artista&&artista.trim()) props['Artista']={'rich_text':[{'text':{'content':artista.trim()}}]};
  if(split&&split.trim()) props['Split Autoral']={'rich_text':[{'text':{'content':split.trim()}}]};
  var res=UrlFetchApp.fetch('https://api.notion.com/v1/pages',{method:'post',contentType:'application/json',headers:H_(),payload:JSON.stringify({parent:{database_id:C.dbCatalogo},properties:props}),muteHttpExceptions:true});
  var j=JSON.parse(res.getContentText()); if(j.object==='error') throw new Error(j.message||'Error al crear');
  var carpeta=''; try{ carpeta=makeSongFolder_(C, cancion.trim()); }catch(e){}
  if(carpeta){ try{ patch_(j.id,{'Carpeta':{'url':carpeta}}); }catch(e){} }
  try{ SpreadsheetApp.getActive().getSheetByName(T_CAT).appendRow([cancion.trim(),artista||'',prod,'⚫️ Pipeline','Producción','','','','','','','','no','no','','','','','',split||'','','','','','','','','no','','no','','no','','no','','no','','','','','',carpeta,j.id]); }catch(e){}
  log_(cancion.trim(),'Canción nueva','',who);
  return {ok:true,row:{Cancion:cancion.trim(),Artista:artista||'',Productor:prod,Estatus:'⚫️ Pipeline',Fase:'Producción',Grabacion:'',Prod:'',Mezcla:'',DocMezcla:'',Feedback:'',Master:'',DocMaster:'',VersAlt:'no',Stems:'no',DocStems:'',LabelCopy:'',DocLabelCopy:'',Contratos:'',DocContratos:'',Split:split||'',DocSplit:'',Release:'',DocRelease:'',Pagado:'',DocPagado:'',Registrado:'',DocRegistrado:'',CoverArt:'no',DocCover:'',Canva:'no',DocCanva:'',Videos:'no',DocVideos:'',Fotos:'no',DocFotos:'',Redes:'no',DocRedes:'',Fee:'',Saldo:'',Notas:'',Link:'',Carpeta:carpeta,Id:j.id}};
}
function createGasto(songId, cancion, concepto, monto, fase, fecha, who){
  var C=cfg_();
  if(!C.dbGastos) throw new Error('Falta la base de gastos en Config.');
  if(!concepto||!concepto.trim()) throw new Error('Falta el concepto del gasto.');
  var m=parseFloat(String(monto).replace(/[^0-9.\-]/g,'')); if(isNaN(m)) m=0;
  var props={ 'Concepto':{'title':[{'text':{'content':concepto.trim().substring(0,1900)}}]}, 'Monto USD':{'number':m}, 'Fase':{'select':{'name':fase||'Otros'}} };
  if(cancion) props['Canción']={'rich_text':[{'text':{'content':String(cancion).substring(0,1900)}}]};
  if(songId)  props['SongId']={'rich_text':[{'text':{'content':String(songId)}}]};
  if(fecha)   props['Fecha']={'date':{'start':fecha}};
  var res=UrlFetchApp.fetch('https://api.notion.com/v1/pages',{method:'post',contentType:'application/json',headers:H_(),payload:JSON.stringify({parent:{database_id:C.dbGastos},properties:props}),muteHttpExceptions:true});
  var j=JSON.parse(res.getContentText()); if(j.object==='error') throw new Error(j.message||'Error al crear gasto');
  try{ SpreadsheetApp.getActive().getSheetByName(T_GAS).appendRow([concepto.trim(), cancion||'', songId||'', fase||'Otros', m, fecha||'', '', j.id]); }catch(e){}
  log_(cancion||'-', 'Gasto: '+concepto.trim()+' $'+m+' ('+(fase||'Otros')+')','',who);
  return {ok:true,row:{Concepto:concepto.trim(),Cancion:cancion||'',SongId:songId||'',Fase:fase||'Otros',Monto:m,Fecha:fecha||'',Notas:'',Id:j.id}};
}
function delGasto(id, who){
  if(!id) return {ok:false};
  try{ UrlFetchApp.fetch('https://api.notion.com/v1/pages/'+id,{method:'patch',contentType:'application/json',headers:H_(),payload:JSON.stringify({archived:true}),muteHttpExceptions:true}); }catch(e){}
  try{ var sh=SpreadsheetApp.getActive().getSheetByName(T_GAS);
    if(sh){ var v=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues(); var idc=v[0].indexOf('Id');
      for(var r=v.length-1;r>=1;r--){ if(String(v[r][idc])===String(id)){ sh.deleteRow(r+1); break; } } } }catch(e){}
  log_('-','Gasto eliminado','',who);
  return {ok:true};
}
function createCredito(songId, cancion, nombre, rol, acreditado, obs, who){
  var C=cfg_();
  if(!C.dbCreditos) throw new Error('Falta la base de créditos en Config.');
  if(!nombre||!nombre.trim()) throw new Error('Falta el nombre del colaborador.');
  var ac=(acreditado===true||acreditado==='sí'||acreditado==='true');
  var props={ 'Nombre':{'title':[{'text':{'content':nombre.trim().substring(0,1900)}}]}, 'Acreditado':{'checkbox':ac} };
  if(rol) props['Rol']={'select':{'name':rol}};
  if(obs) props['Observaciones']={'rich_text':[{'text':{'content':String(obs).substring(0,1900)}}]};
  if(cancion) props['Canción']={'rich_text':[{'text':{'content':String(cancion).substring(0,1900)}}]};
  if(songId)  props['SongId']={'rich_text':[{'text':{'content':String(songId)}}]};
  var res=UrlFetchApp.fetch('https://api.notion.com/v1/pages',{method:'post',contentType:'application/json',headers:H_(),payload:JSON.stringify({parent:{database_id:C.dbCreditos},properties:props}),muteHttpExceptions:true});
  var j=JSON.parse(res.getContentText()); if(j.object==='error') throw new Error(j.message||'Error al crear crédito');
  try{ SpreadsheetApp.getActive().getSheetByName(T_CRE).appendRow([nombre.trim(), rol||'', ac?'sí':'no', obs||'', cancion||'', songId||'', j.id]); }catch(e){}
  log_(cancion||'-', 'Crédito: '+nombre.trim()+' ('+(rol||'—')+', '+(ac?'acreditado':'sin crédito')+')','',who);
  return {ok:true,row:{Nombre:nombre.trim(),Rol:rol||'',Acreditado:ac?'sí':'no',Obs:obs||'',Cancion:cancion||'',SongId:songId||'',Id:j.id}};
}
function toggleCredito(id, acreditado, who){
  if(!id) return {ok:false};
  var ac=(acreditado===true||acreditado==='sí'||acreditado==='true');
  try{ UrlFetchApp.fetch('https://api.notion.com/v1/pages/'+id,{method:'patch',contentType:'application/json',headers:H_(),payload:JSON.stringify({properties:{'Acreditado':{'checkbox':ac}}}),muteHttpExceptions:true}); }catch(e){}
  try{ setCell_(T_CRE,'Id',id,'Acreditado',ac?'sí':'no'); }catch(e){}
  return {ok:true};
}
function delCredito(id, who){
  if(!id) return {ok:false};
  try{ UrlFetchApp.fetch('https://api.notion.com/v1/pages/'+id,{method:'patch',contentType:'application/json',headers:H_(),payload:JSON.stringify({archived:true}),muteHttpExceptions:true}); }catch(e){}
  try{ var sh=SpreadsheetApp.getActive().getSheetByName(T_CRE);
    if(sh){ var v=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues(); var idc=v[0].indexOf('Id');
      for(var r=v.length-1;r>=1;r--){ if(String(v[r][idc])===String(id)){ sh.deleteRow(r+1); break; } } } }catch(e){}
  log_('-','Crédito eliminado','',who);
  return {ok:true};
}
function catSubFolder(parent, nombre){
  var id=folderId_(parent); if(!id) throw new Error('El tema no tiene carpeta.');
  var f; try{ f=DriveApp.getFolderById(id); }catch(e){ throw new Error('No puedo abrir la carpeta del tema.'); }
  var it=f.getFoldersByName(nombre);
  var sub = it.hasNext() ? it.next() : f.createFolder(nombre);
  return {url:sub.getUrl()};
}
function probarCarpeta(){
  var ui=SpreadsheetApp.getUi();
  var C=cfg_();
  var id=folderId_(C.driveCat);
  if(!id){ ui.alert('No pude extraer un ID de:\n"'+C.driveCat+'"'); return; }
  var quien=''; try{ quien=Session.getActiveUser().getEmail(); }catch(e){}
  var f;
  try{ f=DriveApp.getFolderById(id); var nombre=f.getName(); }
  catch(e){ ui.alert('ID extraído: '+id+'\nCuenta del script: '+(quien||'(desconocida)')+'\n\nError al abrir la carpeta:\n'+(e&&e.message?e.message:e)+'\n\nSuele ser que esa carpeta no está compartida con la cuenta del script, o el ID no corresponde a una carpeta.'); return; }
  try{
    var sub=f.createFolder('PRUEBA — borrar '+new Date().toLocaleTimeString());
    ui.alert('✅ Todo bien.\nCarpeta padre: "'+nombre+'"\nCuenta del script: '+(quien||'(desconocida)')+'\n\nSubcarpeta de prueba creada:\n'+sub.getUrl()+'\n\nBorrá esa prueba. Ahora los temas nuevos crearán su carpeta.');
  }catch(e2){ ui.alert('Puedo ver la carpeta "'+nombre+'" pero no crear dentro. Error:\n'+(e2&&e2.message?e2.message:e2)+'\n\nLa cuenta del script ('+(quien||'?')+') necesita permiso de Editor en esa carpeta.'); }
}
function vaciarCatalogo(){
  var C=cfg_(), ui=SpreadsheetApp.getUi();
  if(!C.dbCatalogo){ ui.alert('Falta el ID de la base de catálogo en Config.'); return; }
  var r=ui.alert('Vaciar catálogo', 'Esto ARCHIVA todas las canciones del catálogo en Notion (recuperables 30 días desde la papelera) y limpia la hoja. Vas a recargarlas a mano. ¿Seguimos?', ui.ButtonSet.YES_NO);
  if(r!==ui.Button.YES) return;
  var pages=queryDB_(C.dbCatalogo), n=0;
  pages.forEach(function(p){
    try{ UrlFetchApp.fetch('https://api.notion.com/v1/pages/'+p.id,{method:'patch',contentType:'application/json',headers:H_(),payload:JSON.stringify({archived:true}),muteHttpExceptions:true}); n++; Utilities.sleep(120); }catch(e){}
  });
  try{ writeTab_(T_CAT, ['Cancion','Artista','Productor','Estatus','Fase','Grabacion','Prod','Mezcla','DocMezcla','Feedback','Master','DocMaster','VersAlt','Stems','DocStems','LabelCopy','DocLabelCopy','Contratos','DocContratos','Split','DocSplit','Release','DocRelease','Pagado','DocPagado','Registrado','DocRegistrado','CoverArt','DocCover','Canva','DocCanva','Videos','DocVideos','Fotos','DocFotos','Redes','DocRedes','Fee','Saldo','Notas','Link','Carpeta','Id'], []); }catch(e){}
  log_('-', 'Catálogo vaciado ('+n+' canciones archivadas)','','sistema');
  ui.alert('Listo. Canciones archivadas: '+n+'.\n\nLa hoja del catálogo quedó vacía. Recargá la app y empezá a cargar los temas a mano con el co-manager.');
}
function crearCarpetasFaltantes(){
  var C=cfg_();
  if(!driveCatParent_(C)) throw new Error('No encuentro la carpeta de Drive del catálogo. Pega su ID o link en Config.');
  var rows=readTab_(T_CAT), n=0;
  for(var i=0;i<rows.length;i++){
    var r=rows[i]; if(!r.Cancion || (r.Carpeta&&String(r.Carpeta).trim())) continue;
    var url=''; try{ url=makeSongFolder_(C, r.Cancion); }catch(e){}
    if(url){ try{ if(r.Id) patch_(r.Id,{'Carpeta':{'url':url}}); }catch(e){} setCell_(T_CAT,'Id',r.Id,'Carpeta',url); n++; Utilities.sleep(120); }
  }
  SpreadsheetApp.getUi().alert('Carpetas creadas: '+n);
}

/* ================== SAMPLY (webhook → tema nuevo) ================== */
var SAMPLY_API='https://samply.app/api/v0';
function jsonOut_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function cleanSongName_(s){ s=String(s||'').trim(); s=s.replace(/\.(wav|mp3|aiff?|flac|m4a|ogg|wma|aac|mp4|mov)$/i,''); return s.replace(/\s+/g,' ').trim(); }
function songExists_(name){
  try{ var rows=readTab_(T_CAT), n=cleanSongName_(name).toLowerCase();
    for(var i=0;i<rows.length;i++){ if(cleanSongName_(rows[i].Cancion||'').toLowerCase()===n) return true; } }catch(e){}
  return false;
}
function doPost(e){
  try{
    var C=cfg_();
    var got=(e&&e.parameter&&e.parameter.key)||'';
    if((C.samplySecret||'') && got!==C.samplySecret) return jsonOut_({ok:false,error:'clave inválida'});
    var raw=(e&&e.postData&&e.postData.contents)||'';
    try{ PropertiesService.getScriptProperties().setProperty('SAMPLY_LAST', raw.substring(0,4000)); }catch(_){}
    var body={}; try{ body=JSON.parse(raw); }catch(_){}
    var type=body.type||'';
    var projectid=body.projectid||(body.data&&body.data.after&&body.data.after.projectid)||'';
    if(C.samplyProject && projectid && projectid!==C.samplyProject) return jsonOut_({ok:true,skip:'otro proyecto'});
    var name='';
    if(type==='upload.completed'){ name=body.boxName||(body.data&&body.data.after&&body.data.after.name)||body.name||''; }
    else if(type==='project.created'){ name=body.projectName||(body.data&&body.data.after&&body.data.after.name)||body.name||''; }
    else return jsonOut_({ok:true,ignored:type});
    name=cleanSongName_(name);
    if(!name) return jsonOut_({ok:false,error:'sin nombre'});
    if(songExists_(name)) return jsonOut_({ok:true,dup:name});
    var r=createSong(name, C.artista, '', '', 'Samply');
    var link=body.playerUrl||body.url||'';
    if(link && r&&r.row&&r.row.Id){ try{ setCatField(r.row.Id, name, 'Links', 'url', link, 'Samply'); }catch(_){}
    }
    log_(name,'Tema creado desde Samply ('+type+')','', 'Samply');
    return jsonOut_({ok:true,created:name});
  }catch(err){ return jsonOut_({ok:false,error:String(err)}); }
}
function registrarWebhookSamply(){
  var C=cfg_(), ui=SpreadsheetApp.getUi();
  if(!C.samplyToken){ ui.alert('Falta el "Token de Samply (co-manager)" en Config.'); return; }
  var url=ScriptApp.getService().getUrl();
  if(!url){ ui.alert('Primero publicá el web app (Implementar → Nueva implementación) y volvé a intentar.'); return; }
  var hook=url + (url.indexOf('?')>=0?'&':'?') + 'key=' + encodeURIComponent(C.samplySecret||'reelian33');
  var res=UrlFetchApp.fetch(SAMPLY_API+'/webhooks',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+C.samplyToken},payload:JSON.stringify({label:'E33 Reelian tablero',url:hook,events:['upload.completed','project.created']}),muteHttpExceptions:true});
  ui.alert('Registrar webhook Samply\n\nURL enviada:\n'+hook+'\n\nRespuesta '+res.getResponseCode()+':\n'+res.getContentText().substring(0,1200));
}
function verProyectosSamply(){
  var C=cfg_(), ui=SpreadsheetApp.getUi();
  if(!C.samplyToken){ ui.alert('Falta el token de Samply en Config.'); return; }
  var res=UrlFetchApp.fetch(SAMPLY_API+'/projects',{headers:{Authorization:'Bearer '+C.samplyToken},muteHttpExceptions:true});
  ui.alert('Proyectos de Samply ('+res.getResponseCode()+'):\n\n'+res.getContentText().substring(0,1800));
}
function verUltimoSamply(){
  var last=PropertiesService.getScriptProperties().getProperty('SAMPLY_LAST')||'(sin datos aún)';
  SpreadsheetApp.getUi().alert('Último payload recibido de Samply:\n\n'+last);
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
    sc.setColumnWidth(1,330); sc.setColumnWidth(2,460);
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

  SpreadsheetApp.getUi().alert('Listo. Revisa la hoja "Config" (incluye ahora el ID de la base Catálogo) y luego sincroniza.');
}
function showStatus(){
  var C=cfg_(), p=PropertiesService.getScriptProperties();
  SpreadsheetApp.getUi().alert(
    'Artista: '+C.artista+
    '\nCliente (central): '+(C.cliente||'(vacío)')+
    '\nBase central: '+(C.dbCentral||'(falta)')+
    '\nBase artista: '+(C.dbArtista||'(falta)')+
    '\nBase catálogo: '+(C.dbCatalogo||'(falta)')+
    '\nEquipo: '+C.equipo.map(function(e){return e.nombre;}).join(', ')+
    '\nToken Notion: '+(p.getProperty('NOTION_TOKEN')?'ok':'FALTA')+
    '\nÚltima sync: '+(p.getProperty('LAST_SYNC')||'(nunca)')
  );
}
