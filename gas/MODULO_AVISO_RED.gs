/**
 * ================================================================
 *  MODULO_AVISO_RED.gs  (15-oct-2026)
 *  Correo «Abastecimiento en red» a cada Unidad Ejecutora.
 *
 *  Lee abastecimiento_red.json (lo publica MODULO_ABASTECIMIENTO_RED en la
 *  etapa anterior del ciclo) y envía UN correo por unidad con:
 *    · lo que debe pedir a un CEDIS (Forma 8),
 *    · lo que puede recibir de otras unidades de su red (Forma 7),
 *    · lo que se le propone ceder a otras unidades (Forma 7),
 *    · lo que no tiene fuente en la red,
 *  con un Excel adjunto (una hoja por bloque) y el enlace a su vista del tablero.
 *
 *  Control en CONFIG_ALERTAS (filas AVISO_RED_*, se crean solas, arranca APAGADO):
 *    AVISO_RED_ACTIVO    SI/NO
 *    AVISO_RED_PRUEBA_A  correo: si tiene valor, TODO va solo a ese correo (máx. 3 unidades)
 *    AVISO_RED_DIAS      días de envío: L, M, X, J, V, S, D separados por coma, o L-V, o DIARIO
 *    AVISO_RED_MAX_DEST  tope de destinatarios por envío (protege la cuota de Brevo, 300/día)
 *  Una vez por fecha de saldos. Envío en tramos encadenados (8 unidades por
 *  ejecución) con vigilante, igual que el aviso de reposición por unidad.
 *
 *  REUTILIZA (ya están en el proyecto): enviarCorreoCSS_, _auDirectorio_, _auDirDe_
 *  (MODULO_AVISO_UE) y ABRED_CFG (MODULO_ABASTECIMIENTO_RED).
 *  Todo lo propio usa el prefijo ARV / _arv para no chocar con otros archivos.
 *
 *  Funciones para ejecutar a mano desde el editor:
 *    vistaPreviaAvisoRed()   → le llega a usted el correo de la unidad con más movimientos
 *    enviarAvisoRedManual()  → arranca el envío ya (respeta ACTIVO / PRUEBA_A)
 *    cancelarAvisoRed()      → detiene un envío en curso
 *  Etapa del ciclo (MODULO_INGESTA): iniciarAvisoRedCadena_()
 * ================================================================ */
var ARV = {
  CFG_SH:'CONFIG_ALERTAS',
  URL_TABLERO:'https://torredecontrolcss.github.io/REPOSICION_UE_CSS/abastecimiento_red.html',
  LOGO:'https://raw.githubusercontent.com/TorredeControlCSS/REPOSICION_UE_CSS/main/icons/icon-192.png',
  REMITENTE:'Torre de Control · Dirección Nacional de Logística',
  HANDLER:'arvWorker_', LOTE:8, VENCE_MS:120*60*1000, MAX_PRUEBA:3, TOP:20,
  P:{ULTIMO:'ARV_ULTIMO_ENVIADO', FECHA:'ARV_FECHA', COLA:'ARV_COLA', POS:'ARV_POS', ACTIVO:'ARV_EN_CURSO', INICIO:'ARV_INICIO', ENV:'ARV_ENVIADOS', FALLA:'ARV_FALLAS'},
  FILAS:[
    ['AVISO_RED_ACTIVO','NO','SI/NO — correo «Abastecimiento en red» por Unidad Ejecutora: qué pedir al CEDIS (F8), qué recibir y qué ceder entre unidades (F7), con Excel adjunto'],
    ['AVISO_RED_PRUEBA_A','','Correo de prueba: si tiene valor, TODOS los correos de abastecimiento en red van solo a este correo (máximo 3 unidades). Vaciar para enviar a las unidades.'],
    ['AVISO_RED_DIAS','L','Días de envío del abastecimiento en red: L,M,X,J,V,S,D separados por coma, L-V, o DIARIO'],
    ['AVISO_RED_MAX_DEST','120','Máximo de destinatarios (Para + CC) por envío del abastecimiento en red; si se necesitan más, no se envía y se avisa'],
    ['AVISO_RED_JEFES_TO','','Jefaturas que reciben UN correo de resumen nacional al terminar el envío (unidades, Formas 8 y 7, faltantes y a quién no se pudo enviar)'],
    ['AVISO_RED_JEFES_CC','','Copia del resumen nacional del abastecimiento en red'],
    ['AVISO_RED_CC','','Copia en CADA correo por unidad (ojo: suma a la cuota una vez por unidad; preferir el resumen de jefaturas)']
  ]
};
var ARV_CLASE = ['sobre el objetivo','alto (7–12 meses)','exceso (>12 meses)','sin movimiento (≥90 días)'];
var ARV_CEDIS = {'CEDIS PANAMA':'CEDIS Panamá','CEDIS DIVISA':'CEDIS Divisa','CEDIS CHIRIQUI':'CEDIS Chiriquí'};

/* ---------- utilidades propias ---------- */
function _arvProp_(){ return PropertiesService.getScriptProperties(); }
function _arvFmt_(n){ return (Math.round(n||0)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function _arvFecha_(iso){ var M=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'], p=String(iso||'').slice(0,10).split('-'); return p.length===3?(+p[2])+'-'+M[+p[1]-1]+'-'+p[0]:String(iso||''); }
function _arvEsc_(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function _arvLog_(m){ Logger.log('AVISO RED: '+m); }
function _arvToast_(m,t){ try{ SpreadsheetApp.getActive().toast(m,t||'Abastecimiento en red',-1); }catch(e){} _arvLog_(m); }
/* Avisos internos: a quien ejecuta el proyecto, nunca a las unidades. */
function _arvAvisar_(asunto,cuerpo){ try{ var yo=Session.getEffectiveUser().getEmail(); if(yo) GmailApp.sendEmail(yo,asunto,cuerpo,{name:ARV.REMITENTE}); }catch(e){ _arvLog_('no se pudo avisar: '+e); } }

/* CONFIG_ALERTAS: asegura las filas AVISO_RED_* y las lee. */
function _arvConfig_(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(ARV.CFG_SH);
  if(!sh){ sh=ss.insertSheet(ARV.CFG_SH); sh.getRange(1,1,1,3).setValues([['PARAMETRO','VALOR','DESCRIPCION']]).setFontWeight('bold').setBackground('#0C447C').setFontColor('#FFFFFF'); }
  var v=sh.getDataRange().getValues(), o={}, hay={};
  for(var i=1;i<v.length;i++){ var k=String(v[i][0]||'').trim(); if(k){ hay[k]=true; o[k]=v[i][1]; } }
  var faltan=ARV.FILAS.filter(function(f){ return !hay[f[0]]; });
  if(faltan.length){ sh.getRange(sh.getLastRow()+1,1,faltan.length,3).setValues(faltan); faltan.forEach(function(f){ o[f[0]]=f[1]; }); }
  var prueba=String(o['AVISO_RED_PRUEBA_A']||'').trim(), max=parseInt(o['AVISO_RED_MAX_DEST'],10);
  var lista=function(k){ return String(o[k]||'').split(/[,;\s]+/).map(function(x){ return x.trim().toLowerCase(); }).filter(function(e){ return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); }); };
  return { activo:String(o['AVISO_RED_ACTIVO']||'NO').toUpperCase().indexOf('S')===0,
           prueba:/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(prueba)?prueba:'',
           dias:String(o['AVISO_RED_DIAS']==null?'L':o['AVISO_RED_DIAS']).toUpperCase().trim(),
           maxDest:isFinite(max)&&max>0?max:120,
           jefesTo:lista('AVISO_RED_JEFES_TO'), jefesCc:lista('AVISO_RED_JEFES_CC'), cc:lista('AVISO_RED_CC') };
}
/* ¿Hoy (hora de Panamá) es día de envío? */
function _arvEsDia_(dias){
  var d=+Utilities.formatDate(new Date(),'America/Panama','u');   // 1 = lunes … 7 = domingo
  if(!dias||dias==='DIARIO') return true;
  if(dias==='L-V') return d<=5;
  var L={L:1,M:2,X:3,MI:3,J:4,V:5,S:6,D:7};
  return dias.split(/[,;\s]+/).some(function(x){ return L[x]===d; });
}

/* Lee archivos del repositorio por la API de GitHub (contenido recién publicado, sin caché). */
function _arvLeer_(paths){
  var token=_arvProp_().getProperty('GH_TOKEN'); if(!token) throw new Error('No hay token de GitHub (GH_TOKEN).');
  var rs=UrlFetchApp.fetchAll(paths.map(function(p){ return {url:'https://api.github.com/repos/'+ABRED_CFG.GH_OWNER+'/'+ABRED_CFG.GH_REPO+'/contents/'+p+'?ref=main',
    headers:{Authorization:'Bearer '+token, Accept:'application/vnd.github.raw', 'X-GitHub-Api-Version':'2022-11-28'}, muteHttpExceptions:true}; }));
  var out={};
  rs.forEach(function(r,k){ if(r.getResponseCode()!==200) throw new Error('No se pudo leer '+paths[k]+' (HTTP '+r.getResponseCode()+').'); out[paths[k]]=JSON.parse(r.getContentText()); });
  return out;
}
function _arvDatos_(){
  var d=_arvLeer_([ABRED_CFG.PATH,'renglones_dic.json']);
  return { J:d[ABRED_CFG.PATH], dic:(d['renglones_dic.json']||{}).r||{} };
}

/* Arma lo de una unidad (por índice en J.ues). */
function _arvUE_(J, dic, ui){
  var nom=function(c){ return (dic[c]&&dic[c].n)||c; }, vit=function(c){ return !!(dic[c]&&dic[c].vit); };
  var ueN=function(i){ return J.ues[i]?J.ues[i][0]:String(i); };
  var nec={}; J.n.forEach(function(r){ if(r[0]===ui) nec[r[1]]=r; });
  var f8=[], f7r=[], f7c=[], sin=[];
  J.m.forEach(function(m){
    if(m[0]===ui && m[2]===8) f8.push({cedis:ARV_CEDIS[m[4]]||m[4], otro:m[3]===1, c:m[1], n:nom(m[1]), v:vit(m[1]), q:m[5], nec:nec[m[1]]?nec[m[1]][9]:null});
    else if(m[0]===ui && m[2]===7) f7r.push({ue:ueN(m[4]), c:m[1], n:nom(m[1]), v:vit(m[1]), q:m[5], cl:ARV_CLASE[m[6]]});
    else if(m[2]===7 && m[4]===ui) f7c.push({ue:ueN(m[0]), c:m[1], n:nom(m[1]), v:vit(m[1]), q:m[5], cl:ARV_CLASE[m[6]], dsm:m[7]});
  });
  Object.keys(nec).forEach(function(c){ var r=nec[c]; if(r[12]>0) sin.push({c:c, n:nom(c), v:vit(c), q:r[9], f:r[12], cub:r[10]+r[11]}); });
  var orden=function(a,b){ return (b.v-a.v)||(b.q-a.q); };
  f8.sort(function(a,b){ return a.cedis.localeCompare(b.cedis)||orden(a,b); });
  f7r.sort(function(a,b){ return a.ue.localeCompare(b.ue)||orden(a,b); });
  f7c.sort(function(a,b){ return a.ue.localeCompare(b.ue)||orden(a,b); });
  sin.sort(function(a,b){ return (b.v-a.v)||(b.f-a.f); });
  return { ue:ueN(ui), red:J.ues[ui]?J.ues[ui][1]:'', f8:f8, f7r:f7r, f7c:f7c, sin:sin,
           docsF8:Object.keys(f8.reduce(function(o,x){ o[x.cedis]=1; return o; },{})).length,
           docsF7r:Object.keys(f7r.reduce(function(o,x){ o[x.ue]=1; return o; },{})).length,
           docsF7c:Object.keys(f7c.reduce(function(o,x){ o[x.ue]=1; return o; },{})).length };
}
function _arvTieneAlgo_(A){ return A.f8.length+A.f7r.length+A.f7c.length>0; }

/* ---------- Excel adjunto (una hoja por bloque) ---------- */
function _arvExcel_(A, J){
  var nombre='AbastecimientoRed_'+A.ue.replace(/[^A-Za-z0-9]+/g,'_')+'_'+J.corte;
  var t=SpreadsheetApp.create(nombre), id=t.getId();
  try{
    var hojas=[
      ['Pedir a CEDIS (F8)','Formas 8: pedir al CEDIS',['CEDIS','Tipo','Renglón','Descripción','Vital','Necesidad total','Cantidad a pedir a este CEDIS'],
        A.f8.map(function(x){ return [x.cedis,x.otro?'Otro CEDIS':'CEDIS que le surte',x.c,x.n,x.v?'Sí':'',x.nec,x.q]; }),[150,130,90,380,45,110,130]],
      ['Recibir de UE (F7)','Formas 7: recibir de otras unidades de su red',['Unidad que cede','Renglón','Descripción','Vital','Inventario que se cede','Cantidad a recibir'],
        A.f7r.map(function(x){ return [x.ue,x.c,x.n,x.v?'Sí':'',x.cl,x.q]; }),[280,90,380,45,150,120]],
      ['Ceder a UE (F7)','Formas 7: ceder a otras unidades de su red',['Unidad que recibe','Renglón','Descripción','Vital','Su inventario es','Días sin movimiento','Cantidad a ceder'],
        A.f7c.map(function(x){ return [x.ue,x.c,x.n,x.v?'Sí':'',x.cl,x.dsm==null?'':x.dsm,x.q]; }),[280,90,380,45,150,110,120]],
      ['Sin fuente en la red','Renglones sin fuente suficiente en CEDIS ni en otras unidades',['Renglón','Descripción','Vital','Necesidad','Cubierto por la red','Faltante'],
        A.sin.map(function(x){ return [x.c,x.n,x.v?'Sí':'',x.q,x.cub,x.f]; }),[90,420,45,100,120,100]]
    ];
    var primera=true;
    hojas.forEach(function(h){
      var sh=primera?t.getSheets()[0].setName(h[0]):t.insertSheet(h[0]); primera=false;
      sh.getRange('A1').setValue('CAJA DE SEGURO SOCIAL · DIRECCIÓN NACIONAL DE LOGÍSTICA · ABASTECIMIENTO EN RED');
      sh.getRange(1,1,1,h[2].length).merge().setFontWeight('bold').setFontSize(12).setFontColor('#0C447C');
      sh.getRange('A2').setValue(A.ue+' · '+h[1]+' · saldos al '+_arvFecha_(J.corte)+' · propuesta sujeta a validación de existencias físicas');
      sh.getRange(2,1,1,h[2].length).merge().setFontSize(10).setFontColor('#555555');
      sh.getRange(4,1,1,h[2].length).setValues([h[2]]).setBackground('#1E5A9E').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
      if(h[3].length){
        sh.getRange(5,1,h[3].length,h[2].length).setValues(h[3]);
        var c0=h[0]==='Sin fuente en la red'?4:h[2].length;
        sh.getRange(5,c0,h[3].length,h[2].length-c0+1).setNumberFormat('#,##0');
        if(h[0]==='Pedir a CEDIS (F8)') sh.getRange(5,6,h[3].length,1).setNumberFormat('#,##0');
        sh.getRange(4,1,h[3].length+1,h[2].length).createFilter();
      } else sh.getRange(5,1).setValue('Sin renglones en este bloque.').setFontColor('#5f6b78');
      h[4].forEach(function(w,i){ sh.setColumnWidth(i+1,w); }); sh.setFrozenRows(4);
    });
    SpreadsheetApp.flush();
    return UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/'+id+'/export?format=xlsx',{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}}).getBlob().setName(nombre+'.xlsx');
  } finally { try{ DriveApp.getFileById(id).setTrashed(true); }catch(e){} }
}

/* ---------- cuerpo del correo ---------- */
function _arvHtml_(A, J, banner){
  var td='padding:5px 8px;border-bottom:1px solid #e3e6ea;font:12px Segoe UI,Arial,sans-serif;vertical-align:top;',
      tdn=td+'text-align:right;font-family:Consolas,monospace;',
      th='padding:6px 8px;background:#1E5A9E;color:#fff;font:bold 11px Segoe UI,Arial,sans-serif;text-align:left;',
      thn=th+'text-align:right;';
  var ren=function(x){ return '<b style="color:#1E5A9E;font-family:Consolas,monospace">'+_arvEsc_(x.c)+'</b> '+_arvEsc_(String(x.n).slice(0,60))+(x.v?' <b style="color:#a23224">VITAL</b>':''); };
  var sumQ=function(L){ return L.reduce(function(s,x){ return s+x.q; },0); };
  var bloque=function(titulo, color, fondo, intro, L, cab, fila){
    if(!L.length) return '';
    var h='<div style="margin:14px 0 4px;padding:6px 10px;background:'+fondo+';border-radius:8px;font:bold 13px Segoe UI,Arial,sans-serif;color:'+color+'">'+titulo+'</div>'+
          '<p style="margin:4px 0 6px;font:12px Segoe UI,Arial,sans-serif;color:#444">'+intro+'</p>'+
          '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>'+cab+'</tr>'+L.slice().sort(function(x,y){ return (y.v-x.v)||((y.f||y.q)-(x.f||x.q)); }).slice(0,ARV.TOP).map(fila).join('')+'</table>';
    if(L.length>ARV.TOP) h+='<p style="color:#5f6b78;font-size:11.5px;margin:4px 0">Se muestran los '+ARV.TOP+' principales (vitales y mayores cantidades primero). Los '+L.length+' renglones están en el Excel adjunto.</p>';
    return h;
  };
  var caja=function(v,l,bg){ return '<td style="padding:4px"><div style="background:'+bg+';border-radius:10px;padding:8px 10px;font:12px Segoe UI,Arial,sans-serif;color:#333"><div style="font:bold 18px Consolas,monospace;color:#1d2733">'+v+'</div>'+l+'</div></td>'; };
  var h='<div style="max-width:860px;margin:0 auto;border:1px solid #a8a8a8;background:#fff">'+
    (banner?'<div style="background:#b8860b;color:#fff;padding:6px 14px;font:bold 12px Segoe UI,Arial,sans-serif">'+_arvEsc_(banner)+'</div>':'')+
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#0C447C"><tr><td style="padding:12px 0 12px 16px;width:52px"><img src="'+ARV.LOGO+'" width="44" height="44" alt="CSS"></td>'+
    '<td style="padding:12px 16px;color:#fff;font:14px Segoe UI,Arial,sans-serif"><div style="font-size:17px;font-weight:bold">Abastecimiento en red · '+_arvEsc_(A.ue)+'</div>'+
    '<div style="color:#cfe0f2;font-size:12px">Caja de Seguro Social · Dirección Nacional de Logística · Saldos al '+_arvFecha_(J.corte)+'</div></td></tr></table>'+
    '<div style="padding:12px 16px;font:13px Segoe UI,Arial,sans-serif;color:#333">'+
    '<p style="margin:0 0 8px">Para cubrir los renglones de su unidad que están bajo la línea de pedido se usa primero el inventario de la red: su CEDIS, otro CEDIS que tenga sobre su resguardo y otras unidades de su misma red con exceso o inventario sin movimiento. Esta es la propuesta para su unidad:</p>'+
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>'+
      caja(A.f8.length,'renglones para pedir al CEDIS ('+A.docsF8+' Forma'+(A.docsF8===1?'':'s')+' 8)','#e3edf8')+
      caja(A.f7r.length,'renglones para recibir de '+A.docsF7r+' unidad'+(A.docsF7r===1?'':'es')+' (Forma 7)','#e2f3e7')+
      caja(A.f7c.length,'renglones para ceder a '+A.docsF7c+' unidad'+(A.docsF7c===1?'':'es')+' (Forma 7)','#fde9d6')+
      caja(A.sin.length,'renglones sin fuente suficiente en la red','#fbe4e1')+
    '</tr></table>'+
    bloque('1 · Pedir al CEDIS (Forma 8)','#0C447C','#e3edf8','Una Forma 8 por CEDIS. Total: <b>'+_arvFmt_(sumQ(A.f8))+'</b> unidades. Cuando el CEDIS que le surte no tiene, se propone otro CEDIS que tiene inventario por encima de su resguardo.',
      A.f8,'<th style="'+th+'">CEDIS</th><th style="'+th+'">Renglón</th><th style="'+thn+'">Pedir</th>',
      function(x){ return '<tr><td style="'+td+'white-space:nowrap">'+_arvEsc_(x.cedis)+(x.otro?'<div style="font-size:10.5px;color:#7a5a0b">otro CEDIS</div>':'')+'</td><td style="'+td+'">'+ren(x)+'</td><td style="'+tdn+'font-weight:bold">'+_arvFmt_(x.q)+'</td></tr>'; })+
    bloque('2 · Recibir de otras unidades (Forma 7)','#1f6b36','#e2f3e7','Estas unidades de su red tienen el renglón por encima de lo que necesitan (sobre su objetivo, en exceso o sin movimiento). Coordine con ellas el traslado. Total: <b>'+_arvFmt_(sumQ(A.f7r))+'</b> unidades.',
      A.f7r,'<th style="'+th+'">Unidad que cede</th><th style="'+th+'">Renglón</th><th style="'+thn+'">Recibir</th>',
      function(x){ return '<tr><td style="'+td+'">'+_arvEsc_(x.ue)+'<div style="font-size:10.5px;color:#5f6b78">inventario '+_arvEsc_(x.cl)+'</div></td><td style="'+td+'">'+ren(x)+'</td><td style="'+tdn+'font-weight:bold">'+_arvFmt_(x.q)+'</td></tr>'; })+
    bloque('3 · Ceder a otras unidades (Forma 7)','#8a4b12','#fde9d6','Su unidad tiene estos renglones por encima de su nivel objetivo más 20 días de demanda, o sin movimiento: otras unidades de su red los necesitan. Lo que se propone ceder ya deja cubierta su propia demanda. Total: <b>'+_arvFmt_(sumQ(A.f7c))+'</b> unidades.',
      A.f7c,'<th style="'+th+'">Unidad que recibe</th><th style="'+th+'">Renglón</th><th style="'+thn+'">Ceder</th>',
      function(x){ return '<tr><td style="'+td+'">'+_arvEsc_(x.ue)+'</td><td style="'+td+'">'+ren(x)+'<div style="font-size:10.5px;color:#5f6b78">su inventario: '+_arvEsc_(x.cl)+(x.dsm!=null?' · '+_arvFmt_(x.dsm)+' días sin movimiento':'')+'</div></td><td style="'+tdn+'font-weight:bold">'+_arvFmt_(x.q)+'</td></tr>'; })+
    bloque('4 · Sin fuente suficiente en la red','#a23224','#fbe4e1','Ni los CEDIS ni otras unidades de su red tienen lo suficiente para estos renglones. Se reportan a Planificación.',
      A.sin,'<th style="'+th+'">Renglón</th><th style="'+thn+'">Necesidad</th><th style="'+thn+'">Faltante</th>',
      function(x){ return '<tr><td style="'+td+'">'+ren(x)+'</td><td style="'+tdn+'">'+_arvFmt_(x.q)+'</td><td style="'+tdn+'font-weight:bold">'+_arvFmt_(x.f)+'</td></tr>'; })+
    '<p style="text-align:center;margin:18px 0"><a href="'+ARV.URL_TABLERO+'?ue='+encodeURIComponent(A.ue)+'" style="background:#0C447C;color:#fff;padding:9px 18px;text-decoration:none;font-weight:bold;border-radius:6px">Ver el abastecimiento de mi unidad</a></p>'+
    '<p style="color:#5f6b78;font-size:11px">Es una propuesta calculada con Saldos en línea y Formas 8 de SALMI: valide existencias físicas antes de emitir los formularios. Necesidad = nivel objetivo − saldo − por recibir, para los renglones bajo su línea de pedido (demanda × '+((J.params&&J.params.L)||20)+' días de entrega + inventario crítico). Los insumos médico-quirúrgicos de unidades sin el módulo de inventario MQ en SALMI no se pueden calcular.</p>'+
    '</div></div>';
  return h;
}
function _arvEnviarUno_(A, J, to, cc, banner){
  if(!_arvTieneAlgo_(A)) return false;
  var partes=[]; if(A.f8.length) partes.push(A.f8.length+' a pedir al CEDIS'); if(A.f7r.length) partes.push(A.f7r.length+' a recibir'); if(A.f7c.length) partes.push(A.f7c.length+' a ceder');
  var asunto=(banner?'[PRUEBA] ':'')+'Abastecimiento en red · '+A.ue+' · '+partes.join(' · ')+' · saldos al '+_arvFecha_(J.corte);
  var op={htmlBody:_arvHtml_(A,J,banner), name:ARV.REMITENTE, attachments:[_arvExcel_(A,J)]};
  if(cc&&cc.length) op.cc=cc.join(',');
  enviarCorreoCSS_(to.join(','), asunto, 'Este correo requiere un cliente con soporte HTML. Vea su unidad en: '+ARV.URL_TABLERO+'?ue='+encodeURIComponent(A.ue), op);
  return true;
}

/* ---------- cadena de envío (tramos, con vigilante) ---------- */
function _arvBorrarTriggers_(){ ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()===ARV.HANDLER) ScriptApp.deleteTrigger(t); }); }
function _arvProgramar_(ms){ _arvBorrarTriggers_(); ScriptApp.newTrigger(ARV.HANDLER).timeBased().after(ms||30000).create(); }
function _arvEnCurso_(){ var p=_arvProp_(); if(p.getProperty(ARV.P.ACTIVO)!=='1') return false;
  var ini=+(p.getProperty(ARV.P.INICIO)||0); if(ini&&Date.now()-ini>ARV.VENCE_MS){ p.deleteProperty(ARV.P.ACTIVO); _arvBorrarTriggers_(); return false; } return true; }

/* ETAPA del ciclo: arranca el envío una vez por fecha de saldos, en los días configurados. */
function iniciarAvisoRedCadena_(opts){
  opts=opts||{};
  if(typeof enviarCorreoCSS_!=='function'||typeof _auDirectorio_!=='function') return {ok:false,omitido:'falta MODULO_AVISO_UE en el proyecto'};
  var cfg=_arvConfig_(), p=_arvProp_();
  if(!cfg.activo && !(opts.forzar && cfg.prueba)){ _arvLog_('AVISO_RED_ACTIVO=NO, no se envía.'); return {ok:true,omitido:'inactivo (active AVISO_RED_ACTIVO o ponga un correo en AVISO_RED_PRUEBA_A)'}; }
  if(!opts.forzar && !_arvEsDia_(cfg.dias)){ _arvLog_('hoy no es día de envío ('+cfg.dias+').'); return {ok:true,omitido:'no es día de envío ('+cfg.dias+')'}; }
  if(_arvEnCurso_()) return {ok:false,omitido:'en curso'};
  var D=_arvDatos_(), J=D.J, fecha=J.corte;
  if(!opts.forzar && p.getProperty(ARV.P.ULTIMO)===fecha){ _arvLog_('ya se envió el de los saldos del '+fecha+'.'); return {ok:true,omitido:'ya enviado'}; }
  var dir=cfg.prueba?null:_auDirectorio_(), sinDir=[], cola=[];
  J.ues.forEach(function(u,i){ if(_arvTieneAlgo_(_arvUE_(J,D.dic,i))) cola.push(i); });
  if(cfg.prueba){
    cola.sort(function(a,b){ var A=_arvUE_(J,D.dic,a), B=_arvUE_(J,D.dic,b); return (B.f8.length+B.f7r.length+B.f7c.length)-(A.f8.length+A.f7r.length+A.f7c.length); });
    cola=cola.slice(0,ARV.MAX_PRUEBA);
  } else cola=cola.filter(function(i){ var ok=!!_auDirDe_(dir,J.ues[i][0]); if(!ok) sinDir.push(J.ues[i][0]); return ok; });
  var necesita=cfg.prueba?cola.length+1:cola.reduce(function(s,i){ var d=_auDirDe_(dir,J.ues[i][0]); return s+d.to.length+d.cc.length+cfg.cc.filter(function(e){ return d.to.indexOf(e)<0&&d.cc.indexOf(e)<0; }).length; },0)+cfg.jefesTo.length+cfg.jefesCc.length;
  if(necesita>cfg.maxDest){
    var m='Se necesitan '+necesita+' destinatarios y AVISO_RED_MAX_DEST es '+cfg.maxDest+'. No se envió el abastecimiento en red (saldos del '+fecha+'). Suba el tope solo si la cuota diaria de Brevo (300) alcanza junto con los demás correos del día.';
    _arvAvisar_('[Abastecimiento en red] No se envió: tope de destinatarios',m); _arvLog_(m); return {ok:false,omitido:'tope de destinatarios'};
  }
  var props={}; props[ARV.P.FECHA]=fecha; props[ARV.P.COLA]=JSON.stringify(cola); props[ARV.P.POS]='0'; props[ARV.P.ACTIVO]='1'; props[ARV.P.INICIO]=String(Date.now()); props[ARV.P.ENV]='0';
  props[ARV.P.FALLA]=JSON.stringify(sinDir.map(function(u){ return u+' (sin correo en DIRECTORIO_UE)'; }));
  p.setProperties(props);
  _arvProgramar_(5000);
  _arvLog_('envío en marcha · '+cola.length+' unidades'+(cfg.prueba?' (MODO PRUEBA → '+cfg.prueba+')':'')+' · sin correo: '+sinDir.length);
  return {ok:true, unidades:cola.length, sinDir:sinDir.length, prueba:!!cfg.prueba};
}
function arvWorker_(){
  _arvBorrarTriggers_();
  var p=_arvProp_(); if(p.getProperty(ARV.P.ACTIVO)!=='1') return;
  _arvProgramar_(25*60*1000);   // vigilante: si Google corta esta ejecución, se retoma
  var cfg=_arvConfig_(), fecha=p.getProperty(ARV.P.FECHA), D=_arvDatos_(), J=D.J;
  var cola=JSON.parse(p.getProperty(ARV.P.COLA)||'[]'), fallas=JSON.parse(p.getProperty(ARV.P.FALLA)||'[]');
  if(J.corte!==fecha){   // se publicó un cálculo nuevo a mitad del envío: no se mezclan fechas
    _arvBorrarTriggers_(); p.deleteProperty(ARV.P.ACTIVO);
    _arvAvisar_('[Abastecimiento en red] Envío detenido','Se publicó un cálculo con saldos del '+J.corte+' mientras se enviaba el del '+fecha+'. Se detuvo en la unidad '+(+(p.getProperty(ARV.P.POS)||0))+' de '+cola.length+'.');
    return;
  }
  var dir=cfg.prueba?null:_auDirectorio_();
  var pos=+(p.getProperty(ARV.P.POS)||0), env=+(p.getProperty(ARV.P.ENV)||0), fin=Math.min(cola.length,pos+ARV.LOTE);
  for(;pos<fin;pos++){
    var i=cola[pos], u=J.ues[i]?J.ues[i][0]:String(i);
    try{
      var d=cfg.prueba?{to:[cfg.prueba],cc:[]}:_auDirDe_(dir,u);
      if(d&&!cfg.prueba&&cfg.cc.length) d={to:d.to, cc:d.cc.concat(cfg.cc.filter(function(e){ return d.to.indexOf(e)<0&&d.cc.indexOf(e)<0; }))};
      if(d && _arvEnviarUno_(_arvUE_(J,D.dic,i),J,d.to,d.cc,cfg.prueba?'PRUEBA · en producción iría a los correos de '+u+' en DIRECTORIO_UE':'')) env++;
    }catch(e){ fallas.push(u+': '+e.message); _arvLog_('falló '+u+' → '+e); }
    var pr={}; pr[ARV.P.POS]=String(pos+1); pr[ARV.P.ENV]=String(env); pr[ARV.P.FALLA]=JSON.stringify(fallas); p.setProperties(pr);   // avance guardado por unidad
  }
  if(pos<cola.length){ _arvProgramar_(20000); return; }
  _arvBorrarTriggers_(); p.deleteProperty(ARV.P.ACTIVO);
  if(!cfg.prueba) p.setProperty(ARV.P.ULTIMO,fecha);
  try{ _arvResumenJefes_(cfg,J,D.dic,cola,env,fallas); }catch(e){ fallas.push('Resumen a jefaturas: '+e.message); _arvLog_('falló el resumen a jefaturas → '+e); }
  var m=env+' correos de abastecimiento en red enviados'+(cfg.prueba?' (modo prueba, a '+cfg.prueba+')':'')+' · saldos del '+fecha+(fallas.length?'\n\nNo enviados ('+fallas.length+'):\n- '+fallas.join('\n- '):'');
  _arvAvisar_('[Abastecimiento en red] '+env+' enviados'+(fallas.length?' · '+fallas.length+' sin enviar':''), m);
  _arvLog_('TERMINADO · '+m);
}

/* ---------- para ejecutar a mano ---------- */
function vistaPreviaAvisoRed(){
  try{
    _arvConfig_();   // crea las filas AVISO_RED_* si faltan
    var D=_arvDatos_(), J=D.J, mejor=-1, max=-1;
    J.ues.forEach(function(u,i){ var A=_arvUE_(J,D.dic,i), k=A.f8.length+A.f7r.length+A.f7c.length; if(k>max){ max=k; mejor=i; } });
    var A=_arvUE_(J,D.dic,mejor), yo=Session.getEffectiveUser().getEmail(), d=_auDirDe_(_auDirectorio_(),A.ue);
    _arvEnviarUno_(A,J,[yo],[],'VISTA PREVIA · solo para usted · iría a: '+(d?d.to.concat(d.cc).join(', '):'(unidad sin correo en DIRECTORIO_UE)'));
    _arvToast_('Vista previa de '+A.ue+' enviada a '+yo+'.');
  }catch(e){ _arvToast_('No se pudo: '+e.message,'Abastecimiento en red · ERROR'); }
}
function enviarAvisoRedManual(){
  try{ var r=iniciarAvisoRedCadena_({forzar:true});
    _arvToast_(r.ok&&!r.omitido?'Envío en marcha: '+r.unidades+' unidades'+(r.prueba?' (modo prueba)':'')+'. Al terminar le llega un resumen.':'No se inició: '+r.omitido);
  }catch(e){ _arvToast_('No se pudo: '+e.message,'Abastecimiento en red · ERROR'); }
}
function cancelarAvisoRed(){ var p=_arvProp_(); p.deleteProperty(ARV.P.ACTIVO); _arvBorrarTriggers_(); _arvToast_('Envío del abastecimiento en red cancelado.'); }

/* ---------- resumen nacional para jefaturas (UN correo al terminar) ---------- */
function _arvResumenJefes_(cfg, J, dic, cola, env, fallas){
  var to=cfg.prueba?[cfg.prueba]:cfg.jefesTo; if(!to.length) return;
  var enviadas={}; cola.forEach(function(i){ enviadas[i]=1; });
  var filas=[]; J.ues.forEach(function(u,i){ var A=_arvUE_(J,dic,i); if(!_arvTieneAlgo_(A)) return;
    filas.push({ue:u[0], red:ARV_CEDIS[u[1]]?ARV_CEDIS[u[1]].replace('CEDIS','Red'):u[1], f8:A.f8.length, f7r:A.f7r.length, f7c:A.f7c.length, sin:A.sin.length, ok:!!enviadas[i]&&!fallas.some(function(f){ return String(f).indexOf(u[0]+': ')===0; })}); });
  filas.sort(function(a,b){ return a.red.localeCompare(b.red)||a.ue.localeCompare(b.ue); });
  var k=J.kpi, td='padding:4px 8px;border-bottom:1px solid #e3e6ea;font:12px Segoe UI,Arial,sans-serif;', tdn=td+'text-align:right;font-family:Consolas,monospace;',
      th='padding:6px 8px;background:#1E5A9E;color:#fff;font:bold 11px Segoe UI,Arial,sans-serif;text-align:left;', thn=th+'text-align:right;';
  var caja=function(v,l,bg){ return '<td style="padding:4px"><div style="background:'+bg+';border-radius:10px;padding:8px 10px;font:12px Segoe UI,Arial,sans-serif;color:#333"><div style="font:bold 18px Consolas,monospace;color:#1d2733">'+v+'</div>'+l+'</div></td>'; };
  var noEnv=filas.filter(function(f){ return !f.ok; });
  var h='<div style="max-width:860px;margin:0 auto;border:1px solid #a8a8a8;background:#fff">'+
    (cfg.prueba?'<div style="background:#b8860b;color:#fff;padding:6px 14px;font:bold 12px Segoe UI,Arial,sans-serif">PRUEBA · en producción iría a AVISO_RED_JEFES_TO / AVISO_RED_JEFES_CC</div>':'')+
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#0C447C"><tr><td style="padding:12px 0 12px 16px;width:52px"><img src="'+ARV.LOGO+'" width="44" height="44" alt="CSS"></td>'+
    '<td style="padding:12px 16px;color:#fff;font:14px Segoe UI,Arial,sans-serif"><div style="font-size:17px;font-weight:bold">Abastecimiento en red · resumen nacional</div>'+
    '<div style="color:#cfe0f2;font-size:12px">Caja de Seguro Social · Dirección Nacional de Logística · Saldos al '+_arvFecha_(J.corte)+'</div></td></tr></table>'+
    '<div style="padding:12px 16px;font:13px Segoe UI,Arial,sans-serif;color:#333">'+
    '<p style="margin:0 0 8px">Se envió a <b>'+env+'</b> unidades ejecutoras su propuesta de abastecimiento con el inventario de la red (CEDIS y otras unidades)'+(noEnv.length&&!cfg.prueba?'; <b>'+noEnv.length+'</b> no la recibieron (sin correo en DIRECTORIO_UE o error de envío)':'')+'.</p>'+
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>'+
      caja(Math.round(100*k.cubiertas/k.unidades)+'%','de lo necesario se cubre con la red','#fbf3d6')+
      caja(_arvFmt_(k.formas8),'Formas 8 a CEDIS','#e3edf8')+
      caja(_arvFmt_(k.formas7),'Formas 7 entre unidades','#e2f3e7')+
      caja(_arvFmt_(k.vitalesCompletas)+' / '+_arvFmt_(k.vitales),'líneas vitales cubiertas completas','#fbe4e1')+
    '</tr></table>'+
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:10px"><tr><th style="'+th+'">Unidad ejecutora</th><th style="'+th+'">Red</th><th style="'+thn+'">Pedir a CEDIS</th><th style="'+thn+'">Recibir de UE</th><th style="'+thn+'">Ceder a UE</th><th style="'+thn+'">Sin fuente</th><th style="'+th+'">Correo</th></tr>'+
    filas.map(function(f){ return '<tr><td style="'+td+'">'+_arvEsc_(f.ue)+'</td><td style="'+td+'white-space:nowrap">'+_arvEsc_(f.red)+'</td><td style="'+tdn+'">'+f.f8+'</td><td style="'+tdn+'">'+f.f7r+'</td><td style="'+tdn+'">'+f.f7c+'</td><td style="'+tdn+'">'+f.sin+'</td><td style="'+td+(f.ok?'color:#1f6b36">Enviado':cfg.prueba?'color:#5f6b78">No (prueba)':'color:#a23224;font-weight:bold">No enviado')+'</td></tr>'; }).join('')+'</table>'+
    '<p style="color:#5f6b78;font-size:11px">Cantidades en renglones. El detalle por unidad, las Formas 8 y 7 completas y el Excel nacional están en el tablero.</p>'+
    '<p style="text-align:center;margin:16px 0"><a href="'+ARV.URL_TABLERO+'" style="background:#0C447C;color:#fff;padding:9px 18px;text-decoration:none;font-weight:bold;border-radius:6px">Abrir el tablero de abastecimiento en red</a></p></div></div>';
  var op={htmlBody:h, name:ARV.REMITENTE}; if(!cfg.prueba&&cfg.jefesCc.length) op.cc=cfg.jefesCc.join(',');
  enviarCorreoCSS_(to.join(','), (cfg.prueba?'[PRUEBA] ':'')+'Abastecimiento en red · resumen nacional · '+env+' unidades notificadas · saldos al '+_arvFecha_(J.corte), 'Este correo requiere un cliente con soporte HTML. Tablero: '+ARV.URL_TABLERO, op);
}
