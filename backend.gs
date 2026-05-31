var SHEET_NAME    = "Clientes";
var TORNEOS_INDEX = "Torneos_Index";
var SYNC_SHEET    = "Sync_Activo";
var MASTER_CODE   = "CarlosPN2024";

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); if (headers) s.appendRow(headers); }
  return s;
}

function getSyncSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(SYNC_SHEET);
  if (!s) {
    s = ss.insertSheet(SYNC_SHEET);
    s.getRange(1,1).setValue('{}');
    s.getRange(1,2).setValue('');
    s.getRange(1,3).setValue('');
    s.getRange(1,4).setValue('');
    s.getRange(1,5).setValue('');
    s.getRange(1,6).setValue('');
    s.getRange(1,7).setValue('false');
  }
  return s;
}

function simpleHash_(str) {
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash).toString(36).toUpperCase();
}

function readSession_(s) {
  return {
    keyHash: String(s.getRange(1,4).getValue()||''),
    owner:   String(s.getRange(1,5).getValue()||''),
    start:   String(s.getRange(1,6).getValue()||''),
    locked:  String(s.getRange(1,7).getValue()||'false') === 'true'
  };
}

// ── ROUTER ──────────────────────────────────────────────────────────────────
// Lecturas siempre por GET.
// Escrituras: payloads pequeños por GET (?action=write&op=...&data=...),
//             payloads grandes (saveTournament, syncPush con muchos clientes)
//             por POST con data en el body y action/op en la URL.

function doGet(e) {
  try {
    var action = e.parameter.action || '';

    if (action === 'getClients') {
      var sheet = getOrCreateSheet_(SHEET_NAME, ['nombre','apellido','rut']);
      var data  = sheet.getDataRange().getValues();
      if (data.length < 2) return createJsonResponse({status:'success',data:[]});
      data.shift();
      var clients = data.map(function(r){
        return {nombre:r[0]||'',apellido:r[1]||'',rut:r[2]||''};
      }).filter(function(c){ return c.rut; });
      return createJsonResponse({status:'success',data:clients});
    }

    if (action === 'syncPull') {
      var s  = getSyncSheet_();
      var jv = s.getRange(1,1).getValue();
      var ts = s.getRange(1,2).getValue();
      var la = s.getRange(1,3).getValue();
      var sess = readSession_(s);
      if (!jv || jv === '{}') {
        return createJsonResponse({status:'success',data:{empty:true,timestamp:'',session:sess}});
      }
      var resp = {state:JSON.parse(jv), timestamp:String(ts), session:sess};
      if (la) { try { resp.lastAction = JSON.parse(la); } catch(x){} }
      return createJsonResponse({status:'success',data:resp});
    }

    if (action === 'getSession') {
      var s2 = getSyncSheet_();
      return createJsonResponse({status:'success',data:readSession_(s2)});
    }

    if (action === 'listTournaments') {
      var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TORNEOS_INDEX);
      if (!idx || idx.getLastRow() < 2) return createJsonResponse({status:'success',data:[]});
      var rows = idx.getDataRange().getValues(); rows.shift();
      var list = rows.map(function(r){
        return {sheetName:r[0]||'',tournamentType:r[1]||'',date:r[2]||'',
                director:r[3]||'',savedAt:r[4]||'',players:r[5]||0};
      }).filter(function(r){ return r.sheetName; });
      return createJsonResponse({status:'success',data:list});
    }

    if (action === 'getTournament') {
      var sn = e.parameter.sheetName;
      var ts2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sn);
      if (!ts2) return createJsonResponse({status:'error',message:'No encontrada: '+sn});
      return createJsonResponse({status:'success',data:JSON.parse(ts2.getRange(1,1).getValue())});
    }

    if (action === 'write') {
      var op      = e.parameter.op || '';
      var dataStr = e.parameter.data || '';
      if (!op)      return createJsonResponse({status:'error',message:'Sin operación.'});
      if (!dataStr) return createJsonResponse({status:'error',message:'Sin datos para: '+op});
      var data = JSON.parse(dataStr);
      return processWrite_(op, data);
    }

    return createJsonResponse({status:'error',message:'Acción no válida: '+action});
  } catch(err) {
    return createJsonResponse({status:'error',message:'doGet: '+err.message});
  }
}

// doPost — maneja payloads grandes enviados con data en el body
// y action/op en los parámetros de la URL.
function doPost(e) {
  try {
    var action  = e.parameter.action || '';
    var dataStr = e.parameter.data   || '';
    var op      = e.parameter.op     || action;

    // Leer del body ANTES de cualquier return
    // (el data viene en e.postData.contents, no en e.parameter.data)
    if (!dataStr && e.postData && e.postData.contents) {
      var pairs = e.postData.contents.split('&');
      pairs.forEach(function(pair) {
        var kv = pair.split('=');
        if (kv.length >= 2) {
          var k = decodeURIComponent(kv[0]);
          var v = decodeURIComponent(kv.slice(1).join('='));
          if (k === 'action') action = v;
          if (k === 'op')     op = v;
          if (k === 'data')   dataStr = v;
        }
      });
    }

    if (!op || !dataStr) return createJsonResponse({status:'error', message:'Sin op/data en POST.'});
    return processWrite_(op, JSON.parse(dataStr));
  } catch(err) {
    return createJsonResponse({status:'error', message:'doPost: '+err.message});
  }
}

function processWrite_(op, data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(SHEET_NAME, ['nombre','apellido','rut']);
  var RUT_COL = 2;

  // ── syncPush ─────────────────────────────────────────────────────────────
  if (op === 'syncPush') {
    var syncS = getSyncSheet_();
    var sess  = readSession_(syncS);
    if (sess.locked) {
      var inH = simpleHash_(data.sessionKey||'');
      var mH  = simpleHash_(MASTER_CODE);
      if (inH !== sess.keyHash && inH !== mH) {
        return createJsonResponse({status:'error',
          message:'Torneo bloqueado. Solo el Supervisor activo puede publicar.'});
      }
    }
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    syncS.getRange(1,1).setValue(JSON.stringify(data));
    syncS.getRange(1,2).setValue(ts);
    syncS.getRange(1,3).setValue(data.lastAction ? JSON.stringify(data.lastAction) : '');
    return createJsonResponse({status:'success',data:{timestamp:ts}});
  }

  // ── openSession ──────────────────────────────────────────────────────────
  if (op === 'openSession') {
    var syncS2 = getSyncSheet_();
    var sess2  = readSession_(syncS2);
    if (sess2.keyHash) {
      var inH2 = simpleHash_(data.key||'');
      var mH2  = simpleHash_(MASTER_CODE);
      if (inH2 !== sess2.keyHash && inH2 !== mH2) {
        return createJsonResponse({status:'error',
          message:'Sesión activa de "'+sess2.owner+'". Usa esa clave o el código maestro.'});
      }
    } else {
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      syncS2.getRange(1,4).setValue(simpleHash_(data.key));
      syncS2.getRange(1,5).setValue(data.owner||'Supervisor');
      syncS2.getRange(1,6).setValue(now);
      syncS2.getRange(1,7).setValue('false');
    }
    return createJsonResponse({status:'success',data:{
      owner:  syncS2.getRange(1,5).getValue(),
      start:  syncS2.getRange(1,6).getValue(),
      locked: String(syncS2.getRange(1,7).getValue()) === 'true'
    }});
  }

  // ── closeSession ─────────────────────────────────────────────────────────
  if (op === 'closeSession') {
    var syncS3 = getSyncSheet_();
    var sess3  = readSession_(syncS3);
    if (!sess3.keyHash) return createJsonResponse({status:'success',data:{message:'Sin sesión activa.'}});
    var inH3 = simpleHash_(data.key||'');
    var mH3  = simpleHash_(MASTER_CODE);
    if (inH3 !== sess3.keyHash && inH3 !== mH3) {
      return createJsonResponse({status:'error',message:'Clave incorrecta.'});
    }
    syncS3.getRange(1,4).setValue('');
    syncS3.getRange(1,5).setValue('');
    syncS3.getRange(1,6).setValue('');
    syncS3.getRange(1,7).setValue('false');
    return createJsonResponse({status:'success',data:{message:'Sesión cerrada.'}});
  }

  // ── lockTournament ───────────────────────────────────────────────────────
  if (op === 'lockTournament') {
    var syncS4 = getSyncSheet_();
    var sess4  = readSession_(syncS4);
    var inH4 = simpleHash_(data.key||'');
    var mH4  = simpleHash_(MASTER_CODE);
    if (!sess4.keyHash || (inH4 !== sess4.keyHash && inH4 !== mH4)) {
      return createJsonResponse({status:'error',message:'Clave incorrecta.'});
    }
    syncS4.getRange(1,7).setValue(data.locked ? 'true' : 'false');
    return createJsonResponse({status:'success',data:{
      locked:  data.locked,
      message: data.locked ? 'Torneo bloqueado.' : 'Torneo desbloqueado.'
    }});
  }

  // ── saveTournament ───────────────────────────────────────────────────────
  if (op === 'saveTournament') {
    var tType  = data.tournamentType||'torneo';
    var tDate  = data.date||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
    var saved  = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
    var players= (data.clients||[]).filter(function(c){return !c.isTableBreak;}).length;
    var sName  = ('T_'+tDate.replace(/-/g,'')+'_'+tType).substring(0,90);
    if (ss.getSheetByName(sName)) {
      var sfx=2; while(ss.getSheetByName(sName+'_'+sfx))sfx++; sName=sName+'_'+sfx;
    }
    var newS=ss.insertSheet(sName);
    newS.getRange(1,1).setValue(JSON.stringify(data));
    newS.getRange(1,1).setWrap(false);
    newS.hideSheet();
    var idx2=ss.getSheetByName(TORNEOS_INDEX);
    if (!idx2) {
      idx2=ss.insertSheet(TORNEOS_INDEX);
      idx2.appendRow(['Pestaña','Tipo','Fecha Torneo','Director','Guardado','Jugadores']);
      idx2.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1a1b22').setFontColor('#00daf3');
    }
    idx2.appendRow([sName,tType,tDate,data.director||'',saved,players]);
    return createJsonResponse({status:'success',data:{message:'Guardado.',sheetName:sName}});
  }

  // ── addClient ────────────────────────────────────────────────────────────
  if (op === 'addClient') {
    var rows=sheet.getDataRange().getValues();
    var rn=String(data.rut).trim().toLowerCase();
    var nn=String(data.nombre).trim().toLowerCase();
    var an=String(data.apellido).trim().toLowerCase();
    var dup=rows.slice(1).some(function(r){
      return String(r[2]).trim().toLowerCase()===rn||
        (String(r[0]).trim().toLowerCase()===nn&&String(r[1]).trim().toLowerCase()===an);
    });
    if(dup) return createJsonResponse({status:'error',message:'Cliente duplicado.'});
    sheet.appendRow([data.nombre,data.apellido,data.rut]);
    return createJsonResponse({status:'success',data:{message:'Cliente añadido'}});
  }

  if (op === 'updateClient') {
    var ad=sheet.getDataRange().getValues(); var ri=-1;
    for(var i=0;i<ad.length;i++){if(ad[i][RUT_COL]==data.rut){ri=i+1;break;}}
    if(ri>0){sheet.getRange(ri,1).setValue(data.nombre);sheet.getRange(ri,2).setValue(data.apellido);
      return createJsonResponse({status:'success',data:{message:'Actualizado'}});}
    return createJsonResponse({status:'error',message:'No encontrado: '+data.rut});
  }

  if (op === 'deleteClient') {
    var dv=sheet.getDataRange().getValues();
    for(var j=dv.length-1;j>0;j--){
      if(dv[j][RUT_COL]==data.rut){sheet.deleteRow(j+1);
        return createJsonResponse({status:'success',data:{message:'Eliminado'}});}
    }
    return createJsonResponse({status:'error',message:'No encontrado: '+data.rut});
  }

  if (op === 'importClients') {
    var nc=data; var ed=sheet.getDataRange().getValues();
    var er=new Set(ed.slice(1).map(function(r){return r[RUT_COL];})); var ta=[];
    nc.forEach(function(c){
      if(c&&c.rut&&c.nombre&&c.apellido&&!er.has(c.rut)){ta.push([c.nombre,c.apellido,c.rut]);er.add(c.rut);}
    });
    if(ta.length>0) sheet.getRange(sheet.getLastRow()+1,1,ta.length,3).setValues(ta);
    return createJsonResponse({status:'success',data:{message:ta.length+' importados.'}});
  }

  if (op === 'removeDuplicates') {
    var raw=sheet.getDataRange().getValues(); var hdr=raw[0]; var dr=raw.slice(1);
    var sr=new Set(); var sn=new Set(); var uniq=[]; var rem=0;
    dr.forEach(function(r){
      var rut=String(r[2]).trim().toLowerCase();
      var key=String(r[0]).trim().toLowerCase()+'|'+String(r[1]).trim().toLowerCase();
      if(!rut)return;
      if(sr.has(rut)||sn.has(key)){rem++;}else{sr.add(rut);sn.add(key);uniq.push(r);}
    });
    sheet.clearContents(); sheet.appendRow(hdr);
    if(uniq.length>0) sheet.getRange(2,1,uniq.length,3).setValues(uniq);
    return createJsonResponse({status:'success',data:{message:'Eliminados '+rem+' duplicados.'}});
  }

  return createJsonResponse({status:'error',message:'Operación no válida: '+op});
}
