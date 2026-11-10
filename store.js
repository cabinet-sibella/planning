/* Connexion Microsoft (MSAL) et données dans SharePoint (Microsoft Graph).
   Les données tiennent dans un fichier JSON du dossier « Planning » de l'équipe. L'écriture est conditionnelle (If-Match) :
   si deux personnes enregistrent en même temps, la seconde relit la version de la première et rejoue sa modification. */
(function(){
  var C = window.PLANNING_CONFIG;
  var BASE = 'https://graph.microsoft.com/v1.0';
  var SCOPES = ['User.Read', 'Files.ReadWrite.All'];
  var COLLS = ['planning_taches', 'planning_ressources', 'planning_activites', 'planning_config'];
  var msal = null, local = vide(), etagConnu = null, absent = false;
  var refs = {genere: '', rows: []}, etagRefs = null;
  var pending = [], flushing = null, subs = [], droitsRefuses = false, timer = null, tick = 0;
  var statutCb = function(){};

  function vide(){ var d = {v: 1}; COLLS.forEach(function(c){ d[c] = {}; }); return d; }
  function normaliser(d){ var o = vide(); if (d && typeof d === 'object') COLLS.forEach(function(c){ if (d[c] && typeof d[c] === 'object') o[c] = d[c]; }); return o; }
  function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
  function erreur(msg, status){ var e = new Error(msg); e.status = status; return e; }

  // ---------- authentification
  function compte(){ return msal ? (msal.getActiveAccount() || msal.getAllAccounts()[0] || null) : null; }
  async function init(){
    msal = new window.msal.PublicClientApplication({
      auth: {clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: C.redirectUri, navigateToLoginRequestUrl: false},
      cache: {cacheLocation: 'localStorage'}
    });
    await msal.initialize();
    var r = await msal.handleRedirectPromise();
    if (r && r.account) msal.setActiveAccount(r.account);
    var a = compte(); if (a) msal.setActiveAccount(a);
    return a;
  }
  function login(){ return msal.loginRedirect({scopes: SCOPES, prompt: 'select_account'}); }
  function logout(){ var a = compte(); return msal.logoutRedirect({account: a, postLogoutRedirectUri: C.redirectUri}); }
  async function jeton(){
    var a = compte(); if (!a) throw erreur('Non connecté', 401);
    try { return (await msal.acquireTokenSilent({scopes: SCOPES, account: a})).accessToken; }
    catch (e){ await msal.acquireTokenRedirect({scopes: SCOPES, account: a}); throw erreur('Reconnexion nécessaire', 401); }
  }
  function email(){ var a = compte(); return a ? String(a.username || '').toLowerCase() : ''; }
  function nom(){ var a = compte(); return a ? (a.name || a.username) : ''; }
  function canWrite(){ return !droitsRefuses && C.editeurs.map(function(x){ return x.toLowerCase(); }).indexOf(email()) >= 0; }

  // ---------- Graph
  async function graph(chemin, opt){
    opt = opt || {}; var h = Object.assign({Authorization: 'Bearer ' + await jeton()}, opt.headers || {});
    var r = await fetch(BASE + chemin, {method: opt.method || 'GET', headers: h, body: opt.body});
    if (r.status === 403 && (opt.method === 'PUT' || opt.method === 'POST')){ droitsRefuses = true; statutCb(); }
    if (!r.ok){ var t = ''; try { t = await r.text(); } catch (e){} throw erreur('Graph ' + r.status + ' ' + t.slice(0, 200), r.status); }
    return r;
  }
  function url(nomFichier, dos){ return '/groups/' + C.groupId + '/drive/root:/' + encodeURIComponent(dos || C.dossier) + '/' + encodeURIComponent(nomFichier); }
  async function meta(nomFichier, dos){
    try { return await (await graph(url(nomFichier, dos) + '?$select=id,eTag,size')).json(); }
    catch (e){ if (e.status === 404) return null; throw e; }
  }
  async function contenu(nomFichier, dos){ return JSON.parse(await (await graph(url(nomFichier, dos) + ':/content')).text()); }
  async function dossier(dos){
    try { await graph('/groups/' + C.groupId + '/drive/root/children', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({name: dos || C.dossier, folder: {}, '@microsoft.graph.conflictBehavior': 'fail'})}); }
    catch (e){ if (e.status !== 409) throw e; }
  }
  async function ecrire(nomFichier, obj, etag, dos){
    var corps = JSON.stringify(obj);
    if (corps.length > 3800000) throw erreur('Fichier trop volumineux : archivez d\'anciennes tâches.', 413);
    var h = {'Content-Type': 'application/json'}, q = '';
    if (etag) h['If-Match'] = etag; else q = '?@microsoft.graph.conflictBehavior=fail';
    try { return await (await graph(url(nomFichier, dos) + ':/content' + q, {method: 'PUT', headers: h, body: corps})).json(); }
    catch (e){
      if (e.status === 404 && !etag){ await dossier(dos); return await (await graph(url(nomFichier, dos) + ':/content' + q, {method: 'PUT', headers: h, body: corps})).json(); }
      throw e;
    }
  }

  // ---------- base locale et abonnements
  function appliquer(d, c, id, v){ if (v == null) delete d[c][id]; else d[c][id] = v; }
  function comparer(a, op, b){ if (op === '==') return a === b; if (op === '>=') return a != null && a >= b; if (op === '<=') return a != null && a <= b; return false; }
  function instantane(c, filtres, lim){
    var tout = local[c] || {}, docs = [];
    Object.keys(tout).forEach(function(id){ if (filtres.every(function(f){ return comparer(tout[id][f[0]], f[1], f[2]); })) docs.push({id: id, data: function(){ return tout[id]; }}); });
    docs = docs.slice(0, lim || 100); return {docs: docs, size: docs.length};
  }
  var notifTimer = null;
  function notifier(){ if (notifTimer) return; notifTimer = setTimeout(function(){ notifTimer = null; subs.slice().forEach(function(s){ try { s.cb(); } catch (e){ console.error(e); } }); }, 0); }
  function abonner(fn){ var s = {cb: fn}; subs.push(s); setTimeout(function(){ try { fn(); } catch (e){ console.error(e); } }, 0); return function(){ subs = subs.filter(function(x){ return x !== s; }); }; }
  function requete(c){
    var filtres = [], lim = 100;
    var q = {
      where: function(f, op, v){ filtres.push([f, op, v]); return q; },
      limit: function(n){ lim = n; return q; },
      get: function(){ return Promise.resolve(instantane(c, filtres, lim)); },
      onSnapshot: function(ok){ return abonner(function(){ ok(instantane(c, filtres, lim)); }); }
    };
    return q;
  }
  function doc(chemin){
    var p = chemin.split('/'), c = p[0], id = p.slice(1).join('/');
    return {
      set: function(v){ return muter(c, id, JSON.parse(JSON.stringify(v))); },
      delete: function(){ return muter(c, id, null); },
      onSnapshot: function(ok){ return abonner(function(){ var x = (local[c] || {})[id]; ok({exists: !!x, data: function(){ return x; }}); }); }
    };
  }
  var db = {collection: requete, doc: doc};

  // ---------- écriture groupée avec reprise en cas de conflit
  function muter(c, id, v){
    if (!canWrite()) return Promise.reject(erreur('Lecture seule', 403));
    appliquer(local, c, id, v); notifier();
    return new Promise(function(ok, ko){ pending.push({c: c, id: id, v: v, ok: ok, ko: ko}); planifier(); });
  }
  function planifier(){ if (flushing) return; flushing = vider().then(function(){ flushing = null; if (pending.length) planifier(); }, function(){ flushing = null; }); }
  async function vider(){
    await sleep(150);
    var lot = pending.splice(0), derniere = null;
    for (var i = 0; i < 5; i++){
      try {
        var m = await meta(C.fichierDonnees), base = vide(), etag = null;
        if (m){ base = normaliser(await contenu(C.fichierDonnees)); etag = m.eTag; }
        lot.forEach(function(o){ appliquer(base, o.c, o.id, o.v); });
        var r = await ecrire(C.fichierDonnees, base, etag);
        etagConnu = r.eTag; absent = false; local = base; pending.forEach(function(o){ appliquer(local, o.c, o.id, o.v); }); notifier();
        lot.forEach(function(o){ o.ok(); }); return;
      } catch (e){
        derniere = e;
        if (e.status === 412 || e.status === 409){ await sleep(300 * (i + 1)); continue; }
        break;
      }
    }
    lot.forEach(function(o){ o.ko(derniere || erreur('Conflit', 412)); });
    try { await rafraichir(true); } catch (e){}
  }

  // ---------- lecture
  async function rafraichir(force){
    var m = await meta(C.fichierDonnees);
    if (!m){ absent = true; if (force){ local = vide(); etagConnu = null; notifier(); } return; }
    absent = false;
    if (force || m.eTag !== etagConnu){ local = normaliser(await contenu(C.fichierDonnees)); etagConnu = m.eTag; pending.forEach(function(o){ appliquer(local, o.c, o.id, o.v); }); notifier(); }
  }
  var refsRefusees = false;
  async function rafraichirRefs(){
    var m;
    try { m = await meta(C.fichierRefs, C.dossierRefs); }
    catch (e){ if (e.status === 403 || e.status === 401){ refsRefusees = true; return; } throw e; }
    if (!m){ return; }
    if (m.eTag !== etagRefs){ var d = await contenu(C.fichierRefs, C.dossierRefs); refs = {genere: d.genere || '', rows: Array.isArray(d.rows) ? d.rows : []}; etagRefs = m.eTag; if (window.SIBELLA_refs) window.SIBELLA_refs(refs); }
  }
  async function demarrer(){
    await rafraichir(true); await rafraichirRefs();
    clearInterval(timer);
    timer = setInterval(async function(){
      if (document.hidden || flushing || pending.length) return;
      tick++;
      try { await rafraichir(false); if (tick % 10 === 0) await rafraichirRefs(); statutCb(''); } catch (e){ statutCb('Synchronisation interrompue : ' + (e && e.message ? e.message : e)); }
    }, Math.max(10, C.sondage || 30) * 1000);
    document.addEventListener('visibilitychange', function(){ if (!document.hidden) rafraichir(false).catch(function(){}); });
  }

  // ---------- import initial (réservé aux éditeurs)
  async function importer(obj){
    if (!canWrite()) throw erreur('Lecture seule', 403);
    if (obj.planning){
      var m = await meta(C.fichierDonnees), d = normaliser(obj.planning);
      var r = await ecrire(C.fichierDonnees, d, m ? m.eTag : null); etagConnu = r.eTag; local = d; notifier();
    }
    if (obj.references){
      var m2 = await meta(C.fichierRefs, C.dossierRefs);
      await ecrire(C.fichierRefs, obj.references, m2 ? m2.eTag : null, C.dossierRefs); etagRefs = null; await rafraichirRefs();
    }
  }

  window.Store = {
    init: init, login: login, logout: logout, nom: nom, email: email, canWrite: canWrite,
    db: db, demarrer: demarrer, importer: importer, refs: function(){ return refs; },
    donneesAbsentes: function(){ return absent; }, refsRefusees: function(){ return refsRefusees; }, vide: function(){ return !Object.keys(local.planning_ressources).length; },
    onStatut: function(fn){ statutCb = fn; }
  };
})();
