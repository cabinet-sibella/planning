/* Colle entre la connexion Microsoft (store.js) et le planning (planning.js). */
(function(){
  var $ = function(id){ return document.getElementById(id); };
  var refs = {genere: '', rows: []}, rows = [], toastT = null;

  function norm(s){ return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
  function toast(s){ var t = $('toast'); t.textContent = s; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(function(){ t.hidden = true; }, 4000); }

  // Même règle que le registre : une affaire est « à programmer » si son avancement le dit, sauf avis contraire (champ po).
  var reTodo = /\b(terrain|drone|releve|releves|leve|vol|vols|intervention|bornage|implantation|reperage|debroussaillage|demaquisage)\b[^.;]{0,30}\ba (faire|refaire|programmer|planifier|organiser|prevoir|reprendre)\b|\bprogrammation\b[^.;]{0,20}\ba (faire|prevoir)\b|\ba (programmer|programme|programmee|planifier)\b/;
  var reBesoin = /\b(necessite|releve|topo|drone|gps|navis|rtc|duster|piquetage|epure|orthophoto|lidar|terrain)\b/;
  var reExclu = /\d{1,2}\s*\/\s*\d{1,2}|\b(lundi|mardi|mercredi|jeudi|vendredi|livre|transmis|envoye|realise|realises|fait|effectue|valide|signe|dossier|monter|pas de)\b/;
  function aFaire(r){
    if (r.livre) return false;
    var t = norm(r.avancement);
    if (reTodo.test(t)) return true;
    return Number(r.ref) >= 15590 && reBesoin.test(t) && !reExclu.test(t);
  }

  function indexer(r){
    refs = r || refs;
    rows = refs.rows.map(function(x){ var lib = x.lib || String(x.ref); return Object.assign({}, x, {lib: lib, _s: norm([lib, x.client, x.commune, x.mission].join(' '))}); });
    if (window.SIBELLA_rows) window.SIBELLA_rows();
  }
  window.SIBELLA_refs = indexer;
  window.SIBELLA = {
    rows: function(){ return rows; },
    find: function(id){ id = String(id); for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i]; for (var j = 0; j < rows.length; j++) if (String(rows[j].ref) === id) return rows[j]; return null; },
    isProg: function(r){ if (typeof r.prog === 'boolean') return r.prog; return r.po === true || r.po === false ? r.po : aFaire(r); },
    db: function(){ return window.Store.db; },
    canWrite: function(){ return window.Store.canWrite(); },
    downloads: function(){ return null; },
    equipe: function(){ return []; },
    genere: function(){ return refs.genere; },
    toast: toast
  };

  function etatDroits(){
    var S = window.Store;
    $('qui').textContent = S.nom();
    $('droits').textContent = S.canWrite() ? 'Modification autorisée' : 'Consultation seule';
    $('btnImport').hidden = !S.canWrite();
  }
  function message(t){ var e = $('plMsgGlobal'); e.textContent = t || ''; e.hidden = !t; }

  async function demarrer(){
    var S = window.Store;
    $('loginBox').hidden = true; $('btnOut').hidden = false; etatDroits();
    S.onStatut(function(t){ etatDroits(); message(t); });
    try { await S.demarrer(); }
    catch (e){
      message('Impossible de lire le planning dans SharePoint : ' + (e && e.message ? e.message : e));
      return;
    }
    $('planningView').hidden = false;
    window.PLANNING_START();
    indexer(S.refs());
    if (S.refsRefusees() && S.canWrite()) message('Vous n\'avez pas accès au dossier « Planning - Références » : la liste « À programmer » et la recherche de références ne fonctionneront pas.');
    if (S.vide()) message(S.canWrite() ? 'Aucune donnée dans SharePoint pour l\'instant : utilisez « Importer des données » avec le fichier initial.' : 'Le planning n\'est pas encore initialisé. Revenez un peu plus tard.');
  }

  function brancher(){
    var S = window.Store;
    $('btnLogin').onclick = function(){ S.login(); };
    $('btnOut').onclick = function(){ S.logout(); };
    $('btnImport').onclick = function(){ $('fichierImport').click(); };
    $('fichierImport').onchange = function(){
      var f = this.files[0]; this.value = ''; if (!f) return;
      var lec = new FileReader();
      lec.onload = function(){
        var o; try { o = JSON.parse(lec.result); } catch (e){ toast('Ce fichier n\'est pas un JSON valide.'); return; }
        if (!o || (!o.planning && !o.references)){ toast('Fichier non reconnu : il doit contenir « planning » et/ou « references ».'); return; }
        if (o.planning && !S.vide() && !window.confirm('Des données existent déjà dans SharePoint. Les remplacer par celles de ce fichier ? (toutes les tâches saisies seront perdues)')) return;
        S.importer(o).then(function(){ toast('Import terminé.'); message(''); }, function(e){ toast('Import impossible : ' + (e && e.message ? e.message : e)); });
      };
      lec.readAsText(f);
    };
  }

  async function lancer(){
    brancher();
    if (!window.msal){ $('loginBox').hidden = false; message('La bibliothèque de connexion Microsoft n\'a pas pu être chargée.'); return; }
    var a;
    try { a = await window.Store.init(); }
    catch (e){ $('loginBox').hidden = false; message('Connexion Microsoft impossible : ' + (e && e.message ? e.message : e)); return; }
    if (a) demarrer(); else $('loginBox').hidden = false;
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', lancer); else lancer();
})();
