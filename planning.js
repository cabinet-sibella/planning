/* Planning du Cabinet Sibella : ressources (personnel, véhicules, matériel) en lignes, demi-journées en colonnes.
   Version autonome : les données viennent de SharePoint via store.js (même interface db que la version intégrée au registre). */
(function(){
  var $ = function(id){ return document.getElementById(id); };
  var S = null, db = null, root = null;
  var taches = {}, res = {}, act = {}, unsub = null, fen = '';
  var sites = [], sitesCoul = {};
  // couleurs des sites, par nom ; PAL sert aux sites ajoutés plus tard
  var COUL = {'Ajaccio': '#f7c794', 'Bastia': '#b3d3f2', "L'Île-Rousse": '#b7e0ba', 'Porto-Vecchio': '#d6c2ee'};
  var PAL = ['#a9dcd6', '#f3bcd3', '#f3e3a0'];
  var etat = {vue: 'semaine', ancre: lundi(new Date()), samedi: false, types: {personne: true, vehicule: true, materiel: true}, q: '', aside: true, site: ''};
  try { var mem = JSON.parse(localStorage.getItem('sibella-planning') || '{}'); ['vue','samedi','types','aside','site'].forEach(function(k){ if (mem[k] != null) etat[k] = mem[k]; }); } catch (e){}
  function garder(){ try { localStorage.setItem('sibella-planning', JSON.stringify({vue: etat.vue, samedi: etat.samedi, types: etat.types, aside: etat.aside, site: etat.site})); } catch (e){} }

  var TYPES = [['personne', 'Personnel'], ['vehicule', 'Véhicules'], ['materiel', 'Matériel']];
  var JOURS = ['dim.', 'lun.', 'mar.', 'mer.', 'jeu.', 'ven.', 'sam.'];
  var MOISL = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var ST = {prevu: 'Prévu', fait: 'Réalisé', annule: 'Annulé'};

  // ---------- dates : une demi-journée = H (jours depuis 1970 × 2 + 0 matin / 1 après-midi)
  function p2(n){ return (n < 10 ? '0' : '') + n; }
  function iso(d){ return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); }
  function dt(s){ var a = String(s || '').split('-'); return new Date(+a[0], +a[1] - 1, +a[2]); }
  function plus(d, n){ var x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); return x; }
  function lundi(d){ var x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); var w = (x.getDay() + 6) % 7; return plus(x, -w); }
  function jnum(d){ return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5); }
  function H(s, half){ return jnum(dt(s)) * 2 + (half === 'pm' ? 1 : 0); }
  function dH(h){ var d = new Date(Math.floor(h / 2) * 864e5); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
  function ouvre(d){ var w = d.getDay(); return w !== 0 && (w !== 6 || etat.samedi); }
  function pas(h, k){ while (k > 0){ h++; if (ouvre(dH(h))) k--; } return h; }
  function nbDemi(a, b){ var n = 0; for (var h = a + 1; h <= b; h++) if (ouvre(dH(h))) n++; return n; }
  function semIso(d){ var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); var j = (t.getUTCDay() + 6) % 7; t.setUTCDate(t.getUTCDate() - j + 3); var a = new Date(Date.UTC(t.getUTCFullYear(), 0, 4)); return 1 + Math.round(((t - a) / 864e5 - 3 + ((a.getUTCDay() + 6) % 7)) / 7); }
  function frc(s){ var d = dt(s); return p2(d.getDate()) + '/' + p2(d.getMonth() + 1); }
  function jours(){
    var out = [], d;
    if (etat.vue === 'jour') return [etat.ancre];
    if (etat.vue === 'semaine'){ d = lundi(etat.ancre); for (var i = 0; i < 7; i++){ var x = plus(d, i); if (ouvre(x)) out.push(x); } return out; }
    d = new Date(etat.ancre.getFullYear(), etat.ancre.getMonth(), 1);
    while (d.getMonth() === etat.ancre.getMonth()){ if (ouvre(d)) out.push(d); d = plus(d, 1); }
    return out;
  }

  // ---------- utilitaires
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
  function norm(s){ return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
  function nid(p){ return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
  function encre(hex){ var m = /^#?([0-9a-f]{6})$/i.exec(hex || ''); if (!m) return '#141414'; var n = parseInt(m[1], 16), r = n >> 16, g = (n >> 8) & 255, b = n & 255; return (0.299 * r + 0.587 * g + 0.114 * b) > 150 ? '#141414' : '#ffffff'; }
  function toast(t){ if (S && S.toast) S.toast(t); }
  function peut(){ return !!db && S.canWrite(); }
  function affaire(ref){ return ref ? S.find(String(ref)) : null; }
  function rangSite(x){ var i = sites.indexOf(x || ''); return i < 0 ? 99 : i; }
  // toutes les ressources actives d'un type, le site choisi en premier puis par site
  function listeResTous(type){ return Object.keys(res).map(function(k){ return res[k]; }).filter(function(r){ return r.type === type && r.actif !== false; }).sort(function(a, b){ var sa = etat.site && a.site === etat.site ? -1 : rangSite(a.site), sb = etat.site && b.site === etat.site ? -1 : rangSite(b.site); return (sa - sb) || ((a.ordre || 0) - (b.ordre || 0)) || (a.nom < b.nom ? -1 : 1); }); }
  // ressources affichées dans la grille : celles du site choisi (toutes si « Tous les sites »)
  function listeRes(type){ return listeResTous(type).filter(function(r){ return !etat.site || r.site === etat.site; }); }
  function optSites(sel, vide){ return (vide ? '<option value="">' + esc(vide) + '</option>' : '') + sites.map(function(x){ return '<option value="' + esc(x) + '"' + (x === sel ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join(''); }
  function coulSite(x){ var c = sitesCoul[x]; if (/^#[0-9a-f]{6}$/i.test(c || '')) return c; if (COUL[x]) return COUL[x]; var i = sites.indexOf(x); return i < 0 ? '#efece5' : PAL[i % PAL.length]; }
  // couleur d'une tâche : le ou les sites de rattachement des personnes affectées (à défaut, des véhicules et du matériel)
  function coulTache(t){
    var vus = {}, out = [];
    function prendre(l){ (l || []).forEach(function(p){ var x = res[p] && res[p].site; if (x && !vus[x]){ vus[x] = 1; out.push(x); } }); }
    prendre(t.personnes); if (!out.length){ prendre(t.vehicules); prendre(t.materiels); }
    out.sort(function(a, b){ return rangSite(a) - rangSite(b); });
    return out.length ? out.map(coulSite) : ['#efece5'];
  }
  function nomRes(id){ var r = res[id]; return r ? (r.code || r.nom) : '?'; }
  function listeAct(){ return Object.keys(act).map(function(k){ return act[k]; }).filter(function(a){ return a.actif !== false; }).sort(function(a, b){ return ((a.ordre || 0) - (b.ordre || 0)) || (a.nom < b.nom ? -1 : 1); }); }
  function bornes(t){ return [H(t.date_debut, t.demi_debut), H(t.date_fin || t.date_debut, t.demi_fin || 'pm')]; }
  function tousRes(t){ return [].concat(t.personnes || [], t.vehicules || [], t.materiels || []); }
  function libelle(t){
    var a = act[t.activite], r = affaire(t.ref);
    return {act: a ? a.nom : (t.activite ? 'Activité ?' : ''), ref: t.ref ? String(t.ref) : '', txt: t.titre || (r ? r.client : (t.client || '')), lieu: t.lieu || (r ? r.commune : (t.commune || ''))};
  }
  // conflits : une ressource prise deux fois sur une même demi-journée
  function conflits(){
    var occ = {}, out = {};
    Object.keys(taches).forEach(function(id){
      var t = taches[id]; if (t.statut === 'annule') return; var b = bornes(t);
      tousRes(t).forEach(function(r){ (occ[r] = occ[r] || []).push({id: id, a: b[0], b: b[1]}); });
    });
    Object.keys(occ).forEach(function(r){
      var l = occ[r].sort(function(x, y){ return x.a - y.a; });
      for (var i = 0; i < l.length; i++) for (var j = i + 1; j < l.length && l[j].a <= l[i].b; j++){ out[l[i].id + '|' + r] = 1; out[l[j].id + '|' + r] = 1; }
    });
    return out;
  }
  function occupation(rid, a, b, sauf){
    var o = [];
    Object.keys(taches).forEach(function(id){
      if (id === sauf) return; var t = taches[id]; if (t.statut === 'annule' || tousRes(t).indexOf(rid) < 0) return;
      var x = bornes(t); if (x[0] <= b && x[1] >= a){ var l = libelle(t); o.push((act[t.activite] && act[t.activite].absence ? 'absent : ' : '') + ([l.act, l.ref].filter(Boolean).join(' ') || 'tâche') + ' (' + frc(t.date_debut) + ')'); }
    });
    return o;
  }

  // ---------- abonnement : fenêtre de trois mois autour de la date affichée
  function abonner(){
    if (!db) return;
    var a = iso(plus(new Date(etat.ancre.getFullYear(), etat.ancre.getMonth(), 1), -45)), b = iso(plus(new Date(etat.ancre.getFullYear(), etat.ancre.getMonth() + 1, 0), 45));
    var cle = a + '|' + b, js = jours();
    if (fen && fen.split('|')[0] <= iso(js[0]) && fen.split('|')[1] >= iso(js[js.length - 1])) return;
    if (unsub) unsub();
    fen = cle;
    unsub = db.collection('planning_taches').where('date_fin', '>=', a).where('date_debut', '<=', b).limit(1000).onSnapshot(function(s){
      var m = {}; s.docs.forEach(function(d){ m[d.id] = Object.assign({id: d.id}, d.data()); }); taches = m; dessiner();
      if (s.size >= 1000) toast('Plus de 1 000 tâches sur la période : affichage incomplet.');
    }, function(){ root.querySelector('.pl-msg').textContent = 'Le planning n\'a pas pu être chargé. Rechargez la page.'; });
  }

  // ---------- squelette
  function css(){
    var st = document.createElement('style');
    st.textContent = [
      '#planningView{display:flex;flex-direction:column;gap:12px}',
      '.pl-bar{position:sticky;top:env(safe-area-inset-top,0px);z-index:6;background:var(--bg);padding-block:8px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}',
      '.pl-bar .grp{display:inline-flex;border:1px solid var(--line);border-radius:4px;overflow:hidden;background:var(--surface)}',
      '.pl-bar .grp button{font:inherit;font-size:14px;padding:8px 12px;border:0;background:none;color:var(--ink);cursor:pointer}',
      '.pl-bar .grp button+button{border-left:1px solid var(--line)}',
      '.pl-bar .grp button[aria-pressed="true"]{background:var(--ink);color:var(--bg)}',
      '.pl-per{font-weight:700;font-size:17px;min-width:12em;font-variant-numeric:tabular-nums}',
      '.pl-per small{display:block;font-weight:400;font-size:12px;color:var(--muted)}',
      '.pl-bar select{padding:9px 10px;border:1px solid var(--line);border-radius:4px;background:var(--surface);color:var(--ink);font:inherit;font-weight:700}',
      '.pl-ss{padding:4px 10px 4px 14px;font-size:11.5px;font-weight:700;color:var(--ink);background:color-mix(in srgb,var(--line) 35%,var(--surface));border-bottom:1px solid var(--line);position:sticky;left:0}',
      '.pl-bar input[type=search]{flex:1 1 200px;min-width:0;padding:9px 12px;border:1px solid var(--line);border-radius:4px;background:var(--surface);color:var(--ink);font:inherit}',
      '.pl-bar .chk{color:var(--muted)}',
      '.pl-leg{display:flex;flex-wrap:wrap;gap:5px 12px;font-size:12.5px;color:var(--muted)}',
      '.pl-leg i{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px;vertical-align:-1px;border:1px solid rgba(0,0,0,.12)}',
      '.pl-main{display:grid;grid-template-columns:minmax(0,1fr) 290px;gap:14px;align-items:start}',
      '.pl-main.ferme{grid-template-columns:minmax(0,1fr)}',
      '@media (max-width:980px){.pl-main{grid-template-columns:minmax(0,1fr)}}',
      '.pl-box{background:var(--surface);border:1px solid var(--line);border-radius:4px;overflow-x:auto}',
      '.pl-grid{min-width:860px}',
      '.pl-grid.mois{min-width:1500px}.pl-grid.jour{min-width:0}',
      '.pl-r{display:grid;grid-template-columns:190px minmax(0,1fr);border-bottom:1px solid var(--line)}',
      '.pl-r:last-child{border-bottom:0}',
      '.pl-n{position:sticky;left:0;z-index:2;background:var(--surface);padding:7px 10px;border-right:1px solid var(--line);display:flex;align-items:center;gap:8px;min-width:0;box-shadow:inset 4px 0 0 var(--cg,transparent)}',
      '.pl-n b{font-size:14px;white-space:nowrap}.pl-n span{font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
      '.pl-t{display:grid;grid-auto-rows:minmax(30px,auto);row-gap:3px;padding-block:3px;position:relative}',
      '.pl-c{border-left:1px solid var(--line);min-height:30px}',
      '.pl-c.pm{border-left-style:dotted}',
      '.pl-c.auj{background:color-mix(in srgb,var(--rouge) 6%,transparent)}',
      '.pl-c.sur{background:color-mix(in srgb,var(--ink) 12%,transparent)}',
      '.pl-c:hover{background:var(--row-hover)}',
      '.pl-b{position:relative;z-index:1;margin:0 2px;border:0;border-radius:3px;padding:3px 6px;font:inherit;font-size:12.5px;line-height:1.25;text-align:left;cursor:pointer;overflow:hidden;min-width:0;display:flex;flex-direction:column;justify-content:center}',
      '.pl-b b,.pl-b span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
      '.pl-b span{font-size:11.5px;opacity:.85}',
      '.pl-b.conf{outline:2px dashed var(--rouge);outline-offset:-2px}',
      '.pl-b.conf b{color:inherit}',
      '.pl-b.fait{opacity:.6}',
      '.pl-b{box-shadow:inset 0 0 0 1px rgba(0,0,0,.14)}',
      '.pl-b.dep{box-shadow:inset 0 0 0 1px rgba(0,0,0,.14),inset 0 -3px 0 #141414}',
      '.pl-b.annule{opacity:.4;text-decoration:line-through}',
      '.pl-b.terne{opacity:.18}',
      '.pl-b:focus-visible{outline:2px solid var(--ink);outline-offset:1px}',
      '.pl-h{position:sticky;top:0;z-index:3;background:var(--surface)}',
      '.pl-h .pl-n{font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);font-weight:700}',
      '.pl-h .pl-t{grid-auto-rows:auto;padding:0}',
      '.pl-d{grid-row:1;border-left:1px solid var(--line);padding:6px 6px 2px;font-size:13px;font-weight:700;white-space:nowrap;overflow:hidden}',
      '.pl-d.auj{color:var(--rouge)}',
      '.pl-d small{font-weight:400;color:var(--muted)}',
      '.pl-hd{grid-row:2;border-left:1px solid var(--line);font-size:10.5px;color:var(--muted);padding:0 6px 5px;text-transform:uppercase;letter-spacing:.06em}',
      '.pl-hd.pm{border-left-style:dotted}',
      '.pl-g{display:flex;justify-content:space-between;align-items:baseline;padding:7px 10px;background:var(--bg);border-bottom:1px solid var(--line);font-size:11px;text-transform:uppercase;letter-spacing:.11em;color:var(--muted);font-weight:700;position:sticky;left:0}',
      '.pl-vide{padding:10px 12px;font-size:13px;color:var(--muted)}',
      '.pl-aside{background:var(--surface);border:1px solid var(--line);border-radius:4px;display:flex;flex-direction:column;max-height:calc(100vh - 140px);min-height:0}',
      '.pl-aside h3{margin:0;padding:12px 14px 4px;font-size:15px}',
      '.pl-aside .note{padding:0 14px 8px;font-size:12.5px}',
      '.pl-aside ol{list-style:none;margin:0;padding:0;overflow-y:auto;border-top:1px solid var(--line)}',
      '.pl-aside li{padding:8px 14px;border-bottom:1px solid var(--line);cursor:grab;display:flex;flex-direction:column;gap:2px;font-size:13px}',
      '.pl-aside li:hover{background:var(--row-hover)}',
      '.pl-aside li .ref{font-size:15px}',
      '.pl-aside li .small{font-size:12px}',
      '.pl-aside li .deja{font-size:12px;color:var(--warn)}',
      '.pl-msg:empty{display:none}',
      '.pl-f{display:flex;flex-direction:column;gap:12px}',
      '.pl-f .ligne{display:flex;flex-wrap:wrap;gap:8px 10px;align-items:end}',
      '.pl-f label.f{flex:1 1 140px;min-width:0}',
      '.pl-f input[type=text],.pl-f input[type=date],.pl-f select,.pl-f textarea,.pl-m input[type=text],.pl-m select{padding:8px 10px;border:1px solid var(--line);border-radius:4px;background:var(--bg);color:var(--ink);font:inherit;font-size:14px;text-transform:none;letter-spacing:0;width:100%;min-width:0}',
      '.pl-f textarea{min-height:64px}',
      '.pl-f .look{font-size:13px;color:var(--muted);min-height:1.2em}',
      '.pl-f .look a,.pl-f .carte{color:var(--lien)}',
      '.pl-f fieldset{border:1px solid var(--line);border-radius:4px;padding:8px 10px;margin:0;min-width:0}',
      '.pl-f legend{font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);padding:0 4px}',
      '.pl-ck{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:4px 10px}',
      '.pl-ck label{display:flex;gap:6px;align-items:flex-start;font-size:13.5px}',
      '.pl-ck input{accent-color:var(--rouge);margin-top:2px}',
      '.pl-ck em{display:block;font-style:normal;font-size:11.5px;color:var(--warn)}',
      '.pl-q{display:flex;flex-wrap:wrap;gap:6px}',
      '.pl-f details>summary{cursor:pointer;color:var(--lien);font-size:13.5px}',
      '.pl-q button{font:inherit;font-size:12.5px;padding:5px 9px;border:1px solid var(--line);border-radius:3px;background:var(--bg);color:var(--ink);cursor:pointer}',
      '.pl-m table{min-width:0;width:100%}',
      '.pl-m td{padding:6px 6px;vertical-align:middle}',
      '.pl-m input[type=color]{width:42px;height:32px;padding:0;border:1px solid var(--line);border-radius:4px;background:none}',
      '.pl-m .mini{font:inherit;font-size:12.5px;padding:4px 8px;border:1px solid var(--line);border-radius:3px;background:var(--bg);color:var(--ink);cursor:pointer}'
    ].join('\n');
    document.head.appendChild(st);
  }
  function squelette(){
    root.innerHTML =
      '<div class="pl-bar">' +
        '<button class="btn" id="plAuj">Aujourd\'hui</button>' +
        '<div class="grp"><button id="plPrev" aria-label="Période précédente">‹</button><button id="plNext" aria-label="Période suivante">›</button></div>' +
        '<div class="pl-per" id="plPer"></div>' +
        '<div class="grp" role="group" aria-label="Vue"><button data-v="jour">Jour</button><button data-v="semaine">Semaine</button><button data-v="mois">Mois</button></div>' +
        '<div class="grp" role="group" aria-label="Ressources affichées">' + TYPES.map(function(t){ return '<button data-ty="' + t[0] + '">' + t[1] + '</button>'; }).join('') + '</div>' +
        '<select id="plSite" aria-label="Site"></select>' +
        '<label class="chk"><input type="checkbox" id="plSam"> Samedi</label>' +
        '<input type="search" id="plQ" placeholder="Réf., client, commune, activité…" autocomplete="off">' +
        '<button class="btn primary" id="plNew">+ Nouvelle tâche</button>' +
        '<button class="btn" id="plRes">Ressources</button>' +
        '<button class="btn" id="plAct">Activités</button>' +
        '<button class="btn" id="plXls" hidden>Exporter la période</button>' +
        '<button class="btn" id="plAsideBtn"></button>' +
      '</div>' +
      '<div class="pl-leg" id="plLeg"></div>' +
      '<div class="msg pl-msg" id="plMsg"></div>' +
      '<div class="pl-main" id="plMain"><div class="pl-box"><div id="plGrid"></div></div><aside class="pl-aside" id="plAside"></aside></div>' +
      '<div class="note">Glissez une tâche pour la déplacer (autre jour, autre personne ou autre matériel), ou glissez une affaire « À programmer » sur une case pour la planifier. Une tâche entourée de rouge signale que la personne, le véhicule ou le matériel de cette ligne est pris deux fois. Un trait sous une tâche et la mention « → site » signalent un déplacement hors du site de rattachement.</div>' +
      '<div class="overlay" id="plOv" hidden><div class="panel" role="dialog" aria-modal="true" aria-labelledby="plOvT"><div class="rowbtns" style="justify-content:space-between;align-items:baseline"><h2 id="plOvT"></h2><button class="btn" id="plOvX">Fermer</button></div><div id="plOvB"></div></div></div>';
    $('plAuj').onclick = function(){ etat.ancre = etat.vue === 'semaine' ? lundi(new Date()) : new Date(); tout(); };
    $('plPrev').onclick = function(){ decaler(-1); };
    $('plNext').onclick = function(){ decaler(1); };
    [].forEach.call(root.querySelectorAll('[data-v]'), function(b){ b.onclick = function(){ etat.vue = b.dataset.v; if (etat.vue === 'semaine') etat.ancre = lundi(etat.ancre); garder(); tout(); }; });
    [].forEach.call(root.querySelectorAll('[data-ty]'), function(b){ b.onclick = function(){ etat.types[b.dataset.ty] = !etat.types[b.dataset.ty]; garder(); dessiner(); }; });
    $('plSite').onchange = function(){ etat.site = $('plSite').value; garder(); dessiner(); };
    $('plSam').onchange = function(){ etat.samedi = $('plSam').checked; garder(); dessiner(); };
    var tq = null; $('plQ').oninput = function(){ clearTimeout(tq); tq = setTimeout(function(){ etat.q = norm($('plQ').value.trim()); dessiner(); }, 120); };
    $('plNew').onclick = function(){ var j = jours()[0], auj = iso(new Date()); var d = (etat.vue !== 'jour' && auj >= iso(j) && auj <= iso(jours()[jours().length - 1])) ? auj : iso(j); editer(null, {date_debut: d, date_fin: d}); };
    $('plRes').onclick = gererRes;
    $('plAct').onclick = gererAct;
    $('plXls').onclick = exporter;
    $('plAsideBtn').onclick = function(){ etat.aside = !etat.aside; garder(); dessiner(); };
    $('plOvX').onclick = fermer;
    $('plOv').addEventListener('click', function(e){ if (e.target === $('plOv')) fermer(); });
    document.addEventListener('keydown', function(e){ if (e.key === 'Escape' && !$('plOv').hidden) fermer(); });
    grilleEvenements();
  }
  function decaler(k){
    var a = etat.ancre;
    etat.ancre = etat.vue === 'jour' ? (function(){ var d = plus(a, k); while (!ouvre(d)) d = plus(d, k); return d; })() : etat.vue === 'semaine' ? plus(a, 7 * k) : new Date(a.getFullYear(), a.getMonth() + k, 1);
    tout();
  }
  function tout(){ abonner(); dessiner(); }

  // ---------- grille
  var vis = [], lignesVis = {};
  function dessiner(){
    if (!root || !$('plGrid')) return;
    var js = jours(); vis = [];
    js.forEach(function(d){ var h = jnum(d) * 2; vis.push({h: h, d: d}, {h: h + 1, d: d}); });
    var N = vis.length, aujH = jnum(new Date()) * 2, conf = conflits(), q = etat.q;
    // en-tête période
    var per = '', sub = '';
    if (etat.vue === 'semaine'){ var l = lundi(etat.ancre); per = 'Semaine ' + semIso(l); sub = p2(l.getDate()) + '/' + p2(l.getMonth() + 1) + ' au ' + frc(iso(plus(l, etat.samedi ? 5 : 4))) + '/' + l.getFullYear(); }
    else if (etat.vue === 'mois'){ per = MOISL[etat.ancre.getMonth()] + ' ' + etat.ancre.getFullYear(); sub = 'semaines ' + semIso(js[0]) + ' à ' + semIso(js[js.length - 1]); }
    else { per = JOURS[etat.ancre.getDay()] + ' ' + etat.ancre.getDate() + ' ' + MOISL[etat.ancre.getMonth()]; sub = 'semaine ' + semIso(etat.ancre) + ' · ' + etat.ancre.getFullYear(); }
    $('plPer').innerHTML = esc(per.charAt(0).toUpperCase() + per.slice(1)) + '<small>' + esc(sub + (etat.site ? ' · ' + etat.site : ' · tous les sites')) + '</small>';
    [].forEach.call(root.querySelectorAll('[data-v]'), function(b){ b.setAttribute('aria-pressed', b.dataset.v === etat.vue ? 'true' : 'false'); });
    [].forEach.call(root.querySelectorAll('[data-ty]'), function(b){ b.setAttribute('aria-pressed', etat.types[b.dataset.ty] ? 'true' : 'false'); });
    $('plSam').checked = etat.samedi;
    if (etat.site && sites.length && sites.indexOf(etat.site) < 0) etat.site = '';
    $('plSite').innerHTML = optSites(etat.site, 'Tous les sites');
    $('plAsideBtn').textContent = etat.aside ? 'Masquer « À programmer »' : 'Afficher « À programmer »';
    $('plXls').hidden = !(S.downloads() && window.XLSX);
    $('plLeg').innerHTML = sites.map(function(x){ return '<span><i style="background:' + esc(coulSite(x)) + '"></i>' + esc(x) + '</span>'; }).join('') + (sites.length > 1 ? '<span><i style="background:linear-gradient(100deg,' + coulSite(sites[0]) + ' 0 50%,' + coulSite(sites[1]) + ' 50% 100%)"></i>brigade de deux sites</span>' : '') + listeAct().filter(function(a){ return a.absence; }).map(function(a){ return '<span><i style="background:' + esc(a.couleur || '#999') + '"></i>' + esc(a.nom) + '</span>'; }).join('') + '<span>✓ réalisé</span><span>« s/réf. » : référence pas encore ouverte</span>';
    // lignes : ressources visibles, plus une ligne pour les tâches sans ressource
    var lignes = [], parRes = {}, sansRes = [];
    Object.keys(taches).forEach(function(id){
      var t = taches[id], b = bornes(t), a = -1, z = -1;
      for (var i = 0; i < N; i++){ if (vis[i].h >= b[0] && vis[i].h <= b[1]){ if (a < 0) a = i; z = i; } }
      if (a < 0) return;
      var item = {t: t, a: a, z: z};
      var rs = tousRes(t).filter(function(r){ return res[r]; });
      if (!rs.length){ if (!etat.site || !t.site || t.site === etat.site) sansRes.push(item); } else rs.forEach(function(r){ (parRes[r] = parRes[r] || []).push(item); });
    });
    var cols = 'grid-template-columns:repeat(' + N + ',minmax(0,1fr))';
    var h = '<div class="pl-grid ' + etat.vue + '" style="--n:' + N + '">';
    h += '<div class="pl-r pl-h"><div class="pl-n">Ressources</div><div class="pl-t" style="' + cols + '">';
    js.forEach(function(d, i){ var a = jnum(d) * 2 === aujH; h += '<div class="pl-d' + (a ? ' auj' : '') + '" style="grid-column:' + (2 * i + 1) + ' / span 2">' + esc(etat.vue === 'mois' ? JOURS[d.getDay()].slice(0, 2) + ' ' + d.getDate() : JOURS[d.getDay()] + ' ' + d.getDate()) + (etat.vue !== 'mois' && i === 0 || d.getDate() === 1 ? ' <small>' + esc(MOISL[d.getMonth()].slice(0, 4) + '.') + '</small>' : '') + '</div>'; });
    if (etat.vue !== 'mois') js.forEach(function(d, i){ h += '<div class="pl-hd" style="grid-column:' + (2 * i + 1) + '">matin</div><div class="pl-hd pm" style="grid-column:' + (2 * i + 2) + '">après-midi</div>'; });
    h += '</div></div>';
    function ligne(id, nom, sous, couleur, items){
      items.sort(function(x, y){ return (x.a - y.a) || (y.z - x.z); });
      var pistes = [], html = '';
      items.forEach(function(it){ var k = 0; while (pistes[k] != null && pistes[k] >= it.a) k++; pistes[k] = it.z; it.k = k; });
      var L = Math.max(1, pistes.length);
      for (var i = 0; i < N; i++) html += '<div class="pl-c' + (i % 2 ? ' pm' : '') + (vis[i].h >> 1 === aujH >> 1 ? ' auj' : '') + '" data-s="' + i + '" style="grid-column:' + (i + 1) + ';grid-row:1 / span ' + L + '"></div>';
      items.forEach(function(it){
        var t = it.t, lb = libelle(t), a = act[t.activite] || {}, cs = a.absence ? [a.couleur || '#141414'] : coulTache(t), c = cs[0], r = affaire(t.ref);
        var fond = cs.length > 1 ? 'linear-gradient(100deg,' + cs.map(function(x, i){ return x + ' ' + Math.round(100 * i / cs.length) + '% ' + Math.round(100 * (i + 1) / cs.length) + '%'; }).join(',') + ')' : c;
        var court = etat.vue === 'mois' && (it.z - it.a) < 3;
        var titre = ([lb.act, lb.ref, lb.txt, lb.lieu].filter(Boolean).join(' · ') || 'Tâche') + ' — ' + frc(t.date_debut) + (t.demi_debut === 'pm' ? ' après-midi' : ' matin') + ' → ' + frc(t.date_fin || t.date_debut) + (t.demi_fin === 'am' ? ' matin' : ' après-midi') + (t.heures ? ' (' + t.heures + ')' : '') + ' · ' + tousRes(t).map(nomRes).join(', ') + (conf[t.id + '|' + id] ? ' · CONFLIT : ' + nomRes(id) + ' est déjà pris' : '') + (t.statut && t.statut !== 'prevu' ? ' · ' + ST[t.statut] : '');
        var hit = !q || norm(titre + ' ' + (r ? r.client + ' ' + r.commune : (t.client || '') + ' ' + (t.commune || '')) + ' ' + (t.note || '')).indexOf(q) >= 0;
        var rr = res[id], dep = rr && t.site && rr.site && t.site !== rr.site ? t.site : '';
        if (dep) titre += ' · en déplacement : ' + dep;
        var brig = (t.personnes || []).filter(function(p){ return res[p]; }).map(nomRes);
        var sansRef = !lb.ref && !a.absence && !!(lb.lieu || lb.txt) && !!(t.personnes || []).length;
        if (sansRef) titre = 'Sans référence (à ouvrir) · ' + titre;
        var l1 = court ? (lb.ref || (sansRef ? 's/réf.' : (lb.act || lb.txt || 'Tâche').slice(0, 5))) : (lb.ref ? lb.ref + ' ' + (lb.lieu || lb.txt) : sansRef ? 's/réf. ' + (lb.lieu || lb.txt) : (lb.act || lb.txt || 'Tâche')) + (dep ? ' → ' + dep : '');
        var l2 = court ? '' : [brig.length ? '(' + brig.join('/') + ')' : '', (lb.ref || sansRef) ? (lb.lieu ? lb.txt : '') : (lb.act ? (lb.txt || lb.lieu) : lb.lieu), (lb.ref || sansRef) ? lb.act : ''].filter(Boolean).join(' · ');
        if (t.statut === 'fait') l1 = '✓ ' + l1;
        html += '<button type="button" class="pl-b' + (conf[t.id + '|' + id] ? ' conf' : '') + (t.statut === 'fait' ? ' fait' : '') + (t.statut === 'annule' ? ' annule' : '') + (hit ? '' : ' terne') + (dep ? ' dep' : '') + '" draggable="true" data-t="' + esc(t.id) + '" data-r="' + esc(id) + '" title="' + esc(titre) + '" style="grid-column:' + (it.a + 1) + ' / span ' + (it.z - it.a + 1) + ';grid-row:' + (it.k + 1) + ';background:' + esc(fond) + ';color:' + encre(c) + '"><b>' + esc(l1) + '</b>' + (l2 ? '<span>' + esc(l2) + '</span>' : '') + '</button>';
      });
      return '<div class="pl-r" data-res="' + esc(id) + '"><div class="pl-n" style="--cg:' + esc(couleur || 'transparent') + '"><b>' + esc(nom) + '</b><span>' + esc(sous || '') + '</span></div><div class="pl-t" style="' + cols + '">' + html + '</div></div>';
    }
    var vide = true;
    if (sansRes.length){ h += '<div class="pl-g"><span>Sans ressource</span><span>' + sansRes.length + '</span></div>' + ligne('', 'À affecter', 'aucune ressource', 'var(--rouge)', sansRes); vide = false; }
    TYPES.forEach(function(ty){
      if (!etat.types[ty[0]]) return;
      var l = listeRes(ty[0]);
      h += '<div class="pl-g"><span>' + ty[1] + '</span><span>' + l.length + '</span></div>';
      if (!l.length) h += '<div class="pl-vide">Aucune ressource de ce type' + (etat.site ? ' rattachée au site ' + esc(etat.site) : '') + '. Ajoutez-en ou rattachez-les avec le bouton « Ressources ».</div>';
      var der = null;
      l.forEach(function(r){
        if (!etat.site && (r.site || '') !== der){ der = r.site || ''; h += '<div class="pl-ss">' + esc(der || 'Site non renseigné') + '</div>'; }
        vide = false; h += ligne(r.id, r.code || r.nom, (r.code && r.nom && r.nom !== r.code ? r.nom : (r.detail || '')) + (r.terrain === false ? ' · pas de terrain' : '') + (r.maintenance ? ' · en maintenance' : ''), r.couleur, parRes[r.id] || []);
      });
      if (etat.site){
        var renf = listeResTous(ty[0]).filter(function(r){ return r.site !== etat.site && (parRes[r.id] || []).some(function(it){ return it.t.statut !== 'annule' && (it.t.site === etat.site || tousRes(it.t).some(function(x){ return res[x] && res[x].site === etat.site; })); }); });
        if (renf.length){
          h += '<div class="pl-ss">Renforts venus d\'un autre site (' + renf.length + ')</div>';
          renf.forEach(function(r){ vide = false; h += ligne(r.id, r.code || r.nom, 'rattaché à ' + (r.site || '?') + (r.maintenance ? ' · en maintenance' : ''), r.couleur, parRes[r.id] || []); });
        }
      }
    });
    h += '</div>';
    $('plGrid').innerHTML = h;
    var voirAside = etat.aside && peut();
    $('plMain').className = 'pl-main' + (voirAside ? '' : ' ferme');
    $('plAside').hidden = !voirAside;
    $('plAsideBtn').hidden = !peut();
    if (voirAside) aside();
    if (!Object.keys(res).length && !Object.keys(act).length) $('plMsg').textContent = 'Aucune ressource ni activité enregistrée pour l\'instant : utilisez « Ressources » et « Activités » pour les créer.';
    else $('plMsg').textContent = '';
  }

  // ---------- affaires à programmer (étoile du registre)
  function aside(){
    var rows = S.rows().filter(function(r){ return !r.livre && S.isProg(r); }).sort(function(a, b){ return b.ref - a.ref; });
    var plan = {}; Object.keys(taches).forEach(function(id){ var t = taches[id]; if (t.ref && t.statut !== 'annule' && t.date_fin >= iso(new Date())) (plan[t.ref] = plan[t.ref] || []).push(t.date_debut); });
    var q = etat.q;
    if (q) rows = rows.filter(function(r){ return r._s.indexOf(q) >= 0; });
    $('plAside').innerHTML = '<h3>À programmer <span class="small">(' + rows.length + ')</span></h3><div class="note">Affaires marquées ★ dans le registre. Glissez-en une sur une case de la grille.</div><ol>' + (rows.length ? rows.map(function(r){
      var d = plan[r.ref]; d = d && d.sort()[0];
      return '<li draggable="true" data-ref="' + esc(r.ref) + '" tabindex="0" title="Glisser sur le planning, ou Entrée pour planifier"><span><span class="ref">' + esc(r.lib) + '</span> <span class="client">' + esc(r.client || '') + '</span></span><span class="small">' + esc([r.commune, r.mission].filter(Boolean).join(' · ')) + '</span><span class="small">' + esc(r.avancement || '') + '</span>' + (d ? '<span class="deja">Déjà planifiée le ' + frc(d) + '</span>' : '') + '</li>';
    }).join('') : '<li style="cursor:default"><span class="small">Aucune affaire en attente de programmation.</span></li>') + '</ol>';
  }

  // ---------- interactions sur la grille
  var glisse = null;
  function cellule(e){ return e.target.closest && e.target.closest('.pl-c'); }
  function grilleEvenements(){
    root.addEventListener('click', function(e){
      var b = e.target.closest('.pl-b'); if (b){ editer(taches[b.dataset.t]); return; }
      var c = cellule(e); if (c){ var row = c.closest('.pl-r'), v = vis[+c.dataset.s]; if (!v) return; var d = iso(v.d), half = (v.h % 2) ? 'pm' : 'am', rid = row.dataset.res; var pre = {date_debut: d, date_fin: d, demi_debut: 'am', demi_fin: 'pm'}; if (rid && res[rid]) pre[{personne: 'personnes', vehicule: 'vehicules', materiel: 'materiels'}[res[rid].type]] = [rid]; editer(null, pre); return; }
      var li = e.target.closest('.pl-aside li[data-ref]'); if (li && e.detail >= 2) editer(null, {ref: +li.dataset.ref});
    });
    root.addEventListener('keydown', function(e){ var li = e.target.closest && e.target.closest('.pl-aside li[data-ref]'); if (li && e.key === 'Enter') editer(null, {ref: +li.dataset.ref}); });
    root.addEventListener('dragstart', function(e){
      var b = e.target.closest && e.target.closest('.pl-b'), li = e.target.closest && e.target.closest('.pl-aside li[data-ref]');
      if (b){ var r = b.getBoundingClientRect(), t = taches[b.dataset.t]; if (!t) return; var span = +b.style.gridColumn.split('span ')[1] || 1; glisse = {t: t.id, de: b.dataset.r, off: Math.min(span - 1, Math.max(0, Math.floor((e.clientX - r.left) / r.width * span)))}; }
      else if (li) glisse = {ref: +li.dataset.ref};
      else return;
      try { e.dataTransfer.setData('text/plain', glisse.t || String(glisse.ref)); e.dataTransfer.effectAllowed = 'move'; } catch (err){}
    });
    root.addEventListener('dragend', function(){ glisse = null; [].forEach.call(root.querySelectorAll('.pl-c.sur'), function(c){ c.classList.remove('sur'); }); });
    root.addEventListener('dragover', function(e){ var c = cellule(e); if (!c || !glisse) return; e.preventDefault(); [].forEach.call(root.querySelectorAll('.pl-c.sur'), function(x){ if (x !== c) x.classList.remove('sur'); }); c.classList.add('sur'); });
    root.addEventListener('drop', function(e){
      var c = cellule(e); if (!c || !glisse) return; e.preventDefault(); c.classList.remove('sur');
      var g = glisse; glisse = null;
      var row = c.closest('.pl-r'), cible = row.dataset.res, si = +c.dataset.s;
      if (g.ref){ var v = vis[si], d = iso(v.d), pre = {ref: g.ref, date_debut: d, date_fin: d, demi_debut: 'am', demi_fin: 'pm'}; if (cible && res[cible]) pre[{personne: 'personnes', vehicule: 'vehicules', materiel: 'materiels'}[res[cible].type]] = [cible]; editer(null, pre); return; }
      deplacer(g, cible, Math.max(0, si - g.off));
    });
  }
  function deplacer(g, cible, si){
    var t = taches[g.t]; if (!t) return;
    if (!peut()){ toast('Lecture seule : vous ne pouvez pas modifier le planning depuis cette vue.'); return; }
    si = Math.min(si, vis.length - 1); var b = bornes(t), n = nbDemi(b[0], b[1]), nh = vis[si].h, fh = pas(nh, n);
    var ch = {date_debut: iso(dH(nh)), demi_debut: nh % 2 ? 'pm' : 'am', date_fin: iso(dH(fh)), demi_fin: fh % 2 ? 'pm' : 'am'};
    var msgR = '';
    if (cible !== g.de && cible && res[cible]){
      if (!g.de || !res[g.de]){ var k0 = {personne: 'personnes', vehicule: 'vehicules', materiel: 'materiels'}[res[cible].type]; ch[k0] = [cible]; }
      else if (res[g.de].type !== res[cible].type){ toast('Déposez la tâche sur une ligne du même type (personnel, véhicule ou matériel).'); return; }
      else { var k = {personne: 'personnes', vehicule: 'vehicules', materiel: 'materiels'}[res[cible].type], l = (t[k] || []).slice(), i = l.indexOf(g.de); if (l.indexOf(cible) < 0){ if (i >= 0){ l[i] = cible; msgR = nomRes(cible) + ' remplace ' + nomRes(g.de) + (l.length > 1 ? ' (les autres restent affectés)' : '') + (k === 'personnes' ? ' ; vérifiez le véhicule' : ''); } else l.push(cible); } ch[k] = l; }
    }
    var avant = taches[g.t], nv = Object.assign({}, avant, ch, {maj_le: new Date().toISOString()}); delete nv.id;
    taches[g.t] = Object.assign({id: g.t}, nv); dessiner();
    db.doc('planning_taches/' + g.t).set(nv).then(function(){ toast('Tâche déplacée au ' + frc(ch.date_debut) + (ch.demi_debut === 'pm' ? ' après-midi' : ' matin') + '.' + (msgR ? ' ' + msgR + '.' : '')); }, function(){ taches[g.t] = avant; dessiner(); toast('Déplacement impossible pour le moment. Réessayez.'); });
  }

  // ---------- fenêtre de saisie
  function ouvrir(titre, html){ $('plOvT').textContent = titre; $('plOvB').innerHTML = html; $('plOv').hidden = false; }
  function fermer(){ $('plOv').hidden = true; $('plOvB').innerHTML = ''; }
  function caseRes(type, cle, sel, a, b, sauf, siteT){
    var l = listeResTous(type), der = null, h = '';
    if (!l.length) return '<div class="small">Aucune ressource de ce type.</div>';
    l.forEach(function(r){
      if ((r.site || '') !== der){ if (der !== null) h += '</div>'; der = r.site || ''; h += '<div class="small" style="margin:6px 0 2px;font-weight:700">' + esc(der || 'Site non renseigné') + '</div><div class="pl-ck">'; }
      var o = occupation(r.id, a, b, sauf);
      h += '<label><input type="checkbox" data-k="' + cle + '" value="' + esc(r.id) + '"' + (sel.indexOf(r.id) >= 0 ? ' checked' : '') + '><span>' + esc(r.code || r.nom) + (r.code && r.nom && r.nom !== r.code ? ' <span class="small">' + esc(r.nom) + '</span>' : '') + (r.terrain === false ? '<em>ne fait pas de terrain</em>' : '') + (r.maintenance ? '<em>en maintenance</em>' : '') + (siteT && r.site && r.site !== siteT ? '<em>rattaché à ' + esc(r.site) + '</em>' : '') + (o.length ? '<em>' + esc(o.join(' ; ')) + '</em>' : '') + '</span></label>';
    });
    return h + '</div>';
  }
  function editer(t, pre){
    var nouveau = !t, f = Object.assign({activite: '', ref: 0, titre: '', lieu: '', date_debut: iso(new Date()), demi_debut: 'am', date_fin: '', demi_fin: 'pm', heures: '', personnes: [], vehicules: [], materiels: [], note: '', statut: 'prevu', site: ''}, t || {}, pre || {});
    if (!f.date_fin) f.date_fin = f.date_debut;
    f._plus = !nouveau && !!(f.activite || f.titre || f.lieu || f.heures || f.note || (f.materiels || []).length);
    if (nouveau){ f.vehicules = (f.vehicules || []).slice(); (f.personnes || []).forEach(function(p){ var v = res[p] && res[p].vehicule_defaut; if (v && res[v] && res[v].actif !== false && f.vehicules.indexOf(v) < 0) f.vehicules.push(v); }); }
    var acts = listeAct();
    if (!f.activite && acts.length) f.activite = '';
    var ro = !peut();
    function corps(){
      var a = H(f.date_debut, f.demi_debut), b = H(f.date_fin, f.demi_fin), r = affaire(f.ref);
      var siteEff = f.site || '';
      var look = f.ref ? (r ? esc(r.client || '') + ' — ' + esc(r.commune || '') + (r.livre ? ' [livré]' : '') + '<br>' + esc(r.avancement || '') : f.client ? esc(f.client) + ' — ' + esc(f.commune || '') : 'Référence absente de l\'index des références.') : '';
      var lieu = f.lieu || (r ? r.commune : '');
      var sel = function(v, x){ return v === x ? ' selected' : ''; };
      var blocTL = '<div class="ligne"><label class="f" for="plFt">' + (r ? 'Libellé (facultatif)' : 'Client / objet') + '<input id="plFt" type="text" value="' + esc(f.titre) + '" placeholder="' + esc(r ? r.client : 'ex. M. Dupont, bornage') + '"></label>' +
        '<label class="f" for="plFe">' + (r ? 'Emplacement' : 'Commune / lieu') + '<input id="plFe" type="text" value="' + esc(f.lieu) + '" placeholder="' + esc(r && r.commune ? r.commune : 'ex. LEVIE') + '"></label>' + (lieu ? '<a class="carte" target="_blank" rel="noopener" href="https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(lieu + ', Corse') + '">Voir sur la carte ↗</a>' : '') + '</div>';
      return '<form class="pl-f" id="plF" novalidate>' +
        (ro ? '<div class="msg">Lecture seule : vous ne pouvez pas modifier le planning depuis cette vue.</div>' : '') +
        '<div class="ligne"><label class="f" for="plFr" style="flex:0 1 140px">Référence<input id="plFr" type="text" inputmode="numeric" maxlength="5" value="' + esc(f.ref || '') + '" placeholder="16xxx"></label><div class="look" id="plFl" style="flex:1 1 220px">' + (look || 'Pas encore de référence ? Laissez vide et indiquez le client et la commune.') + '</div></div>' +
        (r ? '' : blocTL) +
        '<fieldset><legend>Chefs de brigade / personnel</legend>' + caseRes('personne', 'personnes', f.personnes || [], a, b, t && t.id, '') + '</fieldset>' +
        '<div class="ligne" style="flex-wrap:nowrap"><label class="f" for="plFd1" style="flex:1 1 120px">Du<input id="plFd1" type="date" value="' + esc(f.date_debut) + '"></label><label class="f" for="plFh1" style="flex:0 1 112px">&nbsp;<select id="plFh1"><option value="am"' + sel(f.demi_debut !== 'pm', true) + '>matin</option><option value="pm"' + sel(f.demi_debut, 'pm') + '>après-midi</option></select></label>' +
        '<label class="f" for="plFd2" style="flex:1 1 120px">Au<input id="plFd2" type="date" value="' + esc(f.date_fin) + '"></label><label class="f" for="plFh2" style="flex:0 1 112px">&nbsp;<select id="plFh2"><option value="am"' + sel(f.demi_fin, 'am') + '>matin</option><option value="pm"' + sel(f.demi_fin !== 'am', true) + '>après-midi</option></select></label></div>' +
        '<div class="pl-q"><button type="button" data-q="j">Journée</button><button type="button" data-q="am">Matin</button><button type="button" data-q="pm">Après-midi</button><button type="button" data-q="s">Toute la semaine</button></div>' +
        '<details id="plFplus"' + (f._plus ? ' open' : '') + '><summary>Plus d\'options : activité, véhicules, matériel, horaire, remarque…</summary><div class="pl-f" style="margin-top:12px">' +
          '<label class="f" for="plFa">Activité (facultatif)<select id="plFa"><option value="">— aucune —</option>' + acts.map(function(x){ return '<option value="' + esc(x.id) + '"' + sel(x.id, f.activite) + '>' + esc(x.nom) + '</option>'; }).join('') + '</select></label>' +
          (r ? blocTL : '') +
          '<label class="f" for="plFhr">Heures (facultatif)<input id="plFhr" type="text" value="' + esc(f.heures) + '" placeholder="ex. RDV 9 h 30 sur place"></label>' +
          (nouveau ? '<div class="ligne"><label class="f" for="plFrep">Répéter<select id="plFrep"><option value="">Ne pas répéter</option><option value="j">Chaque jour ouvré</option><option value="s">Chaque semaine</option></select></label><label class="f" for="plFjq">Jusqu\'au<input id="plFjq" type="date"></label></div>' : '') +
          '<fieldset><legend>Véhicules</legend>' + caseRes('vehicule', 'vehicules', f.vehicules || [], a, b, t && t.id, siteEff) + '</fieldset>' +
          '<fieldset><legend>Matériel</legend>' + caseRes('materiel', 'materiels', f.materiels || [], a, b, t && t.id, siteEff) + '</fieldset>' +
          '<label class="f" for="plFn">Remarque<textarea id="plFn">' + esc(f.note) + '</textarea></label>' +
          '<div class="ligne"><label class="f" for="plFsi" style="max-width:260px">Site d\'intervention<select id="plFsi">' + optSites(f.site, '— non précisé —') + '</select></label>' +
          '<label class="f" for="plFs" style="max-width:220px">Statut<select id="plFs">' + Object.keys(ST).map(function(k){ return '<option value="' + k + '"' + sel(f.statut, k) + '>' + ST[k] + '</option>'; }).join('') + '</select></label></div>' +
        '</div></details>' +
        '<div class="msg" id="plFm" role="status"></div>' +
        '<div class="rowbtns"><button type="submit" class="btn primary" id="plFok"' + (ro ? ' disabled' : '') + '>' + (nouveau ? 'Créer' : 'Enregistrer') + '</button>' + (nouveau ? '' : '<button type="button" class="btn" data-x="dup"' + (ro ? ' disabled' : '') + '>Dupliquer</button><button type="button" class="btn" data-x="suppr"' + (ro ? ' disabled' : '') + '>Supprimer</button>') + '</div>' +
        '<div class="rowbtns" id="plFsup" hidden></div>' +
        '</form>';
    }
    function lire(){
      var F = $('plF'); if (!F) return;
      f.activite = $('plFa').value; f.ref = /^\d{5}$/.test($('plFr').value.trim()) ? +$('plFr').value.trim() : 0;
      f.titre = $('plFt').value.trim(); f.lieu = $('plFe').value.trim(); f.heures = $('plFhr').value.trim();
      f.date_debut = $('plFd1').value || f.date_debut; f.demi_debut = $('plFh1').value; f.date_fin = $('plFd2').value || f.date_debut; f.demi_fin = $('plFh2').value;
      f.note = $('plFn').value.trim(); f.statut = $('plFs').value; f.site = $('plFsi').value;
      ['personnes', 'vehicules', 'materiels'].forEach(function(k){ f[k] = [].slice.call(F.querySelectorAll('input[data-k="' + k + '"]:checked')).map(function(c){ return c.value; }); });
      f._plus = $('plFplus').open;
      if ($('plFrep')){ f._rep = $('plFrep').value; f._jq = $('plFjq').value; }
    }
    function refaire(){ var y = $('plOvB').scrollTop, ae = document.activeElement, fid = ae && ae.id, pos = ae && ae.selectionStart; var rep = f._rep, jq = f._jq; $('plOvB').innerHTML = corps(); if ($('plFrep')){ $('plFrep').value = rep || ''; $('plFjq').value = jq || ''; } brancher(); $('plOvB').scrollTop = y; var el = fid && $(fid); if (el){ el.focus(); try { if (pos != null) el.setSelectionRange(pos, pos); } catch (e){} } }
    function rafraichir(){ lire(); refaire(); }
    function brancher(){
      var F = $('plF');
      ['plFd1', 'plFh1', 'plFd2', 'plFh2'].forEach(function(i){ $(i).addEventListener('change', function(){ lire(); if (f.date_fin < f.date_debut){ $('plFd2').value = f.date_debut; } rafraichir(); }); });
      F.addEventListener('change', function(e){
        var c = e.target; if (c.dataset.k !== 'personnes') return;
        var v = res[c.value] && res[c.value].vehicule_defaut; if (!v) return;
        var vc = F.querySelector('input[data-k="vehicules"][value="' + v + '"]'); if (!vc) return;
        if (c.checked){ if (!vc.checked){ vc.checked = true; toast((res[v].code || res[v].nom) + ' ajouté avec ' + (res[c.value].code || res[c.value].nom) + '.'); } return; }
        var autre = [].some.call(F.querySelectorAll('input[data-k="personnes"]:checked'), function(x){ return res[x.value] && res[x.value].vehicule_defaut === v; });
        if (!autre) vc.checked = false;
      });
      $('plFsi').addEventListener('change', rafraichir);
      $('plFr').addEventListener('input', function(){ var v = $('plFr').value.trim(); if (/^\d{5}$/.test(v) || !v) rafraichir(); });
      F.addEventListener('click', function(e){
        var q = e.target.closest('[data-q]'), x = e.target.closest('[data-x]');
        if (q){ lire(); var k = q.dataset.q; if (k === 's'){ var l = lundi(dt(f.date_debut)); f.date_debut = iso(l); f.date_fin = iso(plus(l, etat.samedi ? 5 : 4)); f.demi_debut = 'am'; f.demi_fin = 'pm'; } else { f.date_fin = f.date_debut; f.demi_debut = k === 'pm' ? 'pm' : 'am'; f.demi_fin = k === 'am' ? 'am' : 'pm'; } refaire(); return; }
        if (!x) return;
        if (x.dataset.x === 'dup'){ lire(); editer(null, Object.assign({}, f, {id: undefined, serie: undefined})); $('plFm').textContent = 'Copie : choisissez la nouvelle date puis « Créer ».'; return; }
        if (x.dataset.x === 'suppr'){ var n = t.serie ? Object.keys(taches).filter(function(k){ return taches[k].serie === t.serie; }).length : 1; var box = $('plFsup'); box.hidden = false; box.innerHTML = '<button type="button" class="btn primary" data-x="s1">Supprimer cette tâche</button>' + (n > 1 ? '<button type="button" class="btn" data-x="sn">Supprimer toute la série (' + n + ' visibles)</button>' : '') + '<button type="button" class="btn" data-x="s0">Annuler</button>'; return; }
        if (x.dataset.x === 's0'){ $('plFsup').hidden = true; return; }
        if (x.dataset.x === 's1' || x.dataset.x === 'sn'){
          var ids = x.dataset.x === 'sn' ? Object.keys(taches).filter(function(k){ return taches[k].serie === t.serie; }) : [t.id];
          Promise.all(ids.map(function(k){ return db.doc('planning_taches/' + k).delete(); })).then(function(){ fermer(); toast(ids.length > 1 ? ids.length + ' tâches supprimées.' : 'Tâche supprimée.'); }, function(){ $('plFm').textContent = 'Suppression impossible pour le moment. Réessayez.'; });
        }
      });
      F.addEventListener('submit', function(e){
        e.preventDefault(); lire(); var m = $('plFm');
        if (!peut()){ m.textContent = 'Lecture seule.'; return; }
        if (!f.activite && !f.titre && !f.lieu && !$('plFr').value.trim()){ m.textContent = 'Indiquez une référence, ou à défaut le client et la commune.'; return; }
        if (!f.personnes.length && !f.vehicules.length && !f.materiels.length){ m.textContent = 'Cochez au moins une personne.'; return; }
        if ($('plFr').value.trim() && !f.ref){ m.textContent = 'La référence doit compter 5 chiffres.'; return; }
        if (f.ref && S.rows().length && !affaire(f.ref) && f._forcer !== f.ref){ f._forcer = f.ref; m.textContent = 'La référence ' + f.ref + ' ne figure pas dans l\'index des références (mis à jour le ' + (S.genere() || '?') + '). Si elle vient d\'être ouverte, cliquez de nouveau sur le bouton pour l\'enregistrer quand même.'; return; }
        if (H(f.date_fin, f.demi_fin) < H(f.date_debut, f.demi_debut)){ m.textContent = 'La fin est avant le début.'; return; }
        var now = new Date().toISOString(), base = {activite: f.activite, ref: f.ref, titre: f.titre, lieu: f.lieu, date_debut: f.date_debut, demi_debut: f.demi_debut, date_fin: f.date_fin, demi_fin: f.demi_fin, heures: f.heures, personnes: f.personnes, vehicules: f.vehicules, materiels: f.materiels, note: f.note, statut: f.statut, site: f.site, maj_le: now};
        var aff = affaire(f.ref); base.client = aff ? (aff.client || '') : (f.client || ''); base.commune = aff ? (aff.commune || '') : (f.commune || '');
        var ecr = [];
        if (nouveau){
          base.cree_le = now;
          var occ = [{d1: f.date_debut, d2: f.date_fin}];
          if (f._rep){
            if (!f._jq){ m.textContent = 'Indiquez jusqu\'à quelle date répéter.'; return; }
            var ecart = jnum(dt(f.date_fin)) - jnum(dt(f.date_debut)), d = dt(f.date_debut), fin = dt(f._jq);
            occ = [];
            while (d <= fin && occ.length < 60){ if (f._rep === 's' || ouvre(d)) occ.push({d1: iso(d), d2: iso(plus(d, ecart))}); d = plus(d, f._rep === 's' ? 7 : 1); }
            if (!occ.length){ m.textContent = 'Aucune date dans l\'intervalle choisi.'; return; }
            base.serie = nid('s');
          }
          occ.forEach(function(o){ var id = nid('t'); ecr.push(db.doc('planning_taches/' + id).set(Object.assign({}, base, {date_debut: o.d1, date_fin: o.d2}))); });
        } else {
          if (t.serie) base.serie = t.serie; base.cree_le = t.cree_le || now;
          ecr.push(db.doc('planning_taches/' + t.id).set(base));
        }
        $('plFok').disabled = true;
        Promise.all(ecr).then(function(){ fermer(); toast(ecr.length > 1 ? ecr.length + ' tâches créées.' : (nouveau ? 'Tâche créée.' : 'Tâche enregistrée.')); var dd = dt(f.date_debut); if (iso(dd) < iso(jours()[0]) || iso(dd) > iso(jours()[jours().length - 1])){ etat.ancre = etat.vue === 'semaine' ? lundi(dd) : dd; tout(); } }, function(err){ $('plFok').disabled = false; m.textContent = err && err.code === 'quota_exceeded' ? 'La base est pleine : supprimez d\'anciennes tâches avant d\'en créer.' : 'Enregistrement impossible pour le moment. Réessayez.'; });
      });
    }
    ouvrir(nouveau ? 'Nouvelle tâche' : 'Tâche du ' + frc(f.date_debut), corps());
    brancher();
    var fr0 = $('plFr'); if (fr0 && nouveau && !f.ref) fr0.focus();
  }

  // ---------- gestion des ressources
  function gererRes(){
    function corps(){
      return '<div class="pl-m"><div class="note">Personnes, véhicules et matériel qui apparaissent en lignes dans le planning, chacun rattaché à son site. Une ressource masquée garde son historique.</div>' +
        TYPES.map(function(ty){
          var l = Object.keys(res).map(function(k){ return res[k]; }).filter(function(r){ return r.type === ty[0]; }).sort(function(a, b){ return ((a.ordre || 0) - (b.ordre || 0)); });
          return '<h3 style="margin:14px 0 4px;font-size:15px">' + ty[1] + '</h3><div class="tablebox"><table><tbody>' + (l.length ? l.map(function(r){
            return '<tr data-id="' + esc(r.id) + '"><td style="width:42px"><input type="color" aria-label="Couleur" value="' + esc(r.couleur || '#777777') + '" data-f="couleur"></td><td><input type="text" aria-label="Code" value="' + esc(r.code || '') + '" data-f="code" placeholder="code" style="max-width:90px"></td><td><input type="text" aria-label="Nom" value="' + esc(r.nom || '') + '" data-f="nom" placeholder="nom"></td><td><select aria-label="Site" data-f="site">' + optSites(r.site || '', 'site ?') + '</select></td>' + (ty[0] === 'personne' ? '<td style="white-space:nowrap"><label class="chk"><input type="checkbox" data-f="terrain"' + (r.terrain === false ? '' : ' checked') + '> Terrain</label></td><td><select aria-label="Véhicule par défaut" data-f="vehicule_defaut"><option value="">véhicule par défaut ?</option>' + listeResTous('vehicule').map(function(v){ return '<option value="' + esc(v.id) + '"' + (v.id === r.vehicule_defaut ? ' selected' : '') + '>' + esc(v.code || v.nom) + (v.site ? ' · ' + esc(v.site) : '') + '</option>'; }).join('') + '</select></td>' : '<td style="white-space:nowrap"><label class="chk"><input type="checkbox" data-f="maintenance"' + (r.maintenance ? ' checked' : '') + '> En maintenance</label></td>') + '<td style="white-space:nowrap"><button class="mini" data-x="up" title="Monter">↑</button> <button class="mini" data-x="down" title="Descendre">↓</button> <button class="mini" data-x="act">' + (r.actif === false ? 'Réafficher' : 'Masquer') + '</button></td></tr>'; }).join('') : '<tr><td class="small">Aucune.</td></tr>') + '</tbody></table></div>' +
            '<div class="rpf" style="margin-top:6px"><input type="text" id="nr-' + ty[0] + '" placeholder="' + (ty[0] === 'personne' ? 'nom de la personne' : ty[0] === 'vehicule' ? 'ex. véhicule et immatriculation' : 'ex. GPS, station totale, drone') + '" style="flex:1 1 200px"><button class="btn" data-x="add" data-ty="' + ty[0] + '">Ajouter</button></div>';
        }).join('') +
        '<div class="msg" id="plMm" role="status"></div></div>';
    }
    ouvrir('Ressources', corps());
    var B = $('plOvB'), m = function(t){ var e = $('plMm'); if (e) e.textContent = t; };
    function ecrire(id, d){ if (!peut()){ m('Lecture seule.'); return Promise.reject(); } return db.doc('planning_ressources/' + id).set(d).catch(function(){ m('Écriture impossible, réessayez.'); throw 0; }); }
    B.onchange = function(e){ var tr = e.target.closest('tr[data-id]'); if (!tr || !e.target.dataset.f) return; var r = res[tr.dataset.id]; var d = Object.assign({}, r); delete d.id; d[e.target.dataset.f] = e.target.type === 'checkbox' ? e.target.checked : e.target.value.trim(); ecrire(r.id, d).then(function(){ m('Enregistré.'); }, function(){}); };
    B.onclick = function(e){
      var x = e.target.closest('[data-x]'); if (!x) return;
      if (x.dataset.x === 'add'){ var ty = x.dataset.ty, inp = $('nr-' + ty), v = inp.value.trim(); if (!v){ m('Saisissez un nom.'); return; } var id = nid('r'); var n = listeRes(ty).length; var nr = {type: ty, nom: v, code: '', couleur: '#777777', ordre: (n + 1) * 10, actif: true, site: etat.site || ''}; ecrire(id, nr).then(function(){ res[id] = Object.assign({id: id}, nr); B.innerHTML = corps(); m(etat.site ? 'Ajouté au site ' + etat.site + '.' : 'Ajouté : choisissez son site dans la liste.'); }, function(){}); return; }
      if (x.dataset.x === 'equipe'){
        var deja = {}; Object.keys(res).forEach(function(k){ if (res[k].type === 'personne' && res[k].code) deja[res[k].code] = 1; });
        var l = S.equipe().filter(function(p){ return !deja[p.ini]; });
        if (!l.length){ m('Tous les responsables de dossiers figurent déjà dans le planning.'); return; }
        var base0 = listeRes('personne').length;
        Promise.all(l.map(function(p, i){ var id = 'p-' + p.ini; return ecrire(id, {type: 'personne', code: p.aff, nom: p.nom || '', couleur: '#777777', ordre: (base0 + i + 1) * 10, actif: true, site: ''}); })).then(function(){ m(l.length + ' personnes ajoutées.'); }, function(){});
        return;
      }
      var tr = x.closest('tr[data-id]'); if (!tr) return; var r = res[tr.dataset.id], d = Object.assign({}, r); delete d.id;
      if (x.dataset.x === 'act'){ d.actif = r.actif === false; ecrire(r.id, d).then(function(){ r.actif = d.actif; B.innerHTML = corps(); }, function(){}); return; }
      var l2 = Object.keys(res).map(function(k){ return res[k]; }).filter(function(y){ return y.type === r.type; }).sort(function(a, b){ return (a.ordre || 0) - (b.ordre || 0); }), i2 = l2.indexOf(r), j2 = i2 + (x.dataset.x === 'up' ? -1 : 1);
      if (j2 < 0 || j2 >= l2.length) return;
      l2.splice(i2, 1); l2.splice(j2, 0, r);
      Promise.all(l2.map(function(y, k){ var dd = Object.assign({}, y); delete dd.id; dd.ordre = (k + 1) * 10; y.ordre = dd.ordre; return ecrire(y.id, dd); })).then(function(){ B.innerHTML = corps(); }, function(){});
    };
  }

  // ---------- gestion des activités
  function gererAct(){
    function corps(){
      var l = Object.keys(act).map(function(k){ return act[k]; }).sort(function(a, b){ return (a.ordre || 0) - (b.ordre || 0); });
      return '<div class="pl-m"><div class="note">Chaque tâche porte une activité ; sa couleur colore la tâche dans la grille. « Indisponibilité » marque les congés et absences : la personne apparaît alors comme absente dans la saisie.</div><div class="tablebox"><table><tbody>' +
        (l.length ? l.map(function(a){ return '<tr data-id="' + esc(a.id) + '"><td style="width:42px"><input type="color" aria-label="Couleur" value="' + esc(a.couleur || '#777777') + '" data-f="couleur"></td><td><input type="text" aria-label="Nom" value="' + esc(a.nom || '') + '" data-f="nom"></td><td style="white-space:nowrap"><label class="chk"><input type="checkbox" data-f="absence"' + (a.absence ? ' checked' : '') + '> Indisponibilité</label></td><td><button class="mini" data-x="act">' + (a.actif === false ? 'Réafficher' : 'Masquer') + '</button></td></tr>'; }).join('') : '<tr><td class="small">Aucune activité.</td></tr>') +
        '</tbody></table></div><div class="rpf" style="margin-top:8px"><input type="text" id="na" placeholder="nouvelle activité" style="flex:1 1 200px"><button class="btn" data-x="add">Ajouter</button></div><div class="msg" id="plMm" role="status"></div></div>';
    }
    ouvrir('Activités', corps());
    var B = $('plOvB'), m = function(t){ var e = $('plMm'); if (e) e.textContent = t; };
    function ecrire(id, d){ if (!peut()){ m('Lecture seule.'); return Promise.reject(); } return db.doc('planning_activites/' + id).set(d).catch(function(){ m('Écriture impossible, réessayez.'); throw 0; }); }
    B.onchange = function(e){ var tr = e.target.closest('tr[data-id]'); if (!tr || !e.target.dataset.f) return; var a = act[tr.dataset.id], d = Object.assign({}, a); delete d.id; d[e.target.dataset.f] = e.target.type === 'checkbox' ? e.target.checked : e.target.value.trim(); ecrire(a.id, d).then(function(){ m('Enregistré.'); }, function(){}); };
    B.onclick = function(e){
      var x = e.target.closest('[data-x]'); if (!x) return;
      if (x.dataset.x === 'add'){ var v = $('na').value.trim(); if (!v){ m('Saisissez un nom.'); return; } var id = nid('a'), d = {nom: v, couleur: '#777777', absence: false, ordre: (Object.keys(act).length + 1) * 10, actif: true}; ecrire(id, d).then(function(){ act[id] = Object.assign({id: id}, d); B.innerHTML = corps(); }, function(){}); return; }
      var tr = x.closest('tr[data-id]'); if (!tr) return; var a = act[tr.dataset.id], d2 = Object.assign({}, a); delete d2.id; d2.actif = a.actif === false;
      ecrire(a.id, d2).then(function(){ a.actif = d2.actif; B.innerHTML = corps(); }, function(){});
    };
  }

  // ---------- export Excel de la période affichée
  function exporter(){
    var dl = S.downloads(); if (!dl || !window.XLSX) return;
    var js = jours(), entete = ['Ressource'];
    js.forEach(function(d){ entete.push(JOURS[d.getDay()] + ' ' + frc(iso(d)) + ' matin', JOURS[d.getDay()] + ' ' + frc(iso(d)) + ' après-midi'); });
    var data = [entete];
    TYPES.forEach(function(ty){ listeRes(ty[0]).forEach(function(r){
      var ligne = [r.code || r.nom];
      js.forEach(function(d){ [0, 1].forEach(function(k){ var h = jnum(d) * 2 + k, l = [];
        Object.keys(taches).forEach(function(id){ var t = taches[id]; if (t.statut === 'annule' || tousRes(t).indexOf(r.id) < 0) return; var b = bornes(t); if (h >= b[0] && h <= b[1]){ var lb = libelle(t); l.push(lb.act + (lb.ref ? ' ' + lb.ref : '') + (lb.txt ? ' ' + lb.txt : '')); } });
        ligne.push(l.join(' / ')); }); });
      data.push(ligne); }); });
    var ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = [{wch: 16}].concat(entete.slice(1).map(function(){ return {wch: 24}; }));
    var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Planning');
    var nom = 'PLANNING_SIBELLA_' + iso(js[0]).replace(/-/g, '') + '_' + iso(js[js.length - 1]).replace(/-/g, '') + '.xlsx';
    dl.save({filename: nom, data: new Blob([XLSX.write(wb, {bookType: 'xlsx', type: 'array'})])}).catch(function(){});
  }

  // ---------- index des références (rechargé quand le fichier change)
  var tAside = null;
  window.SIBELLA_rows = function(){ if (!root) return; clearTimeout(tAside); tAside = setTimeout(function(){ if ($('plAside')) dessiner(); }, 200); };

  function init(){
    S = window.SIBELLA; root = $('planningView'); if (!S || !root) return;
    db = S.db(); css(); squelette();
    db.collection('planning_ressources').limit(300).onSnapshot(function(s){ var m = {}; s.docs.forEach(function(x){ m[x.id] = Object.assign({id: x.id}, x.data()); }); res = m; dessiner(); });
    db.doc('planning_config/sites').onSnapshot(function(x){ var d = x.exists ? x.data() : null; sites = (d && Array.isArray(d.liste)) ? d.liste.filter(function(v){ return typeof v === 'string' && v; }) : []; sitesCoul = (d && d.couleurs && typeof d.couleurs === 'object') ? d.couleurs : {}; dessiner(); });
    db.collection('planning_activites').limit(200).onSnapshot(function(s){ var m = {}; s.docs.forEach(function(x){ m[x.id] = Object.assign({id: x.id}, x.data()); }); act = m; dessiner(); });
    fen = ''; abonner(); dessiner();
  }
  window.PLANNING_START = init;
})();
