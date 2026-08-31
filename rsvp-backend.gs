/* ==========================================================================
   NATALIE & ERIC - WEDDING BACKEND
   Google Apps Script bound to the "Natalie & Eric - Wedding RSVPs" sheet.

   It serves three things:
     1. the public site content (schedule, FAQ, meal options)
     2. the guest RSVP flow (last-name lookup, party load, save)
     3. the admin portal at natalie-eric.website/admin

   FIRST TIME SETUP
     1. Project Settings -> Script properties -> add  ADMIN_PASSWORD
     2. Run  setup()  once. It builds every tab and seeds the content.
     3. Deploy > New deployment > Web app, Execute as Me, Access Anyone.
     4. Paste the /exec URL into CONFIG.apiUrl in index.html and admin.

   REDEPLOYING AFTER A CODE CHANGE
     Deploy > Manage deployments > pencil > Version: New version > Deploy.
     That keeps the same URL. "New deployment" mints a new one and breaks
     the site.
   ========================================================================== */

var SH = {
  GUESTS:'Guests', PARTIES:'Parties', CONFIG:'Config', SCHEDULE:'Schedule',
  FAQ:'FAQ', MEALS:'Meals', LOG:'Log', TEMPLATES:'Templates'
};

/* column indexes, zero based */
var G = { PARTY:0, FIRST:1, LAST:2, EMAIL:3, PHONE:4, ATTENDING:5, MEAL:6,
          DIET:7, UPDATED:8, GID:9, NOTES:10 };
var P = { PARTY:0, NAME:1, OPEN:2, TOKEN:3, LINK:4, WAVE:5, DEADLINE:6,
          INVITED:7, REPLIED:8, ATTENDING:9, EMAILS:10, NOTE:11, LASTREPLY:12,
          SENT:13, REMINDED:14, ADMIN:15 };
var GUEST_COLS = 11, PARTY_COLS = 16;

var SESSION_HOURS = 12;

/* ==========================================================================
   ROUTING
   ========================================================================== */

function doGet(e){
  var p = (e && e.parameter) || {};
  try{
    switch(p.action || 'party'){
      case 'ping':   return json({ok:true, pong:true, time:new Date().toISOString()});
      case 'site':   return json({ok:true, site:siteContent()});
      case 'find':   return json(findParties(p.q));
      case 'party':  return json(getParty(p.token));
      default:       return json({ok:false, error:'Unknown action'});
    }
  }catch(err){ return json({ok:false, error:msg(err)}); }
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(25000);
    var b = JSON.parse(e.postData.contents);
    var a = b.action || '';

    /* --- public --- */
    if(a === 'rsvp')  return json(saveRsvp(b));
    if(a === 'login') return json(login(b));

    /* --- everything below needs a valid admin session --- */
    if(a.indexOf('admin.') !== 0) return json({ok:false, error:'Unknown action'});
    if(!validSession(b.token))    return json({ok:false, error:'Your session expired. Please sign in again.', reauth:true});

    switch(a){
      case 'admin.data':        adoptSheetEdits(); return json({ok:true, data:adminData(b.light)});
      case 'admin.log':         return json({ok:true, log:readLog()});
      case 'admin.saveGuest':   return json(adminSaveGuest(b));
      case 'admin.deleteGuest': return json(adminDeleteGuest(b));
      case 'admin.saveParty':   return json(adminSaveParty(b));
      case 'admin.savePartyBundle': return json(adminSavePartyBundle(b));
      case 'admin.deleteParty': return json(adminDeleteParty(b));
      case 'admin.saveRows':    return json(adminSaveRows(b));
      case 'admin.saveConfig':  return json(adminSaveConfig(b));
      case 'admin.preview':     return json({ok:true, html:renderEmail(b.subject, b.body, previewParty(b.partyId))});
      case 'admin.send':        return json(adminSend(b));
      case 'admin.logout':      return json(logout(b));
      default:                  return json({ok:false, error:'Unknown action'});
    }
  }catch(err){
    return json({ok:false, error:msg(err)});
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
                       .setMimeType(ContentService.MimeType.JSON);
}
function msg(err){ return String((err && err.message) || err); }

/* ==========================================================================
   AUTH
   Password lives in Script Properties, never in this file or the repo.
   Sessions are random tokens with an expiry, also in Script Properties.
   ========================================================================== */

function props(){ return PropertiesService.getScriptProperties(); }

function login(b){
  var want = String(props().getProperty('ADMIN_PASSWORD') || '');
  var user = String(props().getProperty('ADMIN_USER') || 'wedding');
  if(!want) return {ok:false, error:'No admin password is set on the server yet.'};

  /* crude but effective brute-force brake */
  var fails = Number(props().getProperty('LOGIN_FAILS') || 0);
  var until = Number(props().getProperty('LOGIN_BLOCKED_UNTIL') || 0);
  if(until && Date.now() < until){
    return {ok:false, error:'Too many attempts. Try again in a few minutes.'};
  }

  if(String(b.user||'').trim().toLowerCase() !== user.toLowerCase() ||
     String(b.pass||'') !== want){
    fails++;
    props().setProperty('LOGIN_FAILS', String(fails));
    if(fails >= 6){
      props().setProperty('LOGIN_BLOCKED_UNTIL', String(Date.now() + 15*60*1000));
      props().setProperty('LOGIN_FAILS', '0');
    }
    return {ok:false, error:'That username or password is not right.'};
  }

  props().setProperty('LOGIN_FAILS', '0');
  props().deleteProperty('LOGIN_BLOCKED_UNTIL');

  var token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  var sessions = readSessions();
  sessions[token] = Date.now() + SESSION_HOURS*3600*1000;
  writeSessions(sessions);
  return {ok:true, token:token, expires:sessions[token]};
}

function logout(b){
  var s = readSessions();
  delete s[String(b.token||'')];
  writeSessions(s);
  return {ok:true};
}

function readSessions(){
  try{ return JSON.parse(props().getProperty('SESSIONS') || '{}'); }
  catch(e){ return {}; }
}
function writeSessions(s){
  var now = Date.now(), out = {};
  Object.keys(s).forEach(function(k){ if(s[k] > now) out[k] = s[k]; });
  props().setProperty('SESSIONS', JSON.stringify(out));
}
function validSession(token){
  if(!token) return false;
  var s = readSessions();
  return !!(s[token] && s[token] > Date.now());
}

/* ==========================================================================
   HELPERS
   ========================================================================== */

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(n){
  var s = ss().getSheetByName(n);
  if(!s) throw new Error('Missing tab "'+n+'". Run setup() once.');
  return s;
}
function rows(name){
  var s = sheet(name), last = s.getLastRow();
  if(last < 2) return [];
  return s.getRange(2,1,last-1,s.getLastColumn()).getValues();
}
function pad(v){
  if(v === '' || v === null || v === undefined) return '';
  var s = String(v).trim();
  return /^\d+$/.test(s) ? ('0000'+s).slice(-4) : s;
}
function normAttend(v){
  var s = String(v||'').trim().toLowerCase();
  if(s === 'yes' || s === 'y' || s === 'true'  || s === 'attending') return 'yes';
  if(s === 'no'  || s === 'n' || s === 'false' || s === 'declined')  return 'no';
  return '';
}
function isTrue(v){ return v === true || String(v).trim().toUpperCase() === 'TRUE'; }
function ymd(d){
  if(!d) return '';
  if(d instanceof Date) return Utilities.formatDate(d, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(d).trim();
}
function prettyDate(v){
  var s = ymd(v);
  if(!s) return '';
  var parts = s.split('-');
  if(parts.length !== 3) return s;
  var d = new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]));
  if(isNaN(d)) return s;
  var months = ['January','February','March','April','May','June','July',
                'August','September','October','November','December'];
  return months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}
function cfg(key, dflt){
  var v = rows(SH.CONFIG);
  for(var i=0;i<v.length;i++){
    if(String(v[i][0]).trim().toLowerCase() === String(key).trim().toLowerCase()){
      return (v[i][1] === '' || v[i][1] === null) ? dflt : v[i][1];
    }
  }
  return dflt;
}
function setCfg(key, value){
  var s = sheet(SH.CONFIG), v = s.getDataRange().getValues();
  for(var i=1;i<v.length;i++){
    if(String(v[i][0]).trim().toLowerCase() === String(key).trim().toLowerCase()){
      s.getRange(i+1, 2).setValue(value); return;
    }
  }
  s.appendRow([key, value, '']);
}
function newToken(){
  var abc = 'abcdefghjkmnpqrstuvwxyz23456789', out = '';
  for(var i=0;i<8;i++) out += abc.charAt(Math.floor(Math.random()*abc.length));
  return out;
}
function siteUrl(){
  return String(cfg('Site URL','https://natalie-eric.website')).replace(/\/+$/,'');
}
function sendAsAddress(){
  return String(cfg('Send As','natalie.eric.2027@gmail.com')).trim();
}
function sendAs(){
  var want = sendAsAddress();
  if(!want) return '';
  try{
    var a = GmailApp.getAliases();
    for(var i=0;i<a.length;i++){ if(a[i].toLowerCase() === want.toLowerCase()) return a[i]; }
  }catch(e){}
  return '';
}
function mailOptions(){
  var o = { name:'Natalie & Eric', replyTo: sendAsAddress() };
  var alias = sendAs();
  if(alias) o.from = alias;
  return o;
}
function logChange(party, guest, field, from, to, by){
  try{
    var s = sheet(SH.LOG), row = s.getLastRow() + 1;
    /* columns 2-7 as text, so a value like +55 11 9... is not read as a formula */
    s.getRange(row, 2, 1, 6).setNumberFormat('@');
    s.getRange(row, 1, 1, 7).setValues([[new Date(), pad(party), String(guest||''),
                                         String(field||''), String(from||''),
                                         String(to||''), by || 'guest']]);
  }catch(e){}
}

/* ==========================================================================
   PUBLIC: site content
   The website renders its schedule, FAQ and meal list from here, so Eric can
   change them in the sheet without touching code.
   ========================================================================== */

function siteContent(){
  var schedule = rows(SH.SCHEDULE)
    .filter(function(r){ return String(r[3]).trim() && isTrue(r[7]); })
    .sort(function(a,b){ return Number(a[0]||0) - Number(b[0]||0); })
    .map(function(r){
      return {
        start: String(r[1]).trim(), end: String(r[2]).trim(),
        title: String(r[3]).trim(), place: String(r[4]).trim(),
        note:  String(r[5]).trim(),
        tags:  String(r[6]||'').split(',').map(function(s){return s.trim();}).filter(String)
      };
    });

  var faq = rows(SH.FAQ)
    .filter(function(r){ return String(r[1]).trim() && isTrue(r[3]); })
    .sort(function(a,b){ return Number(a[0]||0) - Number(b[0]||0); })
    .map(function(r){ return { q:String(r[1]).trim(), a:String(r[2]) }; });

  var meals = rows(SH.MEALS)
    .filter(function(r){ return String(r[1]).trim() && isTrue(r[2]); })
    .sort(function(a,b){ return Number(a[0]||0) - Number(b[0]||0); })
    .map(function(r){ return String(r[1]).trim(); });

  return {
    schedule: schedule,
    faq:      faq,
    meals:    meals,
    deadline: ymd(cfg('RSVP Deadline','')),
    venue:    String(cfg('Venue','Riverway Clubhouse')),
    address:  String(cfg('Venue Address','9001 Bill Fox Way, Burnaby, BC V5J 5J3')),
    email:    sendAsAddress()
  };
}

/* ==========================================================================
   PUBLIC: find a party by last name
   Returns every open party containing that surname, each with the full list
   of names in it, so a guest can pick the right household. Parties that are
   not open are invisible, which is also what keeps invite waves private:
   a guest who has not been invited yet simply is not found, and learns
   nothing about who was invited first.
   ========================================================================== */

function findParties(q){
  var needle = String(q||'').trim().toLowerCase();
  if(needle.length < 2) return {ok:true, matches:[]};

  var guests  = rows(SH.GUESTS);
  var parties = rows(SH.PARTIES);

  var members = {};
  guests.forEach(function(g){
    var id = pad(g[G.PARTY]);
    if(!id) return;
    var name = (String(g[G.FIRST]||'').trim()+' '+String(g[G.LAST]||'').trim()).trim();
    if(!name) return;
    (members[id] = members[id] || []).push(name);
  });

  var hit = {};
  guests.forEach(function(g){
    var last  = String(g[G.LAST] ||'').trim().toLowerCase();
    var first = String(g[G.FIRST]||'').trim().toLowerCase();
    var full  = (first+' '+last).trim();
    if(last.indexOf(needle) === 0 || first.indexOf(needle) === 0 || full.indexOf(needle) > -1){
      hit[pad(g[G.PARTY])] = true;
    }
  });

  var out = [];
  parties.forEach(function(r){
    var id = pad(r[P.PARTY]);
    if(!hit[id] || !r[P.TOKEN]) return;
    if(!isTrue(r[P.OPEN])) return;
    out.push({
      code:    String(r[P.TOKEN]),
      label:   String(r[P.NAME] || ('Party '+id)),
      members: members[id] || []
    });
  });
  return {ok:true, matches: out.slice(0, 12)};
}

/* ==========================================================================
   PUBLIC: load one party
   ========================================================================== */

function partyRowByToken(token){
  if(!token) return 0;
  var v = rows(SH.PARTIES);
  for(var i=0;i<v.length;i++){
    if(String(v[i][P.TOKEN]).trim() === String(token).trim()) return i+2;
  }
  return 0;
}

function buildParty(row){
  var pv = sheet(SH.PARTIES).getRange(row,1,1,PARTY_COLS).getValues()[0];
  var id = pad(pv[P.PARTY]);
  var guests = rows(SH.GUESTS);
  var list = [];
  guests.forEach(function(g, i){
    if(pad(g[G.PARTY]) !== id) return;
    var name = (String(g[G.FIRST]||'').trim()+' '+String(g[G.LAST]||'').trim()).trim();
    if(!name) return;
    list.push({
      id:        String(g[G.GID] || ('r'+(i+2))),
      row:       i+2,
      name:      name,
      attending: normAttend(g[G.ATTENDING]),
      meal:      String(g[G.MEAL] || ''),
      diet:      String(g[G.DIET] || '')
    });
  });
  return {
    code:      String(pv[P.TOKEN] || ''),
    partyId:   id,
    label:     String(pv[P.NAME] || ('Party '+id)),
    open:      isTrue(pv[P.OPEN]),
    deadline:  ymd(pv[P.DEADLINE]),
    responded: list.some(function(g){ return g.attending === 'yes' || g.attending === 'no'; }),
    note:      String(pv[P.NOTE] || ''),
    email:     String(pv[P.EMAILS] || '').split(/[,;]/)[0].trim(),
    guests:    list
  };
}

function getParty(token){
  if(!token) return {ok:false, error:'No invitation code was supplied.'};
  var row = partyRowByToken(token);
  if(!row) return {ok:false, error:"We couldn't find that invitation link."};
  return {ok:true, site:siteContent(), party:buildParty(row)};
}

/* ==========================================================================
   PUBLIC: save an RSVP
   Enforces the per-party lock on the server, logs every change, and emails
   a notification.
   ========================================================================== */

function saveRsvp(body){
  var row = partyRowByToken(body.token);
  if(!row) return {ok:false, error:"We couldn't find that invitation link."};

  var before = buildParty(row);
  if(!before.open){
    return {ok:false, locked:true,
            error:'RSVPs for your party are closed. Please email us and we will help.'};
  }

  var gs = sheet(SH.GUESTS), byId = {};
  before.guests.forEach(function(g){ byId[g.id] = g; });

  var stamp = new Date(), changes = [];
  (body.guests || []).forEach(function(sub){
    var t = byId[sub.id];
    if(!t) return;
    var attending = sub.attending === 'yes' ? 'Yes' : sub.attending === 'no' ? 'No' : '';
    var meal = attending === 'Yes' ? String(sub.meal||'') : '';
    var diet = attending === 'Yes' ? String(sub.diet||'') : '';

    if(normAttend(attending) !== t.attending){
      changes.push([t.name,'Attending', t.attending || '(no reply)', normAttend(attending) || '(cleared)']);
    }
    if(meal !== t.meal) changes.push([t.name,'Meal', t.meal || '(none)', meal || '(none)']);
    if(diet !== t.diet) changes.push([t.name,'Dietary notes', t.diet || '(none)', diet || '(none)']);

    gs.getRange(t.row, G.ATTENDING+1, 1, 3).setNumberFormat('@');
    gs.getRange(t.row, G.ATTENDING+1).setValue(attending);
    gs.getRange(t.row, G.MEAL+1).setValue(meal);
    gs.getRange(t.row, G.DIET+1).setValue(diet);
    gs.getRange(t.row, G.UPDATED+1).setValue(stamp);
  });

  var ps = sheet(SH.PARTIES);
  if(body.note)  ps.getRange(row, P.NOTE+1).setNumberFormat('@').setValue(String(body.note).slice(0,1000));
  if(body.email) ps.getRange(row, P.EMAILS+1).setNumberFormat('@').setValue(String(body.email).slice(0,200));
  ps.getRange(row, P.LASTREPLY+1).setValue(stamp);

  SpreadsheetApp.flush();

  var isUpdate = before.responded;
  changes.forEach(function(c){
    logChange(before.partyId, c[0], c[1], c[2], c[3], isUpdate ? 'guest (changed)' : 'guest');
  });

  var after = buildParty(row);
  try{ notifyRsvp(after, body, isUpdate, changes); }catch(e){}
  return {ok:true, party:after};
}

function notifyRsvp(party, body, isUpdate, changes){
  if(String(cfg('Notify On RSVP','TRUE')).toUpperCase() !== 'TRUE') return;
  var to = String(cfg('Notify Email','') || sendAsAddress());
  if(!to) return;

  var yes = party.guests.filter(function(g){ return g.attending === 'yes'; });
  var lines = party.guests.map(function(g){
    return '  ' + g.name + ': ' +
      (g.attending === 'yes'
        ? 'Attending (' + (g.meal || 'no meal chosen') + ')' + (g.diet ? ' [' + g.diet + ']' : '')
        : g.attending === 'no' ? 'Not attending' : 'No reply');
  }).join('\n');

  var changeText = '';
  if(isUpdate && changes.length){
    changeText = '\n\nThey CHANGED an earlier reply:\n' + changes.map(function(c){
      return '  ' + c[0] + ' - ' + c[1] + ': ' + c[2] + '  ->  ' + c[3];
    }).join('\n');
  }

  GmailApp.sendEmail(to,
    (isUpdate ? 'RSVP updated: ' : 'RSVP: ') + party.label + ' (' + yes.length + ' attending)',
    party.label + (isUpdate ? ' changed their reply.' : ' just replied.') + '\n\n' + lines +
      changeText +
      (body.note  ? '\n\nTheir note:\n' + body.note : '') +
      (body.email ? '\n\nContact: ' + body.email : '') +
      '\n\nParty ' + party.partyId + '\n' + ss().getUrl(),
    mailOptions());
}

/* ==========================================================================
   ADMIN: read everything the portal needs in one call
   ========================================================================== */

/* ==========================================================================
   ADOPT WHATEVER WAS TYPED STRAIGHT INTO THE SHEET

   Natalie and Eric can add rows to the Guests tab by hand, which is much
   faster than the portal for a long list. This runs before every admin read
   and quietly finishes the job: gives each guest an id, creates any party a
   guest refers to, and gives every party a name, a link code and an invite
   link. New parties come in CLOSED, exactly like ones made in the portal, so
   a hand-typed row can never leak into an earlier invite wave.

   It only writes when something is actually missing, so the normal case
   costs one read and nothing else.
   ========================================================================== */
function adoptSheetEdits(){
  var wrote = false;

  /* --- guests: fill in ids, remember which parties are referenced --- */
  var gs = sheet(SH.GUESTS), gv = rows(SH.GUESTS), referenced = {};
  for(var i=0;i<gv.length;i++){
    var r = gv[i], row = i+2;
    var name = (String(r[G.FIRST]||'').trim()+' '+String(r[G.LAST]||'').trim()).trim();
    if(!name) continue;

    var party = pad(r[G.PARTY]);
    if(!party){
      /* a guest typed with no party number gets one of their own */
      party = nextPartyId(referenced);
      gs.getRange(row, G.PARTY+1).setNumberFormat('@').setValue(party);
      wrote = true;
    } else if(String(r[G.PARTY]) !== party){
      gs.getRange(row, G.PARTY+1).setNumberFormat('@').setValue(party);   /* 1 -> 0001 */
      wrote = true;
    }
    referenced[party] = true;

    if(!String(r[G.GID]||'').trim()){
      gs.getRange(row, G.GID+1).setNumberFormat('@')
        .setValue('g-' + party + '-' + Utilities.getUuid().slice(0,6));
      wrote = true;
    }
  }

  /* --- parties: create the missing ones, finish the incomplete ones --- */
  var ps = sheet(SH.PARTIES), pv = rows(SH.PARTIES), seen = {}, used = {};
  for(var j=0;j<pv.length;j++){
    var pr = pv[j], prow = j+2;
    var pid = pad(pr[P.PARTY]);
    if(!pid) continue;
    if(String(pr[P.PARTY]) !== pid){
      ps.getRange(prow, P.PARTY+1).setNumberFormat('@').setValue(pid);
      wrote = true;
    }
    seen[pid] = true;
    if(pr[P.TOKEN]) used[String(pr[P.TOKEN])] = true;

    if(!String(pr[P.NAME]||'').trim()){
      ps.getRange(prow, P.NAME+1).setNumberFormat('@').setValue('Party ' + pid);
      wrote = true;
    }
    if(!String(pr[P.TOKEN]||'').trim()){
      var t; do { t = newToken(); } while(used[t]);
      used[t] = true;
      ps.getRange(prow, P.TOKEN+1).setNumberFormat('@').setValue(t);
      ps.getRange(prow, P.LINK+1).setNumberFormat('@').setValue(siteUrl() + '/?i=' + t);
      wrote = true;
    } else if(!String(pr[P.LINK]||'').trim()){
      ps.getRange(prow, P.LINK+1).setNumberFormat('@')
        .setValue(siteUrl() + '/?i=' + String(pr[P.TOKEN]).trim());
      wrote = true;
    }
    if(pr[P.OPEN] !== true && pr[P.OPEN] !== false){
      ps.getRange(prow, P.OPEN+1).insertCheckboxes();
      ps.getRange(prow, P.OPEN+1).setValue(false);        /* closed until invited */
      wrote = true;
    }
    if(pr[P.WAVE] === '' || pr[P.WAVE] === null){
      ps.getRange(prow, P.WAVE+1).setValue(1);
      wrote = true;
    }
  }

  /* a guest can name a party that has no row yet; make it */
  for(var key in referenced){
    if(!seen[key]){ ensureParty(key); wrote = true; }
  }

  if(wrote){ SpreadsheetApp.flush(); refreshPartyMetrics(); }
  return wrote;
}

function readLog(){
  return rows(SH.LOG).slice(-400).reverse().map(function(r){
    return { when: r[0] ? new Date(r[0]).toISOString() : '', party:pad(r[1]),
             guest:String(r[2]||''), field:String(r[3]||''),
             from:String(r[4]||''), to:String(r[5]||''), by:String(r[6]||'') };
  });
}

function nextPartyId(extra){
  var used = {};
  rows(SH.PARTIES).forEach(function(r){ var p = pad(r[P.PARTY]); if(p) used[p] = true; });
  rows(SH.GUESTS).forEach(function(r){ var p = pad(r[G.PARTY]); if(p) used[p] = true; });
  for(var k in (extra||{})) used[k] = true;
  var n = 1, id;
  do { id = ('0000'+n).slice(-4); n++; } while(used[id]);
  return id;
}

/* adminData(light) — light skips the change log, which is by far the biggest
   part of the payload and is only ever looked at on its own tab. */
function adminData(light){
  var guests = rows(SH.GUESTS).map(function(r, i){
    return { row:i+2, party:pad(r[G.PARTY]), first:String(r[G.FIRST]||''),
             last:String(r[G.LAST]||''), email:String(r[G.EMAIL]||''),
             phone:String(r[G.PHONE]||''), attending:normAttend(r[G.ATTENDING]),
             meal:String(r[G.MEAL]||''), diet:String(r[G.DIET]||''),
             updated: r[G.UPDATED] ? new Date(r[G.UPDATED]).toISOString() : '',
             id:String(r[G.GID]||''), notes:String(r[G.NOTES]||'') };
  }).filter(function(g){ return g.first || g.last; });

  var parties = rows(SH.PARTIES).map(function(r, i){
    return { row:i+2, party:pad(r[P.PARTY]), name:String(r[P.NAME]||''),
             open:isTrue(r[P.OPEN]), token:String(r[P.TOKEN]||''),
             link:String(r[P.LINK]||''), wave:String(r[P.WAVE]||''),
             deadline:ymd(r[P.DEADLINE]), emails:String(r[P.EMAILS]||''),
             note:String(r[P.NOTE]||''),
             lastReply: r[P.LASTREPLY] ? new Date(r[P.LASTREPLY]).toISOString() : '',
             sent:      r[P.SENT]      ? new Date(r[P.SENT]).toISOString()      : '',
             reminded:  r[P.REMINDED]  ? new Date(r[P.REMINDED]).toISOString()  : '',
             admin:String(r[P.ADMIN]||'') };
  }).filter(function(p){ return p.party; });

  var log = light ? null : readLog();

  return {
    guests:   guests,
    parties:  parties,
    log:      log,      /* null when light; the Change log tab asks for it */
    schedule: rows(SH.SCHEDULE).map(function(r,i){
                return { row:i+2, order:r[0], start:String(r[1]||''), end:String(r[2]||''),
                         title:String(r[3]||''), place:String(r[4]||''),
                         note:String(r[5]||''), tags:String(r[6]||''), visible:isTrue(r[7]) };
              }).filter(function(x){ return x.title; }),
    faq:      rows(SH.FAQ).map(function(r,i){
                return { row:i+2, order:r[0], q:String(r[1]||''), a:String(r[2]||''), visible:isTrue(r[3]) };
              }).filter(function(x){ return x.q; }),
    meals:    rows(SH.MEALS).map(function(r,i){
                return { row:i+2, order:r[0], meal:String(r[1]||''), visible:isTrue(r[2]) };
              }).filter(function(x){ return x.meal; }),
    config:   rows(SH.CONFIG).map(function(r){
                return { key:String(r[0]||''), value: r[1] instanceof Date ? ymd(r[1]) : String(r[1]===null?'':r[1]),
                         help:String(r[2]||'') };
              }).filter(function(x){ return x.key; }),
    templates: rows(SH.TEMPLATES).map(function(r){
                return { key:String(r[0]||''), subject:String(r[1]||''), body:String(r[2]||'') };
               }).filter(function(x){ return x.key; }),
    sheetUrl: ss().getUrl(),
    siteUrl:  siteUrl()
  };
}

/* ==========================================================================
   ADMIN: guests
   ========================================================================== */

function adminSaveGuest(b){
  var g = b.guest || {}, s = sheet(SH.GUESTS);
  var party = pad(g.party);
  if(!party) return {ok:false, error:'Every guest needs a party number.'};
  if(!String(g.first||'').trim() && !String(g.last||'').trim())
    return {ok:false, error:'Every guest needs a name.'};

  ensureParty(party);

  var row = Number(g.row || 0);
  var isNew = !row;
  var before = null;

  if(isNew){
    row = firstFreeRow(s, 1);
    if(!g.id) g.id = 'g-' + party + '-' + Utilities.getUuid().slice(0,6);
  } else {
    var cur = s.getRange(row,1,1,GUEST_COLS).getValues()[0];
    before = { party:pad(cur[G.PARTY]), first:String(cur[G.FIRST]||''), last:String(cur[G.LAST]||''),
               email:String(cur[G.EMAIL]||''), phone:String(cur[G.PHONE]||''),
               attending:normAttend(cur[G.ATTENDING]), meal:String(cur[G.MEAL]||''),
               diet:String(cur[G.DIET]||'') };
    if(!g.id) g.id = String(cur[G.GID] || ('g-'+party+'-'+Utilities.getUuid().slice(0,6)));
  }

  var name = (String(g.first||'').trim()+' '+String(g.last||'').trim()).trim();
  var attending = g.attending === 'yes' ? 'Yes' : g.attending === 'no' ? 'No' : '';

  /* text format everywhere but the timestamp, so a phone like +55 11 9... */
  /* is stored as text instead of being read as a formula                  */
  s.getRange(row,1,1,G.UPDATED).setNumberFormat('@');
  s.getRange(row,G.GID+1,1,2).setNumberFormat('@');
  s.getRange(row,1,1,GUEST_COLS).setValues([[
    party, String(g.first||'').trim(), String(g.last||'').trim(),
    String(g.email||'').trim(), String(g.phone||'').trim(),
    attending, String(g.meal||''), String(g.diet||''),
    (before && before.attending !== normAttend(attending)) || isNew ? new Date() : (s.getRange(row, G.UPDATED+1).getValue() || ''),
    g.id, String(g.notes||'')
  ]]);

  if(isNew){
    logChange(party, name, 'Guest added', '', name, 'admin');
  } else {
    if(before.party !== party)             logChange(party, name, 'Moved party', before.party, party, 'admin');
    if(before.attending !== normAttend(attending))
      logChange(party, name, 'Attending', before.attending || '(no reply)', normAttend(attending) || '(cleared)', 'admin');
    if(before.meal  !== String(g.meal||''))  logChange(party, name, 'Meal',  before.meal  || '(none)', String(g.meal||'')  || '(none)', 'admin');
    if(before.diet  !== String(g.diet||''))  logChange(party, name, 'Dietary notes', before.diet || '(none)', String(g.diet||'') || '(none)', 'admin');
    if(before.email !== String(g.email||'').trim()) logChange(party, name, 'Email', before.email || '(none)', String(g.email||'') || '(none)', 'admin');
    if(before.phone !== String(g.phone||'').trim()) logChange(party, name, 'Phone', before.phone || '(none)', String(g.phone||'') || '(none)', 'admin');
    var oldName = (before.first+' '+before.last).trim();
    if(oldName !== name) logChange(party, name, 'Name', oldName, name, 'admin');
  }

  refreshPartyMetrics();
  return {ok:true, data:adminData(b.light)};
}

function adminDeleteGuest(b){
  var s = sheet(SH.GUESTS), row = Number(b.row||0);
  if(!row) return {ok:false, error:'No row given.'};
  var cur = s.getRange(row,1,1,GUEST_COLS).getValues()[0];
  var name = (String(cur[G.FIRST]||'')+' '+String(cur[G.LAST]||'')).trim();
  logChange(pad(cur[G.PARTY]), name, 'Guest removed', name, '', 'admin');
  s.deleteRow(row);
  refreshPartyMetrics();
  return {ok:true, data:adminData(b.light)};
}

function firstFreeRow(s, col){
  var n = Math.max(s.getMaxRows() - 1, 1);
  var v = s.getRange(2, col, n, 1).getValues();
  for(var i=0;i<v.length;i++){ if(String(v[i][0]).trim() === '') return i+2; }
  return n + 2;
}

/* ==========================================================================
   ADMIN: parties
   ========================================================================== */

function ensureParty(party){
  var s = sheet(SH.PARTIES), v = rows(SH.PARTIES);
  for(var i=0;i<v.length;i++){ if(pad(v[i][P.PARTY]) === party) return i+2; }

  var used = {}; v.forEach(function(r){ if(r[P.TOKEN]) used[String(r[P.TOKEN])] = true; });
  var t; do { t = newToken(); } while(used[t]);

  var row = firstFreeRow(s, 1);
  s.getRange(row,1).setNumberFormat('@');
  s.getRange(row, P.PARTY+1).setValue(party);
  s.getRange(row, P.NAME+1).setValue('Party ' + party);
  s.getRange(row, P.OPEN+1).insertCheckboxes();
  s.getRange(row, P.OPEN+1).setValue(false);          /* closed until invited */
  s.getRange(row, P.TOKEN+1).setValue(t);
  s.getRange(row, P.LINK+1).setValue(siteUrl() + '/?i=' + t);
  s.getRange(row, P.WAVE+1).setValue(1);
  return row;
}

function partyRowById(party){
  var v = rows(SH.PARTIES);
  for(var i=0;i<v.length;i++){ if(pad(v[i][P.PARTY]) === pad(party)) return i+2; }
  return 0;
}

function adminSaveParty(b){
  var p = b.party || {};
  var id = pad(p.party);
  if(!id) return {ok:false, error:'A party needs a number.'};
  var row = partyRowById(id) || ensureParty(id);
  var s = sheet(SH.PARTIES);
  var cur = s.getRange(row,1,1,PARTY_COLS).getValues()[0];

  if(isTrue(cur[P.OPEN]) !== !!p.open)
    logChange(id, '', 'RSVP open', isTrue(cur[P.OPEN]) ? 'yes' : 'no', p.open ? 'yes' : 'no', 'admin');
  if(ymd(cur[P.DEADLINE]) !== ymd(p.deadline))
    logChange(id, '', 'RSVP deadline', ymd(cur[P.DEADLINE]) || '(none)', ymd(p.deadline) || '(none)', 'admin');

  s.getRange(row, P.NAME+1).setNumberFormat('@').setValue(String(p.name||('Party '+id)));
  s.getRange(row, P.OPEN+1).insertCheckboxes();
  s.getRange(row, P.OPEN+1).setValue(!!p.open);
  s.getRange(row, P.WAVE+1).setValue(p.wave === '' || p.wave === undefined ? '' : Number(p.wave));
  s.getRange(row, P.DEADLINE+1).setValue(ymd(p.deadline));
  s.getRange(row, P.EMAILS+1).setNumberFormat('@').setValue(String(p.emails||''));
  s.getRange(row, P.ADMIN+1).setNumberFormat('@').setValue(String(p.admin||''));
  if(!cur[P.TOKEN]){
    var t = newToken();
    s.getRange(row, P.TOKEN+1).setValue(t);
    s.getRange(row, P.LINK+1).setValue(siteUrl() + '/?i=' + t);
  }
  refreshPartyMetrics();
  return {ok:true, data:adminData(b.light)};
}

/* ==========================================================================
   ONE ROUND TRIP FOR A WHOLE PARTY CARD

   The portal used to save the party and then each guest in turn, which meant
   a separate call to Apps Script per person and a very slow card. This takes
   the lot in one request.
   ========================================================================== */
function adminSavePartyBundle(b){
  var out = {ok:true};
  if(b.party) adminSaveParty({party:b.party, light:true});
  (b.guests || []).forEach(function(g){ adminSaveGuest({guest:g, light:true}); });
  (b.remove || []).forEach(function(row){ adminDeleteGuest({row:row, light:true}); });
  refreshPartyMetrics();
  out.data = adminData(true);
  return out;
}

function adminDeleteParty(b){
  var id = pad(b.party);
  var gs = sheet(SH.GUESTS), gv = rows(SH.GUESTS);
  for(var i=gv.length-1;i>=0;i--){ if(pad(gv[i][G.PARTY]) === id) gs.deleteRow(i+2); }
  var row = partyRowById(id);
  if(row) sheet(SH.PARTIES).deleteRow(row);
  logChange(id, '', 'Party removed', id, '', 'admin');
  refreshPartyMetrics();
  return {ok:true, data:adminData(b.light)};
}

/* Rebuilds the live count formulas on every party row. */
function refreshPartyMetrics(){
  var p = sheet(SH.PARTIES), n = p.getLastRow() - 1;
  if(n < 1) return;
  for(var r = 2; r <= n + 1; r++){
    if(!String(p.getRange(r, P.PARTY+1).getValue()).trim()) continue;
    p.getRange(r, P.INVITED+1)
      .setFormula('=COUNTIF(' + SH.GUESTS + '!$A:$A,TEXT($A' + r + ',"0000"))');
    p.getRange(r, P.REPLIED+1)
      .setFormula('=COUNTIFS(' + SH.GUESTS + '!$A:$A,TEXT($A' + r + ',"0000"),' + SH.GUESTS + '!$F:$F,"<>")');
    p.getRange(r, P.ATTENDING+1)
      .setFormula('=COUNTIFS(' + SH.GUESTS + '!$A:$A,TEXT($A' + r + ',"0000"),' + SH.GUESTS + '!$F:$F,"Yes")');
  }
}

/* ==========================================================================
   ADMIN: content tables (schedule, FAQ, meals) and config
   The portal sends the whole table back and we rewrite it, which keeps
   reordering, adding and deleting rows simple and atomic.
   ========================================================================== */

var TABLES = {
  schedule: { name: SH.SCHEDULE, cols: 8,
              map: function(r,i){ return [i+1, r.start, r.end, r.title, r.place, r.note, r.tags, r.visible !== false]; } },
  faq:      { name: SH.FAQ, cols: 4,
              map: function(r,i){ return [i+1, r.q, r.a, r.visible !== false]; } },
  meals:    { name: SH.MEALS, cols: 3,
              map: function(r,i){ return [i+1, r.meal, r.visible !== false]; } },
  templates:{ name: SH.TEMPLATES, cols: 3,
              map: function(r){ return [r.key, r.subject, r.body]; } }
};

function adminSaveRows(b){
  var t = TABLES[b.table];
  if(!t) return {ok:false, error:'Unknown table.'};
  var s = sheet(t.name);
  var last = s.getMaxRows();
  if(last > 1) s.getRange(2,1,last-1,t.cols).clearContent();

  var list = (b.rows || []).filter(function(r){
    return String(r.title || r.q || r.meal || r.key || '').trim();
  });
  if(list.length){
    s.getRange(2,1,list.length,t.cols).setValues(list.map(t.map));
    if(b.table !== 'templates'){
      var flagCol = t.cols;
      s.getRange(2, flagCol, list.length, 1).insertCheckboxes();
      s.getRange(2, flagCol, list.length, 1).setValues(list.map(function(r){ return [r.visible !== false]; }));
    }
  }
  if(b.table === 'meals') syncMealsToConfig();
  logChange('', '', b.table + ' updated', '', list.length + ' rows', 'admin');
  return {ok:true, data:adminData(b.light)};
}

/* ==========================================================================
   MEALS: the Meals tab and the "Meal Options" setting are the same thing,
   so editing either one rewrites the other. The Meals tab is the richer of
   the two (it has a Show checkbox), so it wins when both could apply.
   ========================================================================== */
function mealList(){
  return rows(SH.MEALS)
    .filter(function(r){ return String(r[1]).trim() && isTrue(r[2]); })
    .sort(function(a,b){ return Number(a[0]||0) - Number(b[0]||0); })
    .map(function(r){ return String(r[1]).trim(); });
}
function syncMealsToConfig(){
  setCfg('Meal Options', mealList().join(', '));
}
function syncConfigToMeals(text){
  var wanted = String(text||'').split(',').map(function(x){ return x.trim(); }).filter(String);
  var have = rows(SH.MEALS).map(function(r){ return String(r[1]).trim(); });
  /* nothing to do if the two already agree */
  var visible = mealList();
  if(visible.length === wanted.length && visible.every(function(m,i){ return m === wanted[i]; })) return;

  var s = sheet(SH.MEALS), last = s.getMaxRows();
  if(last > 1) s.getRange(2,1,last-1,3).clearContent();
  if(wanted.length){
    s.getRange(2,1,wanted.length,3).setValues(wanted.map(function(m,i){ return [i+1, m, true]; }));
    s.getRange(2,3,wanted.length,1).insertCheckboxes();
    s.getRange(2,3,wanted.length,1).setValues(wanted.map(function(){ return [true]; }));
  }
}

function adminSaveConfig(b){
  var meals = null;
  (b.config || []).forEach(function(c){
    var k = String(c.key||'').trim();
    if(!k) return;
    if(k === 'Meal Options'){ meals = c.value; return; }   /* handled below */
    setCfg(k, c.value);
  });
  /* editing Meal Options here rewrites the Meals tab, then we read it back
     so the two can never drift apart */
  if(meals !== null){ syncConfigToMeals(meals); syncMealsToConfig(); }
  logChange('', '', 'Settings updated', '', '', 'admin');
  return {ok:true, data:adminData(b.light)};
}

/* ==========================================================================
   EMAIL
   One wrapper so every message looks like the website. The body accepts
   blank-line paragraphs, {{name}} {{names}} {{link}} {{deadline}}, and
   [[Button label]] which becomes a rose button pointing at that party's
   own invitation link.
   ========================================================================== */

function esc(s){
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function previewParty(partyId){
  var row = partyId ? partyRowById(partyId) : 0;
  if(row) return partyForMail(row);
  return { partyId:'0000', label:'The Sample Family', names:'Alex and Sam',
           link: siteUrl() + '/?i=preview1', deadline:'' };
}

function partyForMail(row){
  var pv = sheet(SH.PARTIES).getRange(row,1,1,PARTY_COLS).getValues()[0];
  var id = pad(pv[P.PARTY]);
  var firsts = rows(SH.GUESTS).filter(function(g){ return pad(g[G.PARTY]) === id; })
                              .map(function(g){ return String(g[G.FIRST]||'').trim(); })
                              .filter(String);
  var names = firsts.length <= 1 ? (firsts[0] || '')
            : firsts.slice(0,-1).join(', ') + ' and ' + firsts[firsts.length-1];
  return {
    row: row, partyId: id,
    label: String(pv[P.NAME] || ('Party '+id)),
    names: names,
    link: String(pv[P.LINK] || (siteUrl() + '/?i=' + pv[P.TOKEN])),
    deadline: prettyDate(pv[P.DEADLINE]),
    emails: String(pv[P.EMAILS]||'').split(/[,;]/).map(function(x){return x.trim();}).filter(String)
  };
}

function fillTokens(text, party){
  return String(text||'')
    .replace(/\{\{\s*name\s*\}\}/gi,     party.label)
    .replace(/\{\{\s*names\s*\}\}/gi,    party.names || party.label)
    .replace(/\{\{\s*link\s*\}\}/gi,     party.link)
    .replace(/\{\{\s*deadline\s*\}\}/gi, party.deadline || 'your earliest convenience');
}

function renderEmail(subject, body, party){
  var filled = fillTokens(body, party);

  var blocks = filled.split(/\n\s*\n/).map(function(chunk){
    var btn = chunk.match(/^\s*\[\[(.+?)\]\]\s*$/);
    if(btn){
      return '<p style="margin:30px 0;text-align:center">'+
        '<a href="'+esc(party.link)+'" style="display:inline-block;background:#9B5F57;color:#ffffff;'+
        'padding:15px 34px;text-decoration:none;letter-spacing:3px;font-size:11px;'+
        'text-transform:uppercase;font-family:Helvetica,Arial,sans-serif">'+esc(btn[1].trim())+'</a></p>';
    }
    return '<p style="margin:0 0 16px;font-size:16px;line-height:1.65;color:#3A2E2B">'+
           esc(chunk.trim()).replace(/\n/g,'<br>')+'</p>';
  }).join('');

  return [
'<div style="margin:0;padding:24px 12px;background:#F4EAE4;font-family:Georgia,\'Times New Roman\',serif">',
'  <div style="max-width:560px;margin:0 auto;background:#FFFCFA;border:1px solid #E3D2B4">',
'    <div style="padding:34px 34px 10px;text-align:center">',
'      <div style="font-family:Georgia,serif;font-size:30px;color:#9B5F57;letter-spacing:1px">',
'        N<span style="color:#C6A97A;font-style:italic"> &amp; </span>E</div>',
'      <div style="margin:16px auto 0;height:1px;width:120px;background:#E3D2B4"></div>',
'    </div>',
'    <div style="padding:14px 34px 30px">', blocks, '</div>',
'    <div style="padding:22px 34px 30px;background:#EFE1DA;text-align:center;',
'                font-family:Helvetica,Arial,sans-serif;font-size:11px;letter-spacing:2px;',
'                text-transform:uppercase;color:#7C6862;line-height:2">',
'      Natalie &amp; Eric &middot; July 11, 2027<br>',
'      Riverway Clubhouse &middot; Burnaby, BC<br>',
'      <a href="mailto:'+esc(sendAsAddress())+'" style="color:#9B5F57;letter-spacing:1px">'+esc(sendAsAddress())+'</a>',
'    </div>',
'  </div>',
'</div>'
  ].join('\n');
}

function plainFallback(body, party){
  return fillTokens(body, party).replace(/\[\[(.+?)\]\]/g, '$1: ' + party.link);
}

/* ==========================================================================
   ADMIN: send email to selected parties
   kind: invite | reminder | custom
   An invite also opens that party's RSVP, which is what makes invite waves
   work: a party stays closed and undiscoverable until its invitation goes out.
   ========================================================================== */

function adminSend(b){
  var ids = (b.partyIds || []).map(pad).filter(String);
  if(!ids.length) return {ok:false, error:'Pick at least one party.'};
  if(!String(b.subject||'').trim()) return {ok:false, error:'The email needs a subject.'};
  if(!String(b.body||'').trim())    return {ok:false, error:'The email needs a body.'};

  var kind = b.kind || 'custom';
  var s = sheet(SH.PARTIES);
  var sent = [], skipped = [], failed = [];

  ids.forEach(function(id){
    var row = partyRowById(id);
    if(!row){ skipped.push(id + ' (no such party)'); return; }
    var party = partyForMail(row);
    if(!party.emails.length){ skipped.push(party.label + ' (no email address)'); return; }

    try{
      var opts = mailOptions();
      opts.htmlBody = renderEmail(b.subject, b.body, party);
      GmailApp.sendEmail(party.emails.join(','),
                         fillTokens(b.subject, party),
                         plainFallback(b.body, party),
                         opts);

      if(kind === 'invite'){
        s.getRange(row, P.SENT+1).setValue(new Date());
        s.getRange(row, P.OPEN+1).insertCheckboxes();
        s.getRange(row, P.OPEN+1).setValue(true);
        logChange(id, '', 'Invitation sent', '', party.emails.join(', '), 'admin');
        logChange(id, '', 'RSVP open', 'no', 'yes', 'admin (invite sent)');
      } else if(kind === 'reminder'){
        s.getRange(row, P.REMINDED+1).setValue(new Date());
        logChange(id, '', 'Reminder sent', '', party.emails.join(', '), 'admin');
      } else {
        logChange(id, '', 'Email sent', String(b.subject).slice(0,80), party.emails.join(', '), 'admin');
      }
      sent.push(party.label);
      Utilities.sleep(300);
    }catch(err){
      failed.push(party.label + ' (' + msg(err) + ')');
    }
  });

  SpreadsheetApp.flush();
  return {ok:true, sent:sent, skipped:skipped, failed:failed, data:adminData(true)};
}

/* ==========================================================================
   SETUP  -  run once, and again after any schema change. Safe to re-run:
   it never destroys guest data and migrates the older Parties layout.
   ========================================================================== */

function setup(){
  var s = ss();
  s.setSpreadsheetTimeZone('America/Vancouver');

  /* ---------- Guests ---------- */
  var g = s.getSheetByName(SH.GUESTS) || s.insertSheet(SH.GUESTS, 0);
  var gHead = ['Party','First Name','Last Name','Email','Phone','Attending','Meal',
               'Dietary Notes','Updated','Guest ID','Admin Notes'];
  if(g.getMaxColumns() < GUEST_COLS) g.insertColumnsAfter(g.getMaxColumns(), GUEST_COLS - g.getMaxColumns());
  g.getRange(1,1,1,GUEST_COLS).setValues([gHead]);
  g.getRange('A2:A').setNumberFormat('@');
  styleHeader(g, GUEST_COLS);
  [70,120,120,220,150,90,150,200,150,150,200].forEach(function(w,i){ g.setColumnWidth(i+1, w); });
  g.setFrozenRows(1);
  g.getRange('F2:F2000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'], true).setAllowInvalid(true).build());

  /* ---------- Parties, with migration from the older column order ---------- */
  var p = s.getSheetByName(SH.PARTIES) || s.insertSheet(SH.PARTIES, 1);
  var pHead = ['Party','Party Name','RSVP Open','Link Code','Invite Link','Wave','RSVP Deadline',
               'Invited','Replied','Attending','Email(s)','Note From Guests','Last Reply',
               'Invite Sent','Reminder Sent','Admin Notes'];
  var oldHead = p.getLastColumn() >= 6 ? String(p.getRange(1,6).getValue()).trim() : '';
  if(oldHead === 'Invited'){
    var old = p.getLastRow() > 1 ? p.getRange(2,1,p.getLastRow()-1,12).getValues() : [];
    var moved = old.filter(function(r){ return String(r[0]).trim(); }).map(function(r){
      return [r[0], r[1], r[2], r[3], r[4], 1, '', '', '', '', r[8], r[9], r[10], r[11], '', ''];
    });
    p.clear();
    if(p.getMaxColumns() < PARTY_COLS) p.insertColumnsAfter(p.getMaxColumns(), PARTY_COLS - p.getMaxColumns());
    p.getRange(1,1,1,PARTY_COLS).setValues([pHead]);
    if(moved.length) p.getRange(2,1,moved.length,PARTY_COLS).setValues(moved);
  } else {
    if(p.getMaxColumns() < PARTY_COLS) p.insertColumnsAfter(p.getMaxColumns(), PARTY_COLS - p.getMaxColumns());
    p.getRange(1,1,1,PARTY_COLS).setValues([pHead]);
  }
  p.getRange('A2:A').setNumberFormat('@');
  styleHeader(p, PARTY_COLS);
  [70,190,95,100,290,60,120,75,75,85,220,240,140,120,130,200]
    .forEach(function(w,i){ p.setColumnWidth(i+1, w); });
  p.setFrozenRows(1);
  var pn = Math.max(p.getLastRow()-1, 1);
  p.getRange(2, P.OPEN+1, pn, 1).insertCheckboxes();
  p.getRange(2, P.OPEN+1, pn, 1).setHorizontalAlignment('center');
  p.getRange(2, P.DEADLINE+1, pn, 1).setNumberFormat('@');

  /* ---------- Config ---------- */
  var c = s.getSheetByName(SH.CONFIG) || s.insertSheet(SH.CONFIG, 2);
  if(c.getLastRow() === 0) c.getRange(1,1,1,3).setValues([['Setting','Value','What it does']]);
  c.getRange(1,1,1,3).setValues([['Setting','Value','What it does']]);
  seedConfig([
    ['RSVP Open','TRUE','Master switch. FALSE closes RSVPs for everyone, whatever the Parties tab says.'],
    ['Site URL','https://natalie-eric.website','Used to build each party invite link.'],
    ['Venue','Riverway Clubhouse','Shown on the schedule and in calendar entries.'],
    ['Venue Address','9001 Bill Fox Way, Burnaby, BC V5J 5J3','Used for the map and calendar entries.'],
    ['Notify Email','natalie.eric.2027@gmail.com','Where RSVP notifications go.'],
    ['Notify On RSVP','TRUE','FALSE stops the notification emails.'],
    ['Send As','natalie.eric.2027@gmail.com','The Gmail alias every email is sent from. Must be a verified alias on this account.']
  ]);
  styleHeader(c, 3);
  [170,330,560].forEach(function(w,i){ c.setColumnWidth(i+1, w); });
  c.setFrozenRows(1);
  c.getRange('C2:C').setWrap(true);

  /* ---------- Schedule ---------- */
  var sc = s.getSheetByName(SH.SCHEDULE) || s.insertSheet(SH.SCHEDULE);
  if(sc.getLastRow() === 0){
    sc.getRange(1,1,1,8).setValues([['Order','Start','End','Title','Location','Note','Tags','Visible']]);
    sc.getRange(2,1,5,8).setValues([
      [1,'11:00','12:00','Chinese Tea Ceremony','Riverway Clubhouse','Family and bridal party only.','Formal attire',true],
      [2,'15:30','17:00','Wedding Ceremony','Riverway Clubhouse','Guests arrive at 3:30pm. The ceremony begins promptly at 4pm.','Formal attire',true],
      [3,'17:00','18:00','Cocktail Hour','Riverway Clubhouse','','Formal attire',true],
      [4,'18:00','21:00','Dinner & Reception','Riverway Clubhouse','','Formal attire',true],
      [5,'21:30','23:59','Dancing','Riverway Clubhouse','','Formal attire',true]
    ]);
  }
  sc.getRange(1,1,1,8).setValues([['Order','Start','End','Title','Location','Note','Tags','Visible']]);
  styleHeader(sc, 8);
  [60,70,70,220,180,380,180,80].forEach(function(w,i){ sc.setColumnWidth(i+1, w); });
  sc.setFrozenRows(1);
  sc.getRange(2,8,Math.max(sc.getLastRow()-1,1),1).insertCheckboxes();
  sc.getRange('B2:C').setNumberFormat('@');

  /* ---------- FAQ ---------- */
  var f = s.getSheetByName(SH.FAQ) || s.insertSheet(SH.FAQ);
  if(f.getLastRow() === 0){
    f.getRange(1,1,1,4).setValues([['Order','Question','Answer','Visible']]);
    f.getRange(2,1,7,4).setValues([
      [1,"What's the RSVP deadline?", 'Please RSVP at your earliest convenience so we can get an accurate headcount.', true],
      [2,'Can I bring a guest?', 'Check your invitation to see if you have a plus one.', true],
      [3,'Are kids welcome?', "We love your little ones. To keep the day intimate, children are welcome only if they're named on your invitation, and you'll find everyone we've saved a seat for listed right there. If you're ever unsure who's included, just reach out and we're happy to help.", true],
      [4,'Where can I park?', 'Riverway Clubhouse has limited parking on site, so we recommend taking an Uber or Lyft.', true],
      [5,'What should I wear?', "We'd love to see you in formal attire, think long dresses for the ladies, and suit and tie for the gentlemen.", true],
      [6,'Is the wedding indoors or outdoors?', 'Our wedding ceremony is outdoors, but our reception will be indoors.', true],
      [7,"I'm travelling from Brazil, do I need a visa for Canada?",
        '<p>Usually an eTA rather than a full visa, as long as you are flying in. Brazilian passport holders can apply for an <b>eTA</b> (about CAD $7, approved online, often within minutes) if you either hold a valid US non-immigrant visa on the day you apply, or have held a Canadian visitor visa in the past 10 years.</p><p>Two things worth knowing:</p><ul><li>The eTA only covers arrival <b>by air</b>. If you are driving or coming by bus, train or boat, you need a visitor visa instead.</li><li>Your US visa needs to be valid on the day you apply for the eTA, but it does not have to still be valid when you travel.</li></ul><p>If neither applies to you, you will need a visitor visa, which takes considerably longer and includes biometrics, so please start early. We are happy to send a letter of invitation if it helps.</p>', true]
    ]);
  }
  f.getRange(1,1,1,4).setValues([['Order','Question','Answer','Visible']]);
  styleHeader(f, 4);
  [60,320,760,80].forEach(function(w,i){ f.setColumnWidth(i+1, w); });
  f.setFrozenRows(1);
  f.getRange(2,4,Math.max(f.getLastRow()-1,1),1).insertCheckboxes();
  f.getRange('C2:C').setWrap(true);

  /* ---------- Meals ---------- */
  var m = s.getSheetByName(SH.MEALS) || s.insertSheet(SH.MEALS);
  if(m.getLastRow() === 0){
    m.getRange(1,1,1,3).setValues([['Order','Meal','Visible']]);
    m.getRange(2,1,3,3).setValues([[1,'Steak & Lobster',true],[2,'Fish',true],[3,'Vegetarian',true]]);
  }
  m.getRange(1,1,1,3).setValues([['Order','Meal','Visible']]);
  styleHeader(m, 3);
  [60,260,80].forEach(function(w,i){ m.setColumnWidth(i+1, w); });
  m.setFrozenRows(1);
  m.getRange(2,3,Math.max(m.getLastRow()-1,1),1).insertCheckboxes();

  /* ---------- Log ---------- */
  var l = s.getSheetByName(SH.LOG) || s.insertSheet(SH.LOG);
  if(l.getLastRow() === 0) l.appendRow(['When','Party','Guest','What changed','From','To','By']);
  l.getRange(1,1,1,7).setValues([['When','Party','Guest','What changed','From','To','By']]);
  styleHeader(l, 7);
  [160,70,190,190,240,240,160].forEach(function(w,i){ l.setColumnWidth(i+1, w); });
  l.setFrozenRows(1);

  /* ---------- Templates ---------- */
  var t = s.getSheetByName(SH.TEMPLATES) || s.insertSheet(SH.TEMPLATES);
  if(t.getLastRow() === 0){
    t.getRange(1,1,1,3).setValues([['Key','Subject','Body']]);
    t.getRange(2,1,3,3).setValues([
      ['invite', "You're invited to Natalie & Eric's wedding",
       "Dear {{names}},\n\nTogether with our families, we would love for you to join us as we get married on Sunday, July 11th 2027 at Riverway Clubhouse in Burnaby, BC.\n\nEverything you need is on our website, and the link below is unique to your party, so you can RSVP for everyone in it.\n\n[[Open Invitation]]\n\nWe cannot wait to celebrate with you.\n\nWith love,\nNatalie & Eric"],
      ['reminder', 'A gentle nudge about our wedding RSVP',
       "Dear {{names}},\n\nWe are starting to firm up numbers for July 11th and noticed we have not heard from you yet. Whenever you have a moment, your RSVP link is below.\n\n[[RSVP Here]]\n\nNo rush at all, and please just reply to this email if anything is unclear.\n\nWith love,\nNatalie & Eric"],
      ['thanks', 'Thank you',
       "Dear {{names}},\n\nThank you so much for celebrating with us. It meant more than we can put in an email.\n\nWith love,\nNatalie & Eric"]
    ]);
  }
  t.getRange(1,1,1,3).setValues([['Key','Subject','Body']]);
  styleHeader(t, 3);
  [120,340,760].forEach(function(w,i){ t.setColumnWidth(i+1, w); });
  t.setFrozenRows(1);
  t.getRange('C2:C').setWrap(true);

  var junk = s.getSheetByName('Sheet1');
  if(junk && s.getSheets().length > 1 && junk.getLastRow() === 0) s.deleteSheet(junk);

  refreshPartyMetrics();
  s.toast('Setup complete. Deploy a new version, then open /admin.','Ready',8);
}

function seedConfig(list){
  var c = sheet(SH.CONFIG);
  var have = {};
  rows(SH.CONFIG).forEach(function(r){ have[String(r[0]).trim().toLowerCase()] = true; });
  list.forEach(function(row){
    if(!have[row[0].toLowerCase()]) c.appendRow(row);
  });
}

function styleHeader(sh, cols){
  sh.getRange(1,1,1,cols)
    .setBackground('#8B534C').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
}

/* Convenience menu inside the spreadsheet. */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('Wedding')
    .addItem('Open the admin portal','openAdmin')
    .addSeparator()
    .addItem('Rebuild counts','refreshPartyMetrics')
    .addItem('Re-run setup','setup')
    .addToUi();
}
function openAdmin(){
  var url = siteUrl() + '/admin/';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<p style="font-family:Arial;font-size:14px">'+
      '<a href="'+url+'" target="_blank">Open the admin portal</a></p>').setHeight(80),
    'Admin');
}
