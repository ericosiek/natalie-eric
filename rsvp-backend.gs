/* ==========================================================================
   NATALIE & ERIC - RSVP BACKEND
   Google Apps Script bound to the "Natalie & Eric - Wedding RSVPs" spreadsheet.

   FIRST TIME SETUP
   1. Run  setup()   once. It builds every tab, header, formula and checkbox.
   2. Deploy > New deployment > Web app
        Execute as:      Me
        Who has access:  Anyone
      Copy the /exec URL and paste it into CONFIG.apiUrl in index.html.
   3. Type your guests into the Guests tab (one row per person, Party column
      holds the party number, e.g. 0037).
   4. Run  refreshParties()  to create a party row + unique link for every
      party number that appears in Guests. Safe to re-run any time; it never
      changes a token that already exists.

   DAY TO DAY
   - To close RSVPs for a party: untick "RSVP Open" on its row in Parties.
     The website reflects it on the guest's next page load, and the server
     refuses any submission for that party even if their page was already open.
   - To reopen or extend: tick it back on.
   - To close everything at once: set "RSVP Open" to FALSE on the Config tab.
   ========================================================================== */

var SHEET_GUESTS  = 'Guests';
var SHEET_PARTIES = 'Parties';
var SHEET_CONFIG  = 'Config';

var G = { PARTY:0, FIRST:1, LAST:2, EMAIL:3, PHONE:4, ATTENDING:5, MEAL:6, DIET:7, UPDATED:8, GID:9 };
var P = { PARTY:0, NAME:1, OPEN:2, TOKEN:3, LINK:4, INVITED:5, REPLIED:6, ATTENDING:7,
          EMAILS:8, NOTE:9, LASTREPLY:10, SENT:11 };

/* ========================== web API ====================================== */

function doGet(e){
  try{
    var a = (e && e.parameter && e.parameter.action) || 'party';
    if(a === 'party') return json(getParty(e.parameter.token));
    if(a === 'find')  return json(findParties(e.parameter.q));
    if(a === 'ping')  return json({ok:true, pong:true, time:new Date().toISOString()});
    return json({ok:false, error:'Unknown action'});
  }catch(err){
    return json({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    if(body.action !== 'rsvp') return json({ok:false, error:'Unknown action'});
    return json(saveRsvp(body));
  }catch(err){
    return json({ok:false, error:String(err && err.message || err)});
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/* ========================== reads ======================================== */

function getParty(token){
  if(!token) return {ok:false, error:'No invitation code was supplied.'};
  var row = findPartyRow(String(token).trim());
  if(!row) return {ok:false, error:"We couldn't find that invitation link."};
  return {ok:true, config:publicConfig(), party:buildParty(row)};
}

function findParties(q){
  if(!q || String(q).trim().length < 2) return {ok:true, matches:[]};
  var needle = String(q).trim().toLowerCase();
  var guests = sheet(SHEET_GUESTS).getDataRange().getValues();
  var ids = {};
  for(var i=1;i<guests.length;i++){
    var last  = String(guests[i][G.LAST]  || '').toLowerCase();
    var first = String(guests[i][G.FIRST] || '').toLowerCase();
    if(!last && !first) continue;
    if(last === needle || first === needle ||
       last.indexOf(needle) === 0 || (first+' '+last).indexOf(needle) > -1){
      ids[pad(guests[i][G.PARTY])] = true;
    }
  }
  var out = [];
  var parties = sheet(SHEET_PARTIES).getDataRange().getValues();
  for(var r=1;r<parties.length;r++){
    var id = pad(parties[r][P.PARTY]);
    if(!ids[id] || !parties[r][P.TOKEN]) continue;
    out.push({ code:String(parties[r][P.TOKEN]), label:String(parties[r][P.NAME] || ('Party '+id)) });
    if(out.length >= 8) break;
  }
  return {ok:true, matches:out};
}

function buildParty(row){
  var pv = sheet(SHEET_PARTIES).getRange(row,1,1,12).getValues()[0];
  var id = pad(pv[P.PARTY]);
  var guests = sheet(SHEET_GUESTS).getDataRange().getValues();
  var list = [];
  for(var i=1;i<guests.length;i++){
    if(pad(guests[i][G.PARTY]) !== id) continue;
    var name = String(guests[i][G.FIRST]||'').trim() + ' ' + String(guests[i][G.LAST]||'').trim();
    if(!name.trim()) continue;
    list.push({
      id:        String(guests[i][G.GID] || ('r'+(i+1))),
      row:       i+1,
      name:      name.trim(),
      attending: normAttend(guests[i][G.ATTENDING]),
      meal:      String(guests[i][G.MEAL] || ''),
      diet:      String(guests[i][G.DIET] || '')
    });
  }
  var globalOpen = cfg('RSVP Open', true) === true || String(cfg('RSVP Open', true)).toUpperCase() === 'TRUE';
  return {
    code:      String(pv[P.TOKEN] || ''),
    partyId:   id,
    label:     String(pv[P.NAME] || ('Party '+id)),
    open:      (pv[P.OPEN] === true || String(pv[P.OPEN]).toUpperCase() === 'TRUE') && globalOpen,
    responded: list.some(function(g){ return g.attending === 'yes' || g.attending === 'no'; }),
    note:      String(pv[P.NOTE] || ''),
    email:     String(pv[P.EMAILS] || '').split(/[,;]/)[0].trim(),
    guests:    list
  };
}

/* ========================== write ======================================== */

function saveRsvp(body){
  var row = findPartyRow(String(body.token||'').trim());
  if(!row) return {ok:false, error:"We couldn't find that invitation link."};

  var party = buildParty(row);
  if(!party.open){
    return {ok:false, locked:true,
            error:'RSVPs for your party are closed. Please email us and we will help.'};
  }

  var gs   = sheet(SHEET_GUESTS);
  var byId = {};
  party.guests.forEach(function(g){ byId[g.id] = g; });

  var stamp = new Date();
  (body.guests || []).forEach(function(sub){
    var target = byId[sub.id];
    if(!target) return;
    var attending = (sub.attending === 'yes') ? 'Yes' : (sub.attending === 'no') ? 'No' : '';
    gs.getRange(target.row, G.ATTENDING+1).setValue(attending);
    gs.getRange(target.row, G.MEAL+1).setValue(attending === 'Yes' ? String(sub.meal||'') : '');
    gs.getRange(target.row, G.DIET+1).setValue(attending === 'Yes' ? String(sub.diet||'') : '');
    gs.getRange(target.row, G.UPDATED+1).setValue(stamp);
  });

  var ps = sheet(SHEET_PARTIES);
  if(body.note)  ps.getRange(row, P.NOTE+1).setValue(String(body.note).slice(0,1000));
  if(body.email) ps.getRange(row, P.EMAILS+1).setValue(String(body.email).slice(0,200));
  ps.getRange(row, P.LASTREPLY+1).setValue(stamp);

  SpreadsheetApp.flush();
  var updated = buildParty(row);
  try{ notify(updated, body); }catch(err){ }
  return {ok:true, party:updated};
}

function notify(party, body){
  if(String(cfg('Notify On RSVP', 'TRUE')).toUpperCase() !== 'TRUE') return;
  var to = String(cfg('Notify Email','') || Session.getEffectiveUser().getEmail());
  if(!to) return;
  var yes = party.guests.filter(function(g){ return g.attending === 'yes'; });
  var lines = party.guests.map(function(g){
    return '  ' + g.name + ': ' +
      (g.attending === 'yes' ? 'Attending (' + (g.meal||'no meal chosen') + ')' +
        (g.diet ? ' [' + g.diet + ']' : '')
       : g.attending === 'no' ? 'Not attending' : 'No reply');
  }).join('\n');
  MailApp.sendEmail({
    to: to,
    subject: 'RSVP: ' + party.label + ' (' + yes.length + ' attending)',
    body: party.label + ' just replied.\n\n' + lines +
          (body.note ? '\n\nTheir note:\n' + body.note : '') +
          (body.email ? '\n\nContact: ' + body.email : '') +
          '\n\nParty ' + party.partyId + '\n' + SpreadsheetApp.getActive().getUrl()
  });
}

/* ========================== helpers ====================================== */

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(n){
  var s = ss().getSheetByName(n);
  if(!s) throw new Error('Missing tab "'+n+'". Run setup() once.');
  return s;
}
function pad(v){
  if(v === '' || v === null || v === undefined) return '';
  var s = String(v).trim();
  return /^\d+$/.test(s) ? ('0000'+s).slice(-4) : s;
}
function normAttend(v){
  var s = String(v||'').trim().toLowerCase();
  if(s === 'yes' || s === 'y' || s === 'true' || s === 'attending') return 'yes';
  if(s === 'no'  || s === 'n' || s === 'false' || s === 'declined') return 'no';
  return '';
}
function findPartyRow(token){
  if(!token) return 0;
  var v = sheet(SHEET_PARTIES).getDataRange().getValues();
  for(var i=1;i<v.length;i++){
    if(String(v[i][P.TOKEN]).trim() === token) return i+1;
  }
  return 0;
}
function cfg(key, dflt){
  var v = sheet(SHEET_CONFIG).getDataRange().getValues();
  for(var i=1;i<v.length;i++){
    if(String(v[i][0]).trim().toLowerCase() === String(key).trim().toLowerCase()){
      return v[i][1] === '' ? dflt : v[i][1];
    }
  }
  return dflt;
}
function publicConfig(){
  var meals = String(cfg('Meal Options','Steak & Lobster, Fish, Vegetarian'))
                .split(',').map(function(s){return s.trim();}).filter(String);
  var dl = cfg('RSVP Deadline','');
  if(dl instanceof Date) dl = Utilities.formatDate(dl, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return { meals:meals, deadline:String(dl||''), rsvpOpen:String(cfg('RSVP Open','TRUE')).toUpperCase()==='TRUE' };
}
function newToken(){
  var abc = 'abcdefghjkmnpqrstuvwxyz23456789';
  var out = '';
  for(var i=0;i<8;i++) out += abc.charAt(Math.floor(Math.random()*abc.length));
  return out;
}

/* ==========================================================================
   SETUP  -  run once from the Apps Script editor
   ========================================================================== */
function setup(){
  var s = ss();
  s.setSpreadsheetTimeZone('America/Vancouver');

  /* ---- Guests ---- */
  var g = s.getSheetByName(SHEET_GUESTS) || s.insertSheet(SHEET_GUESTS, 0);
  if(g.getLastRow() === 0){
    g.getRange(1,1,1,10).setValues([[
      'Party','First Name','Last Name','Email','Phone',
      'Attending','Meal','Dietary Notes','Updated','Guest ID']]);
    g.getRange('A2:A').setNumberFormat('@');
    g.appendRow(['0001','Sample','Guest','','','','','','','g-sample-1']);
    g.appendRow(['0001','Second','Guest','','','','','','','g-sample-2']);
  }
  styleHeader(g, 10);
  g.setColumnWidth(1,70); g.setColumnWidth(2,120); g.setColumnWidth(3,120);
  g.setColumnWidth(4,210); g.setColumnWidth(5,130); g.setColumnWidth(6,90);
  g.setColumnWidth(7,140); g.setColumnWidth(8,200); g.setColumnWidth(9,150); g.setColumnWidth(10,120);
  g.setFrozenRows(1);
  g.getRange('F2:F1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'], true).setAllowInvalid(true).build());
  g.getRange('F2:F1000').setHorizontalAlignment('center');

  /* ---- Parties ---- */
  var p = s.getSheetByName(SHEET_PARTIES) || s.insertSheet(SHEET_PARTIES, 1);
  if(p.getLastRow() === 0){
    p.getRange(1,1,1,12).setValues([[
      'Party','Party Name','RSVP Open','Link Code','Invite Link',
      'Invited','Replied','Attending','Email(s)','Note From Guests','Last Reply','Invite Sent']]);
    p.getRange('A2:A').setNumberFormat('@');
  }
  styleHeader(p, 12);
  p.setColumnWidth(1,70);  p.setColumnWidth(2,190); p.setColumnWidth(3,95);
  p.setColumnWidth(4,100); p.setColumnWidth(5,300); p.setColumnWidth(6,75);
  p.setColumnWidth(7,75);  p.setColumnWidth(8,85);  p.setColumnWidth(9,200);
  p.setColumnWidth(10,260);p.setColumnWidth(11,150);p.setColumnWidth(12,110);
  p.setFrozenRows(1);
  p.getRange('C2:C1000').insertCheckboxes();
  p.getRange('C2:C1000').setHorizontalAlignment('center');

  /* ---- Config ---- */
  var c = s.getSheetByName(SHEET_CONFIG) || s.insertSheet(SHEET_CONFIG, 2);
  if(c.getLastRow() === 0){
    c.getRange(1,1,1,3).setValues([['Setting','Value','What it does']]);
    c.getRange(2,1,6,3).setValues([
      ['RSVP Open','TRUE','Master switch. FALSE closes RSVPs for everyone, whatever the Parties tab says.'],
      ['RSVP Deadline','','Optional. Leave blank for "at your earliest convenience". Otherwise yyyy-mm-dd, e.g. 2027-04-15.'],
      ['Meal Options','Steak & Lobster, Fish, Vegetarian','Comma separated. Changing this changes the dropdown on the website.'],
      ['Site URL','https://natalie-eric.website','Used to build each party invite link.'],
      ['Notify Email','','Where RSVP notifications go. Blank = the account that owns this sheet.'],
      ['Notify On RSVP','TRUE','FALSE stops the notification emails.']
    ]);
  }
  styleHeader(c, 3);
  c.setColumnWidth(1,160); c.setColumnWidth(2,320); c.setColumnWidth(3,560);
  c.setFrozenRows(1);
  c.getRange('C2:C').setWrap(true);

  var junk = s.getSheetByName('Sheet1');
  if(junk && s.getSheets().length > 1 && junk.getLastRow() === 0) s.deleteSheet(junk);

  refreshParties();
  s.toast('Setup complete. Now deploy this script as a web app.','Ready',8);
}

function styleHeader(sh, cols){
  sh.getRange(1,1,1,cols)
    .setBackground('#8B534C').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
}

/* ==========================================================================
   refreshParties()  -  makes a Parties row + unique link for every party
   number that appears in Guests. Re-runnable. Never rewrites a live token.
   ========================================================================== */
function refreshParties(){
  var g  = sheet(SHEET_GUESTS), p = sheet(SHEET_PARTIES);
  var gv = g.getDataRange().getValues();
  var i;

  for(i=1;i<gv.length;i++){
    var hasName = String(gv[i][G.FIRST]||'').trim() || String(gv[i][G.LAST]||'').trim();
    if(hasName && !gv[i][G.GID]){
      var gid = 'g-' + pad(gv[i][G.PARTY]) + '-' + (i+1);
      g.getRange(i+1, G.GID+1).setValue(gid);
      gv[i][G.GID] = gid;
    }
  }

  var order = [], seen = {}, surnames = {};
  for(i=1;i<gv.length;i++){
    var id = pad(gv[i][G.PARTY]);
    if(!id) continue;
    if(!seen[id]){ seen[id] = true; order.push(id); surnames[id] = {}; }
    var ln = String(gv[i][G.LAST]||'').trim();
    if(ln) surnames[id][ln] = true;
  }

  var pv = p.getDataRange().getValues();
  var rowOf = {}, tokens = {};
  for(var r=1;r<pv.length;r++){
    var pid = pad(pv[r][P.PARTY]);
    if(pid) rowOf[pid] = r+1;
    if(pv[r][P.TOKEN]) tokens[String(pv[r][P.TOKEN])] = true;
  }

  var site = String(cfg('Site URL','https://natalie-eric.website')).replace(/\/+$/,'');
  var added = 0;

  order.forEach(function(id){
    var row = rowOf[id];
    if(!row){
      row = p.getLastRow() + 1;
      var names = Object.keys(surnames[id]);
      var label = names.length === 1 ? ('The ' + names[0] + ' Family')
                : names.length  >  1 ? names.join(' & ')
                : ('Party ' + id);
      p.getRange(row,1).setNumberFormat('@');
      p.getRange(row,1).setValue(id);
      p.getRange(row,2).setValue(label);
      p.getRange(row,3).insertCheckboxes();
      p.getRange(row,3).setValue(true);
      added++;
    }
    if(!p.getRange(row, P.TOKEN+1).getValue()){
      var t;
      do { t = newToken(); } while(tokens[t]);
      tokens[t] = true;
      p.getRange(row, P.TOKEN+1).setValue(t);
    }
    var tok = String(p.getRange(row, P.TOKEN+1).getValue());
    p.getRange(row, P.LINK+1).setValue(site + '/?i=' + tok);

    p.getRange(row, P.INVITED+1)
      .setFormula('=COUNTIF(' + SHEET_GUESTS + '!$A:$A,TEXT($A' + row + ',"0000"))');
    p.getRange(row, P.REPLIED+1)
      .setFormula('=COUNTIFS(' + SHEET_GUESTS + '!$A:$A,TEXT($A' + row + ',"0000"),' + SHEET_GUESTS + '!$F:$F,"<>")');
    p.getRange(row, P.ATTENDING+1)
      .setFormula('=COUNTIFS(' + SHEET_GUESTS + '!$A:$A,TEXT($A' + row + ',"0000"),' + SHEET_GUESTS + '!$F:$F,"Yes")');
  });

  SpreadsheetApp.flush();
  ss().toast(added + ' new parties added. Links are in column E.', 'Parties refreshed', 6);
}

/* ==========================================================================
   SHARING  -  ready-to-send text for email, WhatsApp and iMessage
   ========================================================================== */
function buildShareText(){
  var p  = sheet(SHEET_PARTIES);
  var pv = p.getDataRange().getValues();
  var s  = ss();
  var sh = s.getSheetByName('Invite Text') || s.insertSheet('Invite Text');
  sh.clear();
  sh.getRange(1,1,1,5).setValues([['Party','Party Name','Message to send','WhatsApp','iMessage / SMS']]);
  styleHeader(sh, 5);
  sh.setColumnWidth(1,70); sh.setColumnWidth(2,190); sh.setColumnWidth(3,520);
  sh.setColumnWidth(4,300); sh.setColumnWidth(5,300);
  sh.setFrozenRows(1);

  var rows = [];
  for(var r=1;r<pv.length;r++){
    var link = String(pv[r][P.LINK]||'');
    if(!link) continue;
    var msg = "You're invited! Natalie & Eric are getting married on Sunday, July 11, 2027 at "
            + "Riverway Clubhouse in Burnaby, BC.\n\n"
            + "Here is your invitation and RSVP link, it is unique to your party:\n" + link
            + "\n\nWe hope you can join us.";
    rows.push([
      pad(pv[r][P.PARTY]),
      String(pv[r][P.NAME]||''),
      msg,
      'https://wa.me/?text=' + encodeURIComponent(msg),
      'sms:&body=' + encodeURIComponent(msg)
    ]);
  }
  if(rows.length) sh.getRange(2,1,rows.length,5).setValues(rows);
  sh.getRange('C2:C').setWrap(true);
  s.toast('Invite Text tab rebuilt for ' + rows.length + ' parties.','Ready',6);
}

/* ==========================================================================
   EMAIL  -  sends each party its own link. Only parties whose "Invite Sent"
   cell is empty, so it is safe to re-run after adding a second batch.
   ========================================================================== */
function sendInvites(){
  var p  = sheet(SHEET_PARTIES);
  var pv = p.getDataRange().getValues();
  var sent = 0, skipped = 0;

  for(var r=1;r<pv.length;r++){
    var emails = String(pv[r][P.EMAILS]||'').split(/[,;]/).map(function(x){return x.trim();}).filter(String);
    var link   = String(pv[r][P.LINK]||'');
    if(!emails.length || !link || pv[r][P.SENT]){ skipped++; continue; }
    var label = String(pv[r][P.NAME]||'Friends');

    MailApp.sendEmail({
      to: emails.join(','),
      subject: "You're invited to Natalie & Eric's wedding",
      htmlBody:
        '<div style="font-family:Georgia,serif;color:#3A2E2B;max-width:520px;margin:0 auto;padding:28px;'+
        'background:#FCF8F4;border:1px solid #E3D2B4;text-align:center">'+
          '<div style="font-size:13px;letter-spacing:4px;color:#9B5F57;text-transform:uppercase">Together with their families</div>'+
          '<div style="font-size:44px;line-height:1.1;margin:14px 0 6px">Natalie <span style="color:#C6A97A;font-style:italic">&amp;</span> Eric</div>'+
          '<div style="font-size:13px;letter-spacing:3px;color:#7C6862;text-transform:uppercase;line-height:2">'+
            'Sunday, July 11, 2027<br>Riverway Clubhouse &middot; Burnaby, BC</div>'+
          '<p style="font-size:16px;line-height:1.6;margin:24px 0">Dear ' + label + ',<br><br>'+
          'We would love for you to be there. Your invitation and RSVP are at the link below, '+
          'and it is unique to your party.</p>'+
          '<p><a href="' + link + '" style="display:inline-block;background:#9B5F57;color:#fff;'+
          'padding:14px 30px;text-decoration:none;letter-spacing:3px;font-size:12px;'+
          'text-transform:uppercase;font-family:Helvetica,Arial,sans-serif">Open your invitation</a></p>'+
          '<p style="font-size:12px;color:#A2908A;word-break:break-all;margin-top:20px">' + link + '</p>'+
        '</div>',
      name: 'Natalie & Eric'
    });
    p.getRange(r+1, P.SENT+1).setValue(new Date());
    sent++;
    Utilities.sleep(400);
  }
  ss().toast('Sent ' + sent + ', skipped ' + skipped + ' (already sent or no email).','Invites',8);
}

function resendTo(partyNumbers){
  var p = sheet(SHEET_PARTIES), pv = p.getDataRange().getValues();
  var want = {};
  (partyNumbers||[]).forEach(function(n){ want[pad(n)] = true; });
  for(var r=1;r<pv.length;r++){
    if(want[pad(pv[r][P.PARTY])]) p.getRange(r+1, P.SENT+1).clearContent();
  }
}

function openAllRsvps(){  setAllOpen(true);  }
function closeAllRsvps(){ setAllOpen(false); }
function setAllOpen(v){
  var p = sheet(SHEET_PARTIES), n = p.getLastRow()-1;
  if(n > 0) p.getRange(2, P.OPEN+1, n, 1).setValue(v);
  ss().toast('All parties set to ' + (v ? 'OPEN' : 'CLOSED') + '.','RSVPs',5);
}

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Wedding')
    .addItem('Refresh parties & links','refreshParties')
    .addItem('Build invite text (WhatsApp / iMessage)','buildShareText')
    .addSeparator()
    .addItem('Email invites to new parties','sendInvites')
    .addSeparator()
    .addItem('Open RSVPs for everyone','openAllRsvps')
    .addItem('Close RSVPs for everyone','closeAllRsvps')
    .addToUi();
}
