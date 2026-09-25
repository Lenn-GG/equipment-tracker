/* G&G Equipment Requests
 * Sign-in: Microsoft 365 (MSAL). Data: two SharePoint lists read and written through Microsoft Graph.
 * Each list item holds one record: Title = record id, Status = status, Payload = JSON of the record.
 */
(() => {
"use strict";
const CFG = window.GG_CONFIG || {};
const TZ = "America/Los_Angeles";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Sites.ReadWrite.All"];
const SETUP_SCOPES = ["Sites.Manage.All"];

const CATS = { rental:"Rental", owned:"Company tool", material:"Material", pallets:"Blocks & pallets" };
const STATUS = {
  pending:"Waiting approval", approved:"Approved · en route", denied:"Denied",
  received:"On site", released:"Headed to warehouse", returned:"Back in warehouse", consumed:"Used up"
};
const OPEN = ["pending","approved","received","released"];
const STAGES = [
  {key:"open", label:"All open", cls:"all", match:s=>OPEN.includes(s)},
  {key:"pending", label:"Pending", cls:"pend", match:s=>s==="pending"},
  {key:"approved", label:"Approved", cls:"appr", match:s=>s==="approved"},
  {key:"received", label:"On site", cls:"site", match:s=>s==="received"},
  {key:"released", label:"Returning", cls:"trans", match:s=>s==="released"},
  {key:"closed", label:"Closed", cls:"done", match:s=>["returned","denied","consumed"].includes(s)},
];
const CONDITIONS = ["Good","Minor wear","Damaged","Missing parts / short"];

const DEFAULT_TOOLS = `Drills & drivers:
Cordless drill/driver
Impact driver
Hammer drill
Rotary hammer (SDS-Plus)
Rotary hammer (SDS-Max)
Right-angle drill
Magnetic drill press
Cutting:
Circular saw
Reciprocating saw
Jigsaw
Portable band saw
Miter / chop saw
Angle grinder 4-1/2"
Angle grinder 7"
Concrete cut-off saw
Oscillating multi-tool
Door & hardware:
Lockset boring jig kit
Hinge mortise template + router
Mortise lock jig
Continuous hinge drilling jig
Closer / exit device template kit
Power door planer
Hole saw kit
Door jack / door stand
Frame spreaders
Fastening:
Powder-actuated tool (Hilti DX)
Pop rivet gun
Rivet nut (nutsert) tool
Finish nailer
Air compressor
Battery caulk gun
Grout pump (frame grouting)
Layout:
Laser level
Rotary laser
6' level
Concrete scanner
Access & handling:
6' step ladder
8' step ladder
10' step ladder
12' step ladder
Extension ladder
Baker scaffold
Door / panel cart
Pallet jack
Glass suction cups
Power & site:
Generator
100' extension cord
Spider box
LED work lights
HEPA dust extractor
Shop vac
MIG welder
Stick welder + leads`;
const DEFAULT_RETURNS = `Preinstall:
Wood pallets
Wood blocking / dunnage
Frame spreaders
Frame shipping crates
Door shipping crates
Strapping / banding (for recycle)`;

const S = {
  pca:null, account:null, me:null, siteId:null, ready:false,
  isSuper:false, isWarehouse:false, isAdmin:false,
  docs:[], stage:"open", q:"", mine:false, loaded:false, syncedAt:0,
  catalog:{ tools:DEFAULT_TOOLS, returns:DEFAULT_RETURNS }, roles:{ superintendents:[], warehouse:[] },
  configItems:{}, // key -> SharePoint item id
};

const $ = id => document.getElementById(id);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmtTime = ts => ts ? new Date(ts).toLocaleString("en-US",{timeZone:TZ,month:"short",day:"numeric",hour:"numeric",minute:"2-digit",timeZoneName:"short"}) : "";
const fmtDate = d => { if(!d) return ""; const [y,m,dd]=d.split("-").map(Number); return new Date(Date.UTC(y,m-1,dd,12)).toLocaleDateString("en-US",{timeZone:"UTC",weekday:"short",month:"short",day:"numeric"}); };
const todayLA = () => new Date().toLocaleDateString("en-CA",{timeZone:TZ});
const daysBetween = (a,b) => Math.max(0, Math.floor((b-a)/86400000));
const lc = s => String(s||"").trim().toLowerCase();
const emails = text => [...new Set(String(text||"").split(/[\s,;]+/).map(lc).filter(e => e.includes("@")))];

// ---------- Microsoft Graph ----------
async function token(scopes = SCOPES){
  try { return (await S.pca.acquireTokenSilent({ account:S.account, scopes })).accessToken; }
  catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || e.errorCode === "interaction_required" || e.errorCode === "consent_required") {
      await S.pca.acquireTokenRedirect({ account:S.account, scopes });
    }
    throw e;
  }
}
async function g(path, { method="GET", body, scopes } = {}){
  const t = await token(scopes);
  const r = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    method, headers:{ Authorization:"Bearer " + t, "Content-Type":"application/json" },
    body: body ? JSON.stringify(body) : undefined
  });
  if (r.status === 204) return null;
  const j = await r.json().catch(() => null);
  if (!r.ok) { const e = new Error(j?.error?.message || r.statusText); e.status = r.status; e.graph = j?.error; throw e; }
  return j;
}
const listPath = name => `/sites/${S.siteId}/lists/${encodeURIComponent(name)}`;

async function resolveSite(){
  const u = new URL(CFG.sharePointSite);
  const path = u.pathname.replace(/\/+$/,"");
  const site = await g(path ? `/sites/${u.hostname}:${path}` : `/sites/${u.hostname}`);
  S.siteId = site.id;
}
async function listExists(name){
  try { await g(listPath(name)); return true; }
  catch (e) { if (e.status === 404) return false; throw e; }
}
async function allItems(name){
  let url = `${listPath(name)}/items?expand=fields(select=Title,Status,Payload)&$top=999`;
  const out = [];
  while (url) { const j = await g(url); out.push(...(j.value||[])); url = j["@odata.nextLink"] || null; }
  return out;
}
function parseItem(it){
  let data = {};
  try { data = JSON.parse(it.fields?.Payload || "{}"); } catch { data = {}; }
  return { ...data, id: it.fields?.Title || data.id, status: it.fields?.Status || data.status, _item: it.id };
}

const Store = {
  async load(){
    const [reqs, cfg] = await Promise.all([allItems(CFG.requestsList), allItems(CFG.configList)]);
    S.docs = reqs.map(parseItem).filter(d => d.id);
    S.configItems = {};
    cfg.forEach(it => {
      const key = it.fields?.Title; let v = {};
      try { v = JSON.parse(it.fields?.Payload || "{}"); } catch {}
      S.configItems[key] = it.id;
      if (key === "catalog") S.catalog = { tools: v.tools ?? DEFAULT_TOOLS, returns: v.returns ?? DEFAULT_RETURNS };
      if (key === "roles") S.roles = { superintendents: v.superintendents || [], warehouse: v.warehouse || [] };
    });
    computeRoles();
    S.loaded = true; S.syncedAt = Date.now();
  },
  async create(id, data){
    const payload = { ...data, id };
    await g(`${listPath(CFG.requestsList)}/items`, { method:"POST", body:{ fields:{ Title:id, Status:data.status, Payload:JSON.stringify(payload) } } });
  },
  async update(doc, patch){
    // Re-read the latest copy so we don't overwrite someone else's step.
    const it = await g(`${listPath(CFG.requestsList)}/items/${doc._item}?expand=fields(select=Title,Status,Payload)`);
    const cur = parseItem(it); delete cur._item;
    const next = { ...cur, ...patch };
    await g(`${listPath(CFG.requestsList)}/items/${doc._item}/fields`, { method:"PATCH", body:{ Status:next.status, Payload:JSON.stringify(next) } });
  },
  async remove(doc){ await g(`${listPath(CFG.requestsList)}/items/${doc._item}`, { method:"DELETE" }); },
  async setConfig(key, value){
    const fields = { Title:key, Payload:JSON.stringify(value) };
    if (S.configItems[key]) await g(`${listPath(CFG.configList)}/items/${S.configItems[key]}/fields`, { method:"PATCH", body:fields });
    else await g(`${listPath(CFG.configList)}/items`, { method:"POST", body:{ fields } });
  }
};

async function createLists(){
  const mk = name => g(`/sites/${S.siteId}/lists`, { method:"POST", scopes:SETUP_SCOPES, body:{
    displayName:name, list:{ template:"genericList" },
    columns:[ { name:"Status", text:{} }, { name:"Payload", text:{ allowMultipleLines:true, textType:"plain", linesForEditing:6 } } ]
  }});
  if (!(await listExists(CFG.requestsList))) await mk(CFG.requestsList);
  if (!(await listExists(CFG.configList))) await mk(CFG.configList);
  await Store.load();
  if (!S.configItems.catalog) await Store.setConfig("catalog", { tools:DEFAULT_TOOLS, returns:DEFAULT_RETURNS });
  if (!S.configItems.roles) await Store.setConfig("roles", { superintendents:[], warehouse:[] });
}

function computeRoles(){
  const me = S.me?.email;
  S.isAdmin = (CFG.admins||[]).map(lc).includes(me);
  S.isSuper = S.isAdmin || S.roles.superintendents.map(lc).includes(me);
  S.isWarehouse = S.roles.warehouse.map(lc).includes(me);
}
const by = () => ({ by:S.me.email, byName:S.me.name });

// ---------- lists (tool picker) ----------
function parseList(text){
  const groups = []; let cur = null;
  String(text||"").split("\n").map(l=>l.trim()).filter(Boolean).forEach(l => {
    if (l.endsWith(":")) { cur = { name:l.slice(0,-1).trim(), items:[] }; groups.push(cur); }
    else { if (!cur) { cur = { name:"", items:[] }; groups.push(cur); } cur.items.push(l); }
  });
  return groups.filter(g => g.items.length);
}
function optionsHTML(text){
  return parseList(text).map(g => g.name
    ? `<optgroup label="${esc(g.name)}">${g.items.map(i=>`<option>${esc(i)}</option>`).join("")}</optgroup>`
    : g.items.map(i=>`<option>${esc(i)}</option>`).join("")).join("");
}

const who = (o, fallback="") => (o && (o.byName || o.by)) || fallback;
function isMine(d){ return S.me && lc(d.requestedBy) === S.me.email; }

// ---------- rendering ----------
function renderWho(){
  const el = $("who");
  if (!S.me) { el.innerHTML = ""; $("setBtn").hidden = true; $("outBtn").hidden = true; return; }
  const role = S.isSuper ? "Superintendent" : S.isWarehouse ? "Warehouse" : "Crew";
  el.innerHTML = `<span>${esc(S.me.name || S.me.email)}</span><span class="role">${role}</span>`;
  $("setBtn").hidden = !(S.ready && (S.isSuper || S.isWarehouse));
  $("outBtn").hidden = false;
}
function renderPipeline(){
  $("pipeline").hidden = !S.ready;
  $("pipeline").innerHTML = STAGES.map(st => {
    const n = S.docs.filter(d => st.match(d.status) && (!S.mine || isMine(d))).length;
    return `<button class="stage ${st.cls}" data-stage="${st.key}" aria-pressed="${S.stage===st.key}"><span class="n">${n}</span><span class="l">${st.label}</span></button>`;
  }).join("");
}
function stepHTML(label, cls, whoName, when, extra){
  return `<li class="step ${cls}"><div class="sl">${label}</div>${whoName?`<div class="sw">${esc(whoName)}</div>`:""}${when?`<div class="sx">${esc(when)}</div>`:""}${extra||""}</li>`;
}
function condHTML(c){ if(!c) return ""; const bad = c!=="Good" && c!=="Minor wear"; return `<div class="sx">Condition: <span class="cond ${bad?"bad":""}">${esc(c)}</span></div>`; }
const sx = (label, v) => (v!=null && v!=="") ? `<div class="sx">${label}${esc(v)}</div>` : "";

function cardHTML(d){
  const st = d.status;
  const late = st==="approved" && d.neededBy && d.neededBy < todayLA();
  const dec = d.decision || {}, rec = d.received || {}, rel = d.released || {}, ret = d.returned || {};
  const dash = st==="denied" ? "—" : "Not yet";
  const reqName = d.requestedByName || d.requestedBy || "Someone";
  let onRent = "";
  if (d.category==="rental" && rec.at) { const days = daysBetween(rec.at, ret.at || rel.at || Date.now()); onRent = `<span>On rent <b>${days} day${days===1?"":"s"}</b></span>`; }

  const whStep = st==="consumed"
    ? stepHTML("Warehouse","","","No leftover")
    : stepHTML("Warehouse", ret.at?"on":"", ret.at?who(ret):"", ret.at?fmtTime(ret.at):(st==="released"?"In transit":dash),
        condHTML(ret.condition)+sx("Qty back: ",ret.qty)+sx("Off-rent #: ",ret.offRent)+sx("",ret.notes));
  const relExtra = condHTML(rel.condition)+sx(d.category==="rental"?"Off-rent #: ":"Leftover: ", rel.ref)+sx("",rel.notes);

  let steps, n;
  if (d.returnOnly) {
    steps = stepHTML("Sent from site","on",reqName,fmtTime(d.requestedAt),condHTML(rel.condition)) + whStep; n = 2;
  } else {
    steps = [
      stepHTML("Requested","on",reqName,fmtTime(d.requestedAt)),
      st==="denied"
        ? stepHTML("Denied","no",who(dec),fmtTime(dec.at), dec.note?`<div class="sx">“${esc(dec.note)}”</div>`:"")
        : stepHTML("Approved", dec.at?"on":"", dec.at?who(dec):"Superintendent", dec.at?fmtTime(dec.at):"Waiting", (dec.note?`<div class="sx">“${esc(dec.note)}”</div>`:"")+sx("PO / contract: ",dec.po)),
      stepHTML("Received", rec.at?"on":"", rec.at?who(rec):"", rec.at?fmtTime(rec.at):dash, condHTML(rec.condition)+(rec.qty!=null&&rec.qty!==""&&Number(rec.qty)!==Number(d.qty)?sx("Qty received: ",rec.qty):"")+sx("",rec.notes)),
      stepHTML("Released", rel.at?"on":"", rel.at?who(rel):"", rel.at?fmtTime(rel.at):dash, relExtra),
      whStep
    ].join(""); n = 5;
  }

  const acts = [];
  if (S.isSuper) acts.push(`<button class="btn ghost grow" data-act="delete">Delete</button>`);
  if (st==="pending" && S.isSuper) acts.push(`<button class="btn bad" data-act="deny">Deny</button><button class="btn primary" data-act="approve">Approve</button>`);
  if (st==="pending" && !S.isSuper) acts.push(`<span class="hint">Waiting on the superintendent</span>`);
  if (st==="approved") acts.push(`<button class="btn good" data-act="receive">Mark received on site</button>`);
  if (st==="received") acts.push(`<button class="btn primary" data-act="release">Release from site</button>`);
  if (st==="released") acts.push((S.isWarehouse || S.isSuper) ? `<button class="btn good" data-act="checkin">Check in at warehouse</button>` : `<span class="hint">Waiting on warehouse check-in</span>`);

  const qtyUnit = `${esc(d.qty)}${d.unit?" "+esc(d.unit):""}`;
  return `<article class="card" data-s="${esc(st)}" data-id="${esc(d.id)}">
    <div class="row1"><span class="rid">${esc(d.id)}</span><span class="chip">${esc(CATS[d.category]||"Item")}</span>${d.returnOnly?'<span class="chip">Return</span>':""}${late?'<span class="chip late">Past need-by</span>':""}<span class="pill ${esc(st)}">${esc(STATUS[st]||st)}</span></div>
    <h3 class="title"><span class="q">${qtyUnit} ×</span> ${esc(d.item)}</h3>
    <div class="meta"><span>Job <b>${esc(d.jobSite)}</b></span>${d.neededBy?`<span>Needed <b>${esc(fmtDate(d.neededBy))}</b></span>`:""}${d.vendor?`<span>Vendor <b>${esc(d.vendor)}</b></span>`:""}${onRent}</div>
    ${d.notes?`<div class="note">${esc(d.notes)}</div>`:""}
    <ol class="steps" style="--n:${n}">${steps}</ol>
    ${acts.length?`<div class="actions">${acts.join("")}</div>`:""}
  </article>`;
}

function gate(html){ $("gate").innerHTML = html ? `<div class="gate">${html}</div>` : ""; }
const msBtn = label => `<button class="btn primary ms" data-act="signin"><i aria-hidden="true"><b style="background:#F25022"></b><b style="background:#7FBA00"></b><b style="background:#00A4EF"></b><b style="background:#FFB900"></b></i>${esc(label)}</button>`;

function render(){
  renderWho(); renderPipeline();
  const tools = document.querySelector(".tools");
  tools.hidden = !S.ready;
  $("newBtn").hidden = !S.ready;
  $("banner").innerHTML = (S.ready && S.isSuper && !S.roles.warehouse.length)
    ? `<div class="banner"><span>No warehouse manager is set yet, so only superintendents can check items back in.</span><button class="btn" data-act="settings">Add warehouse manager</button></div>` : "";
  const list = $("list");
  if (!S.ready) { list.innerHTML = ""; return; }
  gate("");
  const stage = STAGES.find(s => s.key===S.stage);
  const q = S.q.trim().toLowerCase();
  const rows = S.docs.filter(d => stage.match(d.status) && (!S.mine || isMine(d)) &&
    (!q || [d.id,d.item,d.jobSite,d.vendor,d.notes,d.requestedByName].some(v => String(v||"").toLowerCase().includes(q))));
  const sync = S.syncedAt ? `<p class="sync">Updated ${esc(fmtTime(S.syncedAt))}</p>` : "";
  if (!rows.length) {
    list.innerHTML = `<div class="empty"><h3>${S.docs.length ? "Nothing here" : "No requests yet"}</h3><p>${S.docs.length ? "Try another status or clear the search." : "Tap “+ New” to request a lift, a tool or materials, or to send blocks and pallets back."}</p></div>${sync}`;
    return;
  }
  const mineFirst = d => (S.isSuper && d.status==="pending") || ((S.isWarehouse||S.isSuper) && d.status==="released") ? 0 : 1;
  rows.sort((a,b) => (mineFirst(a)-mineFirst(b)) || (b.requestedAt||0)-(a.requestedAt||0));
  list.innerHTML = rows.map(cardHTML).join("") + sync;
}

// ---------- sheets ----------
function closeSheet(){ $("sheetRoot").innerHTML = ""; }
function toast(msg){ const r=$("toastRoot"); r.innerHTML=`<div class="toast" role="status">${esc(msg)}</div>`; clearTimeout(toast.t); toast.t=setTimeout(()=>r.innerHTML="",2600); }
function seg(name, opts, checked){ return `<div class="seg" role="radiogroup">${opts.map(([v,l])=>`<label><input type="radio" name="${name}" value="${esc(v)}" ${v===checked?"checked":""}><span>${esc(l)}</span></label>`).join("")}</div>`; }
function sheet(title, sub, body, onSubmit, submitLabel, submitCls="primary"){
  $("sheetRoot").innerHTML = `<div class="scrim" id="scrim"><div class="sheet" role="dialog" aria-modal="true" aria-labelledby="shT">
    <h2 id="shT">${esc(title)}</h2><p class="sub">${sub}</p>
    <form class="f" id="shF" novalidate>${body}<div class="err" id="shErr" role="alert"></div>
    <div class="foot"><button type="button" class="btn" id="shCancel">Cancel</button><button type="submit" class="btn ${submitCls}" id="shOk">${esc(submitLabel)}</button></div></form></div></div>`;
  $("scrim").addEventListener("click", e => { if (e.target.id==="scrim") closeSheet(); });
  $("shCancel").onclick = closeSheet;
  const f = $("shF");
  f.addEventListener("submit", async e => {
    e.preventDefault(); const ok=$("shOk"); ok.disabled=true; $("shErr").textContent="";
    try {
      const msg = await onSubmit(new FormData(f));
      if (msg !== false) { closeSheet(); if (msg) toast(msg); await refresh(); }
    } catch(err){ console.error(err); $("shErr").textContent = errText(err); }
    finally { if (document.body.contains(ok)) ok.disabled=false; }
  });
  const first = f.querySelector("input:not([type=radio]):not([type=checkbox]),textarea"); first && setTimeout(()=>first.focus(),30);
}
function errText(err){
  if (err && err.userMsg) return err.userMsg;
  if (err && (err.status===403 || err.status===401)) return "SharePoint refused the change. Ask Lenn to give you Edit access to the site.";
  if (err && err.status===404) return "That request was removed by someone else. Refresh the page.";
  return "Couldn't save. Check your connection and try again.";
}
const fail = msg => { const e = new Error(msg); e.userMsg = msg; throw e; };
const need = (v, msg) => String(v||"").trim() || fail(msg);
const txt = v => String(v||"").trim();

function newId(prefix="EQ"){
  const d = todayLA().slice(2).replace(/-/g,"");
  const r = Math.random().toString(36).slice(2,5).toUpperCase();
  return `${prefix}-${d}-${r}`;
}

function openNew(mode="request"){
  const sites = [...new Set(S.docs.map(d=>d.jobSite).filter(Boolean))].slice(0,40);
  const siteFld = `<div class="fld"><label for="f_site">Job site</label><input id="f_site" name="jobSite" list="siteList" placeholder="e.g. Kaiser Oakland – Bldg C" autocomplete="off"><datalist id="siteList">${sites.map(s=>`<option value="${esc(s)}">`).join("")}</datalist></div>`;
  const modeFld = `<div class="mode">${seg("mode",[["request","Request equipment"],["return","Send back to warehouse"]],mode)}</div>`;
  const toolOpts = optionsHTML(S.catalog.tools), returnOpts = optionsHTML(S.catalog.returns);
  let body, sub, label;
  if (mode==="request") {
    sub = "Your superintendent gets it for approval."; label = "Send request";
    body = modeFld + siteFld + `
      <div class="fld"><span class="lab">Type</span>${seg("category",[["owned","Company tool"],["rental","Rental"],["material","Material"]],"owned")}</div>
      <div class="fld" id="toolFld"><label for="f_tool">Tool</label><select id="f_tool" name="tool"><option value="">Choose a tool…</option>${toolOpts}<option value="__other">Other (type it in)</option></select></div>
      <div class="fld" id="itemFld" hidden><label for="f_item" id="itemLab">What do you need?</label><input id="f_item" name="item" placeholder="e.g. 19' scissor lift, Tapcons 1/4×2-3/4"></div>
      <div class="two"><div class="fld"><label for="f_qty">Qty</label><input id="f_qty" name="qty" type="number" min="1" step="1" value="1" inputmode="numeric"></div>
      <div class="fld"><label for="f_unit">Unit (optional)</label><input id="f_unit" name="unit" placeholder="ea, box, pcs"></div></div>
      <div class="two"><div class="fld"><label for="f_need">Needed by</label><input id="f_need" name="neededBy" type="date" value="${todayLA()}"></div>
      <div class="fld" id="vendorFld" style="visibility:hidden"><label for="f_vendor">Vendor (rentals)</label><input id="f_vendor" name="vendor" placeholder="e.g. Sunbelt, United"></div></div>
      <div class="fld"><label for="f_notes">Notes</label><textarea id="f_notes" name="notes" placeholder="Where to drop it, gate code, contact on site…"></textarea></div>`;
  } else {
    sub = "For blocks, pallets and other preinstall material going back. The warehouse checks it in."; label = "Send to warehouse";
    body = modeFld + siteFld + `
      <div class="fld"><label for="f_ret">What's going back?</label><select id="f_ret" name="retItem"><option value="">Choose…</option>${returnOpts}<option value="__other">Other (type it in)</option></select></div>
      <div class="fld" id="retOtherFld" hidden><label for="f_reto">Describe it</label><input id="f_reto" name="retOther" placeholder="e.g. Frame shipping crate"></div>
      <div class="two"><div class="fld"><label for="f_rqty">Qty</label><input id="f_rqty" name="qty" type="number" min="1" step="1" value="1" inputmode="numeric"></div>
      <div class="fld"><label for="f_runit">Unit (optional)</label><input id="f_runit" name="unit" placeholder="pallets, bundles, pcs"></div></div>
      <div class="fld"><span class="lab">Condition</span>${seg("condition",[["Good","Reusable"],["Damaged","Damaged / scrap"]],"Good")}</div>
      <div class="fld"><label for="f_rnotes">Notes</label><textarea id="f_rnotes" name="notes" placeholder="Left at loading dock, on the 2pm truck…"></textarea></div>`;
  }
  sheet(mode==="request"?"New request":"Send back to warehouse", sub, body, async fd => {
    const qty = Number(fd.get("qty"));
    if (!(qty>0)) fail("Enter a quantity of 1 or more.");
    const jobSite = need(fd.get("jobSite"),"Enter the job site.");
    const base = { jobSite, qty, unit:txt(fd.get("unit")), notes:txt(fd.get("notes")), requestedBy:S.me.email, requestedByName:S.me.name, requestedAt:Date.now() };
    if (mode==="request") {
      const category = fd.get("category")||"owned";
      let item = txt(fd.get("item"));
      if (category==="owned") { const t = fd.get("tool"); if (t && t!=="__other") item = t; }
      if (!item) fail(category==="owned" ? "Choose a tool, or pick Other and type it in." : "Say what you need.");
      const id = newId("EQ");
      await Store.create(id, { ...base, item, category, neededBy:fd.get("neededBy")||"", vendor: category==="rental" ? txt(fd.get("vendor")) : "", status:"pending" });
      S.stage = "open"; return `Request ${id} sent for approval`;
    } else {
      let item = fd.get("retItem"); if (!item || item==="__other") item = txt(fd.get("retOther"));
      if (!item) fail("Choose what's going back, or pick Other and describe it.");
      const id = newId("RT");
      await Store.create(id, { ...base, item, category:"pallets", returnOnly:true, status:"released",
        released:{ ...by(), at:base.requestedAt, condition:fd.get("condition")||"Good", ref:"", notes:"" } });
      S.stage = "released"; return `${id} logged. The warehouse will check it in.`;
    }
  }, label);

  document.querySelectorAll('input[name=mode]').forEach(r => r.addEventListener("change", () => { if (r.checked) openNew(r.value); }));
  if (mode==="request") {
    const sync = () => {
      const cat = document.querySelector('input[name=category]:checked').value;
      const tool = $("f_tool").value;
      $("toolFld").hidden = cat!=="owned";
      $("itemFld").hidden = cat==="owned" && tool!=="__other";
      $("itemLab").textContent = cat==="owned" ? "Tool name" : "What do you need?";
      $("vendorFld").style.visibility = cat==="rental" ? "visible" : "hidden";
    };
    document.querySelectorAll('input[name=category]').forEach(r => r.addEventListener("change", sync));
    $("f_tool").addEventListener("change", sync); sync();
  } else {
    $("f_ret").addEventListener("change", e => { $("retOtherFld").hidden = e.target.value!=="__other"; });
  }
}

const subOf = d => `${esc(d.id)} · ${esc(d.qty)}${d.unit?" "+esc(d.unit):""} × ${esc(d.item)} · ${esc(d.jobSite)}`;

function openDecide(d, approve){
  const body = `${approve && d.category==="rental" ? `<div class="fld"><label for="f_po">PO / rental contract # (optional)</label><input id="f_po" name="po"></div>` : ""}
    <div class="fld"><label for="f_dn">${approve?"Note to crew (optional)":"Reason"}</label><textarea id="f_dn" name="note" placeholder="${approve?"e.g. Delivery Tue AM, call Mike on arrival":"e.g. Use the lift already on the Hayward job"}"></textarea></div>`;
  sheet(approve?"Approve request":"Deny request", subOf(d), body, async fd => {
    const note = txt(fd.get("note"));
    if (!approve && !note) fail("Give the crew a reason.");
    await Store.update(d, { status: approve?"approved":"denied", decision:{ ...by(), at:Date.now(), note, po:txt(fd.get("po")) } });
    return approve ? `${d.id} approved` : `${d.id} denied`;
  }, approve?"Approve":"Deny", approve?"primary":"bad");
}

function openReceive(d){
  const body = `<div class="fld"><span class="lab">Condition on arrival</span>${seg("condition",CONDITIONS.map(c=>[c,c]),"Good")}</div>
    <div class="fld"><label for="f_rq">Qty received</label><input id="f_rq" name="qty" type="number" min="0" step="1" value="${esc(d.qty)}" inputmode="numeric"></div>
    <div class="fld"><label for="f_rn">Notes</label><textarea id="f_rn" name="notes" placeholder="Scratches, missing charger, delivery ticket #…"></textarea></div>`;
  sheet("Receive on site", subOf(d), body, async fd => {
    const condition = fd.get("condition")||"Good", notes = txt(fd.get("notes"));
    if (condition!=="Good" && !notes) fail("Describe the wear, damage or what's missing.");
    await Store.update(d, { status:"received", received:{ ...by(), at:Date.now(), condition, qty:fd.get("qty"), notes } });
    return `${d.id} received on site`;
  }, "Mark received", "good");
}

function openRelease(d){
  const isMat = d.category==="material", isRent = d.category==="rental";
  const body = `<div class="fld"><span class="lab">Condition leaving site</span>${seg("condition",CONDITIONS.map(c=>[c,c]),"Good")}</div>
    ${isMat ? `<label class="check"><input type="checkbox" name="allUsed" id="f_used"> All used on the job. Nothing goes back to the warehouse.</label>
      <div class="fld" id="leftFld"><label for="f_ref">Leftover going back</label><input id="f_ref" name="ref" placeholder="e.g. 2 boxes"></div>` : ""}
    ${isRent ? `<div class="fld"><label for="f_ref">Off-rent / pickup # (if already called off)</label><input id="f_ref" name="ref" placeholder="From the vendor's call-off"></div>` : ""}
    <div class="fld"><label for="f_ln">Notes</label><textarea id="f_ln" name="notes" placeholder="On the 2pm truck, staged at loading dock…"></textarea></div>
    <span class="hint" id="relHint">${isMat?"Leftovers":"It"} will show as headed to the warehouse until the warehouse manager checks it in.</span>`;
  sheet("Release from site", subOf(d), body, async fd => {
    const condition = fd.get("condition")||"Good", notes = txt(fd.get("notes"));
    const used = isMat && fd.get("allUsed");
    if (!used && condition!=="Good" && condition!=="Minor wear" && !notes) fail("Describe the damage or what's missing.");
    await Store.update(d, { status: used?"consumed":"released",
      released:{ ...by(), at:Date.now(), condition: used?"":condition, ref: used?"":txt(fd.get("ref")), notes } });
    S.stage = used ? "closed" : "released";
    return used ? `${d.id} closed out` : `${d.id} released. Headed to warehouse.`;
  }, "Release");
  if (isMat) $("f_used").addEventListener("change", e => { $("leftFld").hidden = e.target.checked; $("relHint").hidden = e.target.checked; });
}

function openCheckin(d){
  const qtyBack = d.category==="material" ? (d.released?.ref || "") : d.qty;
  const body = `<div class="fld"><span class="lab">Condition checked in</span>${seg("condition",CONDITIONS.map(c=>[c,c]),"Good")}</div>
    <div class="fld"><label for="f_cq">Qty received back</label><input id="f_cq" name="qty" value="${esc(qtyBack)}"></div>
    ${d.category==="rental" ? `<div class="fld"><label for="f_or">Off-rent / pickup #</label><input id="f_or" name="offRent" value="${esc(d.released?.ref||"")}"></div>` : ""}
    <div class="fld"><label for="f_cn">Notes</label><textarea id="f_cn" name="notes" placeholder="Put away in cage B, sent out for repair, short 1 bit…"></textarea></div>`;
  sheet("Check in at warehouse", subOf(d), body, async fd => {
    const condition = fd.get("condition")||"Good", notes = txt(fd.get("notes"));
    if (condition!=="Good" && !notes) fail("Describe the wear, damage or what's missing.");
    await Store.update(d, { status:"returned", returned:{ ...by(), at:Date.now(), condition, qty:txt(fd.get("qty")), offRent:txt(fd.get("offRent")), notes } });
    return `${d.id} checked in at warehouse`;
  }, "Check in", "good");
}

function openDelete(d){
  sheet("Delete request?", `${subOf(d)}. This removes it and its history for everyone.`, "", async () => {
    await Store.remove(d); return `${d.id} deleted`;
  }, "Delete", "bad");
}

function openSettings(){
  const body = `
    ${S.isSuper ? `
    <div class="fld"><label for="f_sup">Superintendents</label><textarea id="f_sup" name="sup" placeholder="one email per line">${esc(S.roles.superintendents.join("\n"))}</textarea>
      <span class="hint">They approve and deny requests. ${(CFG.admins||[]).length ? "Always included: " + esc(CFG.admins.join(", ")) + "." : ""}</span></div>
    <div class="fld"><label for="f_wh">Warehouse managers</label><textarea id="f_wh" name="wh" placeholder="one email per line">${esc(S.roles.warehouse.join("\n"))}</textarea>
      <span class="hint">They get the “Check in at warehouse” button.</span></div>` : ""}
    <div class="fld"><label for="f_tools">Company tool list</label><textarea id="f_tools" name="tools" style="min-height:200px">${esc(S.catalog.tools)}</textarea>
      <span class="hint">One tool per line. A line ending in “:” starts a group, like “Drills &amp; drivers:”.</span></div>
    <div class="fld"><label for="f_rets">Items sent back to warehouse</label><textarea id="f_rets" name="returns" style="min-height:120px">${esc(S.catalog.returns)}</textarea>
      <span class="hint">Shown when crew picks “Send back to warehouse”.</span></div>`;
  sheet("Settings", S.isSuper ? "People and pick lists." : "Pick lists for the request form.", body, async fd => {
    await Store.setConfig("catalog", { tools:String(fd.get("tools")||""), returns:String(fd.get("returns")||"") });
    if (S.isSuper) await Store.setConfig("roles", { superintendents:emails(fd.get("sup")), warehouse:emails(fd.get("wh")) });
    return "Settings saved";
  }, "Save");
}

// ---------- events ----------
document.addEventListener("click", e => {
  const st = e.target.closest("[data-stage]"); if (st) { S.stage = st.dataset.stage; render(); return; }
  const b = e.target.closest("[data-act]"); if (!b) return;
  const act = b.dataset.act;
  if (act==="signin") return S.pca.loginRedirect({ scopes:SCOPES });
  if (act==="setup") return runSetup(b);
  if (act==="retry") return start();
  if (act==="settings") return openSettings();
  const card = b.closest(".card"); const d = card && S.docs.find(x=>x.id===card.dataset.id); if (!d) return;
  ({approve:()=>openDecide(d,true), deny:()=>openDecide(d,false), receive:()=>openReceive(d), release:()=>openRelease(d), checkin:()=>openCheckin(d), delete:()=>openDelete(d)})[act]?.();
});
document.addEventListener("keydown", e => { if (e.key==="Escape" && $("sheetRoot").innerHTML) closeSheet(); });
$("q").addEventListener("input", e => { S.q = e.target.value; render(); });
$("mine").addEventListener("change", e => { S.mine = e.target.checked; render(); });
$("newBtn").addEventListener("click", () => openNew("request"));
$("setBtn").addEventListener("click", openSettings);
$("outBtn").addEventListener("click", () => S.pca.logoutRedirect({ account:S.account }));

// ---------- startup ----------
let refreshing = false;
async function refresh(){
  if (!S.ready || refreshing) return;
  refreshing = true;
  try { await Store.load(); if (!$("sheetRoot").innerHTML) render(); else renderPipeline(); }
  catch (e) { console.warn("Refresh failed", e); }
  finally { refreshing = false; }
}
setInterval(() => { if (document.visibilityState === "visible") refresh(); }, Math.max(10, CFG.refreshSeconds || 20) * 1000);
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") refresh(); });

async function runSetup(btn){
  btn.disabled = true; btn.textContent = "Creating lists…";
  try { await createLists(); S.ready = true; render(); toast("SharePoint lists created"); }
  catch (e) {
    console.error(e);
    gate(`<h3>Setup didn't finish</h3><p>${esc(e.message || "Unknown error")}</p><p>You need Owner access to the SharePoint site, and the app needs admin consent for Sites.Manage.All (README step 1).</p><button class="btn" data-act="retry">Try again</button>`);
  }
}

async function start(){
  S.ready = false; render();
  gate(`<h3>Loading…</h3><p>Connecting to SharePoint.</p>`);
  try {
    await resolveSite();
    const ok = (await listExists(CFG.requestsList)) && (await listExists(CFG.configList));
    if (!ok) {
      gate(S.isAdmin
        ? `<h3>One-time setup</h3><p>The app stores requests in two SharePoint lists on <code>${esc(CFG.sharePointSite)}</code>. Create them now?</p><button class="btn primary" data-act="setup">Create SharePoint lists</button>`
        : `<h3>Almost ready</h3><p>The app hasn't been set up yet. Ask ${esc((CFG.admins||[])[0] || "your admin")} to open it and finish setup.</p>`);
      return;
    }
    await Store.load();
    S.ready = true; render();
  } catch (e) {
    console.error(e);
    const denied = e.status === 403 || e.status === 401;
    gate(`<h3>${denied ? "No access to SharePoint" : "Couldn't load requests"}</h3><p>${denied
      ? `Your account can't open <code>${esc(CFG.sharePointSite)}</code>. Ask ${esc((CFG.admins||[])[0] || "your admin")} to add you as a Member of that site.`
      : esc(e.message || "Check your connection.")}</p><button class="btn" data-act="retry">Try again</button>`);
  }
}

async function boot(){
  if (!CFG.clientId || CFG.clientId.startsWith("PASTE")) {
    gate(`<h3>Finish setup</h3><p>Open <code>config.js</code> and paste in the Microsoft app's client ID and tenant ID (README step 2).</p>`);
    return;
  }
  if (!window.msal) { gate(`<h3>Couldn't start</h3><p>The Microsoft sign-in library didn't load. Refresh the page.</p>`); return; }
  S.pca = new msal.PublicClientApplication({
    auth:{ clientId:CFG.clientId, authority:`https://login.microsoftonline.com/${CFG.tenantId}`, redirectUri: location.origin + location.pathname },
    cache:{ cacheLocation:"localStorage" }
  });
  await S.pca.initialize();
  let res = null;
  try { res = await S.pca.handleRedirectPromise(); } catch (e) { console.error(e); }
  S.account = res?.account || S.pca.getAllAccounts()[0] || null;
  if (!S.account) {
    gate(`<h3>Sign in</h3><p>Use your G&amp;G Microsoft 365 account.</p>${msBtn("Sign in with Microsoft")}`);
    render(); return;
  }
  S.pca.setActiveAccount(S.account);
  S.me = { email: lc(S.account.username), name: S.account.name || S.account.username };
  computeRoles();
  await start();
}
boot();
})();
