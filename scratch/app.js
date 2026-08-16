/* Quy hoáº¡ch BÆ°u Cá»¥c Express â front-end deck.gl (tÄ©nh, khÃ´ng backend).
   Äá»c data/out/{hubs,wards,meta}.json (+ wards.geojson khi cÃ³ polygon).
   Hub hiá»n thá» ngay; layer phÆ°á»ng tá»± kÃ­ch hoáº¡t khi cÃ³ wards.geojson. */

const DATA_DIR = "/api";   // data ÄÃ£ lá»c theo vÃ¹ng + cáº§n session (server-side auth)
const REGION_COLORS = {
  BTB:[230,25,75], XBG:[60,180,75], TTB:[255,165,0], TNT:[0,130,200],
  DBB:[145,30,180], TBB:[70,240,240], TNB:[240,50,230], TNG:[210,245,60],
  "ÄCL":[250,190,212], DNB:[0,128,128], HNO:[170,110,40], NTB:[255,215,0],
  HCM:[128,0,0], DSH:[0,0,128],
};
const ROLE_COLORS = {
  territorial:[37,99,235], pickup_only:[234,88,12],
  bulky_delivery:[147,51,234], special_mixed:[13,148,136],
};
const TYPE_COLORS = { B2B:[153,27,27], transit:[107,114,128], sales:[209,213,219],
  warehouse:[120,90,60], other:[156,163,175], test:[200,200,200] };
const HIDDEN_TYPES = new Set(["sales","warehouse","other","test"]); // khÃ´ng váº½ máº·c Äá»nh

let DATA = { hubs:[], wards:[], meta:{}, geojson:null };

// ---------- Toast: thÃ´ng bÃ¡o khÃ´ng cháº·n (thay cho viá»c ghi lá»i vÃ o #meta) ----------
function toast(msg, kind="info", opts={}){
  const wrap = document.getElementById("toast-wrap"); if(!wrap) return;
  const icon = kind==="ok" ? "ic-check" : (kind==="err"||kind==="warn") ? "ic-alert" : "ic-infoc";
  const el = document.createElement("div");
  el.className = "toast " + (kind||"");
  el.setAttribute("role", kind==="err" ? "alert" : "status");
  el.innerHTML = `<svg class="ic"><use href="#${icon}"/></svg><div class="tx"></div><button class="x" aria-label="ÄÃ³ng">Ã</button>`;
  el.querySelector(".tx").textContent = String(msg);            // textContent -> khÃ´ng XSS
  const kill = ()=>{ el.classList.add("hide"); setTimeout(()=>el.remove(), 200); };
  el.querySelector(".x").onclick = kill;
  wrap.appendChild(el);
  const dur = (opts.duration!=null) ? opts.duration : (kind==="err" ? 6000 : 3500);
  if(dur>0) setTimeout(kill, dur);
  return el;
}
window.toast = toast;

// ---------- Modal a11y: focus-trap + Escape + khÃ´i phá»¥c focus khi ÄÃ³ng ----------
const _MODAL_FOCUS = new WeakMap();
const _MODAL_SEL = 'a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled]),iframe,[tabindex]:not([tabindex="-1"])';
function modalOpen(modal){
  if(!modal) return;
  _MODAL_FOCUS.set(modal, document.activeElement);
  modal.classList.remove("hidden");
  const first = [...modal.querySelectorAll(_MODAL_SEL)].find(el=>el.offsetParent!==null);
  if(first) try{ first.focus(); }catch(e){}
  if(!modal._trap){
    modal._trap = e=>{
      if(e.key==="Escape"){ e.stopPropagation(); modalClose(modal); return; }
      if(e.key!=="Tab") return;
      const list = [...modal.querySelectorAll(_MODAL_SEL)].filter(el=>el.offsetParent!==null);
      if(list.length<2) return;
      const a=list[0], b=list[list.length-1];
      if(e.shiftKey && document.activeElement===a){ e.preventDefault(); b.focus(); }
      else if(!e.shiftKey && document.activeElement===b){ e.preventDefault(); a.focus(); }
    };
    modal.addEventListener("keydown", modal._trap);
  }
}
function modalClose(modal){
  if(!modal) return;
  modal.classList.add("hidden");
  const prev = _MODAL_FOCUS.get(modal);
  if(prev && prev.focus){ try{ prev.focus(); }catch(e){} }
}
window.modalOpen = modalOpen; window.modalClose = modalClose;

// ---------- Giá»¯ ná»i dung khÃ´ng bá» topbar ÄÃ¨ khi nÃ³ wrap á» mÃ n háº¹p ----------
// Äo chiá»u cao THá»°C cá»§a topbar (1 hay nhiá»u hÃ ng) -> biáº¿n --topbar-h; panel bÃªn dÆ°á»i neo theo biáº¿n nÃ y.
(function(){
  const tb = document.getElementById("topbar"); if(!tb) return;
  const setH = ()=> document.documentElement.style.setProperty("--topbar-h", tb.offsetHeight + "px");
  setH();
  if(window.ResizeObserver) new ResizeObserver(setH).observe(tb);
  window.addEventListener("resize", setH);
})();

// ---------- Popover cho â (.info): render fixed -> khÃ´ng bá» panel cáº¯t mÃ©p / nÃºt ná»i ÄÃ¨ ----------
(function(){
  let pop;
  const getPop = ()=>{ if(!pop){ pop=document.createElement("div"); pop.className="info-pop"; document.body.appendChild(pop); } return pop; };
  function show(el){
    const src = el.querySelector(".tip"); const txt = src ? src.textContent : "";
    if(!txt) return;
    const p = getPop(); p.textContent = txt; p.style.display="block"; p.style.left="0"; p.style.top="0";
    const r = el.getBoundingClientRect(), m = 8, pw = p.offsetWidth, ph = p.offsetHeight;
    let left = r.left + r.width/2 - pw/2;
    left = Math.max(m, Math.min(left, window.innerWidth - pw - m));   // káº¹p trong mÃ n hÃ¬nh
    let top = r.bottom + m;
    if(top + ph > window.innerHeight - m) top = r.top - ph - m;       // háº¿t chá» dÆ°á»i -> láº­t lÃªn trÃªn
    p.style.left = left+"px"; p.style.top = Math.max(m, top)+"px";
  }
  function hide(){ if(pop) pop.style.display="none"; }
  const near = e => e.target && e.target.closest && e.target.closest(".info");
  document.addEventListener("mouseover", e=>{ if(near(e)) show(near(e)); });
  document.addEventListener("mouseout",  e=>{ if(near(e)) hide(); });
  document.addEventListener("focusin",   e=>{ if(near(e)) show(near(e)); });
  document.addEventListener("focusout",  e=>{ if(near(e)) hide(); });
  window.addEventListener("scroll", hide, true);
})();
let wardByCode = {}, hubByCode = {};
let maxDemand = 1;
let map, overlay;
const ui = {
  colormode:"auto", lod:"region", showLbl:false, zq:0,
  layers:{ terr:true, orphan:true, b2b:false, transit:false, wards:true, newwards:true },
  planMode:false, planUnit:"new", planTarget:null, tab:"map", regionFilter:null, jt:false,
  selHub:null, selWard:null, selWardNew:false,
  optov:{ on:false, markers:true, reassign:false, merge:false, ws4:false, target:"pure" },
  allowed:"*", username:"", role:"",
};
function regOK(r){
  if(ui.regionFilter) return r===ui.regionFilter;
  if(ui.allowed==="*"||!ui.allowed) return true;
  return ui.allowed.includes(r);
}
const REGION_LIST = ["BTB","XBG","TTB","TNT","DBB","TBB","TNB","TNG","ÄCL","DNB","HNO","NTB","HCM","DSH"];
const plan = { selected:new Map(), selNew:new Map() }; // selected: ward_code cÅ© (tÃ­nh cáº§u); selNew: new_ward_code (hiá»n thá»/khung)

// ---------- helpers ----------
const $ = s => document.querySelector(s);
const fmt = n => (n==null?"â":Math.round(n).toLocaleString("vi-VN"));
const sum = (a,f)=>a.reduce((s,x)=>s+(f(x)||0),0);
function clamp(v,lo,hi){ return Math.max(lo,Math.min(hi,v)); }
function demandOf(h){ const t=h.territory_demand||{pv:0,dv:0}; return t.pv+t.dv; }

// ---------- mÃ n táº£i + fetch CÃ TIMEOUT ----------
// Má»i request data pháº£i cÃ³ háº¡n chá»: náº¿u má»t request khÃ´ng bao giá» tráº£ (extension cháº·n, máº¡ng rá»t
// giá»¯a máº» táº£i, origin restart giá»¯a lÃºc stream), Promise.all treo mÃ£i -> trang xoay vÄ©nh viá»n.
// CÃ³ timeout thÃ¬ háº¡n chá» háº¿t -> reject -> hiá»n lá»i + nÃºt Thá»­ láº¡i.
const T_SMALL = 20000;    // 3 file ná»n (hubs/wards/meta), vÃ i trÄm KB
const T_BIG   = 120000;   // wards.geojson 18MB + 5 file lá»n khÃ¡c
const T_ME    = 15000;    // /me

function abortAfter(ms){
  if(typeof AbortSignal!=="undefined" && AbortSignal.timeout) return AbortSignal.timeout(ms);
  const ac=new AbortController(); setTimeout(()=>ac.abort(), ms); return ac.signal;   // fallback
}
function errLabel(e){
  const n=(e && e.name) || "";
  if(n==="TimeoutError") return "quÃ¡ thá»i gian chá»";
  if(n==="AbortError")   return "quÃ¡ thá»i gian chá» (bá» huá»·)";
  if(n==="TypeError")    return "request bá» cháº·n hoáº·c máº¥t máº¡ng";
  return (e && e.message) ? String(e.message).slice(0,90) : "lá»i khÃ´ng rÃµ";
}
// signal phá»§ cáº£ lÃºc Äá»c body -> stream Äá»©t giá»¯a ÄÆ°á»ng cÅ©ng bá» cáº¯t theo háº¡n chá», khÃ´ng treo.
async function fetchJSON(url, ms){
  const r = await fetch(url, {signal:abortAfter(ms)});
  if(!r.ok) throw new Error("HTTP "+r.status);
  return r.json();
}

const bootOv = {
  get el(){ return document.getElementById("boot-ov"); },
  busy(tx, sub){ const o=this.el; if(!o) return;
    o.classList.remove("hidden","err"); document.getElementById("bo-err").classList.add("hidden");
    document.getElementById("bo-tx").textContent = tx;
    document.getElementById("bo-sub").textContent = sub || ""; },
  sub(tx){ const e=document.getElementById("bo-sub"); if(e) e.textContent = tx; },
  fail(msg, retry){ const o=this.el; if(!o){ alert(msg); return; }
    o.classList.remove("hidden"); o.classList.add("err");
    document.getElementById("bo-tx").textContent = "KhÃ´ng táº£i ÄÆ°á»£c á»©ng dá»¥ng";
    document.getElementById("bo-sub").textContent = "";
    document.getElementById("bo-err").classList.remove("hidden");
    document.getElementById("bo-err-tx").textContent = msg;
    const b=document.getElementById("bo-retry");
    b.onclick = ()=>{ this.busy("Äang thá»­ láº¡iâ¦"); retry(); }; },
  done(){ this.el?.classList.add("hidden"); },
};

// ---------- khÃ´i phá»¥c phiÃªn (login á» trang riÃªng /web/login.html) ----------
async function boot(){
  window.__bootStarted = true;    // cá» cho watchdog trong index.html biáº¿t app.js ÄÃ£ cháº¡y
  bootOv.busy("Äang kiá»m tra phiÃªn ÄÄng nháº­pâ¦");
  let d;
  try{
    const r = await fetch("/me", {signal:abortAfter(T_ME)});
    // CHá» 401/403 má»i lÃ  "chÆ°a/háº¿t phiÃªn" -> vá» trang login. Lá»i máº¡ng mÃ  cÅ©ng redirect thÃ¬
    // ngÆ°á»i dÃ¹ng bá» ÄÃ¡ sang login.html rá»i login láº¡i váº«n lá»i, khÃ´ng hiá»u vÃ¬ sao.
    if(r.status===401 || r.status===403){ location.href="/web/login.html"; return; }
    if(!r.ok) throw new Error("HTTP "+r.status);
    d = await r.json();
  }catch(e){
    bootOv.fail("KhÃ´ng káº¿t ná»i ÄÆ°á»£c mÃ¡y chá»§ khi kiá»m tra phiÃªn ÄÄng nháº­p ("+errLabel(e)+").", boot);
    return;
  }
  ui.username=d.user; ui.role=d.role||""; ui.allowed=d.allowed;
  ui.regionFilter=(Array.isArray(ui.allowed)&&ui.allowed.length===1)?ui.allowed[0]:null;
  load();
}
function fitToAllowed(){
  if(ui.allowed==="*"||!ui.allowed) return;
  const pts=DATA.hubs.filter(h=>h.has_geo && ui.allowed.includes(h._region));
  if(pts.length){ const lo=pts.map(h=>h.lng),la=pts.map(h=>h.lat);
    map.fitBounds([[Math.min(...lo),Math.min(...la)],[Math.max(...lo),Math.max(...la)]],{padding:50,duration:0}); }
}

async function load() {
  bootOv.busy("Äang táº£i dá»¯ liá»u ná»nâ¦", "hubs Â· wards Â· meta");
  let hubs, wards, meta;
  try{
    [hubs, wards, meta] = await Promise.all(
      ["hubs.json","wards.json","meta.json"].map(f=>fetchJSON(`${DATA_DIR}/${f}`, T_SMALL)));
  }catch(e){
    // 3 file nÃ y Báº®T BUá»C cÃ³ má»i dá»±ng ÄÆ°á»£c app -> lá»i thÃ¬ dá»«ng háº³n vÃ  bÃ¡o, khÃ´ng Äi tiáº¿p.
    bootOv.fail("KhÃ´ng táº£i ÄÆ°á»£c dá»¯ liá»u ná»n ("+errLabel(e)+").", load);
    return;
  }
  DATA.hubs=hubs; DATA.wards=wards; DATA.meta=meta;
  hubs.forEach(h=>hubByCode[h.hub_code]=h);
  wards.forEach(w=>wardByCode[w.ward_code]=w);
  buildHubColorIdx();   // gÃ¡n mÃ u lÃ£nh thá» á»n Äá»nh theo chá» sá» BC
  // gÃ¡n region + tá»nh cho hub = Äa sá» phÆ°á»ng nÃ³ phá»¥c vá»¥
  const top=o=>Object.entries(o).sort((a,b)=>b[1]-a[1])[0]?.[0]||"";
  hubs.forEach(h=>{ const rc={},pc={}; (h.covered_wards||[]).forEach(c=>{ const w=wardByCode[c]; if(!w)return;
      if(w.region) rc[w.region]=(rc[w.region]||0)+1; if(w.province) pc[w.province]=(pc[w.province]||0)+1; });
    h._region=top(rc); h._province=top(pc); });
  maxDemand = Math.max(1, ...hubs.filter(h=>h.assigned).map(demandOf));
  // 6 file lá»n cÃ²n láº¡i: táº£i SONG SONG (má»i cÃ¡i lá»i -> null, khÃ´ng lÃ m há»ng cáº£ máº»), CÃ timeout,
  // vÃ  bÃ¡o rÃµ tá»p nÃ o thiáº¿u thay vÃ¬ im láº·ng dá»±ng báº£n Äá» khuyáº¿t polygon.
  const BIG = ["wards.geojson","optimizer.json","rezone.json",
               "wards_new.geojson","competitors_jt.json","ward_centroids.json"];
  const failed=[]; let n=0;
  bootOv.busy("Äang táº£i dá»¯ liá»u báº£n Äá»â¦", `0/${BIG.length} tá»p`);
  const [geo, opt, rez, geoNew, jt, cent] = await Promise.all(BIG.map(f =>
    fetchJSON(`${DATA_DIR}/${f}`, T_BIG)
      .catch(e=>{ failed.push(`${f} (${errLabel(e)})`); return null; })
      .finally(()=>{ bootOv.sub(`${++n}/${BIG.length} tá»p`); })));
  if(geo) DATA.geojson=geo;
  if(opt) DATA.opt=opt;
  if(rez){ DATA.rez=rez; DATA.rezByCode={}; rez.new_wards.forEach(w=>DATA.rezByCode[w.new_code]=w); }
  if(geoNew){ DATA.geojsonNew=geoNew; buildNewCent(); }
  if(jt) DATA.jt=jt;
  if(cent) DATA.cent=cent;
  try{
    initMap(); renderTopbar(); renderLegend(); renderDQ(); wireControls();
  }catch(e){
    // dá»±ng UI lá»i -> váº«n pháº£i bá» mÃ n táº£i + nÃ³i ra, chá»© khÃ´ng Äá» xoay mÃ£i trÃªn lá»i ÄÃ£ biáº¿t
    bootOv.fail("Lá»i khi dá»±ng giao diá»n: "+errLabel(e), load); throw e;
  }
  bootOv.done();
  // Thiáº¿u tá»p lá»n thÃ¬ app váº«n cháº¡y ÄÆ°á»£c nhÆ°ng KHUYáº¾T (máº¥t polygon/optimizerâ¦) -> pháº£i nÃ³i rÃµ.
  if(failed.length) toast("Thiáº¿u dá»¯ liá»u: "+failed.join(", ")+". Táº£i láº¡i trang Äá» thá»­ láº¡i.", "warn", {duration:12000});
  { const lbl = ui.allowed==="*"?"toÃ n quá»c":(Array.isArray(ui.allowed)?ui.allowed.join(", "):"");
    $("#userbox").innerHTML=`<button id="userbtn" title="${esc(ui.username)}">ð¤</button>`+
      `<div id="usermenu" class="hidden"><div class="um-info"><b>${esc(ui.username||"?")}</b><br>${esc(lbl)}</div>`+
      `<a href="#" id="logout">ÄÄng xuáº¥t</a></div>`;
    $("#userbtn").onclick=()=>$("#usermenu").classList.toggle("hidden");
    $("#logout").onclick=e=>{ e.preventDefault(); fetch("/logout",{method:"POST"}).finally(()=>location.href="/web/login.html"); };
    document.addEventListener("click",e=>{ if(!$("#userbox").contains(e.target)) $("#usermenu")?.classList.add("hidden"); }); }
  if(ui.allowed==="*"){ const cb=$("#clog-btn"); cb.style.display=""; cb.onclick=()=>switchTab("chatlog"); }  // icon Lá»ch sá»­ chat chá» admin
  if(ui.allowed!=="*"){ $("#meta").style.display="none"; $("#diff").style.display="none"; }  // 2 Ã´ bÃ¡o cÃ¡o chá» admin
  wireChat();
  if(location.hash==="#scorecard") switchTab("scorecard");   // tiá»n test/deep-link
  if(location.hash==="#optimizer") switchTab("optimizer");
  if(location.hash==="#rezone") switchTab("rezone");
}

// ---------- map ----------
function initMap() {
  map = new maplibregl.Map({
    container:"map",
    preserveDrawingBuffer:true,   // cho phÃ©p chá»¥p canvas (map.getCanvas().toDataURL) cho Äá» xuáº¥t
    style:{ version:8,
      sources:{ base:{ type:"raster", tileSize:256,
        tiles:["https://a.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png",
               "https://b.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png",
               "https://c.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png"],
        attribution:"Â© OpenStreetMap, Â© CARTO" } },
      layers:[{ id:"base", type:"raster", source:"base", paint:{"raster-opacity":0.9} }] },
    center:[106.5,16.2], zoom:5,
  });
  overlay = new deck.MapboxOverlay({ interleaved:true, layers:[] });
  map.addControl(overlay);
  map.addControl(new maplibregl.NavigationControl(), "bottom-right");
  map.on("zoom", ()=>{ const z=map.getZoom(); const lod = z<7?"region":z<10?"province":"hub";
    const showLbl = z>=9;                               // hiá»n sá»m; LOD theo Äá»-lá»n-phÆ°á»ng tá»± lá»c
    const zq = Math.round(z*2)/2;                        // bÆ°á»c zoom 0.5 -> cáº­p nháº­t nhÃ£n mÆ°á»£t, Ã­t váº½ láº¡i
    if(lod!==ui.lod || showLbl!==ui.showLbl || (showLbl && zq!==ui.zq)){ ui.lod=lod; ui.showLbl=showLbl; ui.zq=zq; draw(); } });
  map.on("load", ()=>{ fitToAllowed(); draw(); });
}

// ----- mÃ u tÃ´ lÃ£nh thá» phÆ°á»ng (deck GeoJsonLayer) -----
// cháº¿ Äá» mÃ u hiá»u lá»±c: auto + Äang xem 1 vÃ¹ng -> tÃ´ theo BC (tÃ´ theo vÃ¹ng lÃºc ÄÃ³ vÃ´ nghÄ©a, cáº£ mÃ n 1 mÃ u)
function effColorMode(){
  if(ui.colormode!=="auto") return ui.colormode;
  const single = ui.regionFilter || (Array.isArray(ui.allowed) && ui.allowed.length===1);
  return single ? "hub" : ui.lod;
}
function wardFill(feature){
  const code = feature.properties.ward_code || feature.properties.polygon_id_code;
  const w = wardByCode[code]; if(!w) return [0,0,0,0];
  if(ui.planUnit==="old" && plan.selected.has(w.ward_code)) return [220,38,38,210];  // tÃ´ tay theo phÆ°á»ng CÅ¨
  const mode = effColorMode();
  if(mode==="demand"){ const t=clamp((w.pv+w.dv)/400,0,1);
    return [Math.round(255*t),Math.round(120*(1-t)),60,170]; }
  if(mode==="hub"){ return hubTerrColor(w.delivery_hub||""); }
  const c=REGION_COLORS[w.region]||[200,200,200];           // region / auto
  return [c[0],c[1],c[2], mode==="province"?195:165];
}
function hashCode(s){let h=0;for(let i=0;i<s.length;i++)h=(h*31+s.charCodeAt(i))|0;return Math.abs(h);}
function hslColor(seed){const hue=seed%360;return hsl2rgb(hue,55,55).concat(170);}
// bá» prefix loáº¡i ÄÆ¡n vá», chá» giá»¯ TÃN
function stripPrefix(n){ return (n||"").replace(/^(PhÆ°á»ng|XÃ£|Thá» tráº¥n|Äáº·c khu|Thá» xÃ£|Quáº­n|Huyá»n)\s+/i,"").trim(); }
// tÃ¢m polygon (ring lá»n nháº¥t) Äá» Äáº·t nhÃ£n tÃªn phÆ°á»ng má»i â tÃ­nh 1 láº§n, cache
let newCent = {};
function ringArea(r){let a=0;for(let i=0,n=r.length-1;i<n;i++)a+=r[i][0]*r[i+1][1]-r[i+1][0]*r[i][1];return a/2;}
function ringCentroid(r){let x=0,y=0,a=0;for(let i=0,n=r.length-1;i<n;i++){const f=r[i][0]*r[i+1][1]-r[i+1][0]*r[i][1];x+=(r[i][0]+r[i+1][0])*f;y+=(r[i][1]+r[i+1][1])*f;a+=f;}a*=0.5;return a?[x/(6*a),y/(6*a)]:r[0];}
function buildNewCent(){
  newCent={};
  for(const f of (DATA.geojsonNew&&DATA.geojsonNew.features||[])){
    const g=f.geometry; if(!g) continue;
    const polys = g.type==="Polygon"?[g.coordinates]:(g.type==="MultiPolygon"?g.coordinates:[]);
    let best=null,bestA=-1;
    for(const p of polys){ const outer=p[0]; if(!outer||outer.length<4) continue; const ar=Math.abs(ringArea(outer)); if(ar>bestA){bestA=ar;best=outer;} }
    if(best) newCent[f.properties.new_ward_code]={p:ringCentroid(best), a:bestA};
  }
}
// mÃ u lÃ£nh thá» BC: gÃ¡n theo CHá» Sá» BC + gÃ³c vÃ ng 137.5Â° -> hue tráº£i Äá»u, khÃ´ng trÃ¹ng;
// thÃªm biáº¿n thiÃªn Äá» Äáº­m/sÃ¡ng Äá» BC cáº¡nh nhau dÃ¹ gáº§n hue váº«n khÃ¡c sáº¯c -> phÃ¢n biá»t rÃµ lÃ£nh thá».
let hubColorIdx = {};
function buildHubColorIdx(){
  hubColorIdx = {};
  DATA.hubs.map(h=>h.hub_code).sort().forEach((c,i)=>{ hubColorIdx[c]=i; });
}
function hubTerrColor(code){
  const i = hubColorIdx[code];
  if(i==null) return [205,205,205,150];          // khÃ´ng rÃµ BC -> xÃ¡m
  const hue = (i*137.508) % 360;
  const sat = [64,78,52][i%3];
  const lig = [56,45,67][Math.floor(i/3)%3];
  return hsl2rgb(hue,sat,lig).concat(185);
}
function hsl2rgb(h,s,l){s/=100;l/=100;const k=n=>(n+h/30)%12;const a=s*Math.min(l,1-l);
  const f=n=>l-a*Math.max(-1,Math.min(k(n)-3,Math.min(9-k(n),1)));
  return [Math.round(255*f(0)),Math.round(255*f(8)),Math.round(255*f(4))];}

// ---------- layers ----------
function visibleHubs() {
  return DATA.hubs.filter(h=>{
    if(!h.has_geo) return false;
    if(!regOK(h._region)) return false;
    if(h.type==="B2B")      return ui.layers.b2b;
    if(h.type==="transit")  return ui.layers.transit;   // gá»m Trung Chuyá»n + Chuyá»n Tiáº¿p
    if(HIDDEN_TYPES.has(h.type)) return false;           // kho khÃ¡c/sales/test/thanh lÃ½
    // express
    if(h.role==="territorial") return ui.layers.terr;
    return ui.layers.orphan; // chuyÃªn dá»¥ng
  });
}
function hubColor(h){
  if(h.type==="B2B"||h.type==="transit"||h.type==="sales") return TYPE_COLORS[h.type];
  if(h.role!=="territorial") return ROLE_COLORS[h.role];
  if(ui.colormode==="region") return REGION_COLORS[regionOfHub(h)]||[120,120,120];
  if(ui.colormode==="demand"){ const t=clamp(demandOf(h)/maxDemand,0,1);
    return [Math.round(255*t), Math.round(80+100*(1-t)), Math.round(60*(1-t))]; }
  return ROLE_COLORS.territorial;
}
function regionOfHub(h){ const w=h.covered_wards&&h.covered_wards[0]; return w&&wardByCode[w]?wardByCode[w].region:""; }
function hubRadius(h){
  // Äá»ng Äá»u háº¿t: má»i BC cÃ¹ng cá»¡, phÃ¢n biá»t táº£i báº±ng mÃ u (khÃ´ng báº±ng kÃ­ch thÆ°á»c)
  return h.role==="territorial" ? 5 : 4.5;
}

function draw(){
  const layers=[];
  // lÃ£nh thá» phÆ°á»ng (polygon) â dÆ°á»i cÃ¹ng
  if(ui.layers.wards && DATA.geojson){
    let feats=DATA.geojson.features;
    if(ui.regionFilter || ui.allowed!=="*") feats=feats.filter(f=>{const w=wardByCode[f.properties.ward_code];return w&&regOK(w.region);});
    layers.push(new deck.GeoJsonLayer({
      id:"wards", data:{type:"FeatureCollection",features:feats}, pickable:true, stroked:true, filled:true,
      getFillColor:wardFill,
      // viá»n: xÃ¡m cho vÃ¹ng KHÃNG cÃ³ dá»¯ liá»u (Äáº£o/Äáº·c khu â khÃ´ng join wardByCode); tráº¯ng cho phÆ°á»ng thÆ°á»ng
      getLineColor:f=>wardByCode[f.properties.ward_code]?[255,255,255,95]:[120,120,120,170], lineWidthMinPixels:0.5,
      updateTriggers:{ getFillColor:[ui.colormode, ui.lod, plan.selected.size, ui.regionFilter, ui.allowed], data:[ui.regionFilter] },
      onClick: info=>info.object && clickWard(info.object),
    }));
  }
  // viá»n lÃ£nh thá» phÆ°á»ng Má»I (re-zone) â trÃªn fill cÅ©; xÃ© = Äá». Báº¯t click khi ÄÆ¡n vá» = "má»i"
  if(ui.layers.newwards && DATA.geojsonNew){
    let nf=DATA.geojsonNew.features;
    if(ui.regionFilter || ui.allowed!=="*") nf=nf.filter(f=>{const w=DATA.rezByCode&&DATA.rezByCode[f.properties.new_ward_code];return w&&regOK(w.region);});
    layers.push(new deck.GeoJsonLayer({
      id:"newwards", data:{type:"FeatureCollection",features:nf}, pickable: ui.planUnit==="new",
      stroked:true, filled:true,
      getFillColor:f=>plan.selNew.has(f.properties.new_ward_code)?[225,29,42,245]:[0,0,0,0],  // tÃ´ tay phÆ°á»ng Má»I: Äá» Äáº¬M phá»§ rÃµ (khÃ´ng trá»n mÃ u ná»n)
      getLineColor:f=>{ if(plan.selNew.has(f.properties.new_ward_code)) return [185,28,28,255];
        const w=DATA.rezByCode&&DATA.rezByCode[f.properties.new_ward_code];
        return w&&w.status==="split"?[220,38,38,235]:[35,35,35,150]; },
      getLineWidth:f=>{ if(plan.selNew.has(f.properties.new_ward_code)) return 2.4;
        const w=DATA.rezByCode&&DATA.rezByCode[f.properties.new_ward_code];
        return w&&w.status==="split"?2.4:1.0; },
      lineWidthUnits:"pixels", lineWidthMinPixels:1,
      updateTriggers:{ pickable:[ui.planUnit], getFillColor:[plan.selNew.size], getLineColor:[plan.selNew.size], getLineWidth:[plan.selNew.size] },
      onClick: info=>info.object && clickNewWard(info.object),
    }));
    // nhÃ£n TÃN phÆ°á»ng má»i â chá» hiá»n khi phÆ°á»ng Äá»§ TO trÃªn mÃ n Äá» chá»©a chá»¯ (LOD theo zoom, khÃ´ng rá»i)
    if(ui.showLbl){
      const z=map.getZoom(), k=Math.pow(2,z);
      const lbls=nf.map(f=>{ const code=f.properties.new_ward_code; const c=newCent[code];
        if(!c || Math.sqrt(c.a)*k <= 110) return null;     // phÆ°á»ng cÃ²n nhá» trÃªn mÃ n -> chÆ°a hiá»n nhÃ£n
        const w=DATA.rezByCode&&DATA.rezByCode[code];
        const nm=stripPrefix(w?w.name:f.properties.name||"");
        return nm?{position:c.p,text:nm}:null; }).filter(Boolean);
      layers.push(new deck.TextLayer({
        id:"newward-labels", data:lbls, pickable:false,
        getPosition:d=>d.position, getText:d=>d.text,
        characterSet:"auto",   // Äá»§ kÃ½ tá»± tiáº¿ng Viá»t
        getSize:14, sizeUnits:"pixels", sizeMinPixels:12, sizeMaxPixels:20,
        getColor:[17,24,39,255], getTextAnchor:"middle", getAlignmentBaseline:"center",
        fontFamily:'"Be Vietnam Pro", -apple-system, system-ui, "Segoe UI", Arial, sans-serif', fontWeight:700,
        // SDF + atlas Äá» phÃ¢n giáº£i cao + smoothing -> chá»¯ sáº¯c nÃ©t, khá»­ rÄng cÆ°a khi phÃ³ng to
        fontSettings:{sdf:true, fontSize:128, buffer:24, radius:30, cutoff:0.25, smoothing:0.12},
        outlineWidth:6, outlineColor:[255,255,255,255],
        maxWidth:130,
        updateTriggers:{ data:[ui.regionFilter,ui.allowed,ui.zq] },
      }));
    }
  }
  // lá»p tham kháº£o BC Äá»i thá»§ J&T (dÆ°á»i hub GHN) â lá»c theo vÃ¹ng Äang xem
  if(ui.jt && DATA.jt){
    let jts=DATA.jt;
    if(ui.regionFilter || ui.allowed!=="*") jts=jts.filter(d=>regOK(d.region));
    layers.push(new deck.ScatterplotLayer({
      id:"jt", data:jts, pickable:true,
      updateTriggers:{ data:[ui.regionFilter,ui.allowed] },
      getPosition:d=>[d.lng,d.lat], getRadius:3.2, radiusUnits:"pixels",
      getFillColor:[20,20,20,200], getLineColor:[255,255,255,160], stroked:true, lineWidthMinPixels:0.6,
      onClick: info=>info.object && showDetail(`<h3>${esc(info.object.name)}</h3>
        <div style="color:#6b7280;margin-bottom:6px">J&T</div>
        <div>${esc(info.object.addr||info.object.province||"â")}
        &nbsp;<a href="https://www.google.com/maps?q=${encodeURIComponent(info.object.lat)},${encodeURIComponent(info.object.lng)}" target="_blank">(Maps)</a></div>`),
    }));
  }
  // ===== Overlay Network Optimizer (chá» Äá» xem + chá»¥p, khÃ´ng lÆ°u) =====
  if(ui.optov.on && DATA.opt){
    const o=DATA.opt, ov=ui.optov;
    const hc=c=>{ const h=hubByCode[c]; return h&&h.has_geo?[h.lng,h.lat]:null; };
    const wc=c=>{ const p=DATA.cent&&DATA.cent[c]; return p?[p[0],p[1]]:null; };
    // mÅ©i tÃªn reassign phÆ°á»ng â BC (cong); target: gáº§n nháº¥t | gáº§n nháº¥t-cÃ²n-táº£i
    if(ov.reassign){
      const arcs=(o.reassign||[]).filter(x=>regOK(x.region)).map(x=>{
        const src=wc(x.ward), dst=hc(ov.target==="cap"?x.to_cap:x.to_pure);
        return src&&dst?{...x,src,dst}:null; }).filter(Boolean);
      layers.push(new deck.ArcLayer({ id:"opt-reassign", data:arcs, pickable:true,
        getSourcePosition:d=>d.src, getTargetPosition:d=>d.dst,
        getSourceColor:[37,99,235,70], getTargetColor:[37,99,235,235],
        getWidth:1.4, widthUnits:"pixels",
        onClick:info=>info.object&&showDetail(reassignHtml(info.object)),
        updateTriggers:{ data:[ui.regionFilter,ov.target,ui.allowed] } }));
    }
    // line gá»p khi ÄÃ³ng BC (close â merge_to)
    if(ov.merge){
      const lines=(o.close||[]).filter(x=>regOK(x.region)&&x.merge_to).map(x=>{
        const src=hc(x.hub), dst=hc(x.merge_to); return src&&dst?{...x,src,dst}:null; }).filter(Boolean);
      layers.push(new deck.LineLayer({ id:"opt-merge", data:lines, pickable:true,
        getSourcePosition:d=>d.src, getTargetPosition:d=>d.dst,
        getColor:[220,38,38,200], getWidth:1.8, widthUnits:"pixels",
        onClick:info=>info.object&&showDetail(closeHtml(info.object)),
        updateTriggers:{ data:[ui.regionFilter,ui.allowed] } }));
    }
    // ws4: cháº¥m phÆ°á»ng cáº§u cao theo 4 nhÃ³m whitespace Ã J&T
    if(ov.ws4 && o.ws4){
      [["mat_khach",[220,38,38]],["greenfield",[16,185,129]],["doi_dau",[245,158,11]]].forEach(([k,col])=>{
        const pts=(o.ws4[k]||[]).filter(x=>regOK(x.region)).map(x=>{const p=wc(x.ward);return p?{...x,_p:p,_g:k}:null;}).filter(Boolean);
        layers.push(new deck.ScatterplotLayer({ id:"opt-ws4-"+k, data:pts, pickable:true,
          getPosition:d=>d._p, getRadius:5, radiusUnits:"pixels",
          getFillColor:[...col,200], stroked:true, getLineColor:[255,255,255,210], lineWidthMinPixels:0.8,
          onClick:info=>info.object&&showDetail(ws4Html(info.object)),
          updateTriggers:{ data:[ui.regionFilter,ui.allowed] } }));
      });
    }
    // markers BC: ÄÃ³ng (Äá») / tÃ¡ch-má» rá»ng (cam) â vÃ²ng to dÆ°á»i cháº¥m hub
    if(ov.markers){
      const mk=[];
      (o.close||[]).filter(x=>regOK(x.region)).forEach(x=>{const p=hc(x.hub);if(p)mk.push({p,col:[220,38,38],_t:"close",x});});
      (o.split||[]).filter(x=>regOK(x.region)).forEach(x=>{const p=hc(x.hub);if(p)mk.push({p,col:[245,158,11],_t:"split",x});});
      layers.push(new deck.ScatterplotLayer({ id:"opt-mk", data:mk, pickable:false,
        getPosition:d=>d.p, getRadius:9, radiusUnits:"pixels",
        stroked:true, filled:false, getLineColor:d=>[...d.col,235], lineWidthMinPixels:2.4,
        updateTriggers:{ data:[ui.regionFilter,ui.allowed] } }));
    }
  }
  // highlight PHÆ¯á»NG Äang chá»n (viá»n vÃ ng dÃ y) â áº©n khi Äang quy hoáº¡ch Äá» khÃ´ng trá»n vá»i Äá» nhÃ³m chá»n
  if(ui.selWard && !ui.planMode){
    const src = ui.selWardNew ? DATA.geojsonNew : DATA.geojson;
    const key = ui.selWardNew ? "new_ward_code" : "ward_code";
    const f = src && src.features.find(x=>x.properties[key]===ui.selWard);
    if(f) layers.push(new deck.GeoJsonLayer({ id:"sel-ward", data:{type:"FeatureCollection",features:[f]},
      stroked:true, filled:true, getFillColor:[255,193,7,70], getLineColor:[245,158,11,255], lineWidthMinPixels:4,
      updateTriggers:{ data:[ui.selWard,ui.selWardNew] } }));
  }
  // Äiá»m hub â trÃªn cÃ¹ng
  const hubs=visibleHubs();
  layers.push(new deck.ScatterplotLayer({
    id:"hubs", data:hubs, pickable:true, stroked:true,
    getPosition:h=>[h.lng,h.lat], getRadius:hubRadius, radiusUnits:"pixels",
    getFillColor:h=>[...hubColor(h),200],
    getLineColor:h=>h.role==="territorial"?[255,255,255,180]:[20,20,20,220],
    lineWidthMinPixels:1.2,
    updateTriggers:{ getFillColor:[ui.colormode] },
    onClick: info=>info.object && clickHub(info.object),
  }));
  // highlight BC Äang chá»n (vÃ²ng sÃ¡ng to)
  if(ui.selHub && hubByCode[ui.selHub] && hubByCode[ui.selHub].has_geo){
    const h=hubByCode[ui.selHub];
    layers.push(new deck.ScatterplotLayer({ id:"sel-hub", data:[h],
      getPosition:d=>[d.lng,d.lat], getRadius:11, radiusUnits:"pixels",
      stroked:true, filled:false, getLineColor:[255,193,7,255], lineWidthMinPixels:3,
      updateTriggers:{ data:[ui.selHub] } }));
  }
  overlay.setProps({ layers });
  buildTooltip(hubs);
}
function buildTooltip(){
  overlay.setProps({ getTooltip: ({object,layer})=>{
    if(!object) return null;
    if(layer&&layer.id==="hubs"){ return {html:`<b>${esc(object.name||object.hub_code)}</b><br>${esc(object.hub_code)} Â· ${roleLabel(object)}`}; }
    if(layer&&layer.id==="jt"){ return {html:`<b>${esc(object.name)}</b> <span style="color:#9ca3af">J&T</span><br>${esc(object.province)}`}; }
    if(layer&&layer.id==="wards"){ const w=wardByCode[object.properties.ward_code||object.properties.polygon_id_code]; if(!w)return null;
      return {html:`<b>${esc(w.name)}</b><br>${esc(w.district)}, ${esc(w.province)}<br>láº¥y ${fmt(w.pv)} Â· giao ${fmt(w.dv)} ÄÆ¡n/ngÃ y`}; }
    if(layer&&layer.id==="opt-reassign"){ return {html:`<b>${esc(object.name)}</b> â Äá»i BC<br>${esc(object.from_name)} â <b>${esc(ui.optov.target==="cap"?object.to_cap_name:object.to_pure_name)}</b>`}; }
    if(layer&&layer.id==="opt-merge"){ return {html:`ÄÃ³ng <b>${esc(object.name)}</b><br>gá»p vá» ${esc(object.merge_to_name)} (${object.merge_dist} km)`}; }
    if(layer&&String(layer.id).startsWith("opt-ws4-")){ const g={mat_khach:"Khoáº£ng cÃ¡ch xa",greenfield:"Greenfield",doi_dau:"Äá»i Äáº§u J&T"}[object._g];
      return {html:`<b>${esc(object.name)}</b> Â· ${g}<br>cáº§u ${fmt(object.dem)}/ngÃ y Â· GHN ${object.d_ghn}km Â· J&T ${object.d_jt}km`}; }
    return null;
  }});
}

// ---------- chi tiáº¿t overlay optimizer ----------
function reassignHtml(x){
  const cap=ui.optov.target==="cap";
  return `<h3>ð ${esc(x.name)}</h3><div style="color:#6b7280;margin-bottom:6px">${esc(x.province)} Â· ${esc(x.region)} Â· cáº§u ${fmt(x.dem)} ÄÆ¡n/ngÃ y</div>
    <table class="bal"><tr><td>BC hiá»n táº¡i</td><td>${esc(x.from_name)} Â· <b>${x.d_cur} km</b></td></tr>
    <tr><td>Gáº§n nháº¥t</td><td>${esc(x.to_pure_name)} Â· <b>${x.d_pure} km</b></td></tr>
    <tr><td>Gáº§n nháº¥t cÃ²n táº£i</td><td>${esc(x.to_cap_name)} Â· <b>${x.d_cap} km</b></td></tr></table>
    <div class="note">Äang váº½ theo: ${cap?"gáº§n nháº¥t-cÃ²n-táº£i":"gáº§n nháº¥t"}</div>`;
}
function closeHtml(x){
  return `<h3>ð´ ÄÃ³ng/rÃ  soÃ¡t: ${esc(x.name)}</h3><div style="color:#6b7280;margin-bottom:6px">${esc(x.province)} Â· ${esc(x.region)}</div>
    <table class="bal"><tr><td>Äá»ng/ÄÆ¡n</td><td><b>${fmt(x.dpo)}</b></td></tr>
    <tr><td>Kg/mÂ²</td><td>${x.kgm2}</td></tr><tr><td>PhÆ°á»ng cover</td><td>${x.n_wards}</td></tr>
    <tr><td>HÄ cÃ²n</td><td>${x.days==null?"â":x.days+" ngÃ y"}</td></tr>
    <tr><td>HÃ nh Äá»ng</td><td>${esc(x.action)}</td></tr>
    <tr><td>Gá»p vá»</td><td>${esc(x.merge_to_name)} (${x.merge_dist} km)</td></tr></table>`;
}
function ws4Html(x){
  const g={mat_khach:"Khoáº£ng cÃ¡ch xa (GHN xa, J&T gáº§n)",greenfield:"Greenfield (cáº£ 2 Äá»u xa)",doi_dau:"Äá»i Äáº§u J&T (cáº£ 2 Äá»u gáº§n)"}[x._g];
  return `<h3>â­ ${esc(x.name)}</h3><div style="color:#6b7280;margin-bottom:6px">${esc(x.province)} Â· ${esc(x.region)}</div>
    <table class="bal"><tr><td>NhÃ³m</td><td><b>${g}</b></td></tr>
    <tr><td>Cáº§u/ngÃ y</td><td>${fmt(x.dem)}</td></tr>
    <tr><td>GHN gáº§n nháº¥t</td><td>${x.d_ghn} km</td></tr>
    <tr><td>J&T gáº§n nháº¥t</td><td>${x.d_jt} km</td></tr></table>`;
}

// ---------- click detail ----------
function roleLabel(h){
  if(h.type!=="express") return ({B2B:"Kho HÃ ng Náº·ng (B2B)",transit:"Kho Trung Chuyá»n/Chuyá»n Tiáº¿p",
    sales:"Field Sales",warehouse:"Kho khÃ¡c",other:"KhÃ¡c",test:"Test"})[h.type]||h.type;
  return {territorial:"BÆ°u cá»¥c",pickup_only:"ChuyÃªn Láº¤Y",bulky_delivery:"Cá»ng ká»nh/GIAO",special_mixed:"ChuyÃªn dá»¥ng há»n há»£p"}[h.role];
}
function clickHub(h){
  ui.selHub=h.hub_code; ui.selWard=null; draw();
  const re=h.realestate; const ac=h.actual;
  const vol=ac?ac.pv+ac.dv:0, wt=ac?ac.pw+ac.dw:0;
  const dpo = re&&vol>0 ? re.rent/(vol*30) : null;     // tiá»n thuÃª thÃ¡ng Ã· (ÄÆ¡n/ngÃ y Ã 30)
  const dpw = re&&wt>0  ? re.rent/(wt*30)  : null;     // Ã· (kg/ngÃ y Ã 30)
  const pr=prodOf(h), prk=prodKgOf(h), ps=prodStats();  // nÄng suáº¥t NV (ÄÆ¡n/NV/ngÃ y) + ngÆ°á»¡ng vÃ¹ng
  const prTag = pr==null ? "" : (ps.p75&&pr>=ps.p75 ? ' <span class="warn">ð´ thiáº¿u ngÆ°á»i</span>'
                : (ps.p25&&pr<=ps.p25 ? ' <span style="color:#2563eb">ðµ thá»«a ngÆ°á»i</span>' : ''));
  const sb=scoreBench(), sc=bcScore(h);                 // ÄIá»M hiá»u suáº¥t tuyá»t Äá»i (ÄÆ¡n-tÄ/NV)
  const scTag = sc==null ? "" : (sb.p75&&sc>=sb.p75 ? ' <span class="warn">ð´ cao</span>'
                : (sb.p25&&sc<=sb.p25 ? ' <span style="color:#2563eb">ðµ tháº¥p</span>' : ''));
  const maps = h.has_geo ? `&nbsp;<a href="https://www.google.com/maps?q=${encodeURIComponent(h.lat)},${encodeURIComponent(h.lng)}" target="_blank">(Maps)</a>` : "";
  let html=`<h3>${esc(h.name||h.hub_code)}</h3>
    <div><span class="tag" style="background:rgb(${(h.type==="express"?ROLE_COLORS[h.role]:TYPE_COLORS[h.type]).join(',')})">${roleLabel(h)}</span>
    <span style="color:#6b7280">${esc(h.hub_code)} Â· ${esc(regionOfHub(h)||"â")}${maps}</span></div>
    <div class="kv">
      <div>PhÆ°á»ng phá»¥c vá»¥</div><div>${h.n_wards}</div>
      <div>Cáº§u Äá»a bÃ n /ngÃ y</div><div>${fmt(demandOf(h))} ÄÆ¡n Â· ${fmt((h.territory_demand.pw||0)+(h.territory_demand.dw||0))} kg</div>
      <div>Â· Láº¥y / Giao (ÄÆ¡n)</div><div>${fmt(h.territory_demand.pv)} / ${fmt(h.territory_demand.dv)}</div>
      <div>Â· Láº¥y / Giao (kg)</div><div>${fmt(h.territory_demand.pw||0)} / ${fmt(h.territory_demand.dw||0)}</div>
      <div class="grp-top">Sáº£n lÆ°á»£ng BC /ngÃ y</div><div class="grp-top">${ac?fmt(vol)+" ÄÆ¡n Â· "+fmt(wt)+" kg":'<span class="warn">thiáº¿u</span>'}</div>
      ${ac?`<div>Â· Láº¥y / Giao (ÄÆ¡n)</div><div>${fmt(ac.pv)} / ${fmt(ac.dv)}</div>
      <div>Â· Láº¥y / Giao (kg)</div><div>${fmt(ac.pw)} / ${fmt(ac.dw)}</div>`:""}
      <div class="grp-top">NhÃ¢n viÃªn</div><div class="grp-top">${h.staff?fmt(h.staff):'<span class="warn">thiáº¿u</span>'}</div>
      ${pr!=null?`<div>ÄÆ¡n / NV / ngÃ y</div><div>${fmt(pr)}${prTag}</div>
      <div>Kg / NV / ngÃ y</div><div>${fmt(prk)}</div>`:""}
      ${sc!=null?`<div>Äiá»m hiá»u suáº¥t</div><div>${fmt(sc)} Äiá»m/NV${scTag}</div>`:""}
      <div class="grp-top">Máº·t báº±ng</div><div class="grp-top">${re?fmt(re.usable_area)+" mÂ²":'<span class="warn">thiáº¿u</span>'}</div>
      <div>Tiá»n thuÃª</div><div>${re?fmt(re.rent)+" Ä":"â"}</div>
      <div>Háº¡n HÄ</div><div>${re&&re.expiry?re.expiry:"â"}</div>
      <div>Äá»ng/ÄÆ¡n</div><div>${dpo?fmt(dpo)+" Ä":"â"}</div>
      <div>Äá»ng/kg</div><div>${dpw?fmt(dpw)+" Ä":"â"}</div>
    </div>`;
  if(h.missing_geo||h.missing_realestate) html+=`<div class="warn">â  ${[h.missing_geo&&"thiáº¿u toáº¡ Äá»",h.missing_realestate&&"thiáº¿u máº·t báº±ng"].filter(Boolean).join(", ")}</div>`;
  const names=h.covered_wards.slice(0,40).map(c=>esc(wardByCode[c]?wardByCode[c].name:c));
  html+=`<div style="margin-top:8px"><b>PhÆ°á»ng (${h.n_wards}):</b> <span style="color:#6b7280;font-size:12px">${names.join(", ")}${h.n_wards>40?" â¦":""}</span></div>`;
  showDetail(html);
  highlightWards(new Set(h.covered_wards));
}
function clickWard(feature){
  const w=wardByCode[feature.properties.ward_code||feature.properties.polygon_id_code]; if(!w) return;
  if(ui.planMode){ togglePlan(w); return; }
  ui.selHub=null; ui.selWard=w.ward_code; ui.selWardNew=false; draw();
  const ph=hubByCode[w.pick_hub], dh=hubByCode[w.delivery_hub];
  showDetail(`<h3>${esc(w.name)}</h3>
    <div style="color:#6b7280">${esc(w.district)}, ${esc(w.province)} Â· ${esc(w.region)}</div>
    <div class="kv">
      <div>Láº¥y</div><div>${fmt(w.pv)} ÄÆ¡n Â· ${fmt(w.pw)} kg</div>
      <div>Giao</div><div>${fmt(w.dv)} ÄÆ¡n Â· ${fmt(w.dw)} kg</div>
      <div>BC láº¥y</div><div>${esc(ph?ph.name:w.pick_hub||"â")}</div>
      <div>BC giao</div><div>${esc(dh?dh.name:w.delivery_hub||"â")}</div>
    </div>`);
}
function clickNewWard(feature){
  const w=DATA.rezByCode && DATA.rezByCode[feature.properties.new_ward_code];
  if(ui.planMode){   // tÃ´ tay ÄÆ¡n vá» "má»i": chá»n phÆ°á»ng má»i = tÃ´ polygon 3321 + gom cáº§u tá»« phÆ°á»ng cÅ© bÃªn trong
    const codes = w ? (w.old_codes||[]) : [];
    const nc = feature.properties.new_ward_code;
    const all = codes.length && codes.every(c=>plan.selected.has(c));
    codes.forEach(c=> all?plan.selected.delete(c):plan.selected.set(c,true));
    if(all) plan.selNew.delete(nc); else plan.selNew.set(nc,true);   // tÃ´/ÄÃ³ng khung theo polygon phÆ°á»ng má»i
    renderPlan(); draw(); return;
  }
  ui.selHub=null; ui.selWard=feature.properties.new_ward_code; ui.selWardNew=true; draw();
  if(w) showRezoneDetail(w);
  else showDetail(`<h3>${esc(feature.properties.name||"")}</h3><div style="color:#6b7280">PhÆ°á»ng má»i ${esc(feature.properties.new_ward_code)}</div>`);
}
function showDetail(html){ $("#detail-body").innerHTML=html; $("#detail").classList.remove("hidden"); }
function highlightWards(set){ /* hook: khi cÃ³ polygon sáº½ tÃ´; hiá»n chá» no-op náº¿u chÆ°a cÃ³ geojson */ }

// ---------- what-if (tÃ­nh trÃªn Cáº¦U Î£ T1, báº£o toÃ n theo phÆ°á»ng) ----------
function togglePlan(w){
  if(plan.selected.has(w.ward_code)) plan.selected.delete(w.ward_code);
  else plan.selected.set(w.ward_code,true);
  renderPlan(); draw();
}
function demandKgOf(h){ const t=h.territory_demand||{}; return (t.pw||0)+(t.dw||0); }
// ---- nÄng suáº¥t nhÃ¢n viÃªn: ÄÆ¡n/NV/ngÃ y (láº¥y+giao dÃ¹ng chung) ----
function prodOf(h){ const a=h.actual, s=h.staff||0; return (a&&s)?(a.pv+a.dv)/s:null; }
function prodKgOf(h){ const a=h.actual, s=h.staff||0; return (a&&s)?(a.pw+a.dw)/s:null; }
function pctl(a,p){ if(!a.length) return null; const s=[...a].sort((x,y)=>x-y); return s[Math.min(s.length-1,Math.floor(p/100*s.length))]; }
function prodStats(){   // ngÆ°á»¡ng ÄÆ¡n/NV theo cÃ¡c BC express trong pháº¡m vi Äang lá»c
  const v=DATA.hubs.filter(h=>h.type==="express"&&h.assigned&&regOK(h._region)&&prodOf(h)!=null).map(prodOf);
  return {p25:pctl(v,25),p50:pctl(v,50),p75:pctl(v,75),n:v.length};
}
// ===== ÄIá»M HIá»U SUáº¤T (benchmark má»i, KHÃNG Äá»¥ng cÃ¡c chá» sá» hiá»n cÃ³) =====
// Quy Äá»i: 1 ÄÆ¡n láº¥y = 0.4 ÄÆ¡n giao; 1 kg láº¥y = 0.4 kg giao. ÄIá»M TUYá»T Äá»I = nÄng suáº¥t
// "ÄÆ¡n-tÆ°Æ¡ng-ÄÆ°Æ¡ng/NV/ngÃ y" = (ÄÆ¡n QÄ + kg QÄ Ã· K) / NV, vá»i K = kg TB má»i ÄÆ¡n (toÃ n quá»c, cá» Äá»nh
// Äá» quy kg vá» ÄÆ¡n). Äiá»m lÃ  sá» THáº¬T, trung vá» má»i vÃ¹ng khÃ¡c nhau. Cá» ð´/ðµ theo P75/P25 vÃ¹ng Äang lá»c.
function effOrd(h){ const a=h.actual; return a?0.4*a.pv+a.dv:0; }     // ÄÆ¡n quy Äá»i
function effKg(h){ const a=h.actual; return a?0.4*a.pw+a.dw:0; }      // kg quy Äá»i
function _scoreHubs(){ return DATA.hubs.filter(h=>h.type==="express"&&h.assigned&&regOK(h._region)&&h.staff&&h.actual); }
let _kgPerOrd=null;
function kgPerOrd(){   // K toÃ n quá»c = Î£ kg QÄ Ã· Î£ ÄÆ¡n QÄ (cá» Äá»nh, khÃ´ng Äá»i theo vÃ¹ng lá»c)
  if(_kgPerOrd!=null) return _kgPerOrd;
  let so=0,sk=0; DATA.hubs.forEach(h=>{ if(h.type==="express"&&h.assigned&&h.actual){ so+=effOrd(h); sk+=effKg(h); } });
  _kgPerOrd = so>0 ? sk/so : 1; return _kgPerOrd;
}
function bcScore(h){ if(!h.staff||!h.actual) return null;
  const K=kgPerOrd()||1; return (effOrd(h)+effKg(h)/K)/h.staff; }   // ÄÆ¡n-tÆ°Æ¡ng-ÄÆ°Æ¡ng/NV/ngÃ y (tuyá»t Äá»i)
function scoreBench(){ const v=_scoreHubs().map(bcScore).filter(x=>x!=null);
  return { p25:pctl(v,25), p50:pctl(v,50), p75:pctl(v,75), n:v.length }; }
const PROP_BTN=`<button id="save-prop" onclick="saveProposal()" style="width:100%;margin-top:8px;padding:9px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:600;cursor:pointer">ð¾ LÆ°u Äá» xuáº¥t</button>`;
// Æ¯á»c lÆ°á»£ng Äá»nh biÃªn cho Äá» xuáº¥t: NV cáº§n cho cáº§u chuyá»n/nháº­n, theo nÄng suáº¥t trung vá» (ÄÆ¡n/NV/ngÃ y).
function planStaffEstimate(gainDon, totDon, T){
  const med=prodStats().p50; if(!med) return null;
  const th=T?hubByCode[T]:null;
  return { med:Math.round(med), addNV:Math.ceil((T?gainDon:totDon)/med), curStaff: th?(th.staff||0):null, isNew:!T };
}
function staffNoteHTML(e){
  if(!e) return "";
  return e.isNew
    ? `ð· Äá»nh biÃªn: BC má»i gÃ¡nh cá»¥m nÃ y â cáº§n ~<b>${fmt(e.addNV)}</b> NV (nÄng suáº¥t TB ${fmt(e.med)} ÄÆ¡n/NV/ngÃ y).`
    : `ð· Äá»nh biÃªn: BC ÄÃ­ch Äang <b>${fmt(e.curStaff)}</b> NV, nháº­n thÃªm cáº§u â cáº§n bá» sung ~<b>${fmt(e.addNV)}</b> NV (nÄng suáº¥t TB ${fmt(e.med)} ÄÆ¡n/NV/ngÃ y).`;
}
// Quy hoáº¡ch theo PHÆ¯á»NG Má»I: tÃ­nh cáº§u tá»« olds[] (ÄÃ chia 1/n cho phÆ°á»ng cÅ© "Nháº­p má»t pháº§n"), gom theo BC giao
function renderPlanNew(){
  const codes=[...plan.selNew.keys()];
  renderLegend();
  if(!codes.length){ $("#detail").classList.add("hidden"); return; }
  const nws=codes.map(c=>DATA.rezByCode&&DATA.rezByCode[c]).filter(Boolean);
  const T=ui.planTarget||null;
  const loss={}; let totDon=0,totKg=0,gainDon=0,gainKg=0; const oldNames=[],oldCodes=[];
  nws.forEach(w=>(w.olds||[]).forEach(o=>{
    totDon+=o.dem||0; totKg+=o.dem_kg||0; oldNames.push(o.name); oldCodes.push(o.ward);
    if(o.bc && o.bc!==T){ const e=loss[o.bc]||(loss[o.bc]={don:0,kg:0}); e.don+=o.dem||0; e.kg+=o.dem_kg||0; gainDon+=o.dem||0; gainKg+=o.dem_kg||0; }
  }));
  const dcell=(don,kg)=>`${fmt(don)} Ä<br>${fmt(kg)} kg`;
  let rows=Object.entries(loss).sort((a,b)=>b[1].don-a[1].don).map(([hc,o])=>{
    const h=hubByCode[hc], bd=h?demandOf(h):0, bk=h?demandKgOf(h):0;
    return `<tr><td class="l">${esc(h?h.name:hc)}</td><td>${dcell(bd,bk)}</td><td>${dcell(bd-o.don,bk-o.kg)}</td><td class="delta-neg">â${fmt(o.don)} Ä<br>â${fmt(o.kg)} kg</td></tr>`;
  }).join("");
  let tName,tbd=0,tbk=0;
  if(T){const th=hubByCode[T];tName=th?th.name:T;tbd=th?demandOf(th):0;tbk=th?demandKgOf(th):0;} else tName="BC Má»I";
  rows+=`<tr style="font-weight:700"><td class="l">${esc(tName)}</td><td>${dcell(tbd,tbk)}</td><td>${dcell(tbd+gainDon,tbk+gainKg)}</td><td class="delta-pos">+${fmt(gainDon)} Ä<br>+${fmt(gainKg)} kg</td></tr>`;
  const tableHTML=`<table class="bal"><tr><th>BC</th><th>TrÆ°á»c</th><th>Sau</th><th>Î</th></tr>${rows}</table>`;
  const regs=[...new Set(nws.map(w=>w.region).filter(Boolean))];
  const est=planStaffEstimate(gainDon, totDon, T);
  planSnapshot={ selCodes:oldCodes, selNewCodes:codes, selNames:oldNames, newWardNames:nws.map(w=>w.name),
    regions:regs, affectedHubs:[...Object.keys(loss),...(T?[T]:[])], target:tName, totDon, totKg, gainDon, gainKg, tableHTML, staff:est,
    lossByHub:Object.entries(loss).map(([hc,o])=>({hub:hc,don:o.don,kg:o.kg})) };
  showDetail(`<h3>Quy hoáº¡ch (what-if)</h3>
    <div style="color:#6b7280">ÄÃ£ chá»n <b>${nws.length}</b> phÆ°á»ng má»i Â· <b>${fmt(totDon)}</b> ÄÆ¡n Â· <b>${fmt(totKg)}</b> kg/ngÃ y â gÃ¡n cho <b>${esc(tName)}</b></div>
    ${tableHTML}
    ${est?`<div class="note">${staffNoteHTML(est)}</div>`:""}
    <div class="note">Cáº§u phÆ°á»ng cÅ© "Nháº­p má»t pháº§n" ÄÃ£ chia 1/n theo quy Æ°á»c Â· BC = bÆ°u cá»¥c giao.</div>${PROP_BTN}`);
}
function renderPlan(){
  if(ui.planUnit==="new") return renderPlanNew();   // tÃ´ theo phÆ°á»ng má»i -> dÃ¹ng olds[] (1/n)
  const sel=[...plan.selected.keys()].map(c=>wardByCode[c]).filter(Boolean);
  renderLegend();   // cáº­p nháº­t ghi chÃº "phÆ°á»ng Äang chá»n"
  if(!sel.length){ $("#detail").classList.add("hidden"); return; }
  const T = ui.planTarget || null;                  // null = BC má»i; else hub_code ÄÃ­ch
  // TÃ¡ch 2 phÃ­a: pháº§n Láº¤Y (pv/pw) trá»« á» BC láº¥y; pháº§n GIAO (dv/dw) trá»« á» BC giao.
  const loss={};                                    // bc -> {don,kg}
  let gainDon=0, gainKg=0, totDon=0, totKg=0;
  const addLoss=(bc,don,kg)=>{ if(!bc||bc===T) return; const o=loss[bc]||(loss[bc]={don:0,kg:0}); o.don+=don; o.kg+=kg; };
  sel.forEach(w=>{
    totDon+=w.pv+w.dv; totKg+=(w.pw||0)+(w.dw||0);
    if(w.pick_hub!==T){     addLoss(w.pick_hub,     w.pv, w.pw||0); gainDon+=w.pv; gainKg+=w.pw||0; }  // Láº¤Y
    if(w.delivery_hub!==T){ addLoss(w.delivery_hub, w.dv, w.dw||0); gainDon+=w.dv; gainKg+=w.dw||0; }  // GIAO
  });
  const dcell=(don,kg)=>`${fmt(don)} Ä<br>${fmt(kg)} kg`;
  let rows=Object.entries(loss).sort((a,b)=>b[1].don-a[1].don).map(([hc,o])=>{
    const h=hubByCode[hc], bd=h?demandOf(h):0, bk=h?demandKgOf(h):0;
    return `<tr><td class="l">${esc(h?h.name:hc)}</td><td>${dcell(bd,bk)}</td><td>${dcell(bd-o.don,bk-o.kg)}</td>`+
      `<td class="delta-neg">â${fmt(o.don)} Ä<br>â${fmt(o.kg)} kg</td></tr>`;
  }).join("");
  // dÃ²ng ÄÃ­ch
  let tName,tbd=0,tbk=0;
  if(T){ const th=hubByCode[T]; tName=th?th.name:T; tbd=th?demandOf(th):0; tbk=th?demandKgOf(th):0; }
  else tName="BC Má»I";
  rows+=`<tr style="font-weight:700"><td class="l">${esc(tName)}</td><td>${dcell(tbd,tbk)}</td>`+
    `<td>${dcell(tbd+gainDon,tbk+gainKg)}</td><td class="delta-pos">+${fmt(gainDon)} Ä<br>+${fmt(gainKg)} kg</td></tr>`;
  const tableHTML=`<table class="bal"><tr><th>BC</th><th>TrÆ°á»c</th><th>Sau</th><th>Î</th></tr>${rows}</table>`;
  const regs=[...new Set(sel.map(w=>w.region).filter(Boolean))];
  // BC bá» áº£nh hÆ°á»ng (máº¥t cáº§u) + BC ÄÃ­ch -> Äá» ÄÆ°a vÃ o khung chá»¥p map
  const affectedHubs=[...Object.keys(loss), ...(T?[T]:[])];
  // phÆ°á»ng Má»I chá»©a cÃ¡c phÆ°á»ng cÅ© ÄÃ£ chá»n
  const selSet=new Set(sel.map(w=>w.ward_code));
  const newWardNames=[...new Set((DATA.rez&&DATA.rez.new_wards||[])
    .filter(nw=>(nw.old_codes||[]).some(c=>selSet.has(c))).map(nw=>nw.name))];
  // lÆ°u snapshot phá»¥c vá»¥ "LÆ°u Äá» xuáº¥t"
  const est=planStaffEstimate(gainDon, totDon, T);
  planSnapshot={ selCodes:sel.map(w=>w.ward_code), selNewCodes:[...plan.selNew.keys()], selNames:sel.map(w=>w.name),
    newWardNames, regions:regs, affectedHubs, target:tName, totDon, totKg, gainDon, gainKg, tableHTML, staff:est,
    lossByHub:Object.entries(loss).map(([hc,o])=>({hub:hc,don:o.don,kg:o.kg})) };
  showDetail(`<h3>Quy hoáº¡ch (what-if)</h3>
    <div style="color:#6b7280">ÄÃ£ chá»n <b>${sel.length}</b> phÆ°á»ng Â· <b>${fmt(totDon)}</b> ÄÆ¡n Â· <b>${fmt(totKg)}</b> kg/ngÃ y â gÃ¡n cho <b>${esc(tName)}</b></div>
    ${tableHTML}
    ${est?`<div class="note">${staffNoteHTML(est)}</div>`:""}
    <div class="note">TÃ¡ch 2 phÃ­a: pháº§n Láº¤Y trá»« á» BC láº¥y, pháº§n GIAO trá»« á» BC giao.</div>
    <button id="save-prop" onclick="saveProposal()" style="width:100%;margin-top:8px;padding:9px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:600;cursor:pointer">ð¾ LÆ°u Äá» xuáº¥t</button>`);
}
let planSnapshot=null;
// Bounds cÃ¡c phÆ°á»ng ÄÃ£ chá»n â tÃ­nh tá»« POLYGON tháº­t (geojson) Äá» chÃ­nh xÃ¡c & khÃ´ng lá» thuá»c centroid
// (nhiá»u phÆ°á»ng 1Aâ¦ HÃ  Ná»i khÃ´ng cÃ³ trong ward_centroids -> trÆ°á»c ÄÃ¢y bounds null -> khÃ´ng zoom).
// Khung bao cÃ¡c polygon (codes) trong 1 geojson, khá»p theo 'key' (ward_code hoáº·c new_ward_code)
function wardBounds(codes, geo, key){
  const set=new Set(codes||[]); if(!set.size||!geo) return null;
  let lo=[181,91], hi=[-181,-91], n=0;
  const add=pt=>{ n++; if(pt[0]<lo[0])lo[0]=pt[0]; if(pt[1]<lo[1])lo[1]=pt[1]; if(pt[0]>hi[0])hi[0]=pt[0]; if(pt[1]>hi[1])hi[1]=pt[1]; };
  for(const f of geo.features){
    const c=f.properties[key]||f.properties.ward_code||f.properties.polygon_id_code; if(!set.has(c)) continue;
    const g=f.geometry; if(!g) continue;
    const polys=g.type==="Polygon"?[g.coordinates]:(g.type==="MultiPolygon"?g.coordinates:[]);
    for(const poly of polys) for(const ring of poly) for(const pt of ring) add(pt);
  }
  return n?[lo,hi]:null;
}
function selectedBounds(codes){  // phÆ°á»ng cÅ© (fallback centroid náº¿u polygon thiáº¿u)
  let b=wardBounds(codes, DATA.geojson, "ward_code");
  if(b) return b;
  let lo=[181,91],hi=[-181,-91],n=0; (codes||[]).forEach(c=>{const p=DATA.cent&&DATA.cent[c]; if(p){n++;lo[0]=Math.min(lo[0],p[0]);lo[1]=Math.min(lo[1],p[1]);hi[0]=Math.max(hi[0],p[0]);hi[1]=Math.max(hi[1],p[1]);}});
  return n?[lo,hi]:null;
}
function captureMap(cb){
  // ÄÃ³ng khung theo polygon PHÆ¯á»NG Má»I (3321) náº¿u quy hoáº¡ch theo phÆ°á»ng má»i; else theo phÆ°á»ng cÅ©
  const sn=planSnapshot||{};
  const b = (sn.selNewCodes&&sn.selNewCodes.length)
    ? wardBounds(sn.selNewCodes, DATA.geojsonNew, "new_ward_code")
    : selectedBounds(sn.selCodes);
  if(b) map.fitBounds(b,{padding:70,maxZoom:13,duration:0});
  let done=false;
  const cap=()=>{ if(done)return; done=true; try{cb(map.getCanvas().toDataURL("image/png"));}catch(e){cb(null);} };
  map.once("idle",()=>setTimeout(cap,200));   // chá»¥p sau khi map render xong khung má»i
  setTimeout(cap,4500);                         // fallback náº¿u idle cháº­m
}
let proposalHTML="", mapImg=null;
function buildProposalHTML(img, name, comment, editable, review){
  const s=planSnapshot, now=new Date().toLocaleString("vi-VN");
  const title=(name||"").trim()||"Äá» xuáº¥t Quy hoáº¡ch BÆ°u Cá»¥c";
  const reviewBlock=(review||"").trim()
    ? `<h2>ð¤ ÄÃ¡nh giÃ¡ cá»§a AI</h2><div class="airev">${mdSafe(review)}</div>` : "";
  // Má»¥c Ghi chÃº: cháº¿ Äá» editable -> Ã´ contenteditable ngay trong report (gÃµ tháº³ng); else -> render tÄ©nh (cho PDF)
  const cmtBlock = editable
    ? `<h2>Ghi chÃº <span id="cmt-cc">${(comment||"").length}/1000</span></h2>
       <div class="cmt edit" id="cmt-edit" contenteditable="true" data-ph="Nháº­p ghi chÃº / nháº­n xÃ©t táº¡i ÄÃ¢yâ¦ (â¤1000 kÃ½ tá»±)">${esc(comment||"")}</div>`
    : ((comment||"").trim()?`<h2>Ghi chÃº</h2><div class="cmt">${esc(comment.trim())}</div>`:"");
  const cmtScript = editable
    ? `<scr`+`ipt>(function(){var e=document.getElementById('cmt-edit'),cc=document.getElementById('cmt-cc');if(!e)return;
       function upd(){var t=e.innerText.replace(/\\n$/,'');if(t.length>1000){e.innerText=t.slice(0,1000);var r=document.createRange();r.selectNodeContents(e);r.collapse(false);var sel=getSelection();sel.removeAllRanges();sel.addRange(r);t=e.innerText.replace(/\\n$/,'');}if(cc)cc.textContent=t.length+'/1000';}
       e.addEventListener('input',upd);upd();})();</scr`+`ipt>`
    : "";
  return `<!doctype html><html lang="vi"><head><meta charset="utf-8"><style>
    body{font-family:"Be Vietnam Pro",system-ui,Arial,sans-serif;color:#0f172a;margin:26px;font-size:13px;line-height:1.5}
    h1{font-size:20px;margin:0 0 2px} h2{font-size:14px;margin:18px 0 6px;border-bottom:2px solid #2563eb;padding-bottom:3px}
    .meta{color:#475569;font-size:12px} table{border-collapse:collapse;width:100%;margin:6px 0;font-size:12px}
    th,td{border:1px solid #e5e7eb;padding:5px 8px;text-align:right} th:first-child,td:first-child{text-align:left}
    th{background:#f1f5f9} .delta-neg{color:#dc2626}.delta-pos{color:#16a34a}
    img{max-width:100%;border:1px solid #e5e7eb;border-radius:8px;margin-top:4px}
    .sum{display:flex;gap:12px;flex-wrap:wrap;margin:8px 0}
    .sum>div{background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;padding:8px 14px;min-width:120px}
    .sum b{font-size:17px;display:block;color:#111827} .sum span{font-size:11px;color:#6b7280}
    .wards{color:#475569;font-size:11.5px}
    .cmt{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:8px 12px;white-space:pre-wrap}
    .cmt.edit{outline:none;min-height:20px;cursor:text} .cmt.edit:focus{border-color:#f59e0b;box-shadow:0 0 0 3px rgba(245,158,11,.25)}
    .cmt.edit:empty:before{content:attr(data-ph);color:#b45309;opacity:.65}
    #cmt-cc{float:right;font-size:11px;color:#94a3b8;font-weight:400}
    .airev{background:#faf5ff;border:1px solid #e9d5ff;border-left:4px solid #7c3aed;border-radius:8px;padding:8px 14px;font-size:12.5px}
    .airev p{margin:4px 0} .airev ul{margin:4px 0;padding-left:20px} .airev li{margin:2px 0}
  </style></head><body>
  <h1>${esc(title)}</h1>
  <div class="meta">NgÆ°á»i láº­p: <b>${esc(ui.username||"")}</b> Â· ${now} Â· VÃ¹ng: ${esc(s.regions.join(", ")||"â")}</div>
  <h2>TÃ³m táº¯t</h2>
  <div class="sum">
    <div><span>PhÆ°á»ng tÃ¡i quy hoáº¡ch</span><b>${fmt(s.selCodes.length)}</b></div>
    <div><span>Cáº§u di chuyá»n</span><b>${fmt(s.totDon)} ÄÆ¡n/ngÃ y</b></div>
    <div><span>Khá»i lÆ°á»£ng</span><b>${fmt(s.totKg)} kg/ngÃ y</b></div>
    <div><span>GÃ¡n cho</span><b>${esc(s.target)}</b></div>
    ${s.staff?`<div><span>Äá»nh biÃªn cáº§n</span><b>${s.staff.isNew?"~"+fmt(s.staff.addNV)+" NV (BC má»i)":"+~"+fmt(s.staff.addNV)+" NV"}</b></div>`:""}
  </div>
  <h2>áº¢nh hÆ°á»ng trÆ°á»c / sau khi quy hoáº¡ch (theo BC Â· ÄÆ¡n + kg)</h2>
  ${s.tableHTML}
  <h2>Báº£n Äá» khu vá»±c thay Äá»i</h2>
  ${img?`<img src="${img}">`:"<i>(khÃ´ng chá»¥p ÄÆ°á»£c báº£n Äá»)</i>"}
  <h2>PhÆ°á»ng má»i tÃ¡i quy hoáº¡ch (${(s.newWardNames||[]).length})</h2>
  <div class="wards">${(s.newWardNames||[]).map(esc).join(" Â· ")||"â"}</div>
  <h2>PhÆ°á»ng cÅ© (nguá»n cáº§u) (${s.selCodes.length})</h2>
  <div class="wards">${s.selNames.map(esc).join(" Â· ")}</div>
  ${reviewBlock}
  ${cmtBlock}
  ${cmtScript}
  </body></html>`;
}
function saveProposal(){
  if(!planSnapshot||!planSnapshot.selCodes.length) return;
  if(ui.tab!=="map") switchTab("map");
  const btn=$("#save-prop"); if(btn){btn.disabled=true; btn.textContent="ð¸ Äang chá»¥p báº£n Äá»â¦";}
  captureMap(img=>{
    mapImg=img;
    const nameEl=$("#prop-name"); if(nameEl) nameEl.value="";   // Äá» trá»ng, user tá»± Äáº·t tÃªn (báº¯t buá»c)
    aiReview=null;   // reset ÄÃ¡nh giÃ¡ AI cho láº§n má» má»i
    wirePropInputs(); renderPreviewEditable();   // dá»±ng report vá»i Ã´ Ghi chÃº gÃµ tháº³ng ÄÆ°á»£c
    $("#prop-status").textContent="";
    lastPdf=null;   // reset nÃºt xÃ¡c nháº­n (láº§n má» má»i)
    const rv=$("#prop-review"); if(rv){ rv.disabled=false; rv.style.display=""; rv.textContent="ð¤ AI ÄÃ¡nh giÃ¡ Äá» xuáº¥t"; }
    const c=$("#prop-confirm"); if(c){ c.disabled=true; c.textContent="â XÃ¡c nháº­n Äá» xuáº¥t (gá»­i OA)"; c.style.background="#16a34a"; c.onclick=confirmProposal; }  // khoÃ¡ Äáº¿n khi Äá»§ tÃªn + ÄÃ¡nh giÃ¡ AI
    modalOpen($("#prop-modal"));
    if(btn){btn.disabled=false; btn.textContent="ð¾ LÆ°u Äá» xuáº¥t";}
  });
}
let _propWired=false, aiReview=null;
// Dá»±ng báº£n xem trÆ°á»c cÃ³ Ã´ Ghi chÃº contenteditable. Chá» gá»i khi Má» dialog -> ghi chÃº rá»ng, chÆ°a cÃ³ ÄÃ¡nh giÃ¡.
function renderPreviewEditable(){
  const fr=$("#prop-frame"); if(fr) fr.srcdoc=buildProposalHTML(mapImg, ($("#prop-name")||{}).value||"", "", true, null);
}
// Dá»±ng láº¡i preview GIá»® ghi chÃº Äang gÃµ + kÃ¨m ÄÃ¡nh giÃ¡ AI (gá»i sau khi cÃ³ ÄÃ¡nh giÃ¡).
function rebuildPropPreview(){
  const cmt=readComment(), fr=$("#prop-frame");
  if(fr) fr.srcdoc=buildProposalHTML(mapImg, ($("#prop-name")||{}).value||"", cmt, true, aiReview);
}
// NÃºt XÃ¡c nháº­n chá» má» khi Äá»¦ tÃªn + ÄÃ£ cÃ³ ÄÃ¡nh giÃ¡ AI (báº¯t buá»c).
function updatePropConfirm(){
  const c=$("#prop-confirm"); if(!c||lastPdf) return;
  c.disabled = !((($("#prop-name")||{}).value||"").trim() && aiReview);
}
// Gá»£i Ã½ vá» trÃ­: trá»ng tÃ¢m ÄÆ¡n hÃ ng (centroid phÆ°á»ng cÅ© weighted theo cáº§u) + phÆ°á»ng cáº§u cao
function planGeoHint(s){
  const wd=c=>wardByCode[c];
  let sx=0,sy=0,sw=0;
  (s.selCodes||[]).forEach(c=>{ const p=DATA.cent&&DATA.cent[c], w=wd(c);
    if(p&&w){ const d=(w.pv+w.dv)||0; sx+=p[0]*d; sy+=p[1]*d; sw+=d; } });
  let center=null;
  if(sw){ const C=[sx/sw,sy/sw]; let bd=Infinity;
    (s.selNewCodes||[]).forEach((code,i)=>{ const nc=newCent[code]; if(!nc) return;
      const dx=nc.p[0]-C[0],dy=nc.p[1]-C[1],d=dx*dx+dy*dy; if(d<bd){bd=d; center=(s.newWardNames||[])[i]||code;} });
    if(!center){ let b2=Infinity; (s.selCodes||[]).forEach(c=>{ const p=DATA.cent&&DATA.cent[c]; if(!p) return;
      const dx=p[0]-C[0],dy=p[1]-C[1],d=dx*dx+dy*dy; if(d<b2){b2=d; center=(wd(c)||{}).name||c;} }); }
  }
  const top=(s.selCodes||[]).map(c=>({n:(wd(c)||{}).name||c, d:((wd(c)||{}).pv||0)+((wd(c)||{}).dv||0)}))
    .sort((a,b)=>b.d-a.d).slice(0,4).filter(x=>x.d>0).map(x=>`${x.n} (${fmt(x.d)} ÄÆ¡n)`);
  return {center, top};
}
async function reviewProposal(){
  if(!planSnapshot) return;
  const btn=$("#prop-review"), st=$("#prop-status");
  if(btn){ btn.disabled=true; btn.textContent="â³ AI Äang ÄÃ¡nh giÃ¡â¦"; }
  if(st) st.textContent="";
  const s=planSnapshot, geo=planGeoHint(s);
  const payload={ target:s.target, totDon:s.totDon, totKg:s.totKg, gainDon:s.gainDon, gainKg:s.gainKg,
    newWardCount:(s.newWardNames||[]).length, oldWardCount:s.selCodes.length,
    staffAdd:s.staff?s.staff.addNV:0, staffMed:s.staff?s.staff.med:0, staffCur:s.staff?s.staff.curStaff:null,
    isNew:s.staff?s.staff.isNew:(s.target==="BC Má»I"), affectedHubs:s.affectedHubs||[], losses:s.lossByHub||[],
    geoCenter:geo.center||"", topWards:geo.top };
  try{
    const r=await fetch("/proposal/review",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify(payload)}).then(r=>r.json());
    if(r.review){ aiReview=r.review; rebuildPropPreview(); toast("AI ÄÃ£ ÄÃ¡nh giÃ¡ xong Äá» xuáº¥t.","ok"); }
    else if(st){ st.textContent="â  "+(r.error||"lá»i ÄÃ¡nh giÃ¡"); toast("ÄÃ¡nh giÃ¡ tháº¥t báº¡i: "+(r.error||"lá»i"),"err"); }
  }catch(e){ if(st) st.textContent="â  Lá»i káº¿t ná»i server."; toast("Lá»i káº¿t ná»i server khi ÄÃ¡nh giÃ¡.","err"); }
  // ÄÃ¡nh giÃ¡ CHá» 1 Láº¦N: thÃ nh cÃ´ng -> áº¨N nÃºt (gá»p cÃ²n 1 chá» bÃ¡o á» status, chá»«a chá» cho Ã´ tÃªn).
  // Muá»n ÄÃ¡nh giÃ¡ láº¡i pháº£i thoÃ¡t dialog & táº¡o láº¡i Äá» xuáº¥t.
  if(btn){ if(aiReview){ btn.style.display="none"; if(st) st.textContent="ð¤ ÄÃ£ ÄÃ¡nh giÃ¡ â"; }
           else { btn.disabled=false; btn.style.display=""; btn.textContent="ð¤ AI ÄÃ¡nh giÃ¡ Äá» xuáº¥t"; } }
  updatePropConfirm();
}
// Äá»c ghi chÃº trá»±c tiáº¿p tá»« Ã´ contenteditable trong iframe (same-origin srcdoc).
function readComment(){
  try{ const d=$("#prop-frame").contentDocument, e=d&&d.getElementById("cmt-edit");
    return e ? e.innerText.replace(/Â /g," ").replace(/\n+$/,"").trim().slice(0,1000) : ""; }
  catch(e){ return ""; }
}
function wirePropInputs(){
  if(_propWired) return; _propWired=true;
  // Äá»i tÃªn -> chá» vÃ¡ tiÃªu Äá» <h1> trong iframe (KHÃNG rerender, káº»o máº¥t ghi chÃº Äang gÃµ) + báº­t/khoÃ¡ nÃºt xÃ¡c nháº­n
  const nm=$("#prop-name"); if(nm) nm.oninput=()=>{
    updatePropConfirm();
    try{ const h=$("#prop-frame").contentDocument.querySelector("h1"); if(h) h.textContent=nm.value.trim()||"Äá» xuáº¥t Quy hoáº¡ch BÆ°u Cá»¥c"; }catch(e){}
  };
}
let lastPdf=null;   // {b64, name} cá»§a Äá» xuáº¥t vá»«a táº¡o
async function confirmProposal(){
  const btn=$("#prop-confirm"), st=$("#prop-status");
  if(btn){btn.disabled=true; btn.textContent="â³ Äang táº¡o PDF & gá»­iâ¦";}
  if(st) st.textContent="";
  const title=(($("#prop-name")||{}).value||"").trim();
  if(!title){ if(st) st.textContent="â  Vui lÃ²ng nháº­p tÃªn Äá» xuáº¥t."; const n=$("#prop-name"); if(n) n.focus();
    if(btn){btn.disabled=false; btn.textContent="â XÃ¡c nháº­n Äá» xuáº¥t (gá»­i OA)";} return; }
  if(!aiReview){ if(st) st.textContent="â  Cáº§n báº¥m 'ð¤ AI ÄÃ¡nh giÃ¡ Äá» xuáº¥t' trÆ°á»c khi gá»­i.";
    if(btn){btn.disabled=false; btn.textContent="â XÃ¡c nháº­n Äá» xuáº¥t (gá»­i OA)";} return; }
  proposalHTML=buildProposalHTML(mapImg, title, readComment(), false, aiReview);   // báº£n sáº¡ch (escaped, khÃ´ng editable) cho PDF
  try{
    const r=await fetch("/proposal",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify({html:proposalHTML, title})}).then(r=>r.json());
    if(st) st.textContent = r.ok ? (r.sent? "â ÄÃ£ gá»­i vÃ o OA GTalk." : ("â ÄÃ£ táº¡o PDF. "+(r.note||""))) : ("â  "+(r.error||"lá»i"));
    if(r.ok) toast(r.sent? "ÄÃ£ gá»­i Äá» xuáº¥t vÃ o OA GTalk." : "ÄÃ£ táº¡o PDF Äá» xuáº¥t.","ok");
    else toast("Gá»­i Äá» xuáº¥t tháº¥t báº¡i: "+(r.error||"lá»i"),"err");
    if(r.pdf){   // cÃ³ PDF -> Äá»i nÃºt thÃ nh Download
      lastPdf={b64:r.pdf, name:r.filename||"de-xuat.pdf"};
      if(btn){ btn.disabled=false; btn.textContent="â¬ Download Äá» xuáº¥t"; btn.style.background="#2563eb"; btn.onclick=downloadProposal; }
      return;
    }
  }catch(e){ if(st) st.textContent="â  Lá»i káº¿t ná»i server."; toast("Lá»i káº¿t ná»i server khi gá»­i Äá» xuáº¥t.","err"); }
  if(btn){btn.disabled=false; btn.textContent="â XÃ¡c nháº­n Äá» xuáº¥t (gá»­i OA)";}
}
function downloadProposal(){
  if(!lastPdf) return;
  const bin=atob(lastPdf.b64), arr=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
  const url=URL.createObjectURL(new Blob([arr],{type:"application/pdf"}));
  const a=document.createElement("a"); a.href=url; a.download=lastPdf.name; document.body.appendChild(a); a.click();
  a.remove(); setTimeout(()=>URL.revokeObjectURL(url),4000);
}
// "GÃ¡n cá»¥m cho" = combobox cÃ³ search (native select khÃ´ng tÃ¬m/scroll ná»i ~1166 BC khi admin)
let PT_HUBS=[];
function fillPlanTargets(){
  PT_HUBS=DATA.hubs.filter(h=>h.assigned&&h.type==="express"&&regOK(h._region))
    .map(h=>({code:h.hub_code, name:h.name||h.hub_code})).sort((a,b)=>a.name.localeCompare(b.name));
  if(ui.planTarget && !PT_HUBS.some(h=>h.code===ui.planTarget)) ui.planTarget=null;   // ÄÃ­ch ngoÃ i vÃ¹ng -> reset
  syncPtInput();
}
function syncPtInput(){
  const inp=$("#pt-input"); if(!inp) return;
  const h=ui.planTarget && PT_HUBS.find(x=>x.code===ui.planTarget);
  inp.value = h ? h.name : "";
}
function renderPtList(q){
  const box=$("#pt-list"); if(!box) return;
  q=(q||"").trim().toLowerCase();
  let items=PT_HUBS;
  if(q) items=items.filter(h=>h.name.toLowerCase().includes(q)||h.code.toLowerCase().includes(q));
  const cap=60, shown=items.slice(0,cap);
  box.innerHTML=`<div class="pt-item" data-code="">â BC má»i â</div>`+
    shown.map(h=>`<div class="pt-item" data-code="${esc(h.code)}">${esc(h.name)} <span class="pt-code">${esc(h.code)}</span></div>`).join("")+
    (items.length>cap?`<div class="pt-more">â¦ cÃ²n ${items.length-cap} BC â gÃµ thÃªm Äá» lá»c</div>`:"");
  box.querySelectorAll(".pt-item").forEach(el=>el.onclick=()=>{
    ui.planTarget=el.dataset.code||null; syncPtInput(); hidePtList(); renderPlan();
  });
  box.style.display="block";
}
function hidePtList(){ const b=$("#pt-list"); if(b) b.style.display="none"; syncPtInput(); }

// ---------- UI chrome ----------
function renderTopbar(){
  const m=DATA.meta, c=m.counts||{};
  $("#meta").textContent=`${m.date} Â· ${fmt(c.express_assigned)} BC Â· ${fmt(c.wards)} phÆ°á»ng`;
  $("#meta").title=`${m.date} Â· ${fmt(c.express_assigned)} BC express Â· ${fmt(c.orphan_express)} chuyÃªn dá»¥ng Â· ${fmt(c.wards)} phÆ°á»ng`;
  const d=m.diff||{};
  $("#diff").textContent = d.first_run?"(láº§n cháº¡y Äáº§u)":(d.summary||"");
}
function renderLegend(){
  let h=`<div class="lg"><b style="font-size:12px">ChÃº giáº£i</b></div>`;
  const items = ui.colormode==="region"
    ? Object.entries(REGION_COLORS).map(([k,c])=>[`VÃ¹ng ${k}`,c,"circle"])
    : [["BÆ°u cá»¥c",ROLE_COLORS.territorial,"circle"],["ChuyÃªn Láº¤Y",ROLE_COLORS.pickup_only,"circle"],
       ["Cá»ng ká»nh/GIAO",ROLE_COLORS.bulky_delivery,"circle"],["ChuyÃªn dá»¥ng há»n há»£p",ROLE_COLORS.special_mixed,"circle"],
       ["Kho HÃ ng Náº·ng (B2B)",TYPE_COLORS.B2B,"circle"],["Kho Trung Chuyá»n/Chuyá»n Tiáº¿p",TYPE_COLORS.transit,"circle"]];
  items.forEach(([t,c,sh])=>h+=`<div class="lg"><span class="sw ${sh}" style="background:rgb(${c.join(',')})"></span>${t}</div>`);
  if(effColorMode()==="hub") h+=`<div class="lg" style="color:#6b7280">ð¨ Phá»§ phÆ°á»ng: má»i mÃ u = 1 bÆ°u cá»¥c (lÃ£nh thá» giao)</div>`;
  if(plan.selected.size) h+=`<div class="lg"><span class="sw" style="background:rgb(220,38,38)"></span>PhÆ°á»ng Äang chá»n (tÃ´ tay)</div>`;
  if(ui.jt) h+=`<div class="lg"><span class="sw circle" style="background:rgb(20,20,20)"></span>BC J&T (tham kháº£o)</div>`;
  if(ui.optov.on){
    h+=`<div class="lg"><b style="font-size:12px">âï¸ Optimizer</b></div>`;
    if(ui.optov.markers){ h+=`<div class="lg"><span class="sw circle" style="background:transparent;border:2px solid rgb(220,38,38)"></span>BC nÃªn ÄÃ³ng/rÃ  soÃ¡t</div>`;
      h+=`<div class="lg"><span class="sw circle" style="background:transparent;border:2px solid rgb(245,158,11)"></span>BC nÃªn tÃ¡ch/má» rá»ng</div>`; }
    if(ui.optov.reassign) h+=`<div class="lg"><span class="sw" style="background:rgb(37,99,235)"></span>PhÆ°á»ng nÃªn Äá»i BC â</div>`;
    if(ui.optov.merge) h+=`<div class="lg"><span class="sw" style="background:rgb(220,38,38)"></span>Gá»p khi ÄÃ³ng BC</div>`;
    if(ui.optov.ws4){ h+=`<div class="lg"><span class="sw circle" style="background:rgb(220,38,38)"></span>WS: Khoáº£ng cÃ¡ch xa</div>`;
      h+=`<div class="lg"><span class="sw circle" style="background:rgb(16,185,129)"></span>WS: Greenfield</div>`;
      h+=`<div class="lg"><span class="sw circle" style="background:rgb(245,158,11)"></span>WS: Äá»i Äáº§u J&T</div>`; }
  }
  $("#legend").innerHTML=h;
}
function renderDQ(){
  const m=DATA.meta;
  const noGeo=DATA.hubs.filter(h=>h.assigned&&h.missing_geo).length;
  const noRe=DATA.hubs.filter(h=>h.assigned&&h.missing_realestate).length;
  const orphanNoGeo=DATA.hubs.filter(h=>h.type==="express"&&h.role!=="territorial"&&!h.has_geo).length;
  let s=`<b>Data-quality</b><br>`;
  if(!DATA.geojson) s+=`â¸ ChÆ°a cÃ³ polygon â layer phÆ°á»ng táº¯t. Äáº·t <code>wards.geojson</code> vÃ o data/out.<br>`;
  s+=`â¸ ${noGeo} BC gÃ¡n phÆ°á»ng thiáº¿u toáº¡ Äá»<br>â¸ ${noRe} BC thiáº¿u máº·t báº±ng<br>â¸ ${orphanNoGeo} BC chuyÃªn dá»¥ng thiáº¿u toáº¡ Äá» (khÃ´ng cháº¥m ÄÆ°á»£c)`;
  (m.warnings||[]).forEach(w=>s+=`<br>â  ${w}`);
  $("#dq").innerHTML=s;
}
// ---------- Scorecard theo vÃ¹ng ----------
function daysToExpiry(s){
  const m=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec((s||"").trim()); if(!m) return null;
  if(+m[3]<2000) return null;   // 30/12/1899 = Ã´ ngÃ y trá»ng cá»§a Sheets -> coi nhÆ° khÃ´ng cÃ³ háº¡n
  const d=new Date(+m[3], +m[2]-1, +m[1]);
  const today=new Date(DATA.meta.date||Date.now());
  return Math.round((d-today)/86400000);
}
function aggregateRegions(){
  const wardCnt={}; DATA.wards.forEach(w=>wardCnt[w.region]=(wardCnt[w.region]||0)+1);
  const A={};
  DATA.hubs.forEach(h=>{
    if(!(h.assigned&&h.type==="express")) return; const r=h._region; if(!r) return;
    // volRE/wtRE = vol/kg cá»§a BC CÃ máº·t báº±ng (Äá»ng/ÄÆ¡n, ÄÆ¡n/mÂ², kg/mÂ² Äá»ng bá» tá»­-máº«u, khÃ´ng phá»ng khi thiáº¿u RE)
    const a=A[r]||(A[r]={region:r,bc:0,phuong:wardCnt[r]||0,terr:0,vol:0,volS:0,volRE:0,wt:0,wtRE:0,area:0,rent:0,exp90:0,nore:0,staff:0});
    a.bc++; a.terr+=demandOf(h); a.staff+=h.staff||0;
    if(h.actual){ a.vol+=h.actual.pv+h.actual.dv; a.wt+=h.actual.pw+h.actual.dw;
      if(h.staff) a.volS+=h.actual.pv+h.actual.dv; }   // volS = volume cá»§a BC CÃ staff (Äá» ÄÆ¡n/NV khÃ´ng bá» thá»i)
    if(h.realestate&&h.realestate.usable_area){ a.area+=h.realestate.usable_area; a.rent+=h.realestate.rent||0;
      if(h.actual){ a.volRE+=h.actual.pv+h.actual.dv; a.wtRE+=h.actual.pw+h.actual.dw; }
      const dd=daysToExpiry(h.realestate.expiry); if(dd!=null&&dd>=0&&dd<=90) a.exp90++; }
    else a.nore++;
  });
  return Object.values(A).map(a=>({...a,
    dpo: a.volRE? a.rent/(a.volRE*30):0, dpm: a.area? a.volRE/a.area:0,
    dpw: a.wtRE?  a.rent/(a.wtRE*30):0,  kgm2: a.area? a.wtRE/a.area:0,
    dpn: a.staff? a.volS/a.staff:0 }));   // ÄÆ¡n/NV/ngÃ y (nÄng suáº¥t; chá» vol cá»§a BC cÃ³ staff)
}
const DECIMAL=new Set(["dpm","kgm2"]);   // cá»t sá» láº» 1 chá»¯ sá»
const SC_COLS=[["region","VÃ¹ng"],["bc","BC"],["phuong","PhÆ°á»ng"],
  ["vol","Volume (ÄÆ¡n)/ngÃ y"],["wt","Khá»i lÆ°á»£ng (kg)/ngÃ y"],["area","mÂ²"],["staff","NV"],
  ["dpo","Äá»ng/ÄÆ¡n"],["dpm","ÄÆ¡n/mÂ²"],["dpw","Äá»ng/kg"],["kgm2","Kg/mÂ²"],["dpn","ÄÆ¡n/NV"],
  ["exp90","Há»£p Äá»ng â¤ 90 ngÃ y"]];
let scSort={key:"vol",desc:true};
function renderScorecard(){
  const all=aggregateRegions();
  const rows=all.filter(r=>regOK(r.region));
  const p=(arr,q)=>{const s=arr.slice().sort((a,b)=>a-b);return s[Math.floor(q*(s.length-1))];};
  const TH={};   // ngÆ°á»¡ng p75/p25 tÃ­nh trÃªn TOÃN Bá» vÃ¹ng (khÃ´ng Äá»i theo filter)
  ["dpo","dpm","dpw","kgm2","dpn"].forEach(k=>{const v=all.map(r=>r[k]); TH[k]={hi:p(v,.75),lo:p(v,.25)};});
  rows.sort((a,b)=>{ const k=scSort.key, va=a[k], vb=b[k];
    return (typeof va==="string"? va.localeCompare(vb): va-vb)*(scSort.desc?-1:1); });
  const tot={bc:0,phuong:0,terr:0,vol:0,volS:0,volRE:0,wt:0,wtRE:0,area:0,rent:0,exp90:0,nore:0,staff:0};
  rows.forEach(r=>["bc","phuong","terr","vol","volS","volRE","wt","wtRE","area","rent","exp90","nore","staff"].forEach(k=>tot[k]+=r[k]));
  tot.dpo=tot.volRE?tot.rent/(tot.volRE*30):0; tot.dpm=tot.area?tot.volRE/tot.area:0;
  tot.dpw=tot.wtRE?tot.rent/(tot.wtRE*30):0;   tot.kgm2=tot.area?tot.wtRE/tot.area:0;
  tot.dpn=tot.staff?tot.volS/tot.staff:0;  tot.region="Tá»NG";
  const cell=(r,k)=>{
    if(k==="region") return `<td>${r.region}</td>`;
    let cls="";
    if(k==="dpo"||k==="dpw"){ if(r[k]>=TH[k].hi) cls="hot"; }        // chi phÃ­ cao = xáº¥u
    if(k==="dpm"||k==="kgm2"){ if(r[k]>=TH[k].hi)cls="hot"; else if(r[k]<=TH[k].lo)cls="cold"; } // cháº­t/dÆ°
    if(k==="dpn"){ if(r[k]>=TH[k].hi)cls="hot"; else if(r[k]<=TH[k].lo)cls="cold"; }   // ÄÆ¡n/NV cao=thiáº¿u ngÆ°á»i, tháº¥p=thá»«a
    const txt = DECIMAL.has(k)? r[k].toFixed(1) : fmt(r[k]);
    // kÃªnh dá»± phÃ²ng ngoÃ i mÃ u (a11y mÃ¹ mÃ u): â² cao / â¼ tháº¥p + tooltip
    const mk = cls==="hot" ? '<span class="mk" aria-hidden="true">â²</span>' : cls==="cold" ? '<span class="mk" aria-hidden="true">â¼</span>' : "";
    const ti = cls==="hot" ? ' title="Cao (â¥P75)"' : cls==="cold" ? ' title="Tháº¥p (â¤P25)"' : "";
    return `<td class="${cls}"${ti}>${mk}${txt}</td>`;
  };
  const th=SC_COLS.map(([k,l])=>`<th data-k="${k}">${l}${scSort.key===k?(scSort.desc?" â¾":" â´"):""}</th>`).join("");
  const body=rows.map(r=>`<tr class="region-row" data-r="${r.region}">`+
    SC_COLS.map(([k])=>cell(r,k)).join("")+`</tr>`).join("");
  const totrow=`<tr class="total"><td>Tá»NG (${rows.length} vÃ¹ng)</td>`+
    SC_COLS.slice(1).map(([k])=> DECIMAL.has(k)?`<td>${tot[k].toFixed(1)}</td>`:`<td>${fmt(tot[k])}</td>`).join("")+`</tr>`;
  $("#sc-table").innerHTML=`<table class="sc"><thead><tr>${th}</tr></thead><tbody>${body}${totrow}</tbody></table>`;
  $("#sc-table").querySelectorAll("th").forEach(t=>t.onclick=()=>{ const k=t.dataset.k;
    if(scSort.key===k) scSort.desc=!scSort.desc; else {scSort.key=k;scSort.desc=true;} renderScorecard(); });
  $("#sc-table").querySelectorAll("tr.region-row").forEach(tr=>tr.onclick=()=>filterRegion(tr.dataset.r));
  renderSubTables();
}

// ---------- 3 báº£ng per-BC: Sáº¯p háº¿t háº¡n / Sáº¯p quÃ¡ táº£i / DÆ° thá»«a ----------
function bcList(){
  const K=kgPerOrd()||1;
  return DATA.hubs.filter(h=>h.assigned && h.type==="express").map(h=>{
    const ac=h.actual||{pv:0,pw:0,dv:0,dw:0}, re=h.realestate;
    const vol=ac.pv+ac.dv, wt=ac.pw+ac.dw, area=re?re.usable_area:0;
    const wl=(0.4*ac.pv+ac.dv)+(0.4*ac.pw+ac.dw)/K;   // táº£i quy Äá»i (ÄÆ¡n-tÄ): ÄÆ¡n & kg láº¥y Ã0.4, kg quy ra ÄÆ¡n
    return { code:h.hub_code, name:h.name||h.hub_code, region:h._region, province:h._province,
      vol, wt, area, rent:re?re.rent:0, expiry:re?re.expiry:"", days:re?daysToExpiry(re.expiry):null,
      dpm: area?vol/area:0, kgm2: area?wt/area:0, diemM2: area?wl/area:0, hasRE:!!re };
  });
}
function pct(arr,q){ const s=arr.slice().sort((a,b)=>a-b); return s.length?s[Math.floor(q*(s.length-1))]:0; }
// báº£ng cÃ³ nÃºt thu/xá» + há»p cuá»n (hiá»n ~10 dÃ²ng, cuá»n xem tiáº¿p). startCollapsed=true -> máº·c Äá»nh gáº­p.
function subTable(rows, cols, cap, startCollapsed, tall){
  const th=cols.map(c=>`<th class="${c.l?'l':''}">${c.h}</th>`).join("");
  const body=rows.slice(0,cap).map(r=>`<tr class="region-row" data-code="${esc(r.code)}">`+
    cols.map(c=>`<td class="${c.l?'l':''}">${c.raw?c.f(r):esc(c.f(r))}</td>`).join("")+`</tr>`).join("");  // raw=true: HTML do code kiá»m soÃ¡t
  const cnt = rows.length>cap?`${cap}/${rows.length}`:`${rows.length}`;
  const arrow = startCollapsed?"â¸ Xem chi tiáº¿t Â· ":"â¾ Thu gá»n Â· ";
  const note=`<div class="sc-note">Click 1 dÃ²ng Äá» xem trÃªn báº£n Äá».</div>`;
  const cls = `tbl-body${startCollapsed?' collapsed':''}${tall?' tall':''}`;
  return `<button class="tbl-toggle" data-n="${cnt}" onclick="toggleTbl(this)">${arrow}${cnt} dÃ²ng</button>`+
    `<div class="${cls}">${note}`+
    `<table class="sc"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function toggleTbl(btn){
  const body=btn.parentNode.querySelector(".tbl-body"); if(!body) return;
  const collapsed=body.classList.toggle("collapsed");   // true náº¿u vá»«a gáº­p láº¡i
  btn.textContent=(collapsed?"â¸ Xem chi tiáº¿t Â· ":"â¾ Thu gá»n Â· ")+btn.dataset.n+" dÃ²ng";
}
function renderSubTables(){
  const bcs=bcList().filter(b=>regOK(b.region));
  const withRE=bcs.filter(b=>b.hasRE && b.area>0 && b.vol>0);
  const eHi=pct(withRE.map(b=>b.diemM2),.75), eLo=pct(withRE.map(b=>b.diemM2),.25);   // ngÆ°á»¡ng theo Äiá»m/mÂ²
  const base=[{h:"VÃ¹ng",f:r=>r.region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},
    {h:"ID BC",f:r=>r.code,l:1},{h:"TÃªn BC",f:r=>r.name,l:1}];
  const loadCols=[...base,{h:"Volume (ÄÆ¡n)/ngÃ y",f:r=>fmt(r.vol)},{h:"Khá»i lÆ°á»£ng (kg)/ngÃ y",f:r=>fmt(r.wt)},
    {h:"mÂ²",f:r=>fmt(r.area)},{h:"ÄÆ¡n/mÂ²",f:r=>r.dpm.toFixed(1)},{h:"Kg/mÂ²",f:r=>r.kgm2.toFixed(1)},{h:"Äiá»m/mÂ²",f:r=>r.diemM2.toFixed(1)}];
  const CAP=5000;   // render háº¿t (há»p cuá»n lo pháº§n dÃ i), khÃ´ng cáº¯t dá»¯ liá»u
  // Sáº¯p háº¿t háº¡n: cÃ²n láº¡i â¤ 90 ngÃ y (gá»m ÄÃ£ háº¿t háº¡n)
  const exp=bcs.filter(b=>b.days!=null && b.days<=90).sort((a,b)=>a.days-b.days);
  $("#sc-expire").innerHTML=subTable(exp,[...base,
    {h:"Háº¡n HÄ",f:r=>r.expiry||"â",l:1},
    {h:"CÃ²n láº¡i",raw:true,f:r=>r.days<0?`<span class="warn">ÄÃ£ háº¿t ${-r.days}d</span>`:`${r.days}d`},
    {h:"Tiá»n thuÃª",f:r=>fmt(r.rent)},{h:"mÂ²",f:r=>fmt(r.area)}],CAP,true);
  // Sáº¯p quÃ¡ táº£i: Äiá»m/mÂ² â¥ P75 (táº£i quy Äá»i ÄÆ¡n+kg trÃªn máº·t báº±ng)
  $("#sc-overload").innerHTML=subTable(withRE.filter(b=>b.diemM2>=eHi).sort((a,b)=>b.diemM2-a.diemM2),loadCols,CAP,true);
  // DÆ° thá»«a: Äiá»m/mÂ² â¤ P25
  $("#sc-surplus").innerHTML=subTable(withRE.filter(b=>b.diemM2<=eLo).sort((a,b)=>a.diemM2-b.diemM2),loadCols,CAP,true);
  document.querySelectorAll("#sc-expire .region-row,#sc-overload .region-row,#sc-surplus .region-row")
    .forEach(tr=>tr.onclick=()=>{ const h=hubByCode[tr.dataset.code]; if(!h||!h.has_geo) return;
      switchTab("map"); map.flyTo({center:[h.lng,h.lat],zoom:13}); clickHub(h); });
}
function filterRegion(r){
  ui.regionFilter=r; switchTab("map"); draw();
  const pts=visibleHubs(); if(pts.length){
    const lons=pts.map(h=>h.lng), lats=pts.map(h=>h.lat);
    map.fitBounds([[Math.min(...lons),Math.min(...lats)],[Math.max(...lons),Math.max(...lats)]],{padding:60,duration:800}); }
  showRegionChip();
}
function showRegionChip(){
  $("#diff").innerHTML = ui.regionFilter? `Lá»c vÃ¹ng: <b>${ui.regionFilter}</b> <span class="chip" id="clearf">â bá» lá»c</span>`:"";
  const c=$("#clearf"); if(c) c.onclick=()=>{ ui.regionFilter=null; draw(); renderTopbar(); };
}
function setRegionFilter(r){
  ui.regionFilter = r || null;
  fillPlanTargets();   // danh sÃ¡ch BC ÄÃ­ch thu háº¹p theo vÃ¹ng Äang lá»c
  renderLegend();      // cáº­p nháº­t chÃº giáº£i (1 vÃ¹ng -> tÃ´ theo BC)
  if(map && map.loaded && map.loaded()) draw();
  if(ui.tab==="scorecard") renderScorecard();
  else if(ui.tab==="optimizer") renderOptimizer();
  else if(ui.tab==="rezone") renderRezone();
  else if(ui.regionFilter){ // map: fit tá»i hub cá»§a vÃ¹ng
    const pts=visibleHubs();
    if(pts.length){ const lo=pts.map(h=>h.lng),la=pts.map(h=>h.lat);
      map.fitBounds([[Math.min(...lo),Math.min(...la)],[Math.max(...lo),Math.max(...la)]],{padding:60,duration:700}); }
  }
}
function switchTab(name){
  ui.tab=name;
  document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("active",t.dataset.tab===name));
  $("#scorecard").classList.toggle("hidden", name!=="scorecard");
  $("#optimizer").classList.toggle("hidden", name!=="optimizer");
  $("#rezone").classList.toggle("hidden", name!=="rezone");
  $("#chatlog").classList.toggle("hidden", name!=="chatlog");
  $("#clog-btn")?.classList.toggle("on", name==="chatlog");
  $("#controls").style.display = (name==="map") ? "" : "none";
  const ctgl=$("#ctrl-toggle"); if(ctgl) ctgl.style.display = (name==="map") ? "" : "none";  // nÃºt â¡ chá» á» tab Báº£n Äá»
  if(name!=="map") $("#detail").classList.add("hidden");
  if(name==="scorecard") renderScorecard();
  if(name==="optimizer") renderOptimizer();
  if(name==="rezone") renderRezone();
  if(name==="chatlog") renderChatlog();
}
function renderChatlog(){
  const c=$("#chatlog");
  if(!c.dataset.init){
    c.innerHTML=`<div class="sc-head"><h2>Lá»ch sá»­ chat</h2><div class="sc-note">Thá»ng kÃª tá»ng há»£p + xem cÃ¢u há»i/tráº£ lá»i theo ngÃ y (chá» admin).</div></div>
      <div id="cl-stats"></div>
      <div class="sc-sub">
        <h3>ð Nháº­t kÃ½ theo ngÃ y <button id="cs-ai" style="margin-left:8px">ð¤ AI tá»ng há»£p cÃ¢u há»i</button></h3>
        <div style="margin-bottom:10px">NgÃ y: <input type="date" id="cl-date" value="">
          <button id="cl-load">Táº£i</button>
          <span style="color:#9ca3af;font-size:11px">(Äá» trá»ng = 15 ngÃ y gáº§n nháº¥t)</span></div>
        <div id="cs-ai-box"></div>
        <div id="cl-table"></div></div>`;
    c.dataset.init="1"; $("#cl-load").onclick=loadChatlog;
    const aib=$("#cs-ai"); if(aib) aib.onclick=summarizeQuestions;
  }
  loadChatstats(); loadChatlog();
}
async function loadChatstats(){
  const box=$("#cl-stats"); if(!box) return; box.innerHTML="Äang táº£i thá»ng kÃªâ¦";
  try{
    const j=await fetch('/chatstats').then(r=>r.json());
    if(j.error){ box.innerHTML="â  "+j.error; return; }
    if(!j.total){ box.innerHTML='<div class="sc-note">ChÆ°a cÃ³ dá»¯ liá»u chat Äá» thá»ng kÃª.</div>'; return; }
    const card=(t,v)=>`<div class="stat-card"><div class="sv">${fmt(v)}</div><div class="st">${t}</div></div>`;
    const period=j.period?`${j.period[0]} â ${j.period[1]}`:"";
    const mx=Math.max(1,...j.by_day.map(d=>d[1]));
    const bars=j.by_day.slice(-30).map(([d,c])=>`<div class="bar-row"><span class="bd">${d.slice(5)}</span><span class="bb"><i style="width:${Math.round(c/mx*100)}%"></i></span><span class="bn">${c}</span></div>`).join("");
    const utbl=`<table class="sc"><thead><tr><th class="l">User</th><th>LÆ°á»£t</th></tr></thead><tbody>`+
      j.by_user.map(([u,c])=>`<tr><td class="l">${esc(u)}</td><td>${c}</td></tr>`).join("")+`</tbody></table>`;
    const rtbl=`<table class="sc"><thead><tr><th class="l">VÃ¹ng</th><th>LÆ°á»£t</th></tr></thead><tbody>`+
      j.by_region.map(([r,c])=>`<tr><td class="l">${esc(r)}</td><td>${c}</td></tr>`).join("")+`</tbody></table>`;
    box.innerHTML=`<div class="sc-note">Tá»ng há»£p toÃ n bá» log Äang lÆ°u${period?` (${period})`:""}</div>
      <div class="stat-cards">${card("Tá»ng lÆ°á»£t",j.total)}${card("User hoáº¡t Äá»ng",j.users)}${card("NgÃ y cÃ³ log",j.days)}${card("'KhÃ´ng cÃ³ dá»¯ liá»u'",j.nodata)}${card("Lá»i",j.errors)}</div>
      <div class="stat-grid">
        <div class="stat-col"><h4>LÆ°á»£t theo ngÃ y (30 gáº§n nháº¥t)</h4>${bars||'<div class="sc-note">â</div>'}</div>
        <div class="stat-col"><h4>Theo user</h4>${utbl}</div>
        <div class="stat-col"><h4>Theo vÃ¹ng</h4>${rtbl}</div>
      </div>`;
  }catch(e){ box.innerHTML="â  Lá»i táº£i thá»ng kÃª."; }
}
async function summarizeQuestions(){
  const btn=$("#cs-ai"), out=$("#cs-ai-box"); if(!out) return;
  if(btn){ btn.disabled=true; btn.textContent="ð¤ Äang tá»ng há»£pâ¦"; }
  out.innerHTML=`<div class="sc-note">AI Äang gom nhÃ³m cÃ¢u há»i (15 ngÃ y gáº§n nháº¥t)â¦</div>`;
  try{
    const j=await fetch('/chatsummary?days=15').then(r=>r.json());
    if(j.error){ out.innerHTML="â  "+j.error; }
    else out.innerHTML=`<div class="ai-summary">${mdSafe(j.summary)}</div>`;
  }catch(e){ out.innerHTML="â  Lá»i gá»i AI."; }
  if(btn){ btn.disabled=false; btn.textContent="ð¤ AI tá»ng há»£p cÃ¢u há»i"; }
}
async function loadChatlog(){
  const d=$("#cl-date").value; $("#cl-table").innerHTML="Äang táº£iâ¦";
  const url = d ? `/chatlog?date=${d}&n=1500` : `/chatlog?days=15&n=1500`;  // trá»ng = 15 ngÃ y gáº§n nháº¥t
  try{
    const j=await fetch(url).then(r=>r.json());
    if(j.error){ $("#cl-table").innerHTML="â  "+j.error; return; }
    if(!j.log||!j.log.length){ $("#cl-table").innerHTML=`ChÆ°a cÃ³ chat (${j.date}).`; return; }
    const rows=j.log.slice().reverse().map(e=>`<tr><td class="l">${esc((e.ts||"").replace("T"," "))}</td>`+
      `<td class="l">${esc(e.user||"")}</td><td class="l">${esc(Array.isArray(e.regions)?e.regions.join(","):(e.regions||""))}</td>`+
      `<td class="l">${esc(e.question||"")}</td><td class="l">${esc((e.answer||e.error||"").slice(0,500))}</td></tr>`).join("");
    $("#cl-table").innerHTML=`<div class="sc-note">${j.count} lÆ°á»£t Â· ${j.date}</div>`+
      `<table class="sc"><thead><tr><th class="l">Thá»i gian</th><th class="l">User</th><th class="l">VÃ¹ng</th><th class="l">CÃ¢u há»i</th><th class="l">Tráº£ lá»i</th></tr></thead><tbody>${rows}</tbody></table>`;
  }catch(e){ $("#cl-table").innerHTML="â  Lá»i táº£i log."; }
}

// ---------- Re-zone tab ----------
function renderRezone(){
  const r = DATA.rez;
  if(!r){ $("#rz-stats").innerHTML="ChÆ°a cÃ³ rezone.json (cáº§n merge_ward + cháº¡y build)."; return; }
  const s=r.stats;
  const scoped = r.new_wards.filter(w=>regOK(w.region));   // lá»c theo vÃ¹ng Äang xem
  const filtered = ui.regionFilter || (Array.isArray(ui.allowed) && ui.allowed.length===1);
  if(filtered){
    const cl=scoped.filter(w=>w.status==="clean").length, sp=scoped.filter(w=>w.status==="split").length, em=scoped.filter(w=>w.status==="empty").length;
    $("#rz-stats").innerHTML=`VÃ¹ng Äang xem: <b>${fmt(scoped.length)}</b> phÆ°á»ng má»i Â· <b>${fmt(cl)}</b> sáº¡ch (1 BC) Â· `+
      `<b>${fmt(sp)}</b> bá» xÃ© â auto gÃ¡n theo Äa sá» cáº§u Â· <b>${em}</b> rá»ng cáº§u. NguyÃªn táº¯c: 1 phÆ°á»ng má»i = 1 BC (khÃ´ng xÃ© láº»).`;
  } else {
    $("#rz-stats").innerHTML=`<b>${fmt(s.new_wards)}</b> phÆ°á»ng má»i Â· <b>${fmt(s.clean)}</b> sáº¡ch (1 BC) Â· `+
      `<b>${fmt(s.split)}</b> bá» xÃ© â auto gÃ¡n theo Äa sá» cáº§u (${s.pct_split}%) Â· <b>${s.empty}</b> rá»ng cáº§u Â· `+
      `cáº§u trong nhÃ³m xÃ© <b>${s.split_dem_pct}%</b>. NguyÃªn táº¯c: 1 phÆ°á»ng má»i = 1 BC (khÃ´ng xÃ© láº»).`;
  }
  const split = scoped.filter(w=>w.status==="split")
    .sort((a,b)=> (a.province||"").localeCompare(b.province||"","vi") || (a.name||"").localeCompare(b.name||"","vi")); // theo tá»nh má»i, rá»i tÃªn phÆ°á»ng
  $("#rz-split").innerHTML=subTable(split.map(w=>({...w,code:w.new_code})),[
    {h:"Tá»nh má»i",f:w=>w.province,l:1},
    {h:"PhÆ°á»ng má»i",f:w=>`${w.name} (${w.new_code})`,l:1},
    {h:"#PhÆ°á»ng cÅ©",f:w=>w.n_old},
    {h:"Cáº§u/ngÃ y",f:w=>fmt(w.dem)},
    {h:"BC auto gÃ¡n",f:w=>`${w.assigned_bc_name||w.assigned_bc}`,l:1},
    {h:"% cáº§u",f:w=>`${w.lead_share}%`},
    {h:"BC khÃ¡c",f:w=>w.candidates.slice(1,4).map(c=>`${(c.bc_name||c.bc).slice(0,22)} ${c.share}%`).join(" Â· "),l:1}],5000,true,true);
  // click 1 dÃ²ng -> panel breakdown phÆ°á»ng cÅ© (CÃ¡ch B)
  document.querySelectorAll("#rz-split .region-row").forEach(tr=>tr.onclick=()=>{
    const w=DATA.rez.new_wards.find(x=>x.new_code===tr.dataset.code); if(w) showRezoneDetail(w);
  });
}
function showRezoneDetail(w){
  let rows=(w.olds||[]).map(o=>`<tr><td class="l">${esc(o.name)} (${esc(o.ward)})</td><td class="l">${esc(o.bc_name||o.bc||"â")}</td>`+
    `<td>${fmt(o.dem)}</td><td>${fmt(o.dem_kg)}</td><td>${esc(o.note==="pháº§n"?"má»t pháº§n":o.note)}</td></tr>`).join("");
  let cand=w.candidates.map(c=>`<tr${c.bc===w.assigned_bc?' style="font-weight:700;background:#dcfce7"':''}><td class="l">${esc(c.bc_name||c.bc)}</td><td>${fmt(c.dem)}</td><td>${c.share}%</td></tr>`).join("");
  const tag = w.status==="split" ? '<b style="color:#dc2626">bá» xÃ©</b>' : (w.status==="empty"?"rá»ng cáº§u":'<b style="color:#16a34a">sáº¡ch (1 BC)</b>');
  showDetail(`<h3>${esc(w.name)}</h3><div style="color:#6b7280">${esc(w.province)} Â· ${w.n_old} phÆ°á»ng cÅ© Â· cáº§u ${fmt(w.dem)} ÄÆ¡n Â· ${fmt(w.dem_kg)} kg /ngÃ y Â· ${tag}</div>
    <div style="margin:8px 0"><b>BC ${w.status==="split"?"auto gÃ¡n":"phá»¥ trÃ¡ch"}:</b> ${esc(w.assigned_bc_name||w.assigned_bc||"â")}${w.lead_share?` (${w.lead_share}% cáº§u)`:""}</div>
    ${w.candidates.length>1?`<b style="font-size:12px">á»¨ng viÃªn (theo cáº§u):</b>
    <table class="bal"><tr><th>BC</th><th>Cáº§u</th><th>%</th></tr>${cand}</table>`:""}
    ${(w.olds&&w.olds.length)?`<b style="font-size:12px">ThÃ nh pháº§n phÆ°á»ng cÅ©:</b>
    <table class="bal"><tr><th>PhÆ°á»ng cÅ©</th><th>BC giao</th><th>ÄÆ¡n</th><th>Kg</th><th>Nháº­p</th></tr>${rows}</table>`:`<div class="note">Gá»m ${w.n_old} phÆ°á»ng cÅ© (cÃ¹ng 1 BC).</div>`}
    ${w.status==="split"?'<div class="note">Auto = Äa sá» cáº§u. Override tay sáº½ thÃªm sau.</div>':""}`);
}

// ---------- Network Optimizer tab ----------
function flyToWardOrHub(code, isHub){
  const h = isHub ? hubByCode[code] : null;
  if(h && h.has_geo){ switchTab("map"); map.flyTo({center:[h.lng,h.lat],zoom:12}); clickHub(h); return; }
  // ward: dÃ¹ng centroid hub giao Äá» bay táº¡m (khÃ´ng cÃ³ toáº¡ Äá» ward á» client) -> bay tá»i BC hiá»n táº¡i
}
function renderOptimizer(){
  const o = DATA.opt;
  if(!o){ $("#opt-stats").innerHTML="ChÆ°a cÃ³ optimizer.json (cáº§n polygon + cháº¡y build)."; return; }
  const s=o.stats, t=o.thresholds;
  $("#opt-stats").innerHTML=`Khoáº£ng cÃ¡ch TB phÆ°á»ngâBC giao <b>${s.avg_d_cur}km</b> â náº¿u gÃ¡n gáº§n nháº¥t <b>${s.avg_d_pure}km</b> (giáº£m ${s.pct_saved}%) Â· `+
    `<b>${fmt(s.n_reassign)}</b> phÆ°á»ng nÃªn Äá»i BC (cáº§u ${fmt(s.reassign_dem)}/ngÃ y) Â· <b>${s.n_close}</b> ÄÃ³ng Â· <b>${s.n_split}</b> tÃ¡ch Â· <b>${s.n_far}</b> phÆ°á»ng xa >30km. `+
    `NgÆ°á»¡ng: Äiá»m/mÂ² P75=${t.eHi} P25=${t.eLo} Â· Äá»ng/Äiá»m P75=${t.dpeP75}.`;
  const CAP=5000;   // render háº¿t (há»p cuá»n lo pháº§n dÃ i), khÃ´ng cáº¯t dá»¯ liá»u
  const RG=arr=>arr.filter(x=>regOK(x.region));
  const base=[{h:"VÃ¹ng",f:r=>r.region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},
    {h:"PhÆ°á»ng",f:r=>`${r.name} (${r.ward})`,l:1}];
  // PhÆ°á»ng nÃªn Äá»i: Äá»i chiáº¿u pure vs cap
  $("#opt-reassign").innerHTML=subTable(RG(o.reassign).map(r=>({...r,code:r.from})),[...base,
    {h:"BC hiá»n táº¡i",f:r=>r.from_name||r.from,l:1},
    {h:"d hiá»n (km)",f:r=>r.d_cur},
    {h:"â Gáº§n nháº¥t",f:r=>r.to_pure?`${r.to_pure_name||r.to_pure} Â· ${r.d_pure}km`:"â",l:1},
    {h:"â Gáº§n nháº¥t cÃ²n táº£i",f:r=>r.to_cap?`${r.to_cap_name||r.to_cap} Â· ${r.d_cap}km`:"â",l:1},
    {h:"Cáº§u/ngÃ y",f:r=>fmt(r.dem)}],CAP,true);
  // KÃ©m hiá»u quáº£ (Äáº¯t + dÆ°) â hÃ nh Äá»ng theo háº¡n HÄ
  $("#opt-close").innerHTML=subTable(RG(o.close).map(c=>({...c,code:c.hub})),[
    {h:"VÃ¹ng",f:r=>r.region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},
    {h:"BC",f:r=>`${r.name} (${r.hub})`,l:1},
    {h:"Äá»ng/Äiá»m",f:r=>fmt(r.dpe)},{h:"Äiá»m/mÂ²",f:r=>r.em2},
    {h:"Háº¡n HÄ",raw:true,f:r=>r.days==null?"thiáº¿u":(r.days<0?`<span class="warn">ÄÃ£ háº¿t ${-r.days}d</span>`:`${r.days}d`)},
    {h:"Gá»p vá» (gáº§n nháº¥t)",f:r=>`${r.merge_to_name||r.merge_to} Â· ${r.merge_dist}km`,l:1},
    {h:"HÃ nh Äá»ng",f:r=>r.action,l:1}],CAP,true);
  // QuÃ¡ táº£i â tÃ¡ch / má» rá»ng
  $("#opt-split").innerHTML=subTable(RG(o.split).map(c=>({...c,code:c.hub})),[
    {h:"VÃ¹ng",f:r=>r.region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},
    {h:"BC",f:r=>`${r.name} (${r.hub})`,l:1},
    {h:"PhÆ°á»ng",f:r=>r.n_wards},{h:"Äiá»m/mÂ²",f:r=>r.em2??"â"},{h:"Volume/ngÃ y",f:r=>fmt(r.vol)},
    {h:"HÃ nh Äá»ng",f:r=>r.action,l:1}],CAP,true);
  // Whitespace Ã J&T â 4 nhÃ³m
  const w4=o.ws4;
  if(!w4 || !w4.has_jt){
    $("#opt-white").innerHTML=`<div class="sc-note">Cáº§n dá»¯ liá»u J&T (competitors_jt.json) Äá» phÃ¢n 4 nhÃ³m.</div>`;
  } else {
    const wcol=[{h:"VÃ¹ng",f:r=>r.region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},{h:"PhÆ°á»ng",f:r=>r.name,l:1},
      {h:"Cáº§u/ngÃ y",f:r=>fmt(r.dem)},{h:"GHN gáº§n nháº¥t",f:r=>r.d_ghn+" km"},{h:"J&T gáº§n nháº¥t",f:r=>r.d_jt+" km"}];
    const tbl=arr=>subTable(RG(arr).map(r=>({...r,code:r.ward})),wcol,CAP,true);
    $("#opt-white").innerHTML=
      `<div class="sc-note">Cáº§u cao = P75 (top 25% phÆ°á»ng â hÃ´m nay â¥${fmt(w4.demP75)} ÄÆ¡n/ngÃ y, tá»± Äá»i theo data) Â· GHN váº¯ng = BC gáº§n nháº¥t â¥${w4.ghn_far}km Â· J&T cÃ³ = J&T â¤${w4.jt_near}km.</div>`+
      `<div class="sc-sub"><h3>ð´ Khoáº£ng cÃ¡ch xa <small>(cáº§u cao Â· GHN váº¯ng Â· J&T cÃ³ â Æ°u tiÃªn)</small></h3>${tbl(w4.mat_khach)}</div>`+
      `<div class="sc-sub"><h3>ð¢ Greenfield <small>(cáº§u cao Â· cáº£ 2 váº¯ng â má» chiáº¿m trÆ°á»c)</small></h3>${tbl(w4.greenfield)}</div>`;
  }
  renderOptStaff();
  renderOptScore();
  // click dÃ²ng -> bay tá»i BC liÃªn quan
  document.querySelectorAll("#optimizer .region-row").forEach(tr=>tr.onclick=()=>flyToWardOrHub(tr.dataset.code,true));
}
// ð Báº£ng Äiá»m hiá»u suáº¥t â benchmark BC theo ÄIá»M tuyá»t Äá»i (ÄÆ¡n-tÄ/NV/ngÃ y). KhÃ´ng Äá»¥ng chá» sá» cÅ©.
function renderOptScore(){
  const box=$("#opt-score"); if(!box) return;
  const sb=scoreBench(), hs=_scoreHubs();
  if(!hs.length || !sb.p50){ box.innerHTML=`<div class="sc-note">ChÆ°a Äá»§ dá»¯ liá»u (cáº§n nhÃ¢n viÃªn + sáº£n lÆ°á»£ng).</div>`; return; }
  const rows=hs.map(h=>({h,sc:bcScore(h)})).filter(x=>x.sc!=null).sort((a,b)=>b.sc-a.sc);
  const tag=sc=> sc>=sb.p75?'ð´ quÃ¡ táº£i' : (sc<=sb.p25?'ðµ dÆ°' : 'â');
  box.innerHTML=
    `<div class="sc-note">${sb.n} BC Â· trung vá» <b>${fmt(sb.p50)}</b> Â· P25 ${fmt(sb.p25)} Â· P75 ${fmt(sb.p75)} (Äiá»m/NV). Äiá»m = (ÄÆ¡n QÄ + kg QÄÃ·K)/NV, K=kg TB má»i ÄÆ¡n toÃ n quá»c; láº¥y Ã0.4.</div>`+
    subTable(rows.map(x=>({...x.h,code:x.h.hub_code,_sc:x.sc})),[
      {h:"VÃ¹ng",f:r=>r._region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},{h:"BC",f:r=>`${r.name||r.hub_code} (${r.hub_code})`,l:1},
      {h:"NV",f:r=>fmt(r.staff)},{h:"ÄÆ¡n/ngÃ y",f:r=>fmt(r.actual.pv+r.actual.dv)},{h:"Kg/ngÃ y",f:r=>fmt(r.actual.pw+r.actual.dw)},
      {h:"Äiá»m",f:r=>fmt(r._sc)},{h:"Xáº¿p loáº¡i",f:r=>tag(r._sc),l:1}],5000,true);
}
// ð· Äá»nh biÃªn nhÃ¢n sá»± â tÃ­nh client-side tá»« hubs (ÄÃ£ cÃ³ staff). NÄng suáº¥t = (pv+dv)/staff.
function renderOptStaff(){
  const box=$("#opt-staff"); if(!box) return;
  const hs=DATA.hubs.filter(h=>h.type==="express"&&h.assigned&&regOK(h._region)&&prodOf(h)!=null);
  const st=prodStats();
  if(!hs.length || !st.p50){ box.innerHTML=`<div class="sc-note">ChÆ°a cÃ³ dá»¯ liá»u nhÃ¢n viÃªn cho pháº¡m vi nÃ y.</div>`; return; }
  const med=st.p50;
  const volOf=h=>h.actual.pv+h.actual.dv;
  const need=h=>Math.max(0, Math.ceil(volOf(h)/med)-h.staff);     // NV cáº§n thÃªm Äá» kÃ©o vá» nÄng suáº¥t trung vá»
  const surplus=h=>Math.max(0, h.staff-Math.ceil(volOf(h)/med));  // NV dÆ° so vá»i trung vá»
  const dist=(a,b)=>{ if(!a.has_geo||!b.has_geo) return Infinity;
    const dx=(a.lng-b.lng)*Math.cos(a.lat*Math.PI/180), dy=a.lat-b.lat; return Math.sqrt(dx*dx+dy*dy)*111; };
  const over=hs.filter(h=>prodOf(h)<=st.p25 && surplus(h)>0);      // thá»«a ngÆ°á»i
  const under=hs.filter(h=>prodOf(h)>=st.p75).sort((a,b)=>prodOf(b)-prodOf(a));   // thiáº¿u ngÆ°á»i
  const overSorted=over.slice().sort((a,b)=>surplus(b)-surplus(a));
  const nearSurplus=h=>{ let best=null,bd=Infinity; over.forEach(o=>{ if(o.hub_code===h.hub_code) return;
    const d=dist(h,o); if(d<bd){bd=d;best=o;} }); return best?`${best.name||best.hub_code} Â· ${bd.toFixed(0)}km (dÆ° ${surplus(best)})`:"â"; };
  const totNV=hs.reduce((s,h)=>s+h.staff,0), totVol=hs.reduce((s,h)=>s+volOf(h),0);
  box.innerHTML=
    `<div class="sc-note">Pháº¡m vi: <b>${fmt(totNV)}</b> NV / <b>${fmt(totVol)}</b> ÄÆ¡n/ngÃ y Â· TB <b>${fmt(med)}</b> ÄÆ¡n/NV (P25 ${fmt(st.p25)} Â· P75 ${fmt(st.p75)}).</div>`+
    `<div class="sc-sub"><h3>ð´ Thiáº¿u ngÆ°á»i <small>(ÄÆ¡n/NV â¥ P75 â quÃ¡ táº£i, cáº§n tuyá»n/san táº£i)</small></h3>`+
    subTable(under.map(h=>({...h,code:h.hub_code})),[
      {h:"VÃ¹ng",f:r=>r._region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},{h:"BC",f:r=>`${r.name||r.hub_code} (${r.hub_code})`,l:1},
      {h:"NV",f:r=>fmt(r.staff)},{h:"ÄÆ¡n/ngÃ y",f:r=>fmt(volOf(r))},{h:"ÄÆ¡n/NV",f:r=>fmt(prodOf(r))},
      {h:"Cáº§n thÃªm",f:r=>`+${need(r)} NV`,l:1},{h:"Äiá»u tá»« (BC dÆ° gáº§n nháº¥t)",f:r=>nearSurplus(r),l:1}],5000,true)+`</div>`+
    `<div class="sc-sub"><h3>ðµ Thá»«a ngÆ°á»i <small>(ÄÆ¡n/NV â¤ P25 â dÆ°, Äiá»u chuyá»n/cáº¯t giáº£m)</small></h3>`+
    subTable(overSorted.map(h=>({...h,code:h.hub_code})),[
      {h:"VÃ¹ng",f:r=>r._region,l:1},{h:"Tá»nh",f:r=>r.province,l:1},{h:"BC",f:r=>`${r.name||r.hub_code} (${r.hub_code})`,l:1},
      {h:"NV",f:r=>fmt(r.staff)},{h:"ÄÆ¡n/ngÃ y",f:r=>fmt(volOf(r))},{h:"ÄÆ¡n/NV",f:r=>fmt(prodOf(r))},
      {h:"DÆ°",f:r=>`${surplus(r)} NV`,l:1}],5000,true)+`</div>`;
}

let helpLoaded=false;
function wireHelp(){
  const btn=$("#help-btn"), modal=$("#help-modal"), close=$("#help-close");
  if(!btn||!modal) return;
  const open=async()=>{
    modalOpen(modal);
    if(!helpLoaded){
      try{ const md=await fetch("/web/HUONG_DAN.md").then(r=>r.text());
        $("#help-content").innerHTML=mdSafe(md); helpLoaded=true; }
      catch(e){ $("#help-content").innerHTML="â  KhÃ´ng táº£i ÄÆ°á»£c hÆ°á»ng dáº«n."; }
    }
  };
  btn.onclick=open;
  if(close) close.onclick=()=>modalClose(modal);
  modal.onclick=e=>{ if(e.target===modal) modalClose(modal); };  // báº¥m ná»n tá»i Äá» ÄÃ³ng

  // prop-modal: báº¥m ná»n tá»i Äá» ÄÃ³ng (Escape + focus-trap do modalOpen lo)
  const pm=$("#prop-modal");
  if(pm) pm.addEventListener("click", e=>{ if(e.target===pm) modalClose(pm); });
}
function wireControls(){
  document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>switchTab(t.dataset.tab));
  const ctgl=$("#ctrl-toggle"); if(ctgl) ctgl.onclick=()=>$("#controls").classList.toggle("collapsed");
  wireHelp();
  $("#colormode").onchange=e=>{ ui.colormode=e.target.value; renderLegend(); draw(); };
  const lyMap={"ly-terr":"terr","ly-orphan":"orphan","ly-b2b":"b2b","ly-transit":"transit","ly-wards":"wards","ly-newwards":"newwards"};
  Object.entries(lyMap).forEach(([id,k])=>{ const el=$("#"+id); if(el) el.onchange=e=>{ ui.layers[k]=e.target.checked; draw(); }; });
  const pu=$("#planunit"); if(pu) pu.onchange=e=>{ ui.planUnit=e.target.value; draw(); };
  const pin=$("#pt-input");
  if(pin){
    pin.onfocus=()=>renderPtList(pin.value);
    pin.oninput=()=>renderPtList(pin.value);
    document.addEventListener("click",e=>{ const box=$("#pt-box"); if(box && !box.contains(e.target)) hidePtList(); });
  }
  const pc=$("#planclear"); if(pc) pc.onclick=()=>{ plan.selected.clear(); plan.selNew.clear(); renderPlan(); draw(); };
  fillPlanTargets();
  const jt=$("#ly-jt"); if(jt) jt.onchange=e=>{ ui.jt=e.target.checked; renderLegend(); draw(); };
  // ----- Lá»p Optimizer -----
  const ovOn=$("#ov-on");
  if(ovOn) ovOn.onchange=e=>{ ui.optov.on=e.target.checked;
    const sub=$("#ov-sub"); if(sub){ sub.style.opacity=e.target.checked?"1":".45"; sub.style.pointerEvents=e.target.checked?"auto":"none"; }
    renderLegend(); draw(); };
  const ovMap={"ov-markers":"markers","ov-reassign":"reassign","ov-merge":"merge","ov-ws4":"ws4"};
  Object.entries(ovMap).forEach(([id,k])=>{ const el=$("#"+id); if(el) el.onchange=e=>{ ui.optov[k]=e.target.checked; renderLegend(); draw(); }; });
  const ovt=$("#ov-target"); if(ovt) ovt.onchange=e=>{ ui.optov.target=e.target.value; draw(); };
  const rf=$("#regionfilter");
  if(rf){
    let opts;
    if(ui.allowed==="*") opts=[["","ToÃ n bá»"],...REGION_LIST.map(r=>[r,r])];
    else if(Array.isArray(ui.allowed)&&ui.allowed.length===1) opts=[[ui.allowed[0],ui.allowed[0]]];
    else opts=[["","ToÃ n bá» (vÃ¹ng cá»§a tÃ´i)"],...(ui.allowed||[]).map(r=>[r,r])];
    rf.innerHTML=opts.map(([v,l])=>`<option value="${v}">${l}</option>`).join("");
    rf.value=ui.regionFilter||"";
    rf.disabled = Array.isArray(ui.allowed)&&ui.allowed.length===1;   // 1 vÃ¹ng: khoÃ¡ cá»©ng
    rf.onchange=e=>setRegionFilter(e.target.value);
  }
  $("#planmode").onclick=()=>{ ui.planMode=!ui.planMode; $("#planmode").classList.toggle("on",ui.planMode);
    $("#planmode").textContent=ui.planMode?"â  Táº¯t cháº¿ Äá» Quy hoáº¡ch":"â¿ Báº­t cháº¿ Äá» Quy hoáº¡ch (tÃ´ tay)";
    $("#plan-note").textContent = ui.planMode
      ? `Click phÆ°á»ng ${ui.planUnit==="new"?"má»i":"cÅ©"} Äá» thÃªm/bá»t vÃ o nhÃ³m; xem báº£ng TrÆ°á»c/Sau bÃªn pháº£i.`
      : "";
    if(ui.planMode){ ui.selWard=null; ui.selHub=null; }   // bá» highlight "xem" Äá» khÃ´ng láº«n vá»i Äá» quy hoáº¡ch
    draw(); };
  $("#detail-close").onclick=()=>{ $("#detail").classList.add("hidden"); ui.selHub=null; ui.selWard=null; draw(); };
  $("#search").oninput=e=>{ const q=e.target.value.toLowerCase().trim(); const box=$("#search-results");
    if(q.length<2){box.innerHTML="";return;}
    const hits=DATA.hubs.filter(h=>h.has_geo&&((h.name||"").toLowerCase().includes(q)||h.hub_code.includes(q))).slice(0,12);
    box.innerHTML=hits.map(h=>`<div class="sr" data-c="${esc(h.hub_code)}">${esc(h.name||h.hub_code)} <span style="color:#9ca3af">${esc(h.hub_code)}</span></div>`).join("");
    box.querySelectorAll(".sr").forEach(el=>el.onclick=()=>{ const h=hubByCode[el.dataset.c];
      map.flyTo({center:[h.lng,h.lat],zoom:12}); clickHub(h); }); };
}

// ---------- Chat AI (há»i vá» vÃ¹ng/BC/phÆ°á»ng/Äá» xuáº¥t) ----------
function esc(s){ return String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
// markdown -> HTML ÄÃ£ lá»c XSS (cáº¯t <script>, onerror...); fallback an toÃ n náº¿u thiáº¿u lib
function mdSafe(t){ t=t||""; const html=window.marked?marked.parse(t):esc(t);
  return window.DOMPurify?DOMPurify.sanitize(html):esc(t); }
function chatMsg(role, html){
  const d=document.createElement("div"); d.className="msg "+role; d.innerHTML=html;
  $("#chat-msgs").appendChild(d); $("#chat-msgs").scrollTop=1e9; return d;
}
async function chatAsk(q){
  chatMsg("u", esc(q));
  const w=chatMsg("a", "<i>Äang tráº£ lá»iâ¦</i>");
  try{
    const r=await fetch("/chat",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify({question:q})});
    const d=await r.json();
    w.innerHTML = d.answer ? mdSafe(d.answer) : ("â  "+(d.error||"lá»i"));
  }catch(e){ w.innerHTML="â  Lá»i káº¿t ná»i server."; }
  $("#chat-msgs").scrollTop=1e9;
}
function wireChat(){
  $("#chat-scope").textContent = ui.allowed==="*"?"(toÃ n quá»c)":"("+(Array.isArray(ui.allowed)?ui.allowed.join(", "):"")+")";
  $("#chat-fab").onclick=()=>{ $("#chat").classList.remove("hidden"); $("#chat-fab").classList.add("hidden"); $("#chat-q").focus();
    if(!$("#chat-msgs").children.length) chatMsg("a","ChÃ o báº¡n ð Há»i tÃ´i vá» <b>vÃ¹ng / bÆ°u cá»¥c / phÆ°á»ng / Äá» xuáº¥t</b> trong khu vá»±c báº¡n phá»¥ trÃ¡ch nhÃ©.<br><small style='color:#6b7280'>VD: \"BC nÃ o sáº¯p háº¿t háº¡n há»£p Äá»ng?\", \"phÆ°á»ng nÃ o nÃªn Äá»i bÆ°u cá»¥c?\"</small>"); };
  $("#chat-close").onclick=()=>{ $("#chat").classList.add("hidden"); $("#chat-fab").classList.remove("hidden"); };
  const send=()=>{ const q=$("#chat-q").value.trim(); if(!q) return; $("#chat-q").value=""; chatAsk(q); };
  $("#chat-send").onclick=send;
  $("#chat-q").onkeydown=e=>{ if(e.key==="Enter"&&!e.shiftKey){ e.preventDefault(); send(); } };
}

boot();
