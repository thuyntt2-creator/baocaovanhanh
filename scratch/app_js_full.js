/* Quy hoạch Bưu Cục Express — front-end deck.gl (tĩnh, không backend).
   Đọc data/out/{hubs,wards,meta}.json (+ wards.geojson khi có polygon).
   Hub hiển thị ngay; layer phường tự kích hoạt khi có wards.geojson. */

const DATA_DIR = "/api";   // data đã lọc theo vùng + cần session (server-side auth)
const REGION_COLORS = {
  BTB:[230,25,75], XBG:[60,180,75], TTB:[255,165,0], TNT:[0,130,200],
  DBB:[145,30,180], TBB:[70,240,240], TNB:[240,50,230], TNG:[210,245,60],
  "ĐCL":[250,190,212], DNB:[0,128,128], HNO:[170,110,40], NTB:[255,215,0],
  HCM:[128,0,0], DSH:[0,0,128],
};
const ROLE_COLORS = {
  territorial:[37,99,235], pickup_only:[234,88,12],
  bulky_delivery:[147,51,234], special_mixed:[13,148,136],
};
const TYPE_COLORS = { B2B:[153,27,27], transit:[107,114,128], sales:[209,213,219],
  warehouse:[120,90,60], other:[156,163,175], test:[200,200,200] };
const HIDDEN_TYPES = new Set(["sales","warehouse","other","test"]); // không vẽ mặc định

let DATA = { hubs:[], wards:[], meta:{}, geojson:null };

// ---------- Toast: thông báo không chặn (thay cho việc ghi lỗi vào #meta) ----------
function toast(msg, kind="info", opts={}){
  const wrap = document.getElementById("toast-wrap"); if(!wrap) return;
  const icon = kind==="ok" ? "ic-check" : (kind==="err"||kind==="warn") ? "ic-alert" : "ic-infoc";
  const el = document.createElement("div");
  el.className = "toast " + (kind||"");
  el.setAttribute("role", kind==="err" ? "alert" : "status");
  el.innerHTML = `<svg class="ic"><use href="#${icon}"/></svg><div class="tx"></div><button class="x" aria-label="Đóng">×</button>`;
  el.querySelector(".tx").textContent = String(msg);            // textContent -> không XSS
  const kill = ()=>{ el.classList.add("hide"); setTimeout(()=>el.remove(), 200); };
  el.querySelector(".x").onclick = kill;
  wrap.appendChild(el);
  const dur = (opts.duration!=null) ? opts.duration : (kind==="err" ? 6000 : 3500);
  if(dur>0) setTimeout(kill, dur);
  return el;
}
window.toast = toast;

// ---------- Modal a11y: focus-trap + Escape + khôi phục focus khi đóng ----------
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

// ---------- Giữ nội dung không bị topbar đè khi nó wrap ở màn hẹp ----------
// Đo chiều cao THỰC của topbar (1 hay nhiều hàng) -> biến --topbar-h; panel bên dưới neo theo biến này.
(function(){
  const tb = document.getElementById("topbar"); if(!tb) return;
  const setH = ()=> document.documentElement.style.setProperty("--topbar-h", tb.offsetHeight + "px");
  setH();
  if(window.ResizeObserver) new ResizeObserver(setH).observe(tb);
  window.addEventListener("resize", setH);
})();

// ---------- Popover cho ⓘ (.info): render fixed -> không bị panel cắt mép / nút nổi đè ----------
(function(){
  let pop;
  const getPop = ()=>{ if(!pop){ pop=document.createElement("div"); pop.className="info-pop"; document.body.appendChild(pop); } return pop; };
  function show(el){
    const src = el.querySelector(".tip"); const txt = src ? src.textContent : "";
    if(!txt) return;
    const p = getPop(); p.textContent = txt; p.style.display="block"; p.style.left="0"; p.style.top="0";
    const r = el.getBoundingClientRect(), m = 8, pw = p.offsetWidth, ph = p.offsetHeight;
    let left = r.left + r.width/2 - pw/2;
    left = Math.max(m, Math.min(left, window.innerWidth - pw - m));   // kẹp trong màn hình
    let top = r.bottom + m;
    if(top + ph > window.innerHeight - m) top = r.top - ph - m;       // hết chỗ dưới -> lật lên trên
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
const REGION_LIST = ["BTB","XBG","TTB","TNT","DBB","TBB","TNB","TNG","ĐCL","DNB","HNO","NTB","HCM","DSH"];
const plan = { selected:new Map(), selNew:new Map() }; // selected: ward_code cũ (tính cầu); selNew: new_ward_code (hiển thị/khung)

// ---------- helpers ----------
const $ = s => document.querySelector(s);
const fmt = n => (n==null?"–":Math.round(n).toLocaleString("vi-VN"));
const sum = (a,f)=>a.reduce((s,x)=>s+(f(x)||0),0);
function clamp(v,lo,hi){ return Math.max(lo,Math.min(hi,v)); }
function demandOf(h){ const t=h.territory_demand||{pv:0,dv:0}; return t.pv+t.dv; }

// ---------- màn tải + fetch CÓ TIMEOUT ----------
// Mọi request data phải có hạn chờ: nếu một request không bao giờ trả (extension chặn, mạng rớt
// giữa mẻ tải, origin restart giữa lúc stream), Promise.all treo mãi -> trang xoay vĩnh viễn.
// Có timeout thì hạn chờ hết -> reject -> hiện lỗi + nút Thử lại.
const T_SMALL = 20000;    // 3 file nền (hubs/wards/meta), vài trăm KB
const T_BIG   = 120000;   // wards.geojson 18MB + 5 file lớn khác
const T_ME    = 15000;    // /me

function abortAfter(ms){
  if(typeof AbortSignal!=="undefined" && AbortSignal.timeout) return AbortSignal.timeout(ms);
  const ac=new AbortController(); setTimeout(()=>ac.abort(), ms); return ac.signal;   // fallback
}
function errLabel(e){
  const n=(e && e.name) || "";
  if(n==="TimeoutError") return "quá thời gian chờ";
  if(n==="AbortError")   return "quá thời gian chờ (bị huỷ)";
  if(n==="TypeError")    return "request bị chặn hoặc mất mạng";
  return (e && e.message) ? String(e.message).slice(0,90) : "lỗi không rõ";
}
// signal phủ cả lúc đọc body -> stream đứt giữa đường cũng bị cắt theo hạn chờ, không treo.
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
    document.getElementById("bo-tx").textContent = "Không tải được ứng dụng";
    document.getElementById("bo-sub").textContent = "";
    document.getElementById("bo-err").classList.remove("hidden");
    document.getElementById("bo-err-tx").textContent = msg;
    const b=document.getElementById("bo-retry");
    b.onclick = ()=>{ this.busy("Đang thử lại…"); retry(); }; },
  done(){ this.el?.classList.add("hidden"); },
};

// ---------- khôi phục phiên (login ở trang riêng /web/login.html) ----------
async function boot(){
  window.__bootStarted = true;    // cờ cho watchdog trong index.html biết app.js đã chạy
  bootOv.busy("Đang kiểm tra phiên đăng nhập…");
  let d;
  try{
    const r = await fetch("/me", {signal:abortAfter(T_ME)});
    // CHỈ 401/403 mới là "chưa/hết phiên" -> về trang login. Lỗi mạng mà cũng redirect thì
    // người dùng bị đá sang login.html rồi login lại vẫn lỗi, không hiểu vì sao.
    if(r.status===401 || r.status===403){ location.href="/web/login.html"; return; }
    if(!r.ok) throw new Error("HTTP "+r.status);
    d = await r.json();
  }catch(e){
    bootOv.fail("Không kết nối được máy chủ khi kiểm tra phiên đăng nhập ("+errLabel(e)+").", boot);
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
  bootOv.busy("Đang tải dữ liệu nền…", "hubs · wards · meta");
  let hubs, wards, meta;
  try{
    [hubs, wards, meta] = await Promise.all(
      ["hubs.json","wards.json","meta.json"].map(f=>fetchJSON(`${DATA_DIR}/${f}`, T_SMALL)));
  }catch(e){
    // 3 file này BẮT BUỘC có mới dựng được app -> lỗi thì dừng hẳn và báo, không đi tiếp.
    bootOv.fail("Không tải được dữ liệu nền ("+errLabel(e)+").", load);
    return;
  }
  DATA.hubs=hubs; DATA.wards=wards; DATA.meta=meta;
  hubs.forEach(h=>hubByCode[h.hub_code]=h);
  wards.forEach(w=>wardByCode[w.ward_code]=w);
  buildHubColorIdx();   // gán màu lãnh thổ ổn định theo chỉ số BC
  // gán region + tỉnh cho hub = đa số phường nó phục vụ
  const top=o=>Object.entries(o).sort((a,b)=>b[1]-a[1])[0]?.[0]||"";
  hubs.forEach(h=>{ const rc={},pc={}; (h.covered_wards||[]).forEach(c=>{ const w=wardByCode[c]; if(!w)return;
      if(w.region) rc[w.region]=(rc[w.region]||0)+1; if(w.province) pc[w.province]=(pc[w.province]||0)+1; });
    h._region=top(rc); h._province=top(pc); });
  maxDemand = Math.max(1, ...hubs.filter(h=>h.assigned).map(demandOf));
  // 6 file lớn còn lại: tải SONG SONG (mỗi cái lỗi -> null, không làm hỏng cả mẻ), CÓ timeout,
  // và báo rõ tệp nào thiếu thay vì im lặng dựng bản đồ khuyết polygon.
  const BIG = ["wards.geojson","optimizer.json","rezone.json",
               "wards_new.geojson","competitors_jt.json","ward_centroids.json"];
  const failed=[]; let n=0;
  bootOv.busy("Đang tải dữ liệu bản đồ…", `0/${BIG.length} tệp`);
  const [geo, opt, rez, geoNew, jt, cent] = await Promise.all(BIG.map(f =>
    fetchJSON(`${DATA_DIR}/${f}`, T_BIG)
      .catch(e=>{ failed.push(`${f} (${errLabel(e)})`); return null; })
      .finally(()=>{ bootOv.sub(`${++n}/${BIG.length} tệp`); })));
  if(geo) DATA.geojson=geo;
  if(opt) DATA.opt=opt;
  if(rez){ DATA.rez=rez; DATA.rezByCode={}; rez.new_wards.forEach(w=>DATA.rezByCode[w.new_code]=w); }
  if(geoNew){ DATA.geojsonNew=geoNew; buildNewCent(); }
  if(jt) DATA.jt=jt;
  if(cent) DATA.cent=cent;
  try{
    initMap(); renderTopbar(); renderLegend(); renderDQ(); wireControls();
  }catch(e){
    // dựng UI lỗi -> vẫn phải bỏ màn tải + nói ra, chứ không để xoay mãi trên lỗi đã biết
    bootOv.fail("Lỗi khi dựng giao diện: "+errLabel(e), load); throw e;
  }
  bootOv.done();
  // Thiếu tệp lớn thì app vẫn chạy được nhưng KHUYẾT (mất polygon/optimizer…) -> phải nói rõ.
  if(failed.length) toast("Thiếu dữ liệu: "+failed.join(", ")+". Tải lại trang để thử lại.", "warn", {duration:12000});
  { const lbl = ui.allowed==="*"?"toàn quốc":(Array.isArray(ui.allowed)?ui.allowed.join(", "):"");
    $("#userbox").innerHTML=`<button id="userbtn" title="${esc(ui.username)}">👤</button>`+
      `<div id="usermenu" class="hidden"><div class="um-info"><b>${esc(ui.username||"?")}</b><br>${esc(lbl)}</div>`+
      `<a href="#" id="logout">Đăng xuất</a></div>`;
    $("#userbtn").onclick=()=>$("#usermenu").classList.toggle("hidden");
    $("#logout").onclick=e=>{ e.preventDefault(); fetch("/logout",{method:"POST"}).finally(()=>location.href="/web/login.html"); };
    document.addEventListener("click",e=>{ if(!$("#userbox").contains(e.target)) $("#usermenu")?.classList.add("hidden"); }); }
  if(ui.allowed==="*"){ const cb=$("#clog-btn"); cb.style.display=""; cb.onclick=()=>switchTab("chatlog"); }  // icon Lịch sử chat chỉ admin
  if(ui.allowed!=="*"){ $("#meta").style.display="none"; $("#diff").style.display="none"; }  // 2 ô báo cáo chỉ admin
  wireChat();
  if(location.hash==="#scorecard") switchTab("scorecard");   // tiện test/deep-link
  if(location.hash==="#optimizer") switchTab("optimizer");
  if(location.hash==="#rezone") switchTab("rezone");
}

// ---------- map ----------
function initMap() {
  map = new maplibregl.Map({
    container:"map",
    preserveDrawingBuffer:true,   // cho phép chụp canvas (map.getCanvas().toDataURL) cho đề xuất
    style:{ version:8,
      sources:{ base:{ type:"raster", tileSize:256,
        tiles:["https://a.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png",
               "https://b.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png",
               "https://c.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png"],
        attribution:"© OpenStreetMap, © CARTO" } },
      layers:[{ id:"base", type:"raster", source:"base", paint:{"raster-opacity":0.9} }] },
    center:[106.5,16.2], zoom:5,
  });
  overlay = new deck.MapboxOverlay({ interleaved:true, layers:[] });
  map.addControl(overlay);
  map.addControl(new maplibregl.NavigationControl(), "bottom-right");
  map.on("zoom", ()=>{ const z=map.getZoom(); const lod = z<7?"region":z<10?"province":"hub";
    const showLbl = z>=9;                               // hiện sớm; LOD theo độ-lớn-phường tự lọc
    const zq = Math.round(z*2)/2;                        // bước zoom 0.5 -> cập nhật nhãn mượt, ít vẽ lại
    if(lod!==ui.lod || showLbl!==ui.showLbl || (showLbl && zq!==ui.zq)){ ui.lod=lod; ui.showLbl=showLbl; ui.zq=zq; draw(); } });
  map.on("load", ()=>{ fitToAllowed(); draw(); });
}

// ----- màu tô lãnh thổ phường (deck GeoJsonLayer) -----
// chế độ màu hiệu lực: auto + đang xem 1 vùng -> tô theo BC (tô theo vùng lúc đó vô nghĩa, cả màn 1 màu)
function effColorMode(){
  if(ui.colormode!=="auto") return ui.colormode;
  const single = ui.regionFilter || (Array.isArray(ui.allowed) && ui.allowed.length===1);
  return single ? "hub" : ui.lod;
}
function wardFill(feature){
  const code = feature.properties.ward_code || feature.properties.polygon_id_code;
  const w = wardByCode[code]; if(!w) return [0,0,0,0];
  if(ui.planUnit==="old" && plan.selected.has(w.ward_code)) return [220,38,38,210];  // tô tay theo phường CŨ
  const mode = effColorMode();
  if(mode==="demand"){ const t=clamp((w.pv+w.dv)/400,0,1);
    return [Math.round(255*t),Math.round(120*(1-t)),60,170]; }
  if(mode==="hub"){ return hubTerrColor(w.delivery_hub||""); }
  const c=REGION_COLORS[w.region]||[200,200,200];           // region / auto
  return [c[0],c[1],c[2], mode==="province"?195:165];
}
function hashCode(s){let h=0;for(let i=0;i<s.length;i++)h=(h*31+s.charCodeAt(i))|0;return Math.abs(h);}
function hslColor(seed){const hue=seed%360;return hsl2rgb(hue,55,55).concat(170);}
// bỏ prefix loại đơn vị, chỉ giữ TÊN
function stripPrefix(n){ return (n||"").replace(/^(Phường|Xã|Thị trấn|Đặc khu|Thị xã|Quận|Huyện)\s+/i,"").trim(); }
// tâm polygon (ring lớn nhất) để đặt nhãn tên phường mới — tính 1 lần, cache
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
// màu lãnh thổ BC: gán theo CHỈ SỐ BC + góc vàng 137.5° -> hue trải đều, không trùng;
// thêm biến thiên độ đậm/sáng để BC cạnh nhau dù gần hue vẫn khác sắc -> phân biệt rõ lãnh thổ.
let hubColorIdx = {};
function buildHubColorIdx(){
  hubColorIdx = {};
  DATA.hubs.map(h=>h.hub_code).sort().forEach((c,i)=>{ hubColorIdx[c]=i; });
}
function hubTerrColor(code){
  const i = hubColorIdx[code];
  if(i==null) return [205,205,205,150];          // không rõ BC -> xám
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
    if(h.type==="transit")  return ui.layers.transit;   // gồm Trung Chuyển + Chuyển Tiếp
    if(HIDDEN_TYPES.has(h.type)) return false;           // kho khác/sales/test/thanh lý
    // express
    if(h.role==="territorial") return ui.layers.terr;
    return ui.layers.orphan; // chuyên dụng
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
  // đồng đều hết: mọi BC cùng cỡ, phân biệt tải bằng màu (không bằng kích thước)
  return h.role==="territorial" ? 5 : 4.5;
}

function draw(){
  const layers=[];
  // lãnh thổ phường (polygon) — dưới cùng
  if(ui.layers.wards && DATA.geojson){
    let feats=DATA.geojson.features;
    if(ui.regionFilter || ui.allowed!=="*") feats=feats.filter(f=>{const w=wardByCode[f.properties.ward_code];return w&&regOK(w.region);});
    layers.push(new deck.GeoJsonLayer({
      id:"wards", data:{type:"FeatureCollection",features:feats}, pickable:true, stroked:true, filled:true,
      getFillColor:wardFill,
      // viền: xám cho vùng KHÔNG có dữ liệu (đảo/đặc khu — không join wardByCode); trắng cho phường thường
      getLineColor:f=>wardByCode[f.properties.ward_code]?[255,255,255,95]:[120,120,120,170], lineWidthMinPixels:0.5,
      updateTriggers:{ getFillColor:[ui.colormode, ui.lod, plan.selected.size, ui.regionFilter, ui.allowed], data:[ui.regionFilter] },
      onClick: info=>info.object && clickWard(info.object),
    }));
  }
  // viền lãnh thổ phường MỚI (re-zone) — trên fill cũ; xé = đỏ. Bắt click khi đơn vị = "mới"
  if(ui.layers.newwards && DATA.geojsonNew){
    let nf=DATA.geojsonNew.features;
    if(ui.regionFilter || ui.allowed!=="*") nf=nf.filter(f=>{const w=DATA.rezByCode&&DATA.rezByCode[f.properties.new_ward_code];return w&&regOK(w.region);});
    layers.push(new deck.GeoJsonLayer({
      id:"newwards", data:{type:"FeatureCollection",features:nf}, pickable: ui.planUnit==="new",
      stroked:true, filled:true,
      getFillColor:f=>plan.selNew.has(f.properties.new_ward_code)?[225,29,42,245]:[0,0,0,0],  // tô tay phường MỚI: đỏ ĐẬM phủ rõ (không trộn màu nền)
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
    // nhãn TÊN phường mới — chỉ hiện khi phường đủ TO trên màn để chứa chữ (LOD theo zoom, không rối)
    if(ui.showLbl){
      const z=map.getZoom(), k=Math.pow(2,z);
      const lbls=nf.map(f=>{ const code=f.properties.new_ward_code; const c=newCent[code];
        if(!c || Math.sqrt(c.a)*k <= 110) return null;     // phường còn nhỏ trên màn -> chưa hiện nhãn
        const w=DATA.rezByCode&&DATA.rezByCode[code];
        const nm=stripPrefix(w?w.name:f.properties.name||"");
        return nm?{position:c.p,text:nm}:null; }).filter(Boolean);
      layers.push(new deck.TextLayer({
        id:"newward-labels", data:lbls, pickable:false,
        getPosition:d=>d.position, getText:d=>d.text,
        characterSet:"auto",   // đủ ký tự tiếng Việt
        getSize:14, sizeUnits:"pixels", sizeMinPixels:12, sizeMaxPixels:20,
        getColor:[17,24,39,255], getTextAnchor:"middle", getAlignmentBaseline:"center",
        fontFamily:'"Be Vietnam Pro", -apple-system, system-ui, "Segoe UI", Arial, sans-serif', fontWeight:700,
        // SDF + atlas độ phân giải cao + smoothing -> chữ sắc nét, khử răng cưa khi phóng to
        fontSettings:{sdf:true, fontSize:128, buffer:24, radius:30, cutoff:0.25, smoothing:0.12},
        outlineWidth:6, outlineColor:[255,255,255,255],
        maxWidth:130,
        updateTriggers:{ data:[ui.regionFilter,ui.allowed,ui.zq] },
      }));
    }
  }
  // lớp tham khảo BC đối thủ J&T (dưới hub GHN) — lọc theo vùng đang xem
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
        <div>${esc(info.object.addr||info.object.province||"–")}
        &nbsp;<a href="https://www.google.com/maps?q=${encodeURIComponent(info.object.lat)},${encodeURIComponent(info.object.lng)}" target="_blank">(Maps)</a></div>`),
    }));
  }
  // ===== Overlay Network Optimizer (chỉ để xem + chụp, không lưu) =====
  if(ui.optov.on && DATA.opt){
    const o=DATA.opt, ov=ui.optov;
    const hc=c=>{ const h=hubByCode[c]; return h&&h.has_geo?[h.lng,h.lat]:null; };
    const wc=c=>{ const p=DATA.cent&&DATA.cent[c]; return p?[p[0],p[1]]:null; };
    // mũi tên reassign phường → BC (cong); target: gần nhất | gần nhất-còn-tải
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
    // line gộp khi đóng BC (close → merge_to)
    if(ov.merge){
      const lines=(o.close||[]).filter(x=>regOK(x.region)&&x.merge_to).map(x=>{
        const src=hc(x.hub), dst=hc(x.merge_to); return src&&dst?{...x,src,dst}:null; }).filter(Boolean);
      layers.push(new deck.LineLayer({ id:"opt-merge", data:lines, pickable:true,
        getSourcePosition:d=>d.src, getTargetPosition:d=>d.dst,
        getColor:[220,38,38,200], getWidth:1.8, widthUnits:"pixels",
        onClick:info=>info.object&&showDetail(closeHtml(info.object)),
        updateTriggers:{ data:[ui.regionFilter,ui.allowed] } }));
    }
    // ws4: chấm phường cầu cao theo 4 nhóm whitespace × J&T
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
    // markers BC: đóng (đỏ) / tách-mở rộng (cam) — vòng to dưới chấm hub
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
  // highlight PHƯỜNG đang chọn (viền vàng dày) — ẩn khi đang quy hoạch để không trộn với đỏ nhóm chọn
  if(ui.selWard && !ui.planMode){
    const src = ui.selWardNew ? DATA.geojsonNew : DATA.geojson;
    const key = ui.selWardNew ? "new_ward_code" : "ward_code";
    const f = src && src.features.find(x=>x.properties[key]===ui.selWard);
    if(f) layers.push(new deck.GeoJsonLayer({ id:"sel-ward", data:{type:"FeatureCollection",features:[f]},
      stroked:true, filled:true, getFillColor:[255,193,7,70], getLineColor:[245,158,11,255], lineWidthMinPixels:4,
      updateTriggers:{ data:[ui.selWard,ui.selWardNew] } }));
  }
  // điểm hub — trên cùng
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
  // highlight BC đang chọn (vòng sáng to)
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
    if(layer&&layer.id==="hubs"){ return {html:`<b>${esc(object.name||object.hub_code)}</b><br>${esc(object.hub_code)} · ${roleLabel(object)}`}; }
    if(layer&&layer.id==="jt"){ return {html:`<b>${esc(object.name)}</b> <span style="color:#9ca3af">J&T</span><br>${esc(object.province)}`}; }
    if(layer&&layer.id==="wards"){ const w=wardByCode[object.properties.ward_code||object.properties.polygon_id_code]; if(!w)return null;
      return {html:`<b>${esc(w.name)}</b><br>${esc(w.district)}, ${esc(w.province)}<br>lấy ${fmt(w.pv)} · giao ${fmt(w.dv)} đơn/ngày`}; }
    if(layer&&layer.id==="opt-reassign"){ return {html:`<b>${esc(object.name)}</b> → đổi BC<br>${esc(object.from_name)} → <b>${esc(ui.optov.target==="cap"?object.to_cap_name:object.to_pure_name)}</b>`}; }
    if(layer&&layer.id==="opt-merge"){ return {html:`Đóng <b>${esc(object.name)}</b><br>gộp về ${esc(object.merge_to_name)} (${object.merge_dist} km)`}; }
    if(layer&&String(layer.id).startsWith("opt-ws4-")){ const g={mat_khach:"Khoảng cách xa",greenfield:"Greenfield",doi_dau:"Đối đầu J&T"}[object._g];
      return {html:`<b>${esc(object.name)}</b> · ${g}<br>cầu ${fmt(object.dem)}/ngày · GHN ${object.d_ghn}km · J&T ${object.d_jt}km`}; }
    return null;
  }});
}

// ---------- chi tiết overlay optimizer ----------
function reassignHtml(x){
  const cap=ui.optov.target==="cap";
  return `<h3>🔁 ${esc(x.name)}</h3><div style="color:#6b7280;margin-bottom:6px">${esc(x.province)} · ${esc(x.region)} · cầu ${fmt(x.dem)} đơn/ngày</div>
    <table class="bal"><tr><td>BC hiện tại</td><td>${esc(x.from_name)} · <b>${x.d_cur} km</b></td></tr>
    <tr><td>Gần nhất</td><td>${esc(x.to_pure_name)} · <b>${x.d_pure} km</b></td></tr>
    <tr><td>Gần nhất còn tải</td><td>${esc(x.to_cap_name)} · <b>${x.d_cap} km</b></td></tr></table>
    <div class="note">Đang vẽ theo: ${cap?"gần nhất-còn-tải":"gần nhất"}</div>`;
}
function closeHtml(x){
  return `<h3>🔴 Đóng/rà soát: ${esc(x.name)}</h3><div style="color:#6b7280;margin-bottom:6px">${esc(x.province)} · ${esc(x.region)}</div>
    <table class="bal"><tr><td>Đồng/đơn</td><td><b>${fmt(x.dpo)}</b></td></tr>
    <tr><td>Kg/m²</td><td>${x.kgm2}</td></tr><tr><td>Phường cover</td><td>${x.n_wards}</td></tr>
    <tr><td>HĐ còn</td><td>${x.days==null?"–":x.days+" ngày"}</td></tr>
    <tr><td>Hành động</td><td>${esc(x.action)}</td></tr>
    <tr><td>Gộp về</td><td>${esc(x.merge_to_name)} (${x.merge_dist} km)</td></tr></table>`;
}
function ws4Html(x){
  const g={mat_khach:"Khoảng cách xa (GHN xa, J&T gần)",greenfield:"Greenfield (cả 2 đều xa)",doi_dau:"Đối đầu J&T (cả 2 đều gần)"}[x._g];
  return `<h3>⭐ ${esc(x.name)}</h3><div style="color:#6b7280;margin-bottom:6px">${esc(x.province)} · ${esc(x.region)}</div>
    <table class="bal"><tr><td>Nhóm</td><td><b>${g}</b></td></tr>
    <tr><td>Cầu/ngày</td><td>${fmt(x.dem)}</td></tr>
    <tr><td>GHN gần nhất</td><td>${x.d_ghn} km</td></tr>
    <tr><td>J&T gần nhất</td><td>${x.d_jt} km</td></tr></table>`;
}

// ---------- click detail ----------
function roleLabel(h){
  if(h.type!=="express") return ({B2B:"Kho Hàng Nặng (B2B)",transit:"Kho Trung Chuyển/Chuyển Tiếp",
    sales:"Field Sales",warehouse:"Kho khác",other:"Khác",test:"Test"})[h.type]||h.type;
  return {territorial:"Bưu cục",pickup_only:"Chuyên LẤY",bulky_delivery:"Cồng kềnh/GIAO",special_mixed:"Chuyên dụng hỗn hợp"}[h.role];
}
function clickHub(h){
  ui.selHub=h.hub_code; ui.selWard=null; draw();
  const re=h.realestate; const ac=h.actual;
  const vol=ac?ac.pv+ac.dv:0, wt=ac?ac.pw+ac.dw:0;
  const dpo = re&&vol>0 ? re.rent/(vol*30) : null;     // tiền thuê tháng ÷ (đơn/ngày × 30)
  const dpw = re&&wt>0  ? re.rent/(wt*30)  : null;     // ÷ (kg/ngày × 30)
  const pr=prodOf(h), prk=prodKgOf(h), ps=prodStats();  // năng suất NV (đơn/NV/ngày) + ngưỡng vùng
  const prTag = pr==null ? "" : (ps.p75&&pr>=ps.p75 ? ' <span class="warn">🔴 thiếu người</span>'
                : (ps.p25&&pr<=ps.p25 ? ' <span style="color:#2563eb">🔵 thừa người</span>' : ''));
  const sb=scoreBench(), sc=bcScore(h);                 // ĐIỂM hiệu suất tuyệt đối (đơn-tđ/NV)
  const scTag = sc==null ? "" : (sb.p75&&sc>=sb.p75 ? ' <span class="warn">🔴 cao</span>'
                : (sb.p25&&sc<=sb.p25 ? ' <span style="color:#2563eb">🔵 thấp</span>' : ''));
  const maps = h.has_geo ? `&nbsp;<a href="https://www.google.com/maps?q=${encodeURIComponent(h.lat)},${encodeURIComponent(h.lng)}" target="_blank">(Maps)</a>` : "";
  let html=`<h3>${esc(h.name||h.hub_code)}</h3>
    <div><span class="tag" style="background:rgb(${(h.type==="express"?ROLE_COLORS[h.role]:TYPE_COLORS[h.type]).join(',')})">${roleLabel(h)}</span>
    <span style="color:#6b7280">${esc(h.hub_code)} · ${esc(regionOfHub(h)||"–")}${maps}</span></div>
    <div class="kv">
      <div>Phường phục vụ</div><div>${h.n_wards}</div>
      <div>Cầu địa bàn /ngày</div><div>${fmt(demandOf(h))} đơn · ${fmt((h.territory_demand.pw||0)+(h.territory_demand.dw||0))} kg</div>
      <div>· Lấy / Giao (đơn)</div><div>${fmt(h.territory_demand.pv)} / ${fmt(h.territory_demand.dv)}</div>
      <div>· Lấy / Giao (kg)</div><div>${fmt(h.territory_demand.pw||0)} / ${fmt(h.territory_demand.dw||0)}</div>
      <div class="grp-top">Sản lượng BC /ngày</div><div class="grp-top">${ac?fmt(vol)+" đơn · "+fmt(wt)+" kg":'<span class="warn">thiếu</span>'}</div>
      ${ac?`<div>· Lấy / Giao (đơn)</div><div>${fmt(ac.pv)} / ${fmt(ac.dv)}</div>
      <div>· Lấy / Giao (kg)</div><div>${fmt(ac.pw)} / ${fmt(ac.dw)}</div>`:""}
      <div class="grp-top">Nhân viên</div><div class="grp-top">${h.staff?fmt(h.staff):'<span class="warn">thiếu</span>'}</div>
      ${pr!=null?`<div>Đơn / NV / ngày</div><div>${fmt(pr)}${prTag}</div>
      <div>Kg / NV / ngày</div><div>${fmt(prk)}</div>`:""}
      ${sc!=null?`<div>Điểm hiệu suất</div><div>${fmt(sc)} điểm/NV${scTag}</div>`:""}
      <div class="grp-top">Mặt bằng</div><div class="grp-top">${re?fmt(re.usable_area)+" m²":'<span class="warn">thiếu</span>'}</div>
      <div>Tiền thuê</div><div>${re?fmt(re.rent)+" đ":"–"}</div>
      <div>Hạn HĐ</div><div>${re&&re.expiry?re.expiry:"–"}</div>
      <div>Đồng/đơn</div><div>${dpo?fmt(dpo)+" đ":"–"}</div>
      <div>Đồng/kg</div><div>${dpw?fmt(dpw)+" đ":"–"}</div>
    </div>`;
  if(h.missing_geo||h.missing_realestate) html+=`<div class="warn">⚠ ${[h.missing_geo&&"thiếu toạ độ",h.missing_realestate&&"thiếu mặt bằng"].filter(Boolean).join(", ")}</div>`;
  const names=h.covered_wards.slice(0,40).map(c=>esc(wardByCode[c]?wardByCode[c].name:c));
  html+=`<div style="margin-top:8px"><b>Phường (${h.n_wards}):</b> <span style="color:#6b7280;font-size:12px">${names.join(", ")}${h.n_wards>40?" …":""}</span></div>`;
  showDetail(html);
  highlightWards(new Set(h.covered_wards));
}
function clickWard(feature){
  const w=wardByCode[feature.properties.ward_code||feature.properties.polygon_id_code]; if(!w) return;
  if(ui.planMode){ togglePlan(w); return; }
  ui.selHub=null; ui.selWard=w.ward_code; ui.selWardNew=false; draw();
  const ph=hubByCode[w.pick_hub], dh=hubByCode[w.delivery_hub];
  showDetail(`<h3>${esc(w.name)}</h3>
    <div style="color:#6b7280">${esc(w.district)}, ${esc(w.province)} · ${esc(w.region)}</div>
    <div class="kv">
      <div>Lấy</div><div>${fmt(w.pv)} đơn · ${fmt(w.pw)} kg</div>
      <div>Giao</div><div>${fmt(w.dv)} đơn · ${fmt(w.dw)} kg</div>
      <div>BC lấy</div><div>${esc(ph?ph.name:w.pick_hub||"–")}</div>
      <div>BC giao</div><div>${esc(dh?dh.name:w.delivery_hub||"–")}</div>
    </div>`);
}
function clickNewWard(feature){
  const w=DATA.rezByCode && DATA.rezByCode[feature.properties.new_ward_code];
  if(ui.planMode){   // tô tay đơn vị "mới": chọn phường mới = tô polygon 3321 + gom cầu từ phường cũ bên trong
    const codes = w ? (w.old_codes||[]) : [];
    const nc = feature.properties.new_ward_code;
    const all = codes.length && codes.every(c=>plan.selected.has(c));
    codes.forEach(c=> all?plan.selected.delete(c):plan.selected.set(c,true));
    if(all) plan.selNew.delete(nc); else plan.selNew.set(nc,true);   // tô/đóng khung theo polygon phường mới
    renderPlan(); draw(); return;
  }
  ui.selHub=null; ui.selWard=feature.properties.new_ward_code; ui.selWardNew=true; draw();
  if(w) showRezoneDetail(w);
  else showDetail(`<h3>${esc(feature.properties.name||"")}</h3><div style="color:#6b7280">Phường mới ${esc(feature.properties.new_ward_code)}</div>`);
}
function showDetail(html){ $("#detail-body").innerHTML=html; $("#detail").classList.remove("hidden"); }
function highlightWards(set){ /* hook: khi có polygon sẽ tô; hiện chỉ no-op nếu chưa có geojson */ }

// ---------- what-if (tính trên CẦU Σ T1, bảo toàn theo phường) ----------
function togglePlan(w){
  if(plan.selected.has(w.ward_code)) plan.selected.delete(w.ward_code);
  else plan.selected.set(w.ward_code,true);
  renderPlan(); draw();
}
function demandKgOf(h){ const t=h.territory_demand||{}; return (t.pw||0)+(t.dw||0); }
// ---- năng suất nhân viên: đơn/NV/ngày (lấy+giao dùng chung) ----
function prodOf(h){ const a=h.actual, s=h.staff||0; return (a&&s)?(a.pv+a.dv)/s:null; }
function prodKgOf(h){ const a=h.actual, s=h.staff||0; return (a&&s)?(a.pw+a.dw)/s:null; }
function pctl(a,p){ if(!a.length) return null; const s=[...a].sort((x,y)=>x-y); return s[Math.min(s.length-1,Math.floor(p/100*s.length))]; }
function prodStats(){   // ngưỡng đơn/NV theo các BC express trong phạm vi đang lọc
  const v=DATA.hubs.filter(h=>h.type==="express"&&h.assigned&&regOK(h._region)&&prodOf(h)!=null).map(prodOf);
  return {p25:pctl(v,25),p50:pctl(v,50),p75:pctl(v,75),n:v.length};
}
// ===== ĐIỂM HIỆU SUẤT (benchmark mới, KHÔNG đụng các chỉ số hiện có) =====
// Quy đổi: 1 đơn lấy = 0.4 đơn giao; 1 kg lấy = 0.4 kg giao. ĐIỂM TUYỆT ĐỐI = năng suất
// "đơn-tương-đương/NV/ngày" = (đơn QĐ + kg QĐ ÷ K) / NV, với K = kg TB mỗi đơn (toàn quốc, cố định
// để quy kg về đơn). Điểm là số THẬT, trung vị mỗi vùng khác nhau. Cờ 🔴/🔵 theo P75/P25 vùng đang lọc.
function effOrd(h){ const a=h.actual; return a?0.4*a.pv+a.dv:0; }     // đơn quy đổi
function effKg(h){ const a=h.actual; return a?0.4*a.pw+a.dw:0; }      // kg quy đổi
function _scoreHubs(){ return DATA.hubs.filter(h=>h.type==="express"&&h.assigned&&regOK(h._region)&&h.staff&&h.actual); }
let _kgPerOrd=null;
function kgPerOrd(){   // K toàn quốc = Σ kg QĐ ÷ Σ đơn QĐ (cố định, không đổi theo vùng lọc)
  if(_kgPerOrd!=null) return _kgPerOrd;
  let so=0,sk=0; DATA.hubs.forEach(h=>{ if(h.type==="express"&&h.assigned&&h.actual){ so+=effOrd(h); sk+=effKg(h); } });
  _kgPerOrd = so>0 ? sk/so : 1; return _kgPerOrd;
}
function bcScore(h){ if(!h.staff||!h.actual) return null;
  const K=kgPerOrd()||1; return (effOrd(h)+effKg(h)/K)/h.staff; }   // đơn-tương-đương/NV/ngày (tuyệt đối)
function scoreBench(){ const v=_scoreHubs().map(bcScore).filter(x=>x!=null);
  return { p25:pctl(v,25), p50:pctl(v,50), p75:pctl(v,75), n:v.length }; }
const PROP_BTN=`<button id="save-prop" onclick="saveProposal()" style="width:100%;margin-top:8px;padding:9px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:600;cursor:pointer">💾 Lưu đề xuất</button>`;
// Ước lượng định biên cho đề xuất: NV cần cho cầu chuyển/nhận, theo năng suất trung vị (đơn/NV/ngày).
function planStaffEstimate(gainDon, totDon, T){
  const med=prodStats().p50; if(!med) return null;
  const th=T?hubByCode[T]:null;
  return { med:Math.round(med), addNV:Math.ceil((T?gainDon:totDon)/med), curStaff: th?(th.staff||0):null, isNew:!T };
}
function staffNoteHTML(e){
  if(!e) return "";
  return e.isNew
    ? `👷 Định biên: BC mới gánh cụm này → cần ~<b>${fmt(e.addNV)}</b> NV (năng suất TB ${fmt(e.med)} đơn/NV/ngày).`
    : `👷 Định biên: BC đích đang <b>${fmt(e.curStaff)}</b> NV, nhận thêm cầu → cần bổ sung ~<b>${fmt(e.addNV)}</b> NV (năng suất TB ${fmt(e.med)} đơn/NV/ngày).`;
}
// Quy hoạch theo PHƯỜNG MỚI: tính cầu từ olds[] (ĐÃ chia 1/n cho phường cũ "Nhập một phần"), gom theo BC giao
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
  const dcell=(don,kg)=>`${fmt(don)} đ<br>${fmt(kg)} kg`;
  let rows=Object.entries(loss).sort((a,b)=>b[1].don-a[1].don).map(([hc,o])=>{
    const h=hubByCode[hc], bd=h?demandOf(h):0, bk=h?demandKgOf(h):0;
    return `<tr><td class="l">${esc(h?h.name:hc)}</td><td>${dcell(bd,bk)}</td><td>${dcell(bd-o.don,bk-o.kg)}</td><td class="delta-neg">−${fmt(o.don)} đ<br>−${fmt(o.kg)} kg</td></tr>`;
  }).join("");
  let tName,tbd=0,tbk=0;
  if(T){const th=hubByCode[T];tName=th?th.name:T;tbd=th?demandOf(th):0;tbk=th?demandKgOf(th):0;} else tName="BC MỚI";
  rows+=`<tr style="font-weight:700"><td class="l">${esc(tName)}</td><td>${dcell(tbd,tbk)}</td><td>${dcell(tbd+gainDon,tbk+gainKg)}</td><td class="delta-pos">+${fmt(gainDon)} đ<br>+${fmt(gainKg)} kg</td></tr>`;
  const tableHTML=`<table class="bal"><tr><th>BC</th><th>Trước</th><th>Sau</th><th>Δ</th></tr>${rows}</table>`;
  const regs=[...new Set(nws.map(w=>w.region).filter(Boolean))];
  const est=planStaffEstimate(gainDon, totDon, T);
  planSnapshot={ selCodes:oldCodes, selNewCodes:codes, selNames:oldNames, newWardNames:nws.map(w=>w.name),
    regions:regs, affectedHubs:[...Object.keys(loss),...(T?[T]:[])], target:tName, totDon, totKg, gainDon, gainKg, tableHTML, staff:est,
    lossByHub:Object.entries(loss).map(([hc,o])=>({hub:hc,don:o.don,kg:o.kg})) };
  showDetail(`<h3>Quy hoạch (what-if)</h3>
    <div style="color:#6b7280">Đã chọn <b>${nws.length}</b> phường mới · <b>${fmt(totDon)}</b> đơn · <b>${fmt(totKg)}</b> kg/ngày → gán cho <b>${esc(tName)}</b></div>
    ${tableHTML}
    ${est?`<div class="note">${staffNoteHTML(est)}</div>`:""}
    <div class="note">Cầu phường cũ "Nhập một phần" đã chia 1/n theo quy ước · BC = bưu cục giao.</div>${PROP_BTN}`);
}
function renderPlan(){
  if(ui.planUnit==="new") return renderPlanNew();   // tô theo phường mới -> dùng olds[] (1/n)
  const sel=[...plan.selected.keys()].map(c=>wardByCode[c]).filter(Boolean);
  renderLegend();   // cập nhật ghi chú "phường đang chọn"
  if(!sel.length){ $("#detail").classList.add("hidden"); return; }
  const T = ui.planTarget || null;                  // null = BC mới; else hub_code đích
  // Tách 2 phía: phần LẤY (pv/pw) trừ ở BC lấy; phần GIAO (dv/dw) trừ ở BC giao.
  const loss={};                                    // bc -> {don,kg}
  let gainDon=0, gainKg=0, totDon=0, totKg=0;
  const addLoss=(bc,don,kg)=>{ if(!bc||bc===T) return; const o=loss[bc]||(loss[bc]={don:0,kg:0}); o.don+=don; o.kg+=kg; };
  sel.forEach(w=>{
    totDon+=w.pv+w.dv; totKg+=(w.pw||0)+(w.dw||0);
    if(w.pick_hub!==T){     addLoss(w.pick_hub,     w.pv, w.pw||0); gainDon+=w.pv; gainKg+=w.pw||0; }  // LẤY
    if(w.delivery_hub!==T){ addLoss(w.delivery_hub, w.dv, w.dw||0); gainDon+=w.dv; gainKg+=w.dw||0; }  // GIAO
  });
  const dcell=(don,kg)=>`${fmt(don)} đ<br>${fmt(kg)} kg`;
  let rows=Object.entries(loss).sort((a,b)=>b[1].don-a[1].don).map(([hc,o])=>{
    const h=hubByCode[hc], bd=h?demandOf(h):0, bk=h?demandKgOf(h):0;
    return `<tr><td class="l">${esc(h?h.name:hc)}</td><td>${dcell(bd,bk)}</td><td>${dcell(bd-o.don,bk-o.kg)}</td>`+
      `<td class="delta-neg">−${fmt(o.don)} đ<br>−${fmt(o.kg)} kg</td></tr>`;
  }).join("");
  // dòng đích
  let tName,tbd=0,tbk=0;
  if(T){ const th=hubByCode[T]; tName=th?th.name:T; tbd=th?demandOf(th):0; tbk=th?demandKgOf(th):0; }
  else tName="BC MỚI";
  rows+=`<tr style="font-weight:700"><td class="l">${esc(tName)}</td><td>${dcell(tbd,tbk)}</td>`+
    `<td>${dcell(tbd+gainDon,tbk+gainKg)}</td><td class="delta-pos">+${fmt(gainDon)} đ<br>+${fmt(gainKg)} kg</td></tr>`;
  const tableHTML=`<table class="bal"><tr><th>BC</th><th>Trước</th><th>Sau</th><th>Δ</th></tr>${rows}</table>`;
  const regs=[...new Set(sel.map(w=>w.region).filter(Boolean))];
  // BC bị ảnh hưởng (mất cầu) + BC đích -> để đưa vào khung chụp map
  const affectedHubs=[...Object.keys(loss), ...(T?[T]:[])];
  // phường MỚI chứa các phường cũ đã chọn
  const selSet=new Set(sel.map(w=>w.ward_code));
  const newWardNames=[...new Set((DATA.rez&&DATA.rez.new_wards||[])
    .filter(nw=>(nw.old_codes||[]).some(c=>selSet.has(c))).map(nw=>nw.name))];
  // lưu snapshot phục vụ "Lưu đề xuất"
  const est=planStaffEstimate(gainDon, totDon, T);
  planSnapshot={ selCodes:sel.map(w=>w.ward_code), selNewCodes:[...plan.selNew.keys()], selNames:sel.map(w=>w.name),
    newWardNames, regions:regs, affectedHubs, target:tName, totDon, totKg, gainDon, gainKg, tableHTML, staff:est,
    lossByHub:Object.entries(loss).map(([hc,o])=>({hub:hc,don:o.don,kg:o.kg})) };
  showDetail(`<h3>Quy hoạch (what-if)</h3>
    <div style="color:#6b7280">Đã chọn <b>${sel.length}</b> phường · <b>${fmt(totDon)}</b> đơn · <b>${fmt(totKg)}</b> kg/ngày → gán cho <b>${esc(tName)}</b></div>
    ${tableHTML}
    ${est?`<div class="note">${staffNoteHTML(est)}</div>`:""}
    <div class="note">Tách 2 phía: phần LẤY trừ ở BC lấy, phần GIAO trừ ở BC giao.</div>
    <button id="save-prop" onclick="saveProposal()" style="width:100%;margin-top:8px;padding:9px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:600;cursor:pointer">💾 Lưu đề xuất</button>`);
}
let planSnapshot=null;
// Bounds các phường đã chọn — tính từ POLYGON thật (geojson) để chính xác & không lệ thuộc centroid
// (nhiều phường 1A… Hà Nội không có trong ward_centroids -> trước đây bounds null -> không zoom).
// Khung bao các polygon (codes) trong 1 geojson, khớp theo 'key' (ward_code hoặc new_ward_code)
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
function selectedBounds(codes){  // phường cũ (fallback centroid nếu polygon thiếu)
  let b=wardBounds(codes, DATA.geojson, "ward_code");
  if(b) return b;
  let lo=[181,91],hi=[-181,-91],n=0; (codes||[]).forEach(c=>{const p=DATA.cent&&DATA.cent[c]; if(p){n++;lo[0]=Math.min(lo[0],p[0]);lo[1]=Math.min(lo[1],p[1]);hi[0]=Math.max(hi[0],p[0]);hi[1]=Math.max(hi[1],p[1]);}});
  return n?[lo,hi]:null;
}
function captureMap(cb){
  // đóng khung theo polygon PHƯỜNG MỚI (3321) nếu quy hoạch theo phường mới; else theo phường cũ
  const sn=planSnapshot||{};
  const b = (sn.selNewCodes&&sn.selNewCodes.length)
    ? wardBounds(sn.selNewCodes, DATA.geojsonNew, "new_ward_code")
    : selectedBounds(sn.selCodes);
  if(b) map.fitBounds(b,{padding:70,maxZoom:13,duration:0});
  let done=false;
  const cap=()=>{ if(done)return; done=true; try{cb(map.getCanvas().toDataURL("image/png"));}catch(e){cb(null);} };
  map.once("idle",()=>setTimeout(cap,200));   // chụp sau khi map render xong khung mới
  setTimeout(cap,4500);                         // fallback nếu idle chậm
}
let proposalHTML="", mapImg=null;
function buildProposalHTML(img, name, comment, editable, review){
  const s=planSnapshot, now=new Date().toLocaleString("vi-VN");
  const title=(name||"").trim()||"Đề xuất Quy hoạch Bưu Cục";
  const reviewBlock=(review||"").trim()
    ? `<h2>🤖 Đánh giá của AI</h2><div class="airev">${mdSafe(review)}</div>` : "";
  // Mục Ghi chú: chế độ editable -> ô contenteditable ngay trong report (gõ thẳng); else -> render tĩnh (cho PDF)
  const cmtBlock = editable
    ? `<h2>Ghi chú <span id="cmt-cc">${(comment||"").length}/1000</span></h2>
       <div class="cmt edit" id="cmt-edit" contenteditable="true" data-ph="Nhập ghi chú / nhận xét tại đây… (≤1000 ký tự)">${esc(comment||"")}</div>`
    : ((comment||"").trim()?`<h2>Ghi chú</h2><div class="cmt">${esc(comment.trim())}</div>`:"");
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
  <div class="meta">Người lập: <b>${esc(ui.username||"")}</b> · ${now} · Vùng: ${esc(s.regions.join(", ")||"–")}</div>
  <h2>Tóm tắt</h2>
  <div class="sum">
    <div><span>Phường tái quy hoạch</span><b>${fmt(s.selCodes.length)}</b></div>
    <div><span>Cầu di chuyển</span><b>${fmt(s.totDon)} đơn/ngày</b></div>
    <div><span>Khối lượng</span><b>${fmt(s.totKg)} kg/ngày</b></div>
    <div><span>Gán cho</span><b>${esc(s.target)}</b></div>
    ${s.staff?`<div><span>Định biên cần</span><b>${s.staff.isNew?"~"+fmt(s.staff.addNV)+" NV (BC mới)":"+~"+fmt(s.staff.addNV)+" NV"}</b></div>`:""}
  </div>
  <h2>Ảnh hưởng trước / sau khi quy hoạch (theo BC · đơn + kg)</h2>
  ${s.tableHTML}
  <h2>Bản đồ khu vực thay đổi</h2>
  ${img?`<img src="${img}">`:"<i>(không chụp được bản đồ)</i>"}
  <h2>Phường mới tái quy hoạch (${(s.newWardNames||[]).length})</h2>
  <div class="wards">${(s.newWardNames||[]).map(esc).join(" · ")||"–"}</div>
  <h2>Phường cũ (nguồn cầu) (${s.selCodes.length})</h2>
  <div class="wards">${s.selNames.map(esc).join(" · ")}</div>
  ${reviewBlock}
  ${cmtBlock}
  ${cmtScript}
  </body></html>`;
}
function saveProposal(){
  if(!planSnapshot||!planSnapshot.selCodes.length) return;
  if(ui.tab!=="map") switchTab("map");
  const btn=$("#save-prop"); if(btn){btn.disabled=true; btn.textContent="📸 Đang chụp bản đồ…";}
  captureMap(img=>{
    mapImg=img;
    const nameEl=$("#prop-name"); if(nameEl) nameEl.value="";   // để trống, user tự đặt tên (bắt buộc)
    aiReview=null;   // reset đánh giá AI cho lần mở mới
    wirePropInputs(); renderPreviewEditable();   // dựng report với ô Ghi chú gõ thẳng được
    $("#prop-status").textContent="";
    lastPdf=null;   // reset nút xác nhận (lần mở mới)
    const rv=$("#prop-review"); if(rv){ rv.disabled=false; rv.style.display=""; rv.textContent="🤖 AI đánh giá đề xuất"; }
    const c=$("#prop-confirm"); if(c){ c.disabled=true; c.textContent="✓ Xác nhận đề xuất (gửi OA)"; c.style.background="#16a34a"; c.onclick=confirmProposal; }  // khoá đến khi đủ tên + đánh giá AI
    modalOpen($("#prop-modal"));
    if(btn){btn.disabled=false; btn.textContent="💾 Lưu đề xuất";}
  });
}
let _propWired=false, aiReview=null;
// Dựng bản xem trước có ô Ghi chú contenteditable. Chỉ gọi khi MỞ dialog -> ghi chú rỗng, chưa có đánh giá.
function renderPreviewEditable(){
  const fr=$("#prop-frame"); if(fr) fr.srcdoc=buildProposalHTML(mapImg, ($("#prop-name")||{}).value||"", "", true, null);
}
// Dựng lại preview GIỮ ghi chú đang gõ + kèm đánh giá AI (gọi sau khi có đánh giá).
function rebuildPropPreview(){
  const cmt=readComment(), fr=$("#prop-frame");
  if(fr) fr.srcdoc=buildProposalHTML(mapImg, ($("#prop-name")||{}).value||"", cmt, true, aiReview);
}
// Nút Xác nhận chỉ mở khi ĐỦ tên + đã có đánh giá AI (bắt buộc).
function updatePropConfirm(){
  const c=$("#prop-confirm"); if(!c||lastPdf) return;
  c.disabled = !((($("#prop-name")||{}).value||"").trim() && aiReview);
}
// Gợi ý vị trí: trọng tâm đơn hàng (centroid phường cũ weighted theo cầu) + phường cầu cao
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
    .sort((a,b)=>b.d-a.d).slice(0,4).filter(x=>x.d>0).map(x=>`${x.n} (${fmt(x.d)} đơn)`);
  return {center, top};
}
async function reviewProposal(){
  if(!planSnapshot) return;
  const btn=$("#prop-review"), st=$("#prop-status");
  if(btn){ btn.disabled=true; btn.textContent="⏳ AI đang đánh giá…"; }
  if(st) st.textContent="";
  const s=planSnapshot, geo=planGeoHint(s);
  const payload={ target:s.target, totDon:s.totDon, totKg:s.totKg, gainDon:s.gainDon, gainKg:s.gainKg,
    newWardCount:(s.newWardNames||[]).length, oldWardCount:s.selCodes.length,
    staffAdd:s.staff?s.staff.addNV:0, staffMed:s.staff?s.staff.med:0, staffCur:s.staff?s.staff.curStaff:null,
    isNew:s.staff?s.staff.isNew:(s.target==="BC MỚI"), affectedHubs:s.affectedHubs||[], losses:s.lossByHub||[],
    geoCenter:geo.center||"", topWards:geo.top };
  try{
    const r=await fetch("/proposal/review",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify(payload)}).then(r=>r.json());
    if(r.review){ aiReview=r.review; rebuildPropPreview(); toast("AI đã đánh giá xong đề xuất.","ok"); }
    else if(st){ st.textContent="⚠ "+(r.error||"lỗi đánh giá"); toast("Đánh giá thất bại: "+(r.error||"lỗi"),"err"); }
  }catch(e){ if(st) st.textContent="⚠ Lỗi kết nối server."; toast("Lỗi kết nối server khi đánh giá.","err"); }
  // Đánh giá CHỈ 1 LẦN: thành công -> ẨN nút (gộp còn 1 chỉ báo ở status, chừa chỗ cho ô tên).
  // Muốn đánh giá lại phải thoát dialog & tạo lại đề xuất.
  if(btn){ if(aiReview){ btn.style.display="none"; if(st) st.textContent="🤖 Đã đánh giá ✓"; }
           else { btn.disabled=false; btn.style.display=""; btn.textContent="🤖 AI đánh giá đề xuất"; } }
  updatePropConfirm();
}
// Đọc ghi chú trực tiếp từ ô contenteditable trong iframe (same-origin srcdoc).
function readComment(){
  try{ const d=$("#prop-frame").contentDocument, e=d&&d.getElementById("cmt-edit");
    return e ? e.innerText.replace(/ /g," ").replace(/\n+$/,"").trim().slice(0,1000) : ""; }
  catch(e){ return ""; }
}
function wirePropInputs(){
  if(_propWired) return; _propWired=true;
  // Đổi tên -> chỉ vá tiêu đề <h1> trong iframe (KHÔNG rerender, kẻo mất ghi chú đang gõ) + bật/khoá nút xác nhận
  const nm=$("#prop-name"); if(nm) nm.oninput=()=>{
    updatePropConfirm();
    try{ const h=$("#prop-frame").contentDocument.querySelector("h1"); if(h) h.textContent=nm.value.trim()||"Đề xuất Quy hoạch Bưu Cục"; }catch(e){}
  };
}
let lastPdf=null;   // {b64, name} của đề xuất vừa tạo
async function confirmProposal(){
  const btn=$("#prop-confirm"), st=$("#prop-status");
  if(btn){btn.disabled=true; btn.textContent="⏳ Đang tạo PDF & gửi…";}
  if(st) st.textContent="";
  const title=(($("#prop-name")||{}).value||"").trim();
  if(!title){ if(st) st.textContent="⚠ Vui lòng nhập tên đề xuất."; const n=$("#prop-name"); if(n) n.focus();
    if(btn){btn.disabled=false; btn.textContent="✓ Xác nhận đề xuất (gửi OA)";} return; }
  if(!aiReview){ if(st) st.textContent="⚠ Cần bấm '🤖 AI đánh giá đề xuất' trước khi gửi.";
    if(btn){btn.disabled=false; btn.textContent="✓ Xác nhận đề xuất (gửi OA)";} return; }
  proposalHTML=buildProposalHTML(mapImg, title, readComment(), false, aiReview);   // bản sạch (escaped, không editable) cho PDF
  try{
    const r=await fetch("/proposal",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify({html:proposalHTML, title})}).then(r=>r.json());
    if(st) st.textContent = r.ok ? (r.sent? "✅ Đã gửi vào OA GTalk." : ("✅ Đã tạo PDF. "+(r.note||""))) : ("⚠ "+(r.error||"lỗi"));
    if(r.ok) toast(r.sent? "Đã gửi đề xuất vào OA GTalk." : "Đã tạo PDF đề xuất.","ok");
    else toast("Gửi đề xuất thất bại: "+(r.error||"lỗi"),"err");
    if(r.pdf){   // có PDF -> đổi nút thành Download
      lastPdf={b64:r.pdf, name:r.filename||"de-xuat.pdf"};
      if(btn){ btn.disabled=false; btn.textContent="⬇ Download Đề xuất"; btn.style.background="#2563eb"; btn.onclick=downloadProposal; }
      return;
    }
  }catch(e){ if(st) st.textContent="⚠ Lỗi kết nối server."; toast("Lỗi kết nối server khi gửi đề xuất.","err"); }
  if(btn){btn.disabled=false; btn.textContent="✓ Xác nhận đề xuất (gửi OA)";}
}
function downloadProposal(){
  if(!lastPdf) return;
  const bin=atob(lastPdf.b64), arr=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
  const url=URL.createObjectURL(new Blob([arr],{type:"application/pdf"}));
  const a=document.createElement("a"); a.href=url; a.download=lastPdf.name; document.body.appendChild(a); a.click();
  a.remove(); setTimeout(()=>URL.revokeObjectURL(url),4000);
}
// "Gán cụm cho" = combobox có search (native select không tìm/scroll nổi ~1166 BC khi admin)
let PT_HUBS=[];
function fillPlanTargets(){
  PT_HUBS=DATA.hubs.filter(h=>h.assigned&&h.type==="express"&&regOK(h._region))
    .map(h=>({code:h.hub_code, name:h.name||h.hub_code})).sort((a,b)=>a.name.localeCompare(b.name));
  if(ui.planTarget && !PT_HUBS.some(h=>h.code===ui.planTarget)) ui.planTarget=null;   // đích ngoài vùng -> reset
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
  box.innerHTML=`<div class="pt-item" data-code="">— BC mới —</div>`+
    shown.map(h=>`<div class="pt-item" data-code="${esc(h.code)}">${esc(h.name)} <span class="pt-code">${esc(h.code)}</span></div>`).join("")+
    (items.length>cap?`<div class="pt-more">… còn ${items.length-cap} BC — gõ thêm để lọc</div>`:"");
  box.querySelectorAll(".pt-item").forEach(el=>el.onclick=()=>{
    ui.planTarget=el.dataset.code||null; syncPtInput(); hidePtList(); renderPlan();
  });
  box.style.display="block";
}
function hidePtList(){ const b=$("#pt-list"); if(b) b.style.display="none"; syncPtInput(); }

// ---------- UI chrome ----------
function renderTopbar(){
  const m=DATA.meta, c=m.counts||{};
  $("#meta").textContent=`${m.date} · ${fmt(c.express_assigned)} BC · ${fmt(c.wards)} phường`;
  $("#meta").title=`${m.date} · ${fmt(c.express_assigned)} BC express · ${fmt(c.orphan_express)} chuyên dụng · ${fmt(c.wards)} phường`;
  const d=m.diff||{};
  $("#diff").textContent = d.first_run?"(lần chạy đầu)":(d.summary||"");
}
function renderLegend(){
  let h=`<div class="lg"><b style="font-size:12px">Chú giải</b></div>`;
  const items = ui.colormode==="region"
    ? Object.entries(REGION_COLORS).map(([k,c])=>[`Vùng ${k}`,c,"circle"])
    : [["Bưu cục",ROLE_COLORS.territorial,"circle"],["Chuyên LẤY",ROLE_COLORS.pickup_only,"circle"],
       ["Cồng kềnh/GIAO",ROLE_COLORS.bulky_delivery,"circle"],["Chuyên dụng hỗn hợp",ROLE_COLORS.special_mixed,"circle"],
       ["Kho Hàng Nặng (B2B)",TYPE_COLORS.B2B,"circle"],["Kho Trung Chuyển/Chuyển Tiếp",TYPE_COLORS.transit,"circle"]];
  items.forEach(([t,c,sh])=>h+=`<div class="lg"><span class="sw ${sh}" style="background:rgb(${c.join(',')})"></span>${t}</div>`);
  if(effColorMode()==="hub") h+=`<div class="lg" style="color:#6b7280">🎨 Phủ phường: mỗi màu = 1 bưu cục (lãnh thổ giao)</div>`;
  if(plan.selected.size) h+=`<div class="lg"><span class="sw" style="background:rgb(220,38,38)"></span>Phường đang chọn (tô tay)</div>`;
  if(ui.jt) h+=`<div class="lg"><span class="sw circle" style="background:rgb(20,20,20)"></span>BC J&T (tham khảo)</div>`;
  if(ui.optov.on){
    h+=`<div class="lg"><b style="font-size:12px">⚙️ Optimizer</b></div>`;
    if(ui.optov.markers){ h+=`<div class="lg"><span class="sw circle" style="background:transparent;border:2px solid rgb(220,38,38)"></span>BC nên đóng/rà soát</div>`;
      h+=`<div class="lg"><span class="sw circle" style="background:transparent;border:2px solid rgb(245,158,11)"></span>BC nên tách/mở rộng</div>`; }
    if(ui.optov.reassign) h+=`<div class="lg"><span class="sw" style="background:rgb(37,99,235)"></span>Phường nên đổi BC →</div>`;
    if(ui.optov.merge) h+=`<div class="lg"><span class="sw" style="background:rgb(220,38,38)"></span>Gộp khi đóng BC</div>`;
    if(ui.optov.ws4){ h+=`<div class="lg"><span class="sw circle" style="background:rgb(220,38,38)"></span>WS: Khoảng cách xa</div>`;
      h+=`<div class="lg"><span class="sw circle" style="background:rgb(16,185,129)"></span>WS: Greenfield</div>`;
      h+=`<div class="lg"><span class="sw circle" style="background:rgb(245,158,11)"></span>WS: Đối đầu J&T</div>`; }
  }
  $("#legend").innerHTML=h;
}
function renderDQ(){
  const m=DATA.meta;
  const noGeo=DATA.hubs.filter(h=>h.assigned&&h.missing_geo).length;
  const noRe=DATA.hubs.filter(h=>h.assigned&&h.missing_realestate).length;
  const orphanNoGeo=DATA.hubs.filter(h=>h.type==="express"&&h.role!=="territorial"&&!h.has_geo).length;
  let s=`<b>Data-quality</b><br>`;
  if(!DATA.geojson) s+=`▸ Chưa có polygon → layer phường tắt. Đặt <code>wards.geojson</code> vào data/out.<br>`;
  s+=`▸ ${noGeo} BC gán phường thiếu toạ độ<br>▸ ${noRe} BC thiếu mặt bằng<br>▸ ${orphanNoGeo} BC chuyên dụng thiếu toạ độ (không chấm được)`;
  (m.warnings||[]).forEach(w=>s+=`<br>⚠ ${w}`);
  $("#dq").innerHTML=s;
}
// ---------- Scorecard theo vùng ----------
function daysToExpiry(s){
  const m=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec((s||"").trim()); if(!m) return null;
  if(+m[3]<2000) return null;   // 30/12/1899 = ô ngày trống của Sheets -> coi như không có hạn
  const d=new Date(+m[3], +m[2]-1, +m[1]);
  const today=new Date(DATA.meta.date||Date.now());
  return Math.round((d-today)/86400000);
}
function aggregateRegions(){
  const wardCnt={}; DATA.wards.forEach(w=>wardCnt[w.region]=(wardCnt[w.region]||0)+1);
  const A={};
  DATA.hubs.forEach(h=>{
    if(!(h.assigned&&h.type==="express")) return; const r=h._region; if(!r) return;
    // volRE/wtRE = vol/kg của BC CÓ mặt bằng (đồng/đơn, đơn/m², kg/m² đồng bộ tử-mẫu, không phồng khi thiếu RE)
    const a=A[r]||(A[r]={region:r,bc:0,phuong:wardCnt[r]||0,terr:0,vol:0,volS:0,volRE:0,wt:0,wtRE:0,area:0,rent:0,exp90:0,nore:0,staff:0});
    a.bc++; a.terr+=demandOf(h); a.staff+=h.staff||0;
    if(h.actual){ a.vol+=h.actual.pv+h.actual.dv; a.wt+=h.actual.pw+h.actual.dw;
      if(h.staff) a.volS+=h.actual.pv+h.actual.dv; }   // volS = volume của BC CÓ staff (để đơn/NV không bị thổi)
    if(h.realestate&&h.realestate.usable_area){ a.area+=h.realestate.usable_area; a.rent+=h.realestate.rent||0;
      if(h.actual){ a.volRE+=h.actual.pv+h.actual.dv; a.wtRE+=h.actual.pw+h.actual.dw; }
      const dd=daysToExpiry(h.realestate.expiry); if(dd!=null&&dd>=0&&dd<=90) a.exp90++; }
    else a.nore++;
  });
  return Object.values(A).map(a=>({...a,
    dpo: a.volRE? a.rent/(a.volRE*30):0, dpm: a.area? a.volRE/a.area:0,
    dpw: a.wtRE?  a.rent/(a.wtRE*30):0,  kgm2: a.area? a.wtRE/a.area:0,
    dpn: a.staff? a.volS/a.staff:0 }));   // đơn/NV/ngày (năng suất; chỉ vol của BC có staff)
}
const DECIMAL=new Set(["dpm","kgm2"]);   // cột số lẻ 1 chữ số
const SC_COLS=[["region","Vùng"],["bc","BC"],["phuong","Phường"],
  ["vol","Volume (đơn)/ngày"],["wt","Khối lượng (kg)/ngày"],["area","m²"],["staff","NV"],
  ["dpo","Đồng/đơn"],["dpm","Đơn/m²"],["dpw","Đồng/kg"],["kgm2","Kg/m²"],["dpn","Đơn/NV"],
  ["exp90","Hợp đồng ≤ 90 ngày"]];
let scSort={key:"vol",desc:true};
function renderScorecard(){
  const all=aggregateRegions();
  const rows=all.filter(r=>regOK(r.region));
  const p=(arr,q)=>{const s=arr.slice().sort((a,b)=>a-b);return s[Math.floor(q*(s.length-1))];};
  const TH={};   // ngưỡng p75/p25 tính trên TOÀN BỘ vùng (không đổi theo filter)
  ["dpo","dpm","dpw","kgm2","dpn"].forEach(k=>{const v=all.map(r=>r[k]); TH[k]={hi:p(v,.75),lo:p(v,.25)};});
  rows.sort((a,b)=>{ const k=scSort.key, va=a[k], vb=b[k];
    return (typeof va==="string"? va.localeCompare(vb): va-vb)*(scSort.desc?-1:1); });
  const tot={bc:0,phuong:0,terr:0,vol:0,volS:0,volRE:0,wt:0,wtRE:0,area:0,rent:0,exp90:0,nore:0,staff:0};
  rows.forEach(r=>["bc","phuong","terr","vol","volS","volRE","wt","wtRE","area","rent","exp90","nore","staff"].forEach(k=>tot[k]+=r[k]));
  tot.dpo=tot.volRE?tot.rent/(tot.volRE*30):0; tot.dpm=tot.area?tot.volRE/tot.area:0;
  tot.dpw=tot.wtRE?tot.rent/(tot.wtRE*30):0;   tot.kgm2=tot.area?tot.wtRE/tot.area:0;
  tot.dpn=tot.staff?tot.volS/tot.staff:0;  tot.region="TỔNG";
  const cell=(r,k)=>{
    if(k==="region") return `<td>${r.region}</td>`;
    let cls="";
    if(k==="dpo"||k==="dpw"){ if(r[k]>=TH[k].hi) cls="hot"; }        // chi phí cao = xấu
    if(k==="dpm"||k==="kgm2"){ if(r[k]>=TH[k].hi)cls="hot"; else if(r[k]<=TH[k].lo)cls="cold"; } // chật/dư
    if(k==="dpn"){ if(r[k]>=TH[k].hi)cls="hot"; else if(r[k]<=TH[k].lo)cls="cold"; }   // đơn/NV cao=thiếu người, thấp=thừa
    const txt = DECIMAL.has(k)? r[k].toFixed(1) : fmt(r[k]);
    // kênh dự phòng ngoài màu (a11y mù màu): ▲ cao / ▼ thấp + tooltip
    const mk = cls==="hot" ? '<span class="mk" aria-hidden="true">▲</span>' : cls==="cold" ? '<span class="mk" aria-hidden="true">▼</span>' : "";
    const ti = cls==="hot" ? ' title="Cao (≥P75)"' : cls==="cold" ? ' title="Thấp (≤P25)"' : "";
    return `<td class="${cls}"${ti}>${mk}${txt}</td>`;
  };
  const th=SC_COLS.map(([k,l])=>`<th data-k="${k}">${l}${scSort.key===k?(scSort.desc?" ▾":" ▴"):""}</th>`).join("");
  const body=rows.map(r=>`<tr class="region-row" data-r="${r.region}">`+
    SC_COLS.map(([k])=>cell(r,k)).join("")+`</tr>`).join("");
  const totrow=`<tr class="total"><td>TỔNG (${rows.length} vùng)</td>`+
    SC_COLS.slice(1).map(([k])=> DECIMAL.has(k)?`<td>${tot[k].toFixed(1)}</td>`:`<td>${fmt(tot[k])}</td>`).join("")+`</tr>`;
  $("#sc-table").innerHTML=`<table class="sc"><thead><tr>${th}</tr></thead><tbody>${body}${totrow}</tbody></table>`;
  $("#sc-table").querySelectorAll("th").forEach(t=>t.onclick=()=>{ const k=t.dataset.k;
    if(scSort.key===k) scSort.desc=!scSort.desc; else {scSort.key=k;scSort.desc=true;} renderScorecard(); });
  $("#sc-table").querySelectorAll("tr.region-row").forEach(tr=>tr.onclick=()=>filterRegion(tr.dataset.r));
  renderSubTables();
}

// ---------- 3 bảng per-BC: Sắp hết hạn / Sắp quá tải / Dư thừa ----------
function bcList(){
  const K=kgPerOrd()||1;
  return DATA.hubs.filter(h=>h.assigned && h.type==="express").map(h=>{
    const ac=h.actual||{pv:0,pw:0,dv:0,dw:0}, re=h.realestate;
    const vol=ac.pv+ac.dv, wt=ac.pw+ac.dw, area=re?re.usable_area:0;
    const wl=(0.4*ac.pv+ac.dv)+(0.4*ac.pw+ac.dw)/K;   // tải quy đổi (đơn-tđ): đơn & kg lấy ×0.4, kg quy ra đơn
    return { code:h.hub_code, name:h.name||h.hub_code, region:h._region, province:h._province,
      vol, wt, area, rent:re?re.rent:0, expiry:re?re.expiry:"", days:re?daysToExpiry(re.expiry):null,
      dpm: area?vol/area:0, kgm2: area?wt/area:0, diemM2: area?wl/area:0, hasRE:!!re };
  });
}
function pct(arr,q){ const s=arr.slice().sort((a,b)=>a-b); return s.length?s[Math.floor(q*(s.length-1))]:0; }
// bảng có nút thu/xổ + hộp cuộn (hiện ~10 dòng, cuộn xem tiếp). startCollapsed=true -> mặc định gập.
function subTable(rows, cols, cap, startCollapsed, tall){
  const th=cols.map(c=>`<th class="${c.l?'l':''}">${c.h}</th>`).join("");
  const body=rows.slice(0,cap).map(r=>`<tr class="region-row" data-code="${esc(r.code)}">`+
    cols.map(c=>`<td class="${c.l?'l':''}">${c.raw?c.f(r):esc(c.f(r))}</td>`).join("")+`</tr>`).join("");  // raw=true: HTML do code kiểm soát
  const cnt = rows.length>cap?`${cap}/${rows.length}`:`${rows.length}`;
  const arrow = startCollapsed?"▸ Xem chi tiết · ":"▾ Thu gọn · ";
  const note=`<div class="sc-note">Click 1 dòng để xem trên bản đồ.</div>`;
  const cls = `tbl-body${startCollapsed?' collapsed':''}${tall?' tall':''}`;
  return `<button class="tbl-toggle" data-n="${cnt}" onclick="toggleTbl(this)">${arrow}${cnt} dòng</button>`+
    `<div class="${cls}">${note}`+
    `<table class="sc"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function toggleTbl(btn){
  const body=btn.parentNode.querySelector(".tbl-body"); if(!body) return;
  const collapsed=body.classList.toggle("collapsed");   // true nếu vừa gập lại
  btn.textContent=(collapsed?"▸ Xem chi tiết · ":"▾ Thu gọn · ")+btn.dataset.n+" dòng";
}
function renderSubTables(){
  const bcs=bcList().filter(b=>regOK(b.region));
  const withRE=bcs.filter(b=>b.hasRE && b.area>0 && b.vol>0);
  const eHi=pct(withRE.map(b=>b.diemM2),.75), eLo=pct(withRE.map(b=>b.diemM2),.25);   // ngưỡng theo điểm/m²
  const base=[{h:"Vùng",f:r=>r.region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},
    {h:"ID BC",f:r=>r.code,l:1},{h:"Tên BC",f:r=>r.name,l:1}];
  const loadCols=[...base,{h:"Volume (đơn)/ngày",f:r=>fmt(r.vol)},{h:"Khối lượng (kg)/ngày",f:r=>fmt(r.wt)},
    {h:"m²",f:r=>fmt(r.area)},{h:"Đơn/m²",f:r=>r.dpm.toFixed(1)},{h:"Kg/m²",f:r=>r.kgm2.toFixed(1)},{h:"Điểm/m²",f:r=>r.diemM2.toFixed(1)}];
  const CAP=5000;   // render hết (hộp cuộn lo phần dài), không cắt dữ liệu
  // Sắp hết hạn: còn lại ≤ 90 ngày (gồm đã hết hạn)
  const exp=bcs.filter(b=>b.days!=null && b.days<=90).sort((a,b)=>a.days-b.days);
  $("#sc-expire").innerHTML=subTable(exp,[...base,
    {h:"Hạn HĐ",f:r=>r.expiry||"–",l:1},
    {h:"Còn lại",raw:true,f:r=>r.days<0?`<span class="warn">đã hết ${-r.days}d</span>`:`${r.days}d`},
    {h:"Tiền thuê",f:r=>fmt(r.rent)},{h:"m²",f:r=>fmt(r.area)}],CAP,true);
  // Sắp quá tải: điểm/m² ≥ P75 (tải quy đổi đơn+kg trên mặt bằng)
  $("#sc-overload").innerHTML=subTable(withRE.filter(b=>b.diemM2>=eHi).sort((a,b)=>b.diemM2-a.diemM2),loadCols,CAP,true);
  // Dư thừa: điểm/m² ≤ P25
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
  $("#diff").innerHTML = ui.regionFilter? `Lọc vùng: <b>${ui.regionFilter}</b> <span class="chip" id="clearf">✕ bỏ lọc</span>`:"";
  const c=$("#clearf"); if(c) c.onclick=()=>{ ui.regionFilter=null; draw(); renderTopbar(); };
}
function setRegionFilter(r){
  ui.regionFilter = r || null;
  fillPlanTargets();   // danh sách BC đích thu hẹp theo vùng đang lọc
  renderLegend();      // cập nhật chú giải (1 vùng -> tô theo BC)
  if(map && map.loaded && map.loaded()) draw();
  if(ui.tab==="scorecard") renderScorecard();
  else if(ui.tab==="optimizer") renderOptimizer();
  else if(ui.tab==="rezone") renderRezone();
  else if(ui.regionFilter){ // map: fit tới hub của vùng
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
  const ctgl=$("#ctrl-toggle"); if(ctgl) ctgl.style.display = (name==="map") ? "" : "none";  // nút ≡ chỉ ở tab Bản đồ
  if(name!=="map") $("#detail").classList.add("hidden");
  if(name==="scorecard") renderScorecard();
  if(name==="optimizer") renderOptimizer();
  if(name==="rezone") renderRezone();
  if(name==="chatlog") renderChatlog();
}
function renderChatlog(){
  const c=$("#chatlog");
  if(!c.dataset.init){
    c.innerHTML=`<div class="sc-head"><h2>Lịch sử chat</h2><div class="sc-note">Thống kê tổng hợp + xem câu hỏi/trả lời theo ngày (chỉ admin).</div></div>
      <div id="cl-stats"></div>
      <div class="sc-sub">
        <h3>📄 Nhật ký theo ngày <button id="cs-ai" style="margin-left:8px">🤖 AI tổng hợp câu hỏi</button></h3>
        <div style="margin-bottom:10px">Ngày: <input type="date" id="cl-date" value="">
          <button id="cl-load">Tải</button>
          <span style="color:#9ca3af;font-size:11px">(để trống = 15 ngày gần nhất)</span></div>
        <div id="cs-ai-box"></div>
        <div id="cl-table"></div></div>`;
    c.dataset.init="1"; $("#cl-load").onclick=loadChatlog;
    const aib=$("#cs-ai"); if(aib) aib.onclick=summarizeQuestions;
  }
  loadChatstats(); loadChatlog();
}
async function loadChatstats(){
  const box=$("#cl-stats"); if(!box) return; box.innerHTML="Đang tải thống kê…";
  try{
    const j=await fetch('/chatstats').then(r=>r.json());
    if(j.error){ box.innerHTML="⚠ "+j.error; return; }
    if(!j.total){ box.innerHTML='<div class="sc-note">Chưa có dữ liệu chat để thống kê.</div>'; return; }
    const card=(t,v)=>`<div class="stat-card"><div class="sv">${fmt(v)}</div><div class="st">${t}</div></div>`;
    const period=j.period?`${j.period[0]} → ${j.period[1]}`:"";
    const mx=Math.max(1,...j.by_day.map(d=>d[1]));
    const bars=j.by_day.slice(-30).map(([d,c])=>`<div class="bar-row"><span class="bd">${d.slice(5)}</span><span class="bb"><i style="width:${Math.round(c/mx*100)}%"></i></span><span class="bn">${c}</span></div>`).join("");
    const utbl=`<table class="sc"><thead><tr><th class="l">User</th><th>Lượt</th></tr></thead><tbody>`+
      j.by_user.map(([u,c])=>`<tr><td class="l">${esc(u)}</td><td>${c}</td></tr>`).join("")+`</tbody></table>`;
    const rtbl=`<table class="sc"><thead><tr><th class="l">Vùng</th><th>Lượt</th></tr></thead><tbody>`+
      j.by_region.map(([r,c])=>`<tr><td class="l">${esc(r)}</td><td>${c}</td></tr>`).join("")+`</tbody></table>`;
    box.innerHTML=`<div class="sc-note">Tổng hợp toàn bộ log đang lưu${period?` (${period})`:""}</div>
      <div class="stat-cards">${card("Tổng lượt",j.total)}${card("User hoạt động",j.users)}${card("Ngày có log",j.days)}${card("'Không có dữ liệu'",j.nodata)}${card("Lỗi",j.errors)}</div>
      <div class="stat-grid">
        <div class="stat-col"><h4>Lượt theo ngày (30 gần nhất)</h4>${bars||'<div class="sc-note">—</div>'}</div>
        <div class="stat-col"><h4>Theo user</h4>${utbl}</div>
        <div class="stat-col"><h4>Theo vùng</h4>${rtbl}</div>
      </div>`;
  }catch(e){ box.innerHTML="⚠ Lỗi tải thống kê."; }
}
async function summarizeQuestions(){
  const btn=$("#cs-ai"), out=$("#cs-ai-box"); if(!out) return;
  if(btn){ btn.disabled=true; btn.textContent="🤖 Đang tổng hợp…"; }
  out.innerHTML=`<div class="sc-note">AI đang gom nhóm câu hỏi (15 ngày gần nhất)…</div>`;
  try{
    const j=await fetch('/chatsummary?days=15').then(r=>r.json());
    if(j.error){ out.innerHTML="⚠ "+j.error; }
    else out.innerHTML=`<div class="ai-summary">${mdSafe(j.summary)}</div>`;
  }catch(e){ out.innerHTML="⚠ Lỗi gọi AI."; }
  if(btn){ btn.disabled=false; btn.textContent="🤖 AI tổng hợp câu hỏi"; }
}
async function loadChatlog(){
  const d=$("#cl-date").value; $("#cl-table").innerHTML="Đang tải…";
  const url = d ? `/chatlog?date=${d}&n=1500` : `/chatlog?days=15&n=1500`;  // trống = 15 ngày gần nhất
  try{
    const j=await fetch(url).then(r=>r.json());
    if(j.error){ $("#cl-table").innerHTML="⚠ "+j.error; return; }
    if(!j.log||!j.log.length){ $("#cl-table").innerHTML=`Chưa có chat (${j.date}).`; return; }
    const rows=j.log.slice().reverse().map(e=>`<tr><td class="l">${esc((e.ts||"").replace("T"," "))}</td>`+
      `<td class="l">${esc(e.user||"")}</td><td class="l">${esc(Array.isArray(e.regions)?e.regions.join(","):(e.regions||""))}</td>`+
      `<td class="l">${esc(e.question||"")}</td><td class="l">${esc((e.answer||e.error||"").slice(0,500))}</td></tr>`).join("");
    $("#cl-table").innerHTML=`<div class="sc-note">${j.count} lượt · ${j.date}</div>`+
      `<table class="sc"><thead><tr><th class="l">Thời gian</th><th class="l">User</th><th class="l">Vùng</th><th class="l">Câu hỏi</th><th class="l">Trả lời</th></tr></thead><tbody>${rows}</tbody></table>`;
  }catch(e){ $("#cl-table").innerHTML="⚠ Lỗi tải log."; }
}

// ---------- Re-zone tab ----------
function renderRezone(){
  const r = DATA.rez;
  if(!r){ $("#rz-stats").innerHTML="Chưa có rezone.json (cần merge_ward + chạy build)."; return; }
  const s=r.stats;
  const scoped = r.new_wards.filter(w=>regOK(w.region));   // lọc theo vùng đang xem
  const filtered = ui.regionFilter || (Array.isArray(ui.allowed) && ui.allowed.length===1);
  if(filtered){
    const cl=scoped.filter(w=>w.status==="clean").length, sp=scoped.filter(w=>w.status==="split").length, em=scoped.filter(w=>w.status==="empty").length;
    $("#rz-stats").innerHTML=`Vùng đang xem: <b>${fmt(scoped.length)}</b> phường mới · <b>${fmt(cl)}</b> sạch (1 BC) · `+
      `<b>${fmt(sp)}</b> bị xé → auto gán theo đa số cầu · <b>${em}</b> rỗng cầu. Nguyên tắc: 1 phường mới = 1 BC (không xé lẻ).`;
  } else {
    $("#rz-stats").innerHTML=`<b>${fmt(s.new_wards)}</b> phường mới · <b>${fmt(s.clean)}</b> sạch (1 BC) · `+
      `<b>${fmt(s.split)}</b> bị xé → auto gán theo đa số cầu (${s.pct_split}%) · <b>${s.empty}</b> rỗng cầu · `+
      `cầu trong nhóm xé <b>${s.split_dem_pct}%</b>. Nguyên tắc: 1 phường mới = 1 BC (không xé lẻ).`;
  }
  const split = scoped.filter(w=>w.status==="split")
    .sort((a,b)=> (a.province||"").localeCompare(b.province||"","vi") || (a.name||"").localeCompare(b.name||"","vi")); // theo tỉnh mới, rồi tên phường
  $("#rz-split").innerHTML=subTable(split.map(w=>({...w,code:w.new_code})),[
    {h:"Tỉnh mới",f:w=>w.province,l:1},
    {h:"Phường mới",f:w=>`${w.name} (${w.new_code})`,l:1},
    {h:"#Phường cũ",f:w=>w.n_old},
    {h:"Cầu/ngày",f:w=>fmt(w.dem)},
    {h:"BC auto gán",f:w=>`${w.assigned_bc_name||w.assigned_bc}`,l:1},
    {h:"% cầu",f:w=>`${w.lead_share}%`},
    {h:"BC khác",f:w=>w.candidates.slice(1,4).map(c=>`${(c.bc_name||c.bc).slice(0,22)} ${c.share}%`).join(" · "),l:1}],5000,true,true);
  // click 1 dòng -> panel breakdown phường cũ (Cách B)
  document.querySelectorAll("#rz-split .region-row").forEach(tr=>tr.onclick=()=>{
    const w=DATA.rez.new_wards.find(x=>x.new_code===tr.dataset.code); if(w) showRezoneDetail(w);
  });
}
function showRezoneDetail(w){
  let rows=(w.olds||[]).map(o=>`<tr><td class="l">${esc(o.name)} (${esc(o.ward)})</td><td class="l">${esc(o.bc_name||o.bc||"–")}</td>`+
    `<td>${fmt(o.dem)}</td><td>${fmt(o.dem_kg)}</td><td>${esc(o.note==="phần"?"một phần":o.note)}</td></tr>`).join("");
  let cand=w.candidates.map(c=>`<tr${c.bc===w.assigned_bc?' style="font-weight:700;background:#dcfce7"':''}><td class="l">${esc(c.bc_name||c.bc)}</td><td>${fmt(c.dem)}</td><td>${c.share}%</td></tr>`).join("");
  const tag = w.status==="split" ? '<b style="color:#dc2626">bị xé</b>' : (w.status==="empty"?"rỗng cầu":'<b style="color:#16a34a">sạch (1 BC)</b>');
  showDetail(`<h3>${esc(w.name)}</h3><div style="color:#6b7280">${esc(w.province)} · ${w.n_old} phường cũ · cầu ${fmt(w.dem)} đơn · ${fmt(w.dem_kg)} kg /ngày · ${tag}</div>
    <div style="margin:8px 0"><b>BC ${w.status==="split"?"auto gán":"phụ trách"}:</b> ${esc(w.assigned_bc_name||w.assigned_bc||"–")}${w.lead_share?` (${w.lead_share}% cầu)`:""}</div>
    ${w.candidates.length>1?`<b style="font-size:12px">Ứng viên (theo cầu):</b>
    <table class="bal"><tr><th>BC</th><th>Cầu</th><th>%</th></tr>${cand}</table>`:""}
    ${(w.olds&&w.olds.length)?`<b style="font-size:12px">Thành phần phường cũ:</b>
    <table class="bal"><tr><th>Phường cũ</th><th>BC giao</th><th>Đơn</th><th>Kg</th><th>Nhập</th></tr>${rows}</table>`:`<div class="note">Gồm ${w.n_old} phường cũ (cùng 1 BC).</div>`}
    ${w.status==="split"?'<div class="note">Auto = đa số cầu. Override tay sẽ thêm sau.</div>':""}`);
}

// ---------- Network Optimizer tab ----------
function flyToWardOrHub(code, isHub){
  const h = isHub ? hubByCode[code] : null;
  if(h && h.has_geo){ switchTab("map"); map.flyTo({center:[h.lng,h.lat],zoom:12}); clickHub(h); return; }
  // ward: dùng centroid hub giao để bay tạm (không có toạ độ ward ở client) -> bay tới BC hiện tại
}
function renderOptimizer(){
  const o = DATA.opt;
  if(!o){ $("#opt-stats").innerHTML="Chưa có optimizer.json (cần polygon + chạy build)."; return; }
  const s=o.stats, t=o.thresholds;
  $("#opt-stats").innerHTML=`Khoảng cách TB phường→BC giao <b>${s.avg_d_cur}km</b> → nếu gán gần nhất <b>${s.avg_d_pure}km</b> (giảm ${s.pct_saved}%) · `+
    `<b>${fmt(s.n_reassign)}</b> phường nên đổi BC (cầu ${fmt(s.reassign_dem)}/ngày) · <b>${s.n_close}</b> đóng · <b>${s.n_split}</b> tách · <b>${s.n_far}</b> phường xa >30km. `+
    `Ngưỡng: điểm/m² P75=${t.eHi} P25=${t.eLo} · đồng/điểm P75=${t.dpeP75}.`;
  const CAP=5000;   // render hết (hộp cuộn lo phần dài), không cắt dữ liệu
  const RG=arr=>arr.filter(x=>regOK(x.region));
  const base=[{h:"Vùng",f:r=>r.region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},
    {h:"Phường",f:r=>`${r.name} (${r.ward})`,l:1}];
  // Phường nên đổi: đối chiếu pure vs cap
  $("#opt-reassign").innerHTML=subTable(RG(o.reassign).map(r=>({...r,code:r.from})),[...base,
    {h:"BC hiện tại",f:r=>r.from_name||r.from,l:1},
    {h:"d hiện (km)",f:r=>r.d_cur},
    {h:"→ Gần nhất",f:r=>r.to_pure?`${r.to_pure_name||r.to_pure} · ${r.d_pure}km`:"–",l:1},
    {h:"→ Gần nhất còn tải",f:r=>r.to_cap?`${r.to_cap_name||r.to_cap} · ${r.d_cap}km`:"–",l:1},
    {h:"Cầu/ngày",f:r=>fmt(r.dem)}],CAP,true);
  // Kém hiệu quả (đắt + dư) — hành động theo hạn HĐ
  $("#opt-close").innerHTML=subTable(RG(o.close).map(c=>({...c,code:c.hub})),[
    {h:"Vùng",f:r=>r.region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},
    {h:"BC",f:r=>`${r.name} (${r.hub})`,l:1},
    {h:"Đồng/điểm",f:r=>fmt(r.dpe)},{h:"Điểm/m²",f:r=>r.em2},
    {h:"Hạn HĐ",raw:true,f:r=>r.days==null?"thiếu":(r.days<0?`<span class="warn">đã hết ${-r.days}d</span>`:`${r.days}d`)},
    {h:"Gộp về (gần nhất)",f:r=>`${r.merge_to_name||r.merge_to} · ${r.merge_dist}km`,l:1},
    {h:"Hành động",f:r=>r.action,l:1}],CAP,true);
  // Quá tải → tách / mở rộng
  $("#opt-split").innerHTML=subTable(RG(o.split).map(c=>({...c,code:c.hub})),[
    {h:"Vùng",f:r=>r.region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},
    {h:"BC",f:r=>`${r.name} (${r.hub})`,l:1},
    {h:"Phường",f:r=>r.n_wards},{h:"Điểm/m²",f:r=>r.em2??"–"},{h:"Volume/ngày",f:r=>fmt(r.vol)},
    {h:"Hành động",f:r=>r.action,l:1}],CAP,true);
  // Whitespace × J&T — 4 nhóm
  const w4=o.ws4;
  if(!w4 || !w4.has_jt){
    $("#opt-white").innerHTML=`<div class="sc-note">Cần dữ liệu J&T (competitors_jt.json) để phân 4 nhóm.</div>`;
  } else {
    const wcol=[{h:"Vùng",f:r=>r.region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},{h:"Phường",f:r=>r.name,l:1},
      {h:"Cầu/ngày",f:r=>fmt(r.dem)},{h:"GHN gần nhất",f:r=>r.d_ghn+" km"},{h:"J&T gần nhất",f:r=>r.d_jt+" km"}];
    const tbl=arr=>subTable(RG(arr).map(r=>({...r,code:r.ward})),wcol,CAP,true);
    $("#opt-white").innerHTML=
      `<div class="sc-note">Cầu cao = P75 (top 25% phường — hôm nay ≥${fmt(w4.demP75)} đơn/ngày, tự đổi theo data) · GHN vắng = BC gần nhất ≥${w4.ghn_far}km · J&T có = J&T ≤${w4.jt_near}km.</div>`+
      `<div class="sc-sub"><h3>🔴 Khoảng cách xa <small>(cầu cao · GHN vắng · J&T có → ưu tiên)</small></h3>${tbl(w4.mat_khach)}</div>`+
      `<div class="sc-sub"><h3>🟢 Greenfield <small>(cầu cao · cả 2 vắng → mở chiếm trước)</small></h3>${tbl(w4.greenfield)}</div>`;
  }
  renderOptStaff();
  renderOptScore();
  // click dòng -> bay tới BC liên quan
  document.querySelectorAll("#optimizer .region-row").forEach(tr=>tr.onclick=()=>flyToWardOrHub(tr.dataset.code,true));
}
// 🏅 Bảng điểm hiệu suất — benchmark BC theo ĐIỂM tuyệt đối (đơn-tđ/NV/ngày). Không đụng chỉ số cũ.
function renderOptScore(){
  const box=$("#opt-score"); if(!box) return;
  const sb=scoreBench(), hs=_scoreHubs();
  if(!hs.length || !sb.p50){ box.innerHTML=`<div class="sc-note">Chưa đủ dữ liệu (cần nhân viên + sản lượng).</div>`; return; }
  const rows=hs.map(h=>({h,sc:bcScore(h)})).filter(x=>x.sc!=null).sort((a,b)=>b.sc-a.sc);
  const tag=sc=> sc>=sb.p75?'🔴 quá tải' : (sc<=sb.p25?'🔵 dư' : '—');
  box.innerHTML=
    `<div class="sc-note">${sb.n} BC · trung vị <b>${fmt(sb.p50)}</b> · P25 ${fmt(sb.p25)} · P75 ${fmt(sb.p75)} (điểm/NV). Điểm = (đơn QĐ + kg QĐ÷K)/NV, K=kg TB mỗi đơn toàn quốc; lấy ×0.4.</div>`+
    subTable(rows.map(x=>({...x.h,code:x.h.hub_code,_sc:x.sc})),[
      {h:"Vùng",f:r=>r._region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},{h:"BC",f:r=>`${r.name||r.hub_code} (${r.hub_code})`,l:1},
      {h:"NV",f:r=>fmt(r.staff)},{h:"Đơn/ngày",f:r=>fmt(r.actual.pv+r.actual.dv)},{h:"Kg/ngày",f:r=>fmt(r.actual.pw+r.actual.dw)},
      {h:"Điểm",f:r=>fmt(r._sc)},{h:"Xếp loại",f:r=>tag(r._sc),l:1}],5000,true);
}
// 👷 Định biên nhân sự — tính client-side từ hubs (đã có staff). Năng suất = (pv+dv)/staff.
function renderOptStaff(){
  const box=$("#opt-staff"); if(!box) return;
  const hs=DATA.hubs.filter(h=>h.type==="express"&&h.assigned&&regOK(h._region)&&prodOf(h)!=null);
  const st=prodStats();
  if(!hs.length || !st.p50){ box.innerHTML=`<div class="sc-note">Chưa có dữ liệu nhân viên cho phạm vi này.</div>`; return; }
  const med=st.p50;
  const volOf=h=>h.actual.pv+h.actual.dv;
  const need=h=>Math.max(0, Math.ceil(volOf(h)/med)-h.staff);     // NV cần thêm để kéo về năng suất trung vị
  const surplus=h=>Math.max(0, h.staff-Math.ceil(volOf(h)/med));  // NV dư so với trung vị
  const dist=(a,b)=>{ if(!a.has_geo||!b.has_geo) return Infinity;
    const dx=(a.lng-b.lng)*Math.cos(a.lat*Math.PI/180), dy=a.lat-b.lat; return Math.sqrt(dx*dx+dy*dy)*111; };
  const over=hs.filter(h=>prodOf(h)<=st.p25 && surplus(h)>0);      // thừa người
  const under=hs.filter(h=>prodOf(h)>=st.p75).sort((a,b)=>prodOf(b)-prodOf(a));   // thiếu người
  const overSorted=over.slice().sort((a,b)=>surplus(b)-surplus(a));
  const nearSurplus=h=>{ let best=null,bd=Infinity; over.forEach(o=>{ if(o.hub_code===h.hub_code) return;
    const d=dist(h,o); if(d<bd){bd=d;best=o;} }); return best?`${best.name||best.hub_code} · ${bd.toFixed(0)}km (dư ${surplus(best)})`:"–"; };
  const totNV=hs.reduce((s,h)=>s+h.staff,0), totVol=hs.reduce((s,h)=>s+volOf(h),0);
  box.innerHTML=
    `<div class="sc-note">Phạm vi: <b>${fmt(totNV)}</b> NV / <b>${fmt(totVol)}</b> đơn/ngày · TB <b>${fmt(med)}</b> đơn/NV (P25 ${fmt(st.p25)} · P75 ${fmt(st.p75)}).</div>`+
    `<div class="sc-sub"><h3>🔴 Thiếu người <small>(đơn/NV ≥ P75 — quá tải, cần tuyển/san tải)</small></h3>`+
    subTable(under.map(h=>({...h,code:h.hub_code})),[
      {h:"Vùng",f:r=>r._region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},{h:"BC",f:r=>`${r.name||r.hub_code} (${r.hub_code})`,l:1},
      {h:"NV",f:r=>fmt(r.staff)},{h:"Đơn/ngày",f:r=>fmt(volOf(r))},{h:"Đơn/NV",f:r=>fmt(prodOf(r))},
      {h:"Cần thêm",f:r=>`+${need(r)} NV`,l:1},{h:"Điều từ (BC dư gần nhất)",f:r=>nearSurplus(r),l:1}],5000,true)+`</div>`+
    `<div class="sc-sub"><h3>🔵 Thừa người <small>(đơn/NV ≤ P25 — dư, điều chuyển/cắt giảm)</small></h3>`+
    subTable(overSorted.map(h=>({...h,code:h.hub_code})),[
      {h:"Vùng",f:r=>r._region,l:1},{h:"Tỉnh",f:r=>r.province,l:1},{h:"BC",f:r=>`${r.name||r.hub_code} (${r.hub_code})`,l:1},
      {h:"NV",f:r=>fmt(r.staff)},{h:"Đơn/ngày",f:r=>fmt(volOf(r))},{h:"Đơn/NV",f:r=>fmt(prodOf(r))},
      {h:"Dư",f:r=>`${surplus(r)} NV`,l:1}],5000,true)+`</div>`;
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
      catch(e){ $("#help-content").innerHTML="⚠ Không tải được hướng dẫn."; }
    }
  };
  btn.onclick=open;
  if(close) close.onclick=()=>modalClose(modal);
  modal.onclick=e=>{ if(e.target===modal) modalClose(modal); };  // bấm nền tối để đóng

  // prop-modal: bấm nền tối để đóng (Escape + focus-trap do modalOpen lo)
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
  // ----- Lớp Optimizer -----
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
    if(ui.allowed==="*") opts=[["","Toàn bộ"],...REGION_LIST.map(r=>[r,r])];
    else if(Array.isArray(ui.allowed)&&ui.allowed.length===1) opts=[[ui.allowed[0],ui.allowed[0]]];
    else opts=[["","Toàn bộ (vùng của tôi)"],...(ui.allowed||[]).map(r=>[r,r])];
    rf.innerHTML=opts.map(([v,l])=>`<option value="${v}">${l}</option>`).join("");
    rf.value=ui.regionFilter||"";
    rf.disabled = Array.isArray(ui.allowed)&&ui.allowed.length===1;   // 1 vùng: khoá cứng
    rf.onchange=e=>setRegionFilter(e.target.value);
  }
  $("#planmode").onclick=()=>{ ui.planMode=!ui.planMode; $("#planmode").classList.toggle("on",ui.planMode);
    $("#planmode").textContent=ui.planMode?"■ Tắt chế độ Quy hoạch":"⚿ Bật chế độ Quy hoạch (tô tay)";
    $("#plan-note").textContent = ui.planMode
      ? `Click phường ${ui.planUnit==="new"?"mới":"cũ"} để thêm/bớt vào nhóm; xem bảng Trước/Sau bên phải.`
      : "";
    if(ui.planMode){ ui.selWard=null; ui.selHub=null; }   // bỏ highlight "xem" để không lẫn với đỏ quy hoạch
    draw(); };
  $("#detail-close").onclick=()=>{ $("#detail").classList.add("hidden"); ui.selHub=null; ui.selWard=null; draw(); };
  $("#search").oninput=e=>{ const q=e.target.value.toLowerCase().trim(); const box=$("#search-results");
    if(q.length<2){box.innerHTML="";return;}
    const hits=DATA.hubs.filter(h=>h.has_geo&&((h.name||"").toLowerCase().includes(q)||h.hub_code.includes(q))).slice(0,12);
    box.innerHTML=hits.map(h=>`<div class="sr" data-c="${esc(h.hub_code)}">${esc(h.name||h.hub_code)} <span style="color:#9ca3af">${esc(h.hub_code)}</span></div>`).join("");
    box.querySelectorAll(".sr").forEach(el=>el.onclick=()=>{ const h=hubByCode[el.dataset.c];
      map.flyTo({center:[h.lng,h.lat],zoom:12}); clickHub(h); }); };
}

// ---------- Chat AI (hỏi về vùng/BC/phường/đề xuất) ----------
function esc(s){ return String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
// markdown -> HTML đã lọc XSS (cắt <script>, onerror...); fallback an toàn nếu thiếu lib
function mdSafe(t){ t=t||""; const html=window.marked?marked.parse(t):esc(t);
  return window.DOMPurify?DOMPurify.sanitize(html):esc(t); }
function chatMsg(role, html){
  const d=document.createElement("div"); d.className="msg "+role; d.innerHTML=html;
  $("#chat-msgs").appendChild(d); $("#chat-msgs").scrollTop=1e9; return d;
}
async function chatAsk(q){
  chatMsg("u", esc(q));
  const w=chatMsg("a", "<i>Đang trả lời…</i>");
  try{
    const r=await fetch("/chat",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify({question:q})});
    const d=await r.json();
    w.innerHTML = d.answer ? mdSafe(d.answer) : ("⚠ "+(d.error||"lỗi"));
  }catch(e){ w.innerHTML="⚠ Lỗi kết nối server."; }
  $("#chat-msgs").scrollTop=1e9;
}
function wireChat(){
  $("#chat-scope").textContent = ui.allowed==="*"?"(toàn quốc)":"("+(Array.isArray(ui.allowed)?ui.allowed.join(", "):"")+")";
  $("#chat-fab").onclick=()=>{ $("#chat").classList.remove("hidden"); $("#chat-fab").classList.add("hidden"); $("#chat-q").focus();
    if(!$("#chat-msgs").children.length) chatMsg("a","Chào bạn 👋 Hỏi tôi về <b>vùng / bưu cục / phường / đề xuất</b> trong khu vực bạn phụ trách nhé.<br><small style='color:#6b7280'>VD: \"BC nào sắp hết hạn hợp đồng?\", \"phường nào nên đổi bưu cục?\"</small>"); };
  $("#chat-close").onclick=()=>{ $("#chat").classList.add("hidden"); $("#chat-fab").classList.remove("hidden"); };
  const send=()=>{ const q=$("#chat-q").value.trim(); if(!q) return; $("#chat-q").value=""; chatAsk(q); };
  $("#chat-send").onclick=send;
  $("#chat-q").onkeydown=e=>{ if(e.key==="Enter"&&!e.shiftKey){ e.preventDefault(); send(); } };
}

boot();
