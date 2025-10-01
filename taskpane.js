if (window.__NIS_TASKPANE_JS__) {
  console.debug("NIS: taskpane.js already loaded; skipping re-eval");
} else {
  window.__NIS_TASKPANE_JS__ = true;

// shim to satisfy Office.js loader (avoids noisy console error)
if (typeof window !== 'undefined' && typeof Office !== 'undefined' && !Office.initialize) {
  Office.initialize = function () {};
}

// تأكيد نداء onReady مبكرًا لتفادي رسالة "Office.js has not fully loaded"
try {
  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(() => {});
  }
} catch (_) {}

// ====== (كل كود الملف الحالي يبدأ من هنا كما هو) ======
/* ======== Nano Interactive Slides - taskpane.js (per-slide + linked sequence + nano progress/cancel + caching) ======== */

/* ---------- Keys ---------- */
const NIS_KEY_PREFIX = 'NIS:scene:'; 
function nisKey(k){ return NIS_KEY_PREFIX + k; }

const NIS_STYLE_KEY_PREFIX='NIS:style:'; 
function nmStyleKey(slideKey){ return NIS_STYLE_KEY_PREFIX + (slideKey||'default'); }

const NIS_LINK_KEY_PREFIX='NIS:link:'; // {enabled,next,auto,autoMs,inherit,inheritNm}

const NIS_IMG_CACHE_PREFIX='NIS:img:'; 
function nmCacheKey(s){ return NIS_IMG_CACHE_PREFIX+[s.theme||'',s.prompt||'',String(s.seed||0),s.aspect||'16:9'].join('|'); }

/* ---------- Defaults (Simulation Controls) ---------- */
const NIS_DEFAULT_PARAMS = Object.freeze({
  speed: 50,
  capacity: 100,
  delay: 1,
  projectToSlide: false,
  projectMs: 1000,
  autoStart: false,
  stopOnChange: false
});

const IMG_META_KEY = "NIS_BG_META";
function nisReadAllBgMeta(){ try{return JSON.parse(localStorage.getItem(IMG_META_KEY)||"{}");}catch(e){return {};}}
function nisWriteAllBgMeta(obj){ localStorage.setItem(IMG_META_KEY, JSON.stringify(obj)); }
function nisSetSlideBgMeta(slideId, meta){ const all=nisReadAllBgMeta(); all[slideId]=meta; nisWriteAllBgMeta(all); }
function nisExportBgMeta(){ const data=nisReadAllBgMeta(); const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"}); const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="nis-image-metadata.json"; a.click(); URL.revokeObjectURL(a.href); }
function nisImportBgMetaFromText(text){ const incoming=JSON.parse(text); const current=nisReadAllBgMeta(); const merged={...current,...incoming}; nisWriteAllBgMeta(merged); if(typeof stylesDashRefresh==="function") stylesDashRefresh(); }

/* ---------- Fast in-memory cache ---------- */
const __NIS_STATE_CACHE = new Map();   // slideKey -> params
const __NIS_LINK_CACHE  = new Map();   // slideKey -> link cfg
let   __NIS_ACTIVE_SLIDE_KEY=null;
let   __NIS_GEN_ABORT = null;          // AbortController أثناء توليد الصورة

/* transient flag: when advancing via Linked Sequence, we may inherit on first visit */
let __NIS_INHERIT_NEXT = null;

/* ---------- Debounced settings save (150ms) ---------- */
let __nisSaveTimer=null, __nisSavePending=false;
function nisScheduleSave(delayMs=150){
  if(__nisSaveTimer){ __nisSavePending=true; return; }
  __nisSaveTimer=setTimeout(()=>{
    try{ Office.context.document.settings.saveAsync(()=>{}); }catch(e){}
    __nisSaveTimer=null;
    if(__nisSavePending){ __nisSavePending=false; nisScheduleSave(delayMs); }
  }, delayMs);
}

/* ---------- Persist helpers ---------- */
const NISPersist = {
  saveScene(k,d){ try{ Office.context.document.settings.set(nisKey(k), JSON.stringify(d)); nisScheduleSave(); }catch(e){} },
  loadScene(k){  try{ const r=Office.context.document.settings.get(nisKey(k)); return r?JSON.parse(r):null; }catch(e){ return null; } },
  saveLink(k,d){  try{ Office.context.document.settings.set(NIS_LINK_KEY_PREFIX+k, JSON.stringify(d)); nisScheduleSave(); }catch(e){} },
  loadLink(k){   try{ const r=Office.context.document.settings.get(NIS_LINK_KEY_PREFIX+k); return r?JSON.parse(r):null; }catch(e){ return null; } },
  cacheGet(key){ try{ return Office.context.document.settings.get(key)||null; }catch(e){ return null; } },
  cacheSet(key,v){ try{ Office.context.document.settings.set(key,v); nisScheduleSave(); }catch(e){} }
};

/* ---------- Engine Registry ---------- */
const EngineRegistry=(()=>{ const m=new Map(); 
  return {
    get(k,f){ if(!m.has(k)) m.set(k,f(k)); return m.get(k); },
    dispose(k,d){ if(m.has(k)){ try{ d(m.get(k)); }catch(e){} m.delete(k); } }
  };
})();

/* ---------- Host helpers ---------- */
function q(id){ return document.getElementById(id); }
function hostIsOffice(){ try{ return !!(Office&&Office.context&&Office.context.host); }catch(e){ return false; } }
function hostIsPowerPoint(){ try{ return Office.context.host==='PowerPoint'; }catch(e){ return false; } }

/* Show host hint (restored helper) */
function setHostHint(){
  const h=q('hostHint'); if(!h) return;
  try{
    if(hostIsOffice()) h.textContent='Host: '+Office.context.host+(hostIsPowerPoint()?' (projection enabled)':'');
    else h.textContent='Host: Browser preview';
  }catch(e){ h.textContent='Host: Browser'; }
}

/* ---------- UI params (Simulation Controls) ---------- */
function getUIParams(){
  const s=q('speed'), c=q('capacity'), d=q('delay');
  const pt=q('projectToggle'), pm=q('projectMs');
  const as=q('autoStartToggle'), st=q('stopOnChangeToggle');
  return {
    speed: s?Number(s.value):null,
    capacity: c?Number(c.value):null,
    delay: d?Number(d.value):null,
    projectToSlide: pt?!!pt.checked:false,
    projectMs: pm?Number(pm.value||1000):1000,
    autoStart: as?!!as.checked:false,
    stopOnChange: st?!!st.checked:false
  };
}
function setUIParams(p){
  const s=q('speed'), sv=q('speedVal');
  const c=q('capacity'), cv=q('capacityVal');
  const d=q('delay'), dv=q('delayVal');
  const pt=q('projectToggle'), pm=q('projectMs');
  const as=q('autoStartToggle'), st=q('stopOnChangeToggle');

  if(s && typeof p.speed==='number'){ s.value=String(p.speed); if(sv) sv.textContent=String(p.speed); s.dispatchEvent(new Event('input',{bubbles:true})); }
  if(c && typeof p.capacity==='number'){ c.value=String(p.capacity); if(cv) cv.textContent=String(p.capacity); c.dispatchEvent(new Event('input',{bubbles:true})); }
  if(d && typeof p.delay==='number'){ d.value=String(p.delay); if(dv) dv.textContent=String(p.delay); d.dispatchEvent(new Event('input',{bubbles:true})); }
  if(pt && typeof p.projectToSlide==='boolean'){ pt.checked=p.projectToSlide; }
  if(pm && typeof p.projectMs==='number'){ pm.value=String(p.projectMs); }
  if(as && typeof p.autoStart==='boolean'){ as.checked=p.autoStart; }
  if(st && typeof p.stopOnChange==='boolean'){ st.checked=p.stopOnChange; }

  const e=getActiveEngine();
  if(e){
    e.setSpeed?.(p.speed);
    e.setCapacity?.(p.capacity);
    e.setDelay?.(p.delay);
    e.setProjectToSlide?.(!!p.projectToSlide);
    if(typeof p.projectMs==='number') e.setProjectMs?.(p.projectMs);
  }
}
nmAbortInFlight('slide-change');

/* ---------- Scene persist / restore ---------- */
function persistCurrentSlide(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;
  const p = getUIParams();
  __NIS_STATE_CACHE.set(__NIS_ACTIVE_SLIDE_KEY, {...p});
  NISPersist.saveScene(__NIS_ACTIVE_SLIDE_KEY, p);
}
function restoreCurrentSlide(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;

  const cached = __NIS_STATE_CACHE.get(__NIS_ACTIVE_SLIDE_KEY);
  if(cached) setUIParams(cached);

  const persisted = NISPersist.loadScene(__NIS_ACTIVE_SLIDE_KEY);
  if(persisted){
    setUIParams(persisted);
    __NIS_STATE_CACHE.set(__NIS_ACTIVE_SLIDE_KEY, {...persisted});
    return;
  }

  const defaults={...NIS_DEFAULT_PARAMS};
  setUIParams(defaults);
  __NIS_STATE_CACHE.set(__NIS_ACTIVE_SLIDE_KEY, {...defaults});
  NISPersist.saveScene(__NIS_ACTIVE_SLIDE_KEY, defaults);
}

/* ---------- Nano style helpers (per slide) ---------- */
function nmLoadStyleForSlideKey(k){
  try{ const raw=Office.context.document.settings.get(nmStyleKey(k)); return raw?JSON.parse(raw):null; }catch(e){ return null; }
}
function nmSaveStyleForSlideKey(k,s){
  try{ Office.context.document.settings.set(nmStyleKey(k), JSON.stringify(s)); nisScheduleSave(); }catch(e){}
}

function nmImgMetaKey(slideId){ return `NIS:imgmeta:${slideId}`; }
function nmSaveImageMeta(slideId, meta){
  try { localStorage.setItem(nmImgMetaKey(slideId), JSON.stringify(meta||{})); } catch(e){}
  try { if (typeof nisSetSlideBgMeta === 'function') nisSetSlideBgMeta(String(slideId), meta || {}); } catch(_){}
}
function nmLoadImageMeta(slideId){
  try { return JSON.parse(localStorage.getItem(nmImgMetaKey(slideId)) || '{}'); } catch(e){ return {}; }
}

/* ---------- Fast slide id (race) ---------- */
function captureSlideKeyFast(){
  return new Promise((resolve)=>{
    let done=false; const once=(id)=>{ if(!done){ done=true; resolve(id||'default-slide'); } };

    try{
      if(window.PowerPoint && PowerPoint.run){
        PowerPoint.run(async (ctx)=>{
          const s = ctx.presentation.getSelectedSlides().getItemAt(0);
          s.load("id"); await ctx.sync();
          once(String(s.id||'default-slide'));
        }).catch(()=>{});
      }
    }catch(e){}

    try{
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        r=>{
          if(r.status===Office.AsyncResultStatus.Succeeded &&
             r.value && r.value.slides && r.value.slides[0] && r.value.slides[0].id){
            once(String(r.value.slides[0].id));
          }else{
            once('default-slide');
          }
        }
      );
    }catch(e){}

    setTimeout(()=>once('default-slide'), 250);
  });
}

/* ---------- Slide change wiring ---------- */
let __NIS_LINK_AUTO_TIMER = null;
function linkAutoClear(){
  if(__NIS_LINK_AUTO_TIMER){ clearTimeout(__NIS_LINK_AUTO_TIMER); __NIS_LINK_AUTO_TIMER=null; }
}

function wireSlideChange(){
  try{
    if(!(Office && Office.context && Office.context.document && Office.EventType)) return;

    let lastSlideKey = null;

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async ()=>{
        // persist old slide + clear timers
        persistCurrentSlide();
        linkAutoClear();

        // stop engine if requested
        const prev = getUIParams();
        if(prev.stopOnChange){
          const e = getActiveEngine();
          e?.stop?.();
        }

        const k = await captureSlideKeyFast();
        if(!k) return;
        if(k === lastSlideKey) return;

        lastSlideKey = k;
        __NIS_ACTIVE_SLIDE_KEY = k;

        // --- INHERIT (first visit) if requested ---
        if(__NIS_INHERIT_NEXT && __NIS_INHERIT_NEXT.from){
          const fromKey = __NIS_INHERIT_NEXT.from;
          __NIS_INHERIT_NEXT = null; // consume flag

          const hasScene = NISPersist.loadScene(k);
          const linkCfg  = __NIS_LINK_CACHE.get(fromKey) || linkLoad(fromKey);

          if(!hasScene && linkCfg?.inherit){
            const src = __NIS_STATE_CACHE.get(fromKey) || NISPersist.loadScene(fromKey) || {...NIS_DEFAULT_PARAMS};
            NISPersist.saveScene(k, src);
            __NIS_STATE_CACHE.set(k, {...src});
          }

          if(linkCfg?.inheritNm){
            const tgtStyle = nmLoadStyleForSlideKey(k);
            if(!tgtStyle){
              const srcStyle = nmLoadStyleForSlideKey(fromKey);
              if(srcStyle) nmSaveStyleForSlideKey(k, srcStyle);
            }
          }
        }
        // -------------------------------------------

        getActiveEngine();
        restoreCurrentSlide();
        nmInit();
        linkRestoreForSlide();   // refresh Linked Sequence UI
        stylesDashRefresh();

        const cur = getUIParams();
        if(cur.autoStart){
          const e = getActiveEngine();
          e?.start?.();
          linkAutoArmForActive(); // arm auto-advance if enabled
        }
      }
    );
  }catch(e){}
}

/* ---------- Simple internal preview engine ---------- */
function drawPreview(ctx,state){
  const w=ctx.canvas.width,h=ctx.canvas.height;
  ctx.clearRect(0,0,w,h);
  ctx.fillStyle='#f9fafb'; ctx.fillRect(0,0,w,h);
  ctx.fillStyle='#e5e7eb'; for(let i=0;i<10;i++){ ctx.fillRect(i*(w/10),0,1,h); }
  ctx.fillStyle='#111827'; ctx.fillRect(32,h-40,Math.max(20,Math.min(w-64,state.capacity)),12);
  const y=h/2; ctx.beginPath(); ctx.arc(state.x,y,12,0,Math.PI*2); ctx.fillStyle='#2563eb'; ctx.fill();
  ctx.font='14px system-ui,Segoe UI,Arial'; ctx.fillStyle='#374151';
  ctx.fillText('spd '+state.speed+'  cap '+state.capacity+'  dly '+state.delay,32,28);
}

function createInternalEngine(slideKey){
  // ensure we have a <canvas id="preview">
  let host = q('preview'); 
  let canvas = host;
  if(!canvas || typeof canvas.getContext !== 'function'){
    const c = document.createElement('canvas');
    c.id='preview'; c.width=480; c.height=220;
    if(host){ host.innerHTML=''; host.appendChild(c); }
    canvas = c;
  }
  const ctx = canvas.getContext('2d');

  let running=false, tm=null, lastProject=0;
  let state={ ...NIS_DEFAULT_PARAMS, x:40 };

  const step=()=>{
    if(!running) return;
    const v=Math.max(1,Math.floor((state.speed||50)/3));
    state.x+=v; if(state.x>canvas.width-40) state.x=40;

    if(__NIS_ACTIVE_SLIDE_KEY===slideKey){
      drawPreview(ctx,state);
      const now=Date.now();
      if(state.projectToSlide && hostIsPowerPoint() && now-lastProject>=state.projectMs){
        projectCanvas(canvas); lastProject=now;
      }
    }
    const tickMs=Math.max(5,200-(state.speed||50)*1.5)+(state.delay||0)*100;
    tm=setTimeout(step,tickMs);
  };

  return {
    start(){ if(running) return; running=true; step(); },
    stop(){ running=false; if(tm){ clearTimeout(tm); tm=null; } },
    setSpeed(v){ state.speed=v; },
    setCapacity(v){ state.capacity=v; },
    setDelay(v){ state.delay=v; },
    setProjectToSlide(v){ state.projectToSlide=!!v; },
    setProjectMs(v){ state.projectMs=Number(v)||1000; },
    reset(){ state={ ...NIS_DEFAULT_PARAMS, x:40 }; 
             if(__NIS_ACTIVE_SLIDE_KEY===slideKey) drawPreview(ctx,state); },
    snapshot(){ if(__NIS_ACTIVE_SLIDE_KEY===slideKey){ drawPreview(ctx,state); if(hostIsPowerPoint()) projectCanvas(canvas); } },
    download(){ downloadPNG(); },
    export(){ exportJSON(); },
    import(file){ importJSON(file); }
  };
}

function createEngineForSlide(slideKey){
  if(typeof window.NIS_createEngine==='function') return window.NIS_createEngine(slideKey);
  if(typeof window.createEngine==='function') return window.createEngine(slideKey);
  return createInternalEngine(slideKey);
}
function getActiveEngine(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return null;
  return EngineRegistry.get(__NIS_ACTIVE_SLIDE_KEY, createEngineForSlide);
}

/* ---------- Simulation UI bindings ---------- */
function bindSimUI(){
  const startBtn=q('start'), stopBtn=q('stop'), resetBtn=q('reset');
  const snapBtn=q('snapshot'), expBtn=q('exportJson');
  const impBtn=q('importJson'), impFile=q('jsonFile'), pngBtn=q('downloadPng');
  const s=q('speed'), sv=q('speedVal');
  const c=q('capacity'), cv=q('capacityVal');
  const d=q('delay'), dv=q('delayVal');
  const pt=q('projectToggle'), pm=q('projectMs');

  if(stopBtn){  stopBtn .addEventListener('click',()=>{ const e=getActiveEngine(); e?.stop?.();  linkAutoClear(); }); }
  if(resetBtn){ resetBtn.addEventListener('click',()=>{ setUIParams({...NIS_DEFAULT_PARAMS}); const e=getActiveEngine(); e?.reset?.(); persistCurrentSlide(); }); }

  if(s){ s.addEventListener('input',()=>{ const v=+s.value; if(sv) sv.textContent=String(v); const e=getActiveEngine(); e?.setSpeed?.(v); });
        s.addEventListener('change',()=>{ persistCurrentSlide(); }); }
  if(c){ c.addEventListener('input',()=>{ const v=+c.value; if(cv) cv.textContent=String(v); const e=getActiveEngine(); e?.setCapacity?.(v); });
        c.addEventListener('change',()=>{ persistCurrentSlide(); }); }
  if(d){ d.addEventListener('input',()=>{ const v=+d.value; if(dv) dv.textContent=String(v); const e=getActiveEngine(); e?.setDelay?.(v); });
        d.addEventListener('change',()=>{ persistCurrentSlide(); }); }

  if(pt){ pt.addEventListener('change',()=>{ const v=!!pt.checked; const e=getActiveEngine(); e?.setProjectToSlide?.(v); persistCurrentSlide(); }); }
  if(pm){ pm.addEventListener('input',()=>{ const v=Number(pm.value)||1000; const e=getActiveEngine(); e?.setProjectMs?.(v); });
        pm.addEventListener('change',()=>{ persistCurrentSlide(); }); }

  if(startBtn){
    startBtn.addEventListener('click', ()=>{
      const p = getUIParams();
      const e = getActiveEngine();
      if(e && e.setProjectToSlide) e.setProjectToSlide(!!p.projectToSlide);
      if(e && e.setProjectMs)       e.setProjectMs(Number(p.projectMs||1000));
      e?.start?.();
      linkAutoArmForActive();
    });
  }

  if(snapBtn){ snapBtn.addEventListener('click',()=>{ const e=getActiveEngine(); e?.snapshot?.(); }); }
  if(expBtn){  expBtn .addEventListener('click',()=>{ const e=getActiveEngine(); e?.export?.();  }); }
  if(impBtn){  impBtn .addEventListener('click',()=>{ if(impFile) impFile.click(); }); }
  if(impFile){ impFile.addEventListener('change',ev=>{ const f=ev.target.files&&ev.target.files[0]; if(f){ const e=getActiveEngine(); e?.import?.(f); impFile.value=''; } }); }
  if(pngBtn){  pngBtn .addEventListener('click',()=>{ const e=getActiveEngine(); e?.download?.(); }); }
}

/* ---------- Project/export helpers ---------- */
function projectCanvas(canvas){
  try{
    const dataUrl=canvas.toDataURL('image/png');
    const base64=dataUrl.split(',')[1];
    Office.context.document.setSelectedDataAsync(base64,{coercionType:Office.CoercionType.Image},()=>{});
  }catch(e){}
}
function exportJSON(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;
  const payload={ slideKey:__NIS_ACTIVE_SLIDE_KEY, params:getUIParams(), ts:Date.now() };
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='nis-scene-'+__NIS_ACTIVE_SLIDE_KEY+'.json'; a.click(); URL.revokeObjectURL(a.href);
}
function importJSON(file){
  const r=new FileReader();
  r.onload=()=>{ 
    try{
      const data = JSON.parse(r.result);
      const p = (data && data.params) ? data.params : data;
      const e = getActiveEngine();
      if (e && e.reset) e.reset();           // أولًا: رجّع الحالة الافتراضية
      setUIParams(p);                        // ثانيًا: طبّق الإعدادات المستوردة على الـ UI والموتور
      persistCurrentSlide();                 // وأخيرًا خزّنها
    }catch(_){}
  };
  r.readAsText(file);
}
function downloadPNG(){
  let canvas=q('preview');
  if(canvas && typeof canvas.toDataURL==='function'){
    const a=document.createElement('a'); a.href=canvas.toDataURL('image/png');
    a.download='nis-slide-'+(__NIS_ACTIVE_SLIDE_KEY||'default')+'.png'; a.click();
  }
}

/* ---------- Nano Mode (with Progress/Cancel/Cache) ---------- */
function nmGetStyle(){
  try{
    const key=nmStyleKey(__NIS_ACTIVE_SLIDE_KEY);
    const raw=Office.context.document.settings.get(key);
    return raw ? JSON.parse(raw) : {theme:'',seed:42,prompt:'',aspect:'16:9',caption:false,autoInc:true};
  }catch(e){ return {theme:'',seed:42,prompt:'',aspect:'16:9',caption:false,autoInc:true}; }
}
function nmSaveStyle(s){
  try{
    const key=nmStyleKey(__NIS_ACTIVE_SLIDE_KEY);
    Office.context.document.settings.set(key, JSON.stringify(s));
    nisScheduleSave();
  }catch(e){}
}
function nmReadInputs(){
  return {
    theme: q('nmTheme')?.value||'',
    seed:  q('nmSeed') ? Number(q('nmSeed').value||0) : 0,
    prompt:q('nmPrompt')?.value||'',
    aspect:q('nmAspect')?.value||'16:9',
    caption:q('nmCaption')?q('nmCaption').checked:false,
    autoInc:q('nmAutoInc')?q('nmAutoInc').checked:true
  };
}

// ---- Nano Style Preview Chip ----
function nmUpdateStyleChip(style){
  try{
    const $aspect = document.getElementById('nmChipAspect');
    const $seed   = document.getElementById('nmChipSeed');
    const $theme  = document.getElementById('nmChipTheme');
    const $prompt = document.getElementById('nmChipPrompt');
    const aspect  = style?.aspect || '16:9';
    const seed    = (typeof style?.seed==='number' ? style.seed : Number(style?.seed||0)) || 0;
    const theme   = (style?.theme||'').trim();
    const prompt  = (style?.prompt||'').trim();

    if($aspect) $aspect.textContent = `Aspect ${aspect}`;
    if($seed)   $seed.textContent   = `• Seed #${seed}`;
    if($theme)  $theme.textContent  = theme ? `• ${theme}` : '';
    if($prompt) $prompt.textContent = prompt ? `• ${prompt}` : '';

    // لون بسيط حسب الـ aspect
    const chip = document.getElementById('nmStyleChip');
    if(chip){
      let bg = '#fafafa', bd = '#e5e7eb';
      if(aspect==='1:1'){ bg='#f8fafc'; bd='#e2e8f0'; }
      else if(aspect==='4:3'){ bg='#f9fafb'; bd='#e5e7eb'; }
      else { bg='#fff7ed'; bd='#fed7aa'; } // 16:9
      chip.style.background = bg;
      chip.style.borderColor = bd;
    }
  }catch(e){}
}

// ---- Styles Dashboard helpers ----
function nmMakeMiniChip(style){
  const aspect = style?.aspect || '16:9';
  const seed   = (typeof style?.seed==='number' ? style.seed : Number(style?.seed||0)) || 0;
  const theme  = (style?.theme||'').trim();
  const prompt = (style?.prompt||'').trim();

  const wrap = document.createElement('div');
  wrap.style.cssText = "display:flex;gap:.4rem;align-items:center;border:1px solid #e5e7eb;border-radius:.4rem;padding:.2rem .4rem;background:#fafafa;font:12px system-ui,Segoe UI,Arial;color:#374151;";

  const elA = document.createElement('span'); elA.textContent = aspect; elA.style.fontWeight = "600";
  const elS = document.createElement('span'); elS.textContent = `#${seed}`; elS.style.opacity=".8";
  const elT = document.createElement('span'); elT.textContent = theme;  elT.style.maxWidth="160px"; elT.style.whiteSpace="nowrap"; elT.style.overflow="hidden"; elT.style.textOverflow="ellipsis";
  const elP = document.createElement('span'); elP.textContent = prompt; elP.style.maxWidth="220px"; elP.style.whiteSpace="nowrap"; elP.style.overflow="hidden"; elP.style.textOverflow="ellipsis";

  wrap.appendChild(elA); wrap.appendChild(elS);
  if(theme) wrap.appendChild(elT);
  if(prompt) wrap.appendChild(elP);
  return wrap;
}

// ---- Overlay label (corner tag) ----
const NIS_OVERLAY_NAME = 'NIS_OVERLAY';

function ovBuildBase64(style, opacityPct){
  // نص الشارة
  const aspect = style?.aspect || '16:9';
  const seed   = (typeof style?.seed==='number' ? style.seed : Number(style?.seed||0)) || 0;
  const theme  = (style?.theme||'').trim();

  // كانفاس شفاف
  const W = 460, H = 64, R = 12;
  const c = document.createElement('canvas');
  c.width = W; c.height = H;
  const g = c.getContext('2d');

  // خلفية نصف شفافة + حد خفيف
  const alpha = Math.max(0.2, Math.min(1, (opacityPct||80)/100));
  g.fillStyle = `rgba(255,255,255,${alpha})`;
  g.strokeStyle = 'rgba(0,0,0,0.12)';
  g.lineWidth = 1;

  // rounded rect
  g.beginPath();
  g.moveTo(R,0); g.lineTo(W-R,0); g.quadraticCurveTo(W,0,W,R);
  g.lineTo(W,H-R); g.quadraticCurveTo(W,H,W-R,H);
  g.lineTo(R,H); g.quadraticCurveTo(0,H,0,H-R);
  g.lineTo(0,R); g.quadraticCurveTo(0,0,R,0); g.closePath();
  g.fill(); g.stroke();

  // نص
  g.font = '600 15px Segoe UI, system-ui, Arial';
  g.fillStyle = '#111';
  g.textBaseline = 'middle';
  const parts = [`${aspect}`, `#${seed}`, theme ? theme : ''];
  const txt = parts.filter(Boolean).join('  •  ');
  g.fillText(txt, 16, H/2);

  // حوّل Base64 بدون prefix
  const dataUrl = c.toDataURL('image/png');
  return dataUrl.split(',')[1]; // remove "data:image/png;base64,"
}

async function ovRemoveCurrent(ctx){
  const sel = ctx.presentation.getSelectedSlides();
  sel.load("items");
  await ctx.sync();
  const slide = sel.items?.[0];
  if(!slide) return false;
  const shapes = slide.shapes;
  shapes.load("items/name");
  await ctx.sync();
  let removed = false;
  for(const sh of shapes.items){
    if((sh.name||'') === NIS_OVERLAY_NAME){
      sh.delete();
      removed = true;
    }
  }
  if(removed) await ctx.sync();
  return removed;
}

async function ovApplyToCurrent(opts){
  const style   = nmReadInputs();
  const corner  = (opts?.corner||'tr');
  const opacity = Number(opts?.opacity||80);

  const b64 = ovBuildBase64(style, opacity);

  // 1) امسح أي شارة قديمة أولًا داخل run مستقل
  if (window.PowerPoint && PowerPoint.run){
    try{
      await PowerPoint.run(async (ctx)=>{ await ovRemoveCurrent(ctx); });
    }catch(_) {}
  }

  // 2) أدرج الصورة بضمان (نفس الطريقة المستخدمة للخلفيات)
  await new Promise((res)=> {
    Office.context.document.setSelectedDataAsync(
      b64,
      { coercionType: Office.CoercionType.Image },
      ()=>res()
    );
  });

  // 3) لقّط آخر شيب (المضاف حالًا) واضبط حجمه/مكانه واسمه
  await PowerPoint.run(async (ctx)=>{
    const pres  = ctx.presentation;
    const slide = pres.getSelectedSlides().getItemAt(0);
    slide.load("id");
    const shapes = slide.shapes;
    shapes.load("items");
    await ctx.sync();

    const count = shapes.items.length;
    if(count === 0) return;

    let img = shapes.items[count-1];
    try { img.load(["width","height","left","top","name"]); } catch(_) {}
    await ctx.sync();

    // اسم ثابت للـ overlay
    try { img.name = NIS_OVERLAY_NAME; } catch(_) {}

    // حجم ومكان
    const margin = 12;
    const w = 460, h = 64; // نفس أبعاد الكانفاس
    // مبدئيًا أعلى يسار
    try{ img.left = margin; img.top = margin; img.width = w; img.height = h; }catch(_){}

    // موضع الركن — بدون getSlideSize (بنفس افتراض 960×540 pt)
    const sw = 960, sh = 540;
    try{
      if (corner === 'tr') {
        img.left = Math.max(margin, sw - w - margin); 
        img.top  = margin;
      } else if (corner === 'br') {
        img.left = Math.max(margin, sw - w - margin); 
        img.top  = Math.max(margin, sh - h - margin);
      } else if (corner === 'bl') {
        img.left = margin; 
        img.top  = Math.max(margin, sh - h - margin);
      }
      // 'tl' يظل (margin, margin)
    }catch(_){}

    await ctx.sync();
  });
}

async function stylesDashRefresh(){
  const list = document.getElementById('stylesList');
  if(!list) return;
  list.innerHTML = '';

  if(!(window.PowerPoint && PowerPoint.run)){
    const div = document.createElement('div');
    div.className = 'muted';
    div.textContent = '(Styles Dashboard needs PowerPoint host)';
    list.appendChild(div);
    return;
  }

  await PowerPoint.run(async (ctx)=>{
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    slides.items.forEach(s=>s.load(["id","index","title"]));
    await ctx.sync();

    for(const sl of slides.items){
      const id    = String(sl.id);
      const index = (Number(sl.index) || 0) + 1;
      const title = (sl.title || '').trim();

      const row = document.createElement('div');
      row.style.cssText="display:grid;grid-template-columns:110px 1fr auto auto auto;gap:8px;align-items:center;border:1px solid #eee;border-radius:8px;padding:8px";

      const left = document.createElement('div');
      left.innerHTML = `<div style="font-weight:600">Slide ${index}</div><div class="muted" style="font-size:11px">${title||id}</div>`;

      const style = nmLoadStyleForSlideKey(id) || { theme:'', seed:42, prompt:'', aspect:'16:9' };
      const chip = nmMakeMiniChip(style);
      const meta = nmLoadImageMeta(id);
if(meta && meta.at){
  chip.title = `Provider: ${meta.provider||'-'} | Seed: ${meta.seed} | Aspect: ${meta.aspect} | ${meta.at}`;
}


      const btnApply = document.createElement('button'); btnApply.textContent = 'Apply Background';
      const btnCopy  = document.createElement('button'); btnCopy.textContent  = 'Copy from Current';
      const btnRegen = document.createElement('button'); btnRegen.textContent = 'Regenerate';

      btnApply.addEventListener('click', async ()=>{
        const sendToBack = !!(q('nmBgLock') && q('nmBgLock').checked);
        // جرّب كاش أولًا
        let b64 = nmFindCachedForStyle ? nmFindCachedForStyle(style) : null;
        if(!b64){ b64 = await nmGenerate(style); if(!b64) return; }
        // بدّل التحديد للسلايد الهدف وطبّق
        await PowerPoint.run(async (ctx2)=>{
          ctx2.presentation.setSelectedSlides([id]); await ctx2.sync();
        });
        await nmApplyAsBackground(b64, { sendToBack });
      });

      btnCopy.addEventListener('click', async ()=>{
        // اقرأ الستايل الحالي من Inputs واحفظه على السلايد الهدف
        const cur = nmReadInputs();
        nmSaveStyleForSlideKey(id, cur);
        // حدّث الـ chip
        row.children[1].innerHTML = ''; row.children[1].appendChild(nmMakeMiniChip(cur));
      });

      btnRegen.addEventListener('click', async ()=>{
        let s = nmLoadStyleForSlideKey(id) || nmReadInputs();
        if(s.autoInc!==false){ s.seed = (Number(s.seed)||0)+1; nmSaveStyleForSlideKey(id, s); }
        const sendToBack = !!(q('nmBgLock') && q('nmBgLock').checked);
        const b64 = await nmGenerate(s); if(!b64) return;
        await PowerPoint.run(async (ctx2)=>{
          ctx2.presentation.setSelectedSlides([id]); await ctx2.sync();
        });
        await nmApplyAsBackground(b64, { sendToBack });
        // refresh chip
        row.children[1].innerHTML = ''; row.children[1].appendChild(nmMakeMiniChip(s));
      });

      row.appendChild(left);
      row.appendChild(chip);
      row.appendChild(btnApply);
      row.appendChild(btnCopy);
      row.appendChild(btnRegen);

      list.appendChild(row);
    }
  });
}

async function sdRegenerateApplyForSlides(style, ids, useCacheFirst){
  if(!ids || ids.length===0) return 0;
  let done = 0;
  // progress UI
  if(typeof nmShowBusy==='function') nmShowBusy(true);
  try{
    for(let i=0;i<ids.length;i++){
      const id = String(ids[i]);
      // بدّل التحديد للسلايد الهدف
      await PowerPoint.run(async (ctx)=>{
        ctx.presentation.setSelectedSlides([id]); await ctx.sync();
      });

      // جرّب الكاش أولًا
      let b64 = null;
      if(useCacheFirst && typeof nmFindCachedForStyle==='function'){
        b64 = nmFindCachedForStyle(style);
      }
      // لو مفيش كاش → ولّد
      if(!b64){
        const s = { ...style };
        if(s.autoInc!==false){ s.seed = (Number(s.seed)||0) + 1; }
        nmSaveStyleForSlideKey(id, s);
        b64 = await nmGenerate(s);
      }
      if(b64){
        await nmApplyAsBackground(b64, { sendToBack: !!(q('nmBgLock') && q('nmBgLock').checked) });
        done++;
      }
      if(typeof nmSetProgress==='function'){
        const p = Math.round(((i+1)/ids.length)*100);
        nmSetProgress(p, `Applied ${i+1}/${ids.length}`);
      }
    }
  } finally {
    if(typeof nmShowBusy==='function') nmShowBusy(false);
  }
  return done;
}

function stylesDashInit(){
  // نضمن التشغيل بعد Office.onReady (أو DOM fallback للمعاينة في المتصفح)
  const boot = async ()=>{
    await stylesDashRefresh();

    const btn     = document.getElementById('sdRegenApplySelected');
    const hint    = document.getElementById('sdHint');
    const useCache= document.getElementById('sdUseCacheFirst');

    const imgExport = document.getElementById('imgMetaExport');
    const imgImport = document.getElementById('imgMetaImport');
    const imgFile   = document.getElementById('imgMetaImportFile');

    if (imgExport) imgExport.addEventListener('click', ()=>{ nisExportBgMeta(); });

    if (imgImport) imgImport.addEventListener('click', ()=>{ if(imgFile) imgFile.click(); });

    if (imgFile) {
      imgFile.addEventListener('change', (e)=>{
        const f = e.target.files && e.target.files[0];
        if(!f) return;
        const r = new FileReader();
        r.onload = ()=> { nisImportBgMetaFromText(r.result); };
        r.readAsText(f);
        e.target.value="";
      });
    }

    if (btn){
      btn.addEventListener('click', async ()=>{
        try{
          const s = nmReadInputs();
          nmSaveStyle(s);

          const ids = await ppGetSelectedSlideIds();
          if(!ids || ids.length===0){
            if(hint) hint.textContent = 'No selected slides.';
            return;
          }

          if(hint) hint.textContent = 'Working...';
          const n = await sdRegenerateApplyForSlides(s, ids, !!(useCache && useCache.checked));
          if(hint) hint.textContent = `Done: ${n} slide(s).`;

          // حدث قائمة الشرائح والـ chips بعد التنفيذ
          stylesDashRefresh();
        }catch(e){
          console.error(e);
          if(hint) hint.textContent = 'Failed.';
        }
      });
    }
  };

  if (window.Office && Office.onReady) {
    Office.onReady().then(boot).catch(()=>{ /* ignore */ });
  } else {
    document.addEventListener('DOMContentLoaded', boot);
  }
}

// throttle لتقليل كثرة التحديثات مع الكتابة
let __nmChipTs = 0;
function nmUpdateStyleChipThrottled(style){
  const now = (typeof performance!=='undefined' && performance.now) ? performance.now() : Date.now();
  if(now - __nmChipTs < 60) return;
  __nmChipTs = now;
  nmUpdateStyleChip(style);
}

function nmWriteInputs(s){
  if(q('nmTheme')) q('nmTheme').value = s.theme||'';
  if(q('nmSeed'))  q('nmSeed').value  = String(typeof s.seed==='number'?s.seed:42);
  if(q('nmPrompt'))q('nmPrompt').value= s.prompt||'';
  if(q('nmAspect'))q('nmAspect').value= s.aspect||'16:9';
  if(q('nmCaption'))q('nmCaption').checked=!!s.caption;
  if(q('nmAutoInc'))q('nmAutoInc').checked=(s.autoInc!==false);
  nmUpdateStyleChipThrottled(s);
}
function nmHash(str){ let h=2166136261>>>0; for(let i=0;i<str.length;i++){ h^=str.charCodeAt(i); h=Math.imul(h,16777619); } return h>>>0; }
function nmSize(aspect){
  if (aspect === '4:3') return { w: 1024, h: 768 };   // 0.79 MP
  if (aspect === '1:1') return { w: 1024, h: 1024 };  // 1.05 MP
  return { w: 1280, h: 720 };                         // 0.92 MP (آمن تحت 2MP)
}

// ---- Abort helper: cancel any in-flight generation safely ----
function nmAbortInFlight(reason='user-cancel'){
  try{
    if(__NIS_GEN_ABORT){ __NIS_GEN_ABORT.abort(reason); __NIS_GEN_ABORT=null; }
  }catch(e){}
  nmShowBusy(false);
  const h=q('nmHint');
  if(h){
    if(reason==='slide-change') h.textContent='Canceled (slide changed).';
    else if(reason==='preempt') h.textContent='Canceled (new request started).';
    else h.textContent='Canceled.';
  }
}

// ---- Throttled progress to reduce UI jank ----
let __nmProgTs = 0;
function nmSetProgressThrottled(pct,msg){
  const now = (typeof performance!=='undefined' && performance.now) ? performance.now() : Date.now();
  if(now - __nmProgTs < 50) return;  // ~20fps
  __nmProgTs = now;
  nmSetProgress(pct,msg);
}

// ---- Cache helpers ----
function nmFindCachedForStyle(style){
  try{
    const key = nmCacheKey(style);
    const b64 = NISPersist.cacheGet(key);
    return b64 || null;
  }catch(e){ return null; }
}
function nmShowCachedButton(style){
  const btn = q('nmApplyCached'), hint=q('nmHint');
  if(!btn) return;
  const has = !!nmFindCachedForStyle(style);
  btn.style.display = has ? 'inline-block' : 'none';
  if(hint && has) hint.textContent = 'Cached image available for this style.';
}

/* Placeholder generator (fallback) */
function nmGeneratePNG(style){
  const sz=nmSize(style.aspect||'16:9'); const w=sz.w,h=sz.h;
  const c=document.createElement('canvas'); c.width=w; c.height=h; const ctx=c.getContext('2d');
  const seed=nmHash((style.theme||'')+'|'+(style.prompt||'')+'|'+String(style.seed||0));
  const a1=(seed%360), a2=((seed>>3)%360);
  const g=ctx.createLinearGradient(0,0,w,h); g.addColorStop(0,'hsl('+a1+' 70% 60%)'); g.addColorStop(1,'hsl('+a2+' 70% 40%)');
  ctx.fillStyle=g; ctx.fillRect(0,0,w,h);
  const motif=(seed%3); ctx.save(); ctx.globalAlpha=0.18;
  if(motif===0){ for(let i=0;i<10;i++){ const r=((seed>>i)&255)/255; const x=r*w,y=(1-r)*h; ctx.beginPath(); ctx.arc(x,y,80*(0.3+r),0,Math.PI*2); ctx.fillStyle='#fff'; ctx.fill(); } }
  else if(motif===1){ for(let i=0;i<14;i++){ const r=((seed>>i)&255)/255; ctx.fillStyle='#fff'; ctx.fillRect(r*w,0,6,h); } }
  else{ ctx.translate(w/2,h/2); for(let i=0;i<8;i++){ ctx.rotate(((seed>>i)&7)*0.15); ctx.fillStyle='#fff'; ctx.fillRect(0,0,w*0.35,3); } }
  ctx.restore();
  if(style.caption){ ctx.fillStyle='rgba(0,0,0,0.6)'; ctx.fillRect(0,h-64,w,64);
    ctx.fillStyle='#fff'; ctx.font='bold 22px system-ui,Segoe UI,Arial';
    const text=(style.theme||'')+' | '+(style.prompt||'')+' | #'+String(style.seed||0);
    ctx.fillText(text,24,h-24);
  }
  return c.toDataURL('image/png').split(',')[1];
}

/* UI helpers (busy/progress/cancel) */
function nmShowBusy(on){
  const busy=q('nmBusy'), prog=q('nmProg'), btnCancel=q('nmCancel');
  ['nmStyleSelected','nmRegenSelected','nmSave'].forEach(id=>{ const b=q(id); if(b) b.disabled=on; });
  if(busy) busy.style.display=on?'block':'none';
  if(prog){ prog.style.display=on?'inline-block':'none'; if(!on){ prog.value=0; } }
  if(btnCancel) btnCancel.style.display=on?'inline-block':'none';
}
function nmSetProgress(pct,msg){
  const prog=q('nmProg'), hint=q('nmHint'), busy=q('nmBusy');
  if(prog && typeof pct==='number'){ prog.value=Math.max(0,Math.min(100,Math.floor(pct))); }
  if(busy) busy.textContent = msg ? ('Generating… '+msg) : 'Generating…';
  if(hint && msg) hint.textContent = msg;
}
function nmApplyToSelection(base64){
  try{ Office.context.document.setSelectedDataAsync(base64,{coercionType:Office.CoercionType.Image},()=>{}); }catch(e){}
}

// ---- Get selected slide IDs ----
async function ppGetSelectedSlideIds(){
  if(!(window.PowerPoint && PowerPoint.run)) return null;
  try{
    let ids = [];
    await PowerPoint.run(async (ctx)=>{
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items");
      await ctx.sync();
      ids = (sel.items||[]).map(s=>String(s.id));
    });
    return ids;
  }catch(_){ return null; }
}

// ---- Get ALL slide IDs ----
async function ppGetAllSlideIds(){
  if(!(window.PowerPoint && PowerPoint.run)) return null;
  try{
    let ids = [];
    await PowerPoint.run(async (ctx)=>{
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      slides.items.forEach(s=>s.load("id"));
      await ctx.sync();
      ids = slides.items.map(s=>String(s.id));
    });
    return ids;
  }catch(_){ return null; }
}

// ---- Copy given style to a list of slide IDs (no generation) ----
async function nmCopyStyleToSlides(style, slideIds){
  if(!Array.isArray(slideIds) || slideIds.length===0) return 0;
  let n = 0;
  for(const id of slideIds){
    try{ nmSaveStyleForSlideKey(id, style); n++; }catch(_){}
  }
  return n;
}

// ---- One-shot generate (used by retry wrapper) ----
async function nmGenerateOnce(style, onProgress, ctrl, timeoutMs=45000){
  const provider = typeof window.NIS_generateImage==='function' ? window.NIS_generateImage : null;
  const timeout = new Promise((_,rej)=>setTimeout(()=>rej(new Error('timeout')), timeoutMs));

  if(!provider) return { base64: nmGeneratePNG(style), via: 'fallback' };

  const result = await Promise.race([
    provider(style, onProgress, ctrl.signal),
    timeout
  ]);
  if (typeof result === 'string') {
    // ممكن يرجّع base64 خام أو data URL
    if (result.startsWith('data:image/')) {
      return { base64: result.split(',')[1], via: 'provider' };
    }
    return { base64: result, via: 'provider' };
  }

  if (result && typeof result.base64 === 'string') {
    const b64 = result.base64.startsWith('data:image/')
      ? result.base64.split(',')[1]
      : result.base64;
    return { base64: b64, via: result.via || 'provider' };
  }
}
// ---- Apply generated image as slide background ----
// ---- Apply generated image as slide background (fit & optional send-to-back) ----
async function nmApplyAsBackground(base64, opts={ sendToBack: false }){
  const hint = q('nmHint');

  // 1) guaranteed insert
  try{
    await new Promise((res)=> {
      Office.context.document.setSelectedDataAsync(
        base64,
        { coercionType: Office.CoercionType.Image },
        ()=>res()
      );
    });
  }catch(e){
    if(hint) hint.textContent = 'Inserted image (fallback).';
    return;
  }

  // 2) try to resize & z-order with PowerPoint API
  try{
    if(!(window.PowerPoint && PowerPoint.run)) {
      if(hint) hint.textContent = 'Applied (basic insert).';
      return;
    }

    await PowerPoint.run(async (ctx)=>{
      const pres  = ctx.presentation;
      const slide = pres.getSelectedSlides().getItemAt(0);
      slide.load("id");
      const shapes = slide.shapes;
      shapes.load("items");
      await ctx.sync();

      const count = shapes.items.length;
      if(count === 0) return;

      // heuristics: pick the last shape (just inserted), but confirm it's a picture if possible
      let pic = shapes.items[count-1];
      try { pic.load(["type","width","height","left","top"]); } catch(_) {}
      await ctx.sync();

      try{
        const style = (typeof nmReadInputs === 'function') ? nmReadInputs() : {};
        if (window.PowerPoint && PowerPoint.run){
          await PowerPoint.run(async (ctx2)=>{
            const sel = ctx2.presentation.getSelectedSlides();
            sel.load("items");
            await ctx2.sync();
            const curId = sel.items?.[0]?.id;
            if(curId){
              nmSaveImageMeta(String(curId), {
                provider: style?.provider || 'mock/pixabay',
                seed: style?.seed,
                aspect: style?.aspect,
                theme: style?.theme,
                prompt: style?.prompt,
                at: new Date().toISOString()
              });
            }
          });
        }
      }catch(_){}

      // fallback slide size (pt). Not all hosts expose page size.
      let slideW = 960, slideH = 540;

      try{ pic.left = 0; }catch(_){}
      try{ pic.top  = 0; }catch(_){}
      try{ pic.width  = slideW; }catch(_){}
      try{ pic.height = slideH; }catch(_){}

      if(opts.sendToBack){
        try{ pic.sendToBack(); }catch(_){}
      }

      await ctx.sync();
    });

    if(hint) hint.textContent = opts.sendToBack ? 'Applied as background (fit & back).' : 'Applied as background (fit).';
  }catch(e){
    if(hint) hint.textContent = 'Applied (basic insert).';
  }
}

// ---- Paste image from clipboard ----
async function nmPasteAsBackground(){
  const hint = q('nmHint');
  try{
    if(!navigator.clipboard || !navigator.clipboard.read){
      if(hint) hint.textContent='Clipboard API not supported.';
      return;
    }
    const items = await navigator.clipboard.read();
    for(const item of items){
      for(const type of item.types){
        if(type.startsWith("image/")){
          const blob = await item.getType(type);
          const buf = await blob.arrayBuffer();
          const base64 = btoa(String.fromCharCode(...new Uint8Array(buf)));
          await nmApplyAsBackground(base64, { sendToBack: (q('nmBgLock')?.checked) });
          if(hint) hint.textContent='Pasted image applied as background.';
          return;
        }
      }
    }
    if(hint) hint.textContent='No image found in clipboard.';
  }catch(e){
    if(hint) hint.textContent='Clipboard paste failed.';
  }
}

// ---- Apply to all selected slides (batch) ----
async function nmApplyAsBackgroundBatch(style){
  const hint = q('nmHint');
  const sendToBack = !!(q('nmBgLock') && q('nmBgLock').checked);

  // 1) resolve/calc image once (cache first)
  let base64 = (typeof nmFindCachedForStyle==='function') ? nmFindCachedForStyle(style) : null;
  if(!base64){
    const gen = await nmGenerate(style);
    if(!gen){
      if(hint) hint.textContent = 'Generation failed — batch aborted.';
      return;
    }
    base64 = gen;
  }

  // 2) enumerate selected slides
  if(!(window.PowerPoint && PowerPoint.run)){
    // fallback: apply on current slide only
    await nmApplyAsBackground(base64, { sendToBack });
    if(hint) hint.textContent = 'Host lacks batch API — applied to current slide.';
    return;
  }

  try{
    await PowerPoint.run(async (ctx)=>{
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items");
      await ctx.sync();

      if(!sel.items || sel.items.length===0){
        // no selection, apply to current
        await nmApplyAsBackground(base64, { sendToBack });
        return;
      }

      // loop slides: select → insert → size → back
      for(const s of sel.items){
        try{
          // switch selection to that slide
          ctx.presentation.setSelectedSlides([s.id]);
          await ctx.sync();

          // apply to this slide
          // (we're in run context; nmApplyAsBackground uses Office.context + a new run, which is fine)
          // keep it sequential to avoid race
          // eslint-disable-next-line no-await-in-loop
          await nmApplyAsBackground(base64, { sendToBack });
        }catch(_){}
      }
    });

    if(hint) hint.textContent = `Applied to ${'selected'} slide(s).`;
  }catch(e){
    // fallback to current slide only
    await nmApplyAsBackground(base64, { sendToBack });
    if(hint) hint.textContent = 'Batch failed — applied to current slide.';
  }
}

/* Core generate with provider/timeout/cancel/cache (+ catch for canceled) */
async function nmGenerate(style){
  const key=nmCacheKey(style);

  // أوقف أي عملية توليد سابقة قبل ما نبدأ الجديدة
  nmAbortInFlight('preempt');

  nmShowBusy(true); 
  nmSetProgress(0,'Starting');

  try{
    const cached = NISPersist.cacheGet(key);
    if(cached){ nmSetProgress(100,'Cached'); try{ nmShowCachedButton(style); }catch(e){}; return cached; }

    const ctrl = new AbortController();
    __NIS_GEN_ABORT = ctrl;
    const onProgress = (p,msg)=>{ try{ nmSetProgressThrottled(p,msg||''); }catch(e){} };

    let resBase64=null, attempt=0, maxAttempts=2;
    let currentStyle = {...style};

    while(attempt < maxAttempts){
      try{
        nmSetProgress(5+attempt*5, attempt?('Retry '+attempt):'Generating');
        const { base64 } = await nmGenerateOnce(currentStyle, onProgress, ctrl, 45000);
        resBase64 = base64;
        break; // success
      }catch(err){
        if (ctrl.signal.aborted) throw err; // canceled
        attempt++;
        if(attempt >= maxAttempts) {
          throw err;
        }
        // Retry: زوّد الـ seed وغير الرسالة
        currentStyle = {...currentStyle, seed: (Number(currentStyle.seed)||0)+1};
        nmSetProgress(10+attempt*5, 'Retrying with seed '+currentStyle.seed);
      }
    }

    if(!resBase64){
      nmSetProgress(25,'Fallback generator');
      resBase64 = nmGeneratePNG(currentStyle);
    }

    nmSetProgress(90,'Caching');
    NISPersist.cacheSet(nmCacheKey(currentStyle), resBase64); // أخزّن حسب الستايل المُجرّب فعليًا

    nmSetProgress(100,'Done');
    try{ nmShowCachedButton(currentStyle); }catch(e){}
    return resBase64;
  } catch(err){
    const h=q('nmHint'); 
    if(h) h.textContent=(err && err.message==="canceled") ? "Canceled." : ("Error: "+(err?.message||"failed"));
    return null;
  } finally {
    __NIS_GEN_ABORT = null;
    nmShowBusy(false);
  }
}

/* Bind Nano UI */
function bindNanoUI(){
  const save=q('nmSave'), styleSel=q('nmStyleSelected'), regen=q('nmRegenSelected'), hint=q('nmHint');
  const btnCancel=q('nmCancel');
  const applyCached = q('nmApplyCached');
  const applyBg = q('nmApplyBackground');
  const applyBgBatch = q('nmApplyBackgroundBatch');
  const bgLock       = q('nmBgLock');
  const copySel = q('nmCopyStyleSel');
  const copyAll = q('nmCopyStyleAll');
  const pasteBg = q('nmPasteBackground');
  const expStyles = q('nmExportStyles');
  const impStyles = q('nmImportStyles');
  const impFile   = q('nmImportStylesFile');

    // ---- Overlay UI wiring ----
  const ovCorner   = q('ovCorner');
  const ovOpacity  = q('ovOpacity');
  const ovOpacityVal = q('ovOpacityVal');
  const ovApplyBtn = q('ovApply');
  const ovRemoveBtn= q('ovRemove');
  const ovHint     = q('ovHint');

  if(ovOpacity && ovOpacityVal){
    ovOpacity.addEventListener('input', ()=>{ ovOpacityVal.textContent = `${ovOpacity.value}%`; });
  }
  if(ovApplyBtn){
    ovApplyBtn.addEventListener('click', async ()=>{
      try{
        if(ovHint) ovHint.textContent = 'Applying...';
        await ovApplyToCurrent({ corner: ovCorner?.value || 'tr', opacity: Number(ovOpacity?.value||80) });
        if(ovHint) ovHint.textContent = 'Overlay applied.';
      }catch(e){
        console.error(e);
        if(ovHint) ovHint.textContent = 'Failed to apply overlay.';
      }
    });
  }
  if(ovRemoveBtn){
    ovRemoveBtn.addEventListener('click', async ()=>{
      try{
        await PowerPoint.run(async (ctx)=>{ await ovRemoveCurrent(ctx); });
        if(ovHint) ovHint.textContent = 'Overlay removed (if existed).';
      }catch(e){
        console.error(e);
        if(ovHint) ovHint.textContent = 'Failed to remove overlay.';
      }
    });
  }

if(expStyles){
  expStyles.addEventListener('click', ()=>{ nmExportAllStyles(); });
}
if(impStyles && impFile){
  impStyles.addEventListener('click', ()=>{ impFile.click(); });
  impFile.addEventListener('change', ev=>{
    const f=ev.target.files && ev.target.files[0];
    if(f){ nmImportStylesFile(f); impFile.value=''; }
  });
}

if(pasteBg){
  pasteBg.addEventListener('click', async ()=>{
    await nmPasteAsBackground();
  });
}

  // تحديث البطاقة فور تغيّر أي input
['nmTheme','nmSeed','nmPrompt','nmAspect','nmCaption','nmAutoInc'].forEach(id=>{
  const el = document.getElementById(id);
  if(!el) return;
  const handler = ()=>{
    const s = nmReadInputs();
    nmUpdateStyleChipThrottled(s);
  };
  el.addEventListener('input', handler);
  el.addEventListener('change', handler);
});

// بعد أي حفظ/توليد/إعادة توليد حدّث البطاقة
const _afterGenHint = ()=>{
  const s = nmReadInputs();
  nmUpdateStyleChipThrottled(s);
};

const saveBtn  = document.getElementById('nmSave');
const applyBtn = document.getElementById('nmStyleSelected');
const regenBtn = document.getElementById('nmRegenSelected');
const cacheBtn = document.getElementById('nmApplyCached');

if(saveBtn)  saveBtn .addEventListener('click', _afterGenHint);
if(applyBtn) applyBtn.addEventListener('click', _afterGenHint);
if(regenBtn) regenBtn.addEventListener('click', _afterGenHint);
if(cacheBtn) cacheBtn.addEventListener('click', _afterGenHint);

  if(save){ save.addEventListener('click',()=>{ const s=nmReadInputs(); nmSaveStyle(s); if(hint) hint.textContent='Style saved for this slide'; }); }

  if(styleSel){ styleSel.addEventListener('click',async()=>{
    const s=nmReadInputs(); nmSaveStyle(s); nmShowCachedButton(s);
    try{ const b64=await nmGenerate(s); if(b64) nmApplyToSelection(b64); if(hint && b64) hint.textContent='Applied to selection'; }
    catch(err){ if(hint) hint.textContent=(err&&err.message==='timeout')?'Generation timeout.':'Generation failed.'; }
  }); }

  if(regen){ regen.addEventListener('click',async()=>{
    let s=nmReadInputs();
    if(s.autoInc){ s.seed=(s.seed||0)+1; nmWriteInputs(s); }
    nmSaveStyle(s); nmShowCachedButton(s);
    try{ const b64=await nmGenerate(s); if(b64) nmApplyToSelection(b64); if(hint && b64) hint.textContent=s.autoInc?'Next seed applied':'Re-generated with same seed'; }
    catch(err){ if(hint) hint.textContent=(err&&err.message==='timeout')?'Generation timeout.':'Generation failed.'; }
  }); }

  if(applyCached){
  applyCached.addEventListener('click', ()=>{
    const s = nmReadInputs();
    const b64 = nmFindCachedForStyle(s);
    const hint = q('nmHint');
    if(b64){
      nmApplyToSelection(b64);
      if(hint) hint.textContent = 'Cached image applied.';
    }else{
      if(hint) hint.textContent = 'No cached image for current style.';
      applyCached.style.display = 'none';
    }
  }); }

  // زر Apply as Background (سلايد واحد)
if (applyBg){
  applyBg.addEventListener('click', async ()=>{
    const s   = nmReadInputs();
    nmSaveStyle(s);
    const hint = q('nmHint');

    // جرّب كاش أولاً
    const b64 = (typeof nmFindCachedForStyle === 'function') ? nmFindCachedForStyle(s) : null;
    if (b64){
      await nmApplyAsBackground(b64, { sendToBack: !!(bgLock && bgLock.checked) });
      if (hint) hint.textContent = 'Background applied from cache.';
      return;
    }

    // مفيش كاش: ولّد ثم طبّق
    const gen = await nmGenerate(s);
    if (gen){
      await nmApplyAsBackground(gen, { sendToBack: !!(bgLock && bgLock.checked) });
      if (hint) hint.textContent = 'Background applied.';
    } else {
      if (hint) hint.textContent = 'Generation failed — cannot apply background.';
    }
  });
}

// زر Apply to Selected Slides (Batch)
if (applyBgBatch){
  applyBgBatch.addEventListener('click', async ()=>{
    const s = nmReadInputs();
    nmSaveStyle(s);
    nmShowBusy(true); nmSetProgress(0, 'Batch applying');
    try{
      await nmApplyAsBackgroundBatch(s);
      nmSetProgress(100, 'Done');
    } finally {
      nmShowBusy(false);
    }
  });
}

// Toggle "Send to back"
if (bgLock){
  bgLock.addEventListener('change', ()=>{
    const hint = q('nmHint');
    if (hint) hint.textContent = bgLock.checked ? 'Backgrounds will be sent to back.' : 'Backgrounds inserted in front.';
  });
}

// Copy style → Selected
if (copySel){
  copySel.addEventListener('click', async ()=>{
    const hint = q('nmHint');
    const s = nmReadInputs();
    nmSaveStyle(s);
    const ids = await ppGetSelectedSlideIds();
    if (!ids || ids.length === 0){
      if (hint) hint.textContent = 'No selected slides.';
      return;
    }
    const n = await nmCopyStyleToSlides(s, ids);
    if (hint) hint.textContent = `Style copied to ${n} selected slide(s).`;
  });
}

// Copy style → All
if (copyAll){
  copyAll.addEventListener('click', async ()=>{
    const hint = q('nmHint');
    const s = nmReadInputs();
    nmSaveStyle(s);
    const ids = await ppGetAllSlideIds();
    if (!ids || ids.length === 0){
      if (hint) hint.textContent = 'No slides found.';
      return;
    }
    const n = await nmCopyStyleToSlides(s, ids);
    if (hint) hint.textContent = `Style copied to ${n} slide(s).`;
  });
}


  if(btnCancel){
  btnCancel.addEventListener('click',()=>{
    nmAbortInFlight('user-cancel');
  });
}
}
function nmInit(){
  const s = nmGetStyle();
  nmWriteInputs(s);
  // لو فيه صورة متخزّنة للستايل الحالي، نعرض زرار Apply cached
  nmShowCachedButton(s);
  nmUpdateStyleChipThrottled(s);
}

/* ---------- Linked Sequence (MVP + Auto-advance + Inherit) ---------- */
/* Storage */
function linkLoad(k){
  const persisted = NISPersist.loadLink(k);
  if(persisted && typeof persisted==='object'){
    return {
      enabled: !!persisted.enabled,
      next: persisted.next || null,
      auto: !!persisted.auto,
      autoMs: Number(persisted.autoMs||3000),
      inherit: !!persisted.inherit,
      inheritNm: !!persisted.inheritNm
    };
  }
  return {enabled:false, next:null, auto:false, autoMs:3000, inherit:false, inheritNm:false};
}
function linkSave(k, data){
  const clean={
    enabled:!!data.enabled, next:data.next||null,
    auto:!!data.auto, autoMs:Number(data.autoMs||3000),
    inherit:!!data.inherit, inheritNm:!!data.inheritNm
  };
  __NIS_LINK_CACHE.set(k, clean);
  NISPersist.saveLink(k, clean);
}

// ---- Export all slide styles ----
async function nmExportAllStyles(){
  let styles = {};
  try{
    if(!(window.PowerPoint && PowerPoint.run)){
      return;
    }
    await PowerPoint.run(async (ctx)=>{
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      slides.items.forEach(s=>s.load("id"));
      await ctx.sync();
      for(const sl of slides.items){
        const id = String(sl.id);
        const st = nmLoadStyleForSlideKey(id);
        if(st) styles[id] = st;
      }
    });
  }catch(e){}
  const blob = new Blob([JSON.stringify(styles,null,2)], {type:"application/json"});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='nis-project-styles.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---- Import styles JSON ----
function nmImportStylesFile(f){
  const r=new FileReader();
  r.onload=async ()=>{
    try{
      const data=JSON.parse(r.result);
      for(const [id,st] of Object.entries(data)){
        nmSaveStyleForSlideKey(id, st);
      }
      const hint=q('nmHint'); if(hint) hint.textContent='Styles imported for '+Object.keys(data).length+' slide(s).';
    }catch(e){}
  };
  r.readAsText(f);
}

/* UI helpers */
async function linkPopulateDropdown(){
  const sel=q('linkNext'); if(!sel) return;
  sel.innerHTML='<option value="">— None —</option>';
  if(!hostIsPowerPoint() || !(window.PowerPoint&&PowerPoint.run)){
    const hint=q('linkHint'); if(hint) hint.textContent='(PowerPoint API not available)';
    return;
  }
  try{
    await PowerPoint.run(async (ctx)=>{
      const coll = ctx.presentation.slides;
      coll.load("items"); await ctx.sync();
      coll.items.forEach(s=>s.load("id,index")); 
      await ctx.sync();
      coll.items.forEach(sl=>{
        const opt=document.createElement('option');
        opt.value=String(sl.id);
        opt.textContent='Slide '+(Number(sl.index)+1);
        sel.appendChild(opt);
      });
    });
  }catch(e){
    const hint=q('linkHint'); if(hint) hint.textContent='(Cannot enumerate slides)';
  }
}
function linkWriteToUI(k){
  const state = __NIS_LINK_CACHE.get(k) || linkLoad(k);
  __NIS_LINK_CACHE.set(k,state);
  const en=q('linkEnable'), nextSel=q('linkNext'), au=q('linkAuto'), auMs=q('linkAutoMs');
  const inh=q('linkInherit'), inhNm=q('linkInheritNm');
  if(en) en.checked=!!state.enabled;
  if(nextSel){ nextSel.value = state.next || ""; }
  if(au) au.checked=!!state.auto;
  if(auMs) auMs.value=String(Number(state.autoMs||3000));
  if(inh) inh.checked=!!state.inherit;
  if(inhNm) inhNm.checked=!!state.inheritNm;
}
function linkReadFromUI(){
  const en=q('linkEnable'), nextSel=q('linkNext'), au=q('linkAuto'), auMs=q('linkAutoMs');
  const inh=q('linkInherit'), inhNm=q('linkInheritNm');
  return { 
    enabled: !!(en && en.checked), 
    next: (nextSel && nextSel.value) ? nextSel.value : null,
    auto: !!(au && au.checked),
    autoMs: Number(auMs && auMs.value ? auMs.value : 3000),
    inherit: !!(inh && inh.checked),
    inheritNm: !!(inhNm && inhNm.checked)
  };
}

/* Auto-advance timer (per active slide) */
function linkAutoArmForActive(){
  linkAutoClear();
  const k=__NIS_ACTIVE_SLIDE_KEY; if(!k) return;
  const st = __NIS_LINK_CACHE.get(k) || linkLoad(k);
  if(!st.enabled || !st.next || !st.auto) return;
  const ms = Math.max(200, Number(st.autoMs||3000));
  __NIS_LINK_AUTO_TIMER = setTimeout(async ()=>{
    __NIS_LINK_AUTO_TIMER=null;
    await linkAdvanceFrom(k);
  }, ms);
}
async function linkAdvanceFrom(k){
  const st = __NIS_LINK_CACHE.get(k) || linkLoad(k);
  if(!st.enabled || !st.next) return false;
  if(!(window.PowerPoint&&PowerPoint.run)) return false;

  // mark to inherit on first visit if flags enabled
  __NIS_INHERIT_NEXT = st.inherit || st.inheritNm ? {from:k} : null;

  try{
    await PowerPoint.run(async (ctx)=>{
      ctx.presentation.setSelectedSlides([st.next]); // API set 1.5
      await ctx.sync();
    });
    return true;
  }catch(e){
    const hint=q('linkHint'); if(hint) hint.textContent='(Advance failed)';
    __NIS_INHERIT_NEXT = null;
    return false;
  }
}

/* Bind UI */
function bindLinkUI(){
  const en=q('linkEnable'), nextSel=q('linkNext');
  const play=q('linkPlay'), adv=q('linkAdvance'), hint=q('linkHint');
  const au=q('linkAuto'), auMs=q('linkAutoMs'), inh=q('linkInherit'), inhNm=q('linkInheritNm');

  const onSave=()=>{ if(!__NIS_ACTIVE_SLIDE_KEY) return; const cur=linkReadFromUI(); linkSave(__NIS_ACTIVE_SLIDE_KEY,cur); if(hint) hint.textContent='Saved.'; };

  en?.addEventListener('change', onSave);
  nextSel?.addEventListener('change', onSave);
  au?.addEventListener('change', onSave);
  auMs?.addEventListener('change', onSave);
  inh?.addEventListener('change', onSave);
  inhNm?.addEventListener('change', onSave);

  play?.addEventListener('click',()=>{
    const cur=getUIParams();
    if(cur.autoStart){ const e=getActiveEngine(); e?.start?.(); linkAutoArmForActive(); }
    if(hint) hint.textContent='Sequence ready — auto/advance as set.';
  });
  adv?.addEventListener('click', async ()=>{
    linkAutoClear();
    const ok = await linkAdvanceFrom(__NIS_ACTIVE_SLIDE_KEY);
    if(!ok){ if(hint) hint.textContent='No next slide set for this slide.'; }
    else{ if(hint) hint.textContent='Advanced.'; }
  });
}

/* Restore for current slide */
function linkRestoreForSlide(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;
  linkPopulateDropdown().then(()=>{ linkWriteToUI(__NIS_ACTIVE_SLIDE_KEY); });
}

/* ---------- Boot ---------- */
function initBoot(){
  bindSimUI();
  bindNanoUI();
  bindLinkUI();

  captureSlideKeyFast().then(k=>{
    __NIS_ACTIVE_SLIDE_KEY=k;
    getActiveEngine();
    restoreCurrentSlide();
    nmInit();
    stylesDashInit();
    linkRestoreForSlide();

    const cur = getUIParams();
    if(cur.autoStart){
      const e = getActiveEngine();
      if(e && e.setProjectToSlide) e.setProjectToSlide(!!cur.projectToSlide);
      if(e && e.setProjectMs)       e.setProjectMs(Number(cur.projectMs||1000));
      e?.start?.();
      linkAutoArmForActive();
    }});

  wireSlideChange();
  setHostHint();

  const e=getActiveEngine(); e?.reset?.();

  // Shortcuts: Ctrl+Alt+S toggle, Ctrl+Alt+Right advance
  document.addEventListener('keydown', (ev)=>{
    if(!(ev.ctrlKey && ev.altKey)) return;
    if(ev.code==='KeyS'){
      ev.preventDefault();
      linkAutoClear();
      const e=getActiveEngine(); e?.stop?.(); setTimeout(()=>{ e?.start?.(); linkAutoArmForActive(); },0);
    }else if(ev.code==='ArrowRight'){
      ev.preventDefault();
      linkAutoClear();
      linkAdvanceFrom(__NIS_ACTIVE_SLIDE_KEY);
    }
  });
}

if(window.Office && Office.onReady){
  Office.onReady(()=>{ 
    if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', initBoot); }
    else { initBoot(); }
  });
}else{
  if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', initBoot); }
  else { initBoot(); }
}
// ====== end of taskpane.js ======
} // <--- close guard
