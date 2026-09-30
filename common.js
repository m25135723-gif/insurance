/* 工具箱共用程式：連線設定、API 呼叫、個資遮蔽、日期計算 */
const TB = (() => {
  const CFG_KEY = 'toolbox.tracker.cfg';
  const PRIV_KEY = 'toolbox.privacy';
  const store = {
    get(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } },
    set(k,v){ try{ localStorage.setItem(k,v); }catch(e){} },
    del(k){ try{ localStorage.removeItem(k); }catch(e){} }
  };
  function getCfg(){ try{ const c = JSON.parse(store.get(CFG_KEY)||'null'); return c && c.url && c.key ? c : null; }catch(e){ return null; } }
  function setCfg(url,key){ store.set(CFG_KEY, JSON.stringify({url:url.trim(), key:key.trim()})); }
  function clearCfg(){ store.del(CFG_KEY); }

  const ERR = {
    unauthorized:'密鑰不正確，請到「設定」重新輸入。',
    not_setup:'後端尚未執行 setup，請依架設教學完成第 4 步。',
    bad_request:'送出的資料格式有誤。',
    unknown_action:'後端版本太舊，請更新 Apps Script 程式碼並重新部署。'
  };
  async function api(action, payload, cfgOverride){
    const cfg = cfgOverride || getCfg();
    if(!cfg) throw new Error('尚未連線，請先到「設定」輸入網址和密鑰。');
    let r;
    try{
      r = await fetch(cfg.url, {method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'},
        body: JSON.stringify(Object.assign({key:cfg.key, action}, payload||{}))});
    }catch(e){ throw new Error('連不到後端，請確認網路和網址是否正確。'); }
    let j;
    try{ j = await r.json(); }catch(e){ throw new Error('後端回應不是預期格式，請確認網址是「網頁應用程式」的部署網址。'); }
    if(!j.ok) throw new Error(ERR[j.error] || ('後端錯誤：' + j.error));
    return j.data;
  }

  function isPrivate(){ return store.get(PRIV_KEY) === '1'; }
  function setPrivate(on){ store.set(PRIV_KEY, on ? '1' : '0'); }
  function maskName(s){
    s = String(s||''); const a = Array.from(s);
    if(a.length <= 1) return s;
    if(a.length === 2) return a[0] + '＊';
    return a[0] + '＊'.repeat(a.length-2) + a[a.length-1];
  }
  function maskPlate(s){
    s = String(s||''); const a = Array.from(s);
    if(a.length <= 3) return s;
    return a.map((c,i)=> (i < 2 || i >= a.length-2 || c === '-') ? c : '＊').join('');
  }
  const name = s => isPrivate() ? maskName(s) : String(s||'');
  const plate = s => isPrivate() ? maskPlate(s) : String(s||'');

  function esc(s){ return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

  // 日期一律用 YYYY-MM-DD 字串手動解析，避免 new Date('YYYY-MM-DD') 被當成 UTC 的問題
  function dayNum(s){ const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s||''); return m ? Date.UTC(+m[1], +m[2]-1, +m[3]) / 864e5 : null; }
  function todayStr(){ const d = new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
  function daysUntil(s){ const n = dayNum(s); return n == null ? null : n - dayNum(todayStr()); }
  function fmtDate(s){ return /^\d{4}-\d{2}-\d{2}$/.test(s||'') ? s.replace(/-/g,'/') : '—'; }

  function toast(t){ const el = document.getElementById('toast'); if(!el) return; el.textContent = t; el.classList.add('show'); clearTimeout(toast.h); toast.h = setTimeout(()=>el.classList.remove('show'), 2200); }

  return {getCfg, setCfg, clearCfg, api, isPrivate, setPrivate, maskName, maskPlate, name, plate, esc, dayNum, todayStr, daysUntil, fmtDate, toast};
})();
