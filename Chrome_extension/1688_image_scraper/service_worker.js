function sanitizeName(name){
  name = (name || '1688_product').toString().trim();
  name = name.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g,' ').trim();
  if (!name) name='1688_product';
  if (name.length>80) name=name.slice(0,80);
  return name;
}
function pad3(n){ return String(n).padStart(3,'0'); }
function pad2(n){ return String(n).padStart(2,'0'); }

async function getUniqueFolderName(baseName){
  // We cannot reliably check the OS download directory. Instead, we keep a lightweight
  // per-browser counter so repeated downloads of the same product don't overwrite.
  const base = sanitizeName(baseName);
  const data = await chrome.storage.local.get({folderCounters:{}});
  const counters = data.folderCounters || {};
  const next = (counters[base] || 0) + 1;
  counters[base] = next;
  await chrome.storage.local.set({folderCounters: counters});
  // First time -> no suffix; subsequent runs -> _01, _02...
  if (next === 1) return base;
  return `${base}_${pad2(next-1)}`;
}


async function getActiveTab(){
  const tabs = await chrome.tabs.query({active:true, currentWindow:true});
  return tabs && tabs[0];
}

function uniq(arr){
  const seen=new Set(); const out=[];
  for (const x of arr){ if (!x) continue; const k=String(x); if (seen.has(k)) continue; seen.add(k); out.push(k); }
  return out;
}

async function fetchImageSize(url, timeoutMs=12000){
  const ctrl = new AbortController();
  const to = setTimeout(()=>ctrl.abort('timeout'), timeoutMs);
  try{
    const res = await fetch(url, {signal: ctrl.signal, credentials:'omit', mode:'cors'});
    if (!res.ok) return null;
    const blob = await res.blob();
    // createImageBitmap is available in SW in modern Chromium
    const bmp = await createImageBitmap(blob);
    const w = bmp.width, h = bmp.height;
    bmp.close && bmp.close();
    return {w,h};
  }catch(e){
    return null;
  }finally{
    clearTimeout(to);
  }
}

function guessExt(url){
  const m = url.match(/\.(jpg|jpeg|png|webp|gif)(\?|#|$)/i);
  if (!m) return 'jpg';
  const ext = m[1].toLowerCase();
  return ext === 'jpeg' ? 'jpg' : ext;
}

async function downloadOne(folder, sub, idx, url, nameOverride){
  const ext = guessExt(url);
  const base = nameOverride ? nameOverride : pad3(idx);
  const filename = `${folder}/${sub}/${base}.${ext}`;
  await chrome.downloads.download({url, filename, conflictAction:'uniquify', saveAs:false});
}

function isValidVideoUrl(url){
  if (!url) return false;
  const u = String(url).trim();
  if (!u) return false;
  if (u.startsWith('blob:')) return false;
  if (u.startsWith('data:')) return false;
  if (!/^https?:\/\//i.test(u)) return false;
  // Prevent useless HTML downloads like video.htm when URL points to a page/iframe
  return /\.(mp4|webm|m3u8|mov|mkv)(\?|#|$)/i.test(u);
}

function guessVideoExt(url){
  const u = String(url || '').toLowerCase();
  const m = u.match(/\.(mp4|webm|m3u8|mov|mkv)(\?|#|$)/i);
  return m ? m[1].toLowerCase() : 'mp4';
}

async function downloadVideo(folder, url){
  if (!isValidVideoUrl(url)) return false;
  const ext = guessVideoExt(url);
  const filename = `${folder}/video.${ext}`;
  await chrome.downloads.download({url, filename, conflictAction:'uniquify', saveAs:false});
  return true;
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    try{
      // Utility: fetch arbitrary text (used by content script to retrieve detail HTML via detailUrl/iframes)
      if (msg?.type === 'FETCH_TEXT') {
        const url = String(msg.url || '');
        if (!/^https?:\/\//i.test(url)) { sendResponse({ok:false, error:'Bad url'}); return; }
        try{
          const res = await fetch(url, {credentials:'omit', cache:'no-store'});
          const text = await res.text();
          sendResponse({ok:true, status: res.status, url: res.url, text});
        }catch(e){
          sendResponse({ok:false, error: String(e?.message||e)});
        }
        return;
      }

      if (msg?.type !== 'SCRAPE_DOWNLOAD') return;

      const tab = await getActiveTab();
      if (!tab?.id) { sendResponse({ok:false, error:'No active tab.'}); return; }
      if (!/^https:\/\/detail\.1688\.com\/offer\//.test(tab.url||'')) {
        sendResponse({ok:false, error:'Open a 1688 offer page: https://detail.1688.com/offer/...' });
        return;
      }

      // Ensure content script is ready
      const minSize = Math.max(1, parseInt(msg.minSize||450,10));
      let resp;
      try{
        resp = await chrome.tabs.sendMessage(tab.id, {type:'SCRAPE_COLLECT', minSize});
      }catch(e){
        sendResponse({ok:false, error:'Content script not reachable. Refresh the page once and try again.'});
        return;
      }
      if (!resp?.ok) { sendResponse({ok:false, error: resp?.error || 'Scrape failed'}); return; }

      const productName = sanitizeName(resp.productName);
      const folder = await getUniqueFolderName(productName);
      const groups = resp.groups || {main:[], sku:[], details:[]};
      const skuItems = Array.isArray(resp.skuItems) ? resp.skuItems : [];

      // Dedup
      for (const k of ['main','sku','details']) groups[k] = uniq(groups[k]);

      // If named SKU items exist, use them as authoritative SKU images and remove from main.
      const skuItemUrls = skuItems.map(it => it && it.url).filter(Boolean);
      if (skuItemUrls.length){
        groups.sku = uniq(skuItemUrls);
        const skuSet = new Set(groups.sku);
        groups.main = groups.main.filter(u => !skuSet.has(u));
      }

      // Filter by size
      const filtered = {main:[], sku:[], details:[]};
      const checked = new Map();
      const skuNameByUrl = new Map();
      if (skuItems.length){
        for (let i=0;i<skuItems.length;i++){
          const it=skuItems[i];
          if (!it || !it.url) continue;
          // Preserve page order using the skuItems array order
          const seq = String(i+1).padStart(2,'0');
          skuNameByUrl.set(it.url, `${seq}_${sanitizeName(it.name || 'sku')}`);
        }
      }
      const noteParts=[];
      for (const k of ['details','main','sku']){ // check details first
        for (const u of groups[k]){
          if (!u) continue;
          if (checked.has(u)) {
            const ok = checked.get(u);
            if (ok) filtered[k].push(u);
            continue;
          }
          const sz = await fetchImageSize(u);
          const ok = !!(sz && sz.w>=minSize && sz.h>=minSize);
          checked.set(u, ok);
          if (ok) filtered[k].push(u);
        }
      }

      const counts = {main: filtered.main.length, sku: filtered.sku.length, details: filtered.details.length};

      let downloaded = 0;
      for (const k of ['details','main','sku']){
        let i=1;
        for (const u of filtered[k]){
          const nameOverride = (k==='sku' && skuNameByUrl.has(u)) ? skuNameByUrl.get(u) : null;
          await downloadOne(folder, k, i++, u, nameOverride);
          downloaded++;
        }
      }

      
      let videoDownloaded = false;
      const videoUrl = resp.videoUrl ? String(resp.videoUrl) : '';
      if (videoUrl){
        try{
          videoDownloaded = await downloadVideo(folder, videoUrl);
        }catch(e){
          noteParts.push('Video download failed: ' + String(e?.message||e));
        }
      }

let note='';
      if (noteParts.length) note += (note ? '\n' : '') + noteParts.join('\n');
      if (downloaded===0) {
        note = 'No images >= min size. If the page uses lazy/JS data, try scrolling to the details section and click the 详情/图文详情 tab once, then run again.';
      }

      sendResponse({ok:true, counts, downloaded, videoDownloaded, folder, note});
    }catch(e){
      sendResponse({ok:false, error: String(e?.message||e)});
    }
  })();
  return true;
});
