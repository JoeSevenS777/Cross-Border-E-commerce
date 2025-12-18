/**
 * content.js (v1.0)
 *
 * Goal: classify and return 3 groups of image URLs from a 1688 offer page:
 *  - main: gallery / main image set
 *  - sku: sku/spec thumbnails
 *  - details: description / 图文详情 images
 *
 * Key change vs v0.9: do NOT rely on network hooking (often misses due to timing/CSP).
 * Instead, discover the product-description "detailUrl" from the page's embedded JSON
 * (window.context / __INIT_DATA__-like blobs) and fetch+parse that HTML for <img> URLs.
 */

function uniq(arr){
  const s=new Set();
  const out=[];
  for (const x of arr){
    if (!x) continue;
    const k=String(x);
    if (s.has(k)) continue;
    s.add(k);
    out.push(k);
  }
  return out;
}

function normalizeUrl(u){
  if (!u) return null;
  u = String(u).trim();
  if (!u) return null;
  if (u.startsWith('//')) u = location.protocol + u;
  // strip common thumbnail transforms
  u = u.replace(/(\.(?:jpe?g|png|webp|gif))_.+$/i, '$1');
  u = u.replace(/_(\d+x\d+)(q\d+)?\.(jpe?g|png|webp|gif)(\?|#|$)/i, '.$3$4');
  return u;
}

function extractImgUrlsFromText(text){
  const out=[];
  if (!text) return out;

  // absolute
  const reAbs = /https?:\/\/[^"'\\\s<>]+?\.(?:jpg|jpeg|png|webp)(?:\?[^"'\\\s<>]*)?/ig;
  let m;
  while((m=reAbs.exec(text))!==null){
    out.push(normalizeUrl(m[0]));
  }

  // protocol-relative
  const rePR = /\/\/[^"'\\\s<>]+?\.(?:jpg|jpeg|png|webp)(?:\?[^"'\\\s<>]*)?/ig;
  while((m=rePR.exec(text))!==null){
    out.push(normalizeUrl(m[0]));
  }

  return out.filter(Boolean);
}

function decodeJsonString(s){
  if (s == null) return '';
  const raw = String(s);
  // Many values in embedded JSON are already JSON-escaped.
  try{
    return JSON.parse('"' + raw.replace(/\\/g,'\\\\').replace(/"/g,'\\"') + '"');
  }catch(_){
    return raw
      .replace(/\\u002F/g,'/')
      .replace(/\\\//g,'/')
      .replace(/&amp;/g,'&')
      .trim();
  }
}

function extractFromScripts(pattern){
  try{
    const scripts = Array.from(document.scripts || []);
    for (const s of scripts){
      const txt = s?.textContent || '';
      if (!txt || txt.length < 50) continue;
      const m = txt.match(pattern);
      if (m && m[1]) return decodeJsonString(m[1]);
    }
  }catch(_){}
  return '';
}

function getProductName(){
  // Prefer embedded data (offerTitle/subject) over generic DOM nodes,
  // because on 1688 the first <h1> can be shop navigation.
  const fromScripts =
    extractFromScripts(/\"offerTitle\"\s*:\s*\"([^\"]+)\"/) ||
    extractFromScripts(/\"subject\"\s*:\s*\"([^\"]+)\"/);

  const fromDom =
    document.querySelector('meta[property="og:title"]')?.getAttribute('content') ||
    document.querySelector('meta[name="description"]')?.getAttribute('content') ||
    document.querySelector('[data-spm-anchor-id*="title"] h1')?.textContent ||
    document.querySelector('[class*="title"] h1')?.textContent ||
    document.querySelector('h1[itemprop="name"]')?.textContent ||
    document.title;

  const name = (fromScripts || fromDom || '1688_product').toString().trim();
  // Strip common suffixes
  return name.replace(/\s*-\s*阿里巴巴\s*$/,'').trim() || '1688_product';
}

function getVideoUrl(){
  // Try embedded JSON first
  const v =
    extractFromScripts(/\"videoUrl\"\s*:\s*\"([^\"]+)\"/) ||
    extractFromScripts(/\"cloudVideo\"[\s\S]{0,200}?\"url\"\s*:\s*\"([^\"]+)\"/);

  if (v && /^https?:\/\//i.test(v)) return v;

  // Fallback DOM scan
  const src = document.querySelector('video source')?.src || document.querySelector('video')?.src;
  return (src && /^https?:\/\//i.test(src)) ? src : '';
}


function collectDomImgs(scope){
  const root = scope || document;
  const urls=[];
  root.querySelectorAll('img').forEach(img=>{
    const attrs = ['src','data-src','data-original','data-lazy-src','data-ks-lazyload','data-img','data-url'];
    for (const a of attrs){
      const v = img.getAttribute(a);
      if (v) urls.push(normalizeUrl(v));
    }
    if (img.currentSrc) urls.push(normalizeUrl(img.currentSrc));
  });
  return urls.filter(Boolean);
}

function collectEmbeddedScriptText(){
  const chunks=[];
  document.querySelectorAll('script').forEach(sc=>{
    const t = sc.textContent || '';
    if (t && t.length > 50) chunks.push(t);
  });
  return chunks.join('\n');
}

function extractJsonArrayFromBlob(blob, key){
  // Extract JSON array literal following "<key>": [ ... ] with bracket counting.
  const idx = blob.indexOf(`"${key}"`);
  if (idx === -1) return null;
  const colon = blob.indexOf(':', idx);
  if (colon === -1) return null;
  const lb = blob.indexOf('[', colon);
  if (lb === -1) return null;

  let depth = 0;
  let inStr = false;
  let esc = false;
  for (let i = lb; i < blob.length; i++){
    const ch = blob[i];
    if (inStr){
      if (esc) { esc = false; continue; }
      if (ch === '\\') { esc = true; continue; }
      if (ch === '"') { inStr = false; continue; }
      continue;
    } else {
      if (ch === '"') { inStr = true; continue; }
      if (ch === '[') depth++;
      if (ch === ']') {
        depth--;
        if (depth === 0){
          const raw = blob.slice(lb, i+1);
          try{
            const cleaned = raw
              .replace(/\\u002F/g,'/')
              .replace(/\\\//g,'/')
              .replace(/\u002F/g,'/')
              .replace(/\//g,'/');
            return JSON.parse(cleaned);
          }catch(e){
            return null;
          }
        }
      }
    }
  }
  return null;
}

function collectSkuItemsFirstImageProp(){
  // Return [{name,url}] from the FIRST skuProps dimension that actually provides imageUrl values.
  const blob = collectEmbeddedScriptText();
  const skuProps = extractJsonArrayFromBlob(blob, 'skuProps');
  if (!Array.isArray(skuProps)) return [];
  let chosen = null;
  for (const prop of skuProps){
    const values = prop && prop.value;
    if (!Array.isArray(values) || values.length === 0) continue;
    const hasImg = values.some(v => normalizeUrl(v && v.imageUrl));
    if (hasImg) { chosen = values; break; }
  }
  if (!chosen) return [];
  const items=[];
  for (const v of chosen){
    const name = sanitizeSkuLabel(v && v.name);
    const url = normalizeUrl(v && v.imageUrl);
    if (name && url) items.push({name, url});
  }
  return items;
}



function findDetailUrlFromScripts(){
  const blob = collectEmbeddedScriptText();
  // detailUrl is usually inside a JSON field like: "detailUrl":"https:\/\/itemcdn.tmall.com\/1688offer\/....html"
  const m = blob.match(/"detailUrl"\s*:\s*"([^"]+)"/i);
  if (!m) return null;
  let url = m[1];
  url = url.replace(/\\u002F/g,'/');
  url = url.replace(/\\\//g,'/');
  url = url.replace(/&amp;/g,'&');
  if (url.startsWith('//')) url = location.protocol + url;
  return url;
}

function findDescIframeUrl(){
  const iframes = Array.from(document.querySelectorAll('iframe'));
  for (const fr of iframes){
    const src = fr.getAttribute('src') || '';
    if (/offer_desc\.htm|desc\.htm|description/i.test(src)) {
      return normalizeUrl(src);
    }
  }
  return null;
}

async function fetchTextViaSW(url){
  if (!url) return null;
  return await new Promise((resolve)=>{
    chrome.runtime.sendMessage({type:'FETCH_TEXT', url}, (resp)=>{
      if (!resp?.ok) return resolve(null);
      resolve(resp.text || null);
    });
  });
}

async function collectDetailsImages(){
  // Preferred: detailUrl (itemcdn.tmall.com/1688offer/...)
  const detailUrl = findDetailUrlFromScripts();
  if (detailUrl){
    const html = await fetchTextViaSW(detailUrl);
    if (html) return uniq(extractImgUrlsFromText(html));
  }

  // Fallback: description iframe URL (offer_desc/desc)
  const iframeUrl = findDescIframeUrl();
  if (iframeUrl){
    const html = await fetchTextViaSW(iframeUrl);
    if (html) return uniq(extractImgUrlsFromText(html));
  }

  // Last resort: DOM scan (often contains only placeholders)
  const maybe = collectDomImgs(document);
  // Heuristic: details images frequently are long and not from the main gallery; but we cannot reliably detect.
  return uniq(maybe);
}

function parseGalleryAndSkuFromScripts(){
  const blob = collectEmbeddedScriptText();
  const main=[];
  const sku=[];

  // Pull "mainImage":["url1","url2",...] if present
  const mainMatch = blob.match(/"mainImage"\s*:\s*\[(.*?)\]/s);
  if (mainMatch){
    const block = mainMatch[1];
    const urls = block.match(/"(https?:\\\/\\\/[^\"]+|\\\/\\\/[^\"]+)"/g) || [];
    for (const raw of urls){
      let u = raw.slice(1,-1);
      u = u.replace(/\\u002F/g,'/').replace(/\\\//g,'/');
      main.push(normalizeUrl(u));
    }
  }

  // Pull common sku-image patterns: "imageUrl" or "image" within skuModel
  const skuBlocks = blob.match(/"skuModel"[\s\S]*?\}\s*,\s*"/g) || [];
  const skuText = skuBlocks.join('\n') || blob;
  const skuUrls = skuText.match(/"(?:imageUrl|image|imgUrl)"\s*:\s*"([^"]+)"/ig) || [];
  for (const pair of skuUrls){
    const m = pair.match(/:\s*"([^"]+)"/);
    if (!m) continue;
    let u = m[1].replace(/\\u002F/g,'/').replace(/\\\//g,'/');
    sku.push(normalizeUrl(u));
  }

  return {main: uniq(main).filter(Boolean), sku: uniq(sku).filter(Boolean)};
}

function splitMainSkuFromDom(){
  const mainSet = new Set();
  const skuSet = new Set();

  const gallery = document.querySelector('[class*="gallery"],[class*="swiper"],[class*="slider"],[class*="pic"],[class*="image"],[class*="preview"]');
  collectDomImgs(gallery || document).forEach(u=>mainSet.add(u));

  const skuRoot = document.querySelector('[class*="sku"],[class*="spec"],[class*="prop"],[id*="sku"],[id*="spec"],[class*="od-sku"],[class*="sku-selection"]');
  collectDomImgs(skuRoot || document).forEach(u=>skuSet.add(u));

  return {main: Array.from(mainSet), sku: Array.from(skuSet)};
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    try{
      if (msg?.type !== 'SCRAPE_COLLECT') return;

      const productName = getProductName();

      // Prefer structured script parsing, then merge with DOM-derived URLs.
      const fromScripts = parseGalleryAndSkuFromScripts();
      const fromDom = splitMainSkuFromDom();

      const main = uniq([...(fromScripts.main||[]), ...(fromDom.main||[])]).filter(Boolean);
      const sku = uniq([...(fromScripts.sku||[]), ...(fromDom.sku||[])]).filter(Boolean);

      const details = await collectDetailsImages();

      const skuItems = collectSkuItemsFirstImageProp();
      const videoUrl = getVideoUrl();
      sendResponse({ok:true, productName, groups:{main, sku, details}, skuItems, videoUrl});
    }catch(e){
      sendResponse({ok:false, error: String(e?.message||e)});
    }
  })();
  return true;
});

function sanitizeSkuLabel(text){
  let s=String(text||'').replace(/\s+/g,' ').trim();
  s=s.split(/¥|￥|库存/)[0];
  s=s.replace(/^分类\s*\d+\s*/,'');
  return s.trim();
}
