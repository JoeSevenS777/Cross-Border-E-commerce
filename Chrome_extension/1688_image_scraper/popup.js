const logEl = document.getElementById('log');
function log(msg){
  const t = new Date().toLocaleTimeString();
  logEl.textContent += `[${t}] ${msg}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}
function setMinSize(v){ return chrome.storage.sync.set({minSize: v}); }
function getMinSize(){ return chrome.storage.sync.get({minSize: 450}); }

document.getElementById('reload').addEventListener('click', async () => {
  log('Refreshing page...');
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (tab?.id) {
    chrome.tabs.reload(tab.id);
  }
});


document.getElementById('run').addEventListener('click', async () => {
  const minSize = Math.max(1, parseInt(document.getElementById('minSize').value || '450',10));
  await setMinSize(minSize);
  log('Requesting scrape...');
  chrome.runtime.sendMessage({type:'SCRAPE_DOWNLOAD', minSize}, (resp) => {
    if (chrome.runtime.lastError) {
      log('Failed: ' + chrome.runtime.lastError.message);
      return;
    }
    if (!resp) { log('No response.'); return; }
    if (!resp.ok) { log('Failed: ' + (resp.error||'unknown')); return; }
    log(`Done. details=${resp.counts.details} main=${resp.counts.main} sku=${resp.counts.sku} (downloaded=${resp.downloaded}) folder=${resp.folder || 'n/a'} video=${resp.videoDownloaded ? 'yes' : 'no'}`);
    if (resp.note) log(resp.note);
  });
});

(async ()=>{
  const {minSize} = await getMinSize();
  document.getElementById('minSize').value = minSize;
  log('Ready.');
})();