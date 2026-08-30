(async function(){
  try {
    const src = await fetch('updated-data.js?v=' + Date.now(), {cache:'no-store'}).then(r => {
      if (!r.ok) throw new Error('Could not load updated-data.js: ' + r.status);
      return r.text();
    });
    const m = src.match(/const b64='([^']+)'/);
    if (!m) throw new Error('Embedded dashboard data was not found.');
    const bin = Uint8Array.from(atob(m[1]), c => c.charCodeAt(0));
    let text = null;
    // Try browser-native formats first, then Brotli (the current dataset is Brotli-compressed).
    for (const format of ['gzip','deflate','deflate-raw']) {
      try {
        const ds = new DecompressionStream(format);
        text = await new Response(new Blob([bin]).stream().pipeThrough(ds)).text();
        JSON.parse(text);
        break;
      } catch (_) { text = null; }
    }
    if (!text) {
      const mod = await import('https://unpkg.com/brotli-wasm@3.0.1/index.web.js?module');
      const brotli = await mod.default;
      text = new TextDecoder().decode(brotli.decompress(bin));
    }
    window.S1_DATA = JSON.parse(text);
    window.dispatchEvent(new Event('s1dataready'));
  } catch (err) {
    console.error('S1 dashboard data load failed:', err);
    window.S1_DATA_ERROR = String(err && err.message ? err.message : err);
    window.dispatchEvent(new Event('s1dataerror'));
  }
})();
