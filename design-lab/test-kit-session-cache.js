// A staff-opened Kit Expansion session must be readable by the OWNER-side commit.
// The modal opens as the invoking user; commitKitFromModal hops to the owner via
// /exec (OwnerBridge). A USER cache is per-user, so the owner could never see it.
// Run: node design-lab/test-kit-session-cache.js   (SRC=<dir> to test another copy)
const fs = require('fs'), path = require('path'), vm = require('vm');
const src = fs.readFileSync(path.join(process.env.SRC || path.join(__dirname, '..'), 'KitExpansion.js'), 'utf8');
function grab(name) {
  const i = src.indexOf('function ' + name + '(');
  let d = 0, j = src.indexOf('{', i);
  for (let k = j; k < src.length; k++) { if (src[k] === '{') d++; else if (src[k] === '}' && --d === 0) return src.slice(i, k + 1); }
}
let who = 'staff';
const userStores = {}, scriptStore = {};
const mk = s => ({ get: k => (k in s ? s[k] : null), put: (k, v) => { s[k] = v; }, remove: k => { delete s[k]; } });
const ctx = { CacheService: { getUserCache: () => mk(userStores[who] = userStores[who] || {}), getScriptCache: () => mk(scriptStore) },
  JSON, KIT_MODAL_CACHE_PREFIX: 'KitExpansionModal:', KIT_MODAL_CACHE_TTL: 1800 };
vm.createContext(ctx);
vm.runInContext(grab('_saveKitModalSession') + '\n' + grab('_loadKitModalSession'), ctx);
let fail = 0; const ok = (n, c) => { console.log((c ? '  ✓ ' : '  ✗ ') + n); if (!c) fail++; };
who = 'staff'; ctx._saveKitModalSession('abc', { currentIndex: 0, queue: [1] });
who = 'owner'; const s = ctx._loadKitModalSession('abc');
ok('owner-side commit finds the staff-opened session', !!s && s.queue.length === 1);
ok('unknown session id still misses', ctx._loadKitModalSession('nope') === null);
ok('empty session id misses', ctx._loadKitModalSession('') === null);
ok('opener writes the SCRIPT cache too (openKitExpansionModal)', /cache = CacheService\.getScriptCache\(\);\s*\n\s*cache\.put\(KIT_MODAL_CACHE_PREFIX/.test(src));
ok('no getUserCache call left in the session code', !/CacheService\.getUserCache\(\)/.test(src));
console.log(fail ? fail + ' FAILED' : 'all green'); process.exit(fail ? 1 : 0);
