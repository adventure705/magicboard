'use strict';
/* Memory only. save() resolves to the confirmed field value after readback. */
(function (root, factory) {
  'use strict';
  const api = Object.freeze({ create: factory });
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.MagicBoardAutosave = api;
})(globalThis, function create({ save, canStart = () => true, changed = () => {}, confirmed = () => {},
  delay = 500, schedule = setTimeout, cancel = clearTimeout }) {
  const drafts = new Map();
  const key = (id, field) => JSON.stringify([id, field]);
  let flight = null;
  function stop(d) { if (d.timer != null) cancel(d.timer); d.timer = null; }
  function wake() { queueMicrotask(drain); } // Never start inside a remote state callback/finally.
  async function drain() {
    if (flight || !canStart()) return;
    const d = Array.from(drafts.values()).find(d => d.ready && !d.error);
    if (!d) return;
    flight = d; d.ready = false;
    const sent = d.value;
    try {
      const result = await save({ id: d.id, field: d.field, value: sent, expected: d.expected });
      if (flight !== d) return; // reset, possibly followed by another session
      if (typeof result !== 'string') throw new Error('저장한 값을 확인할 수 없습니다.');
      d.expected = result;
      confirmed(d.id, d.field, result);
      if (d.value === sent || d.value === result) { stop(d); drafts.delete(d.key); }
    } catch (e) {
      if (flight !== d) return;
      stop(d); d.ready = false; d.error = e?.message || '저장하지 못했습니다.';
      d.needsCheck = true; // POST may have succeeded; even reverting locally requires a readback.
    } finally {
      if (flight === d) { flight = null; changed(); wake(); }
    }
  }
  function input(id, field, value, expected, composing = false) {
    const k = key(id, field);
    const d = drafts.get(k) || {key:k,id,field,expected,error:'',needsCheck:false};
    stop(d); d.value = value; d.ready = false;
    if (d !== flight && !d.needsCheck && value === d.expected) drafts.delete(k);
    else {
      drafts.set(k, d);
      if (!composing) d.timer = schedule(() => { d.timer = null; d.ready = true; wake(); }, delay);
    }
    changed();
  }
  function flush(id, field) {
    const d = drafts.get(key(id, field));
    if (d && !d.error) { stop(d); d.ready = true; wake(); }
  }
  function retry(id, field) {
    const d = drafts.get(key(id, field));
    if (d?.error) { d.error = ''; flush(id, field); changed(); }
  }
  function discard(id, field) {
    const d = drafts.get(key(id, field));
    if (d?.error && d !== flight) { stop(d); drafts.delete(d.key); changed(); }
  }
  function reset() { for (const d of drafts.values()) stop(d); drafts.clear(); flight = null; changed(); }
  const snapshot = d => d && ({id:d.id,field:d.field,value:d.value,expected:d.expected,error:d.error});
  return Object.freeze({input,flush,retry,discard,wake,reset,hasUnsaved:()=>drafts.size>0,
    get:(id,field)=>snapshot(drafts.get(key(id,field))),errors:()=>[...drafts.values()].filter(d=>d.error).map(snapshot)});
});
