/* Read-only adapter for header-inclusive Sheets values.batchGet arrays.
 * Use UNFORMATTED_VALUE + FORMATTED_STRING. No storage, network or Google runtime. */
(function (root, factory) {
  'use strict';
  const api = Object.freeze(factory());
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.MagicBoardSheets = api;
})(globalThis, function () {
  'use strict';
  const HOUR = 3600000, DAY = 24 * HOUR, KST = 9 * HOUR;
  const registryHeader = ['샵 ID','샵명','샵 URL','국가','SOLD 공개','활성 리스팅','추적 상태','신규 문서 등록 시각(KST)','기존 등록 시각(KST)','메모','마지막 누적 판매','마지막 누적 즐겨찾기','마지막 관측 시각(KST)','출처','기록 구분'];
  const dailyHeader = ['샵 ID','일자(KST)','수집 시각(KST)','누적 SOLD','일별 SOLD','활성 리스팅','리스팅 증감','누적 샵 즐겨찾기','즐겨찾기 증감','기록 종류','출처 URL','상태','메모'];
  const intervalHeader = ['샵 ID','일자(KST)','구간','예정 수집 시각(KST)','실제 수집 시각(KST)','누적 SOLD','구간 SOLD','활성 리스팅','리스팅 증감','누적 샵 즐겨찾기','즐겨찾기 증감','이전 실제 관측 시각(KST)','출처 URL','상태','집계 구분','메모','이전 예정 시각(KST)'];
  const okStatus = new Set(['정상','성공','완료','부분확인','부분 확인']);
  const dict = () => Object.create(null);
  const text = v => v == null ? '' : String(v);
  const nonempty = v => v != null && v !== '';
  const unsafe = k => k === 'prototype' || Object.hasOwn(Object.prototype, k);
  const fail = message => { throw new Error(message); };

  function cleanShopUrl(value) {
    let s = text(value).trim().replace(/^http:\/\//i, 'https://');
    if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
    const m = /^https:\/\/(?:www\.|m\.)?etsy\.com\/(?:[a-z]{2}(?:-[a-z]{2})?\/)?shop\/([a-z0-9_]+)(?:[/?#]|$)/i.exec(s);
    return m ? 'https://www.etsy.com/shop/' + m[1] : '';
  }
  const urlKey = v => cleanShopUrl(v).toLowerCase();
  function number(v) {
    if (typeof v === 'string') {
      const s = v.trim();
      if (!/^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(s)) return null;
      v = Number(s.replace(/,/g,''));
    }
    return Number.isFinite(v) ? v : null;
  }
  function counter(v) {
    const n = number(v);
    return Number.isSafeInteger(n) && n >= 0 ? n : null;
  }
  // All unzoned sheet timestamps are wall time in Korea, never browser local time.
  function ms(v, allowDateOnly = false) {
    if (v instanceof Date) return v.getTime();
    if (Number.isFinite(v) && v >= 0 && v < 2958466) return Date.UTC(1899,11,30) + Math.round(v * DAY) - KST;
    if (typeof v !== 'string') return NaN;
    const m = /^(\d{4})([-.])\s*(\d{1,2})\2\s*(\d{1,2})\.?(?:(?:[T]|\s+)(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?\s*(AM|PM)?\s*(KST|Z|[+-]\d{2}:\d{2})?)?$/i.exec(v.trim());
    if (!m || (!allowDateOnly && !m[6])) return NaN;
    const y = +m[1], month = +m[3], d = +m[4], midnight = Date.UTC(y,month-1,d);
    const checked = new Date(midnight);
    if (checked.getUTCFullYear() !== y || checked.getUTCMonth() !== month-1 || checked.getUTCDate() !== d) return NaN;
    const meridian = (m[5] || m[10] || '').toUpperCase();
    let h = +(m[6] || 0), min = +(m[7] || 0), s = +(m[8] || 0);
    if (h > 23 || min > 59 || s > 59 || (meridian && (h < 1 || h > 12)) || (m[5] && m[10])) return NaN;
    if (meridian) h = h % 12 + (/^(PM|오후)$/.test(meridian) ? 12 : 0);
    const zone = (m[11] || 'KST').toUpperCase();
    let offset = zone === 'KST' ? 540 : 0;
    if (/^[+-]/.test(zone)) {
      const zh = +zone.slice(1,3), zm = +zone.slice(4);
      if (zh > 23 || zm > 59) return NaN;
      offset = (zone[0] === '-' ? -1 : 1) * (60 * zh + zm);
    }
    return midnight + ((h*60 + min - offset)*60 + s)*1000 + +((m[9] || '').slice(0,3).padEnd(3,'0'));
  }
  function stamp(t) {
    if (!Number.isFinite(t) || !Number.isFinite(new Date(t + KST).getTime())) return '';
    return new Date(t + KST).toISOString().replace('T',' ').replace(/(?:\.000)?Z$/, ' KST');
  }
  const normalizeDay = v => stamp(ms(v,true)).slice(0,10);

  // Verify identities before interpreting cells; Sheets omits trailing empty cells.
  function table(values, schema, name, registry = false) {
    if (!Array.isArray(values) || !Array.isArray(values[0]) || !values.every(Array.isArray)) fail(name + ': 시트 범위를 읽을 수 없습니다.');
    const h = values[0].map(v => text(v).trim());
    const order = schema.map((label,i) => {
      const pos = h.indexOf(label);
      if (pos < 0 || pos !== h.lastIndexOf(label) || (registry && pos !== i)) fail(name + ': 헤더를 확인하세요 (' + label + ').');
      return pos;
    });
    return {h, rows: values.slice(1).filter(row => row.some(nonempty)).map(row => registry ? row : order.map(i => row[i]))};
  }
  function readShops(reg) {
    const ids = new Set(), active = new Map(), columns = [], indexes = [], names = new Set();
    let last = -Infinity;
    reg.h.slice(15).forEach((name,j) => {
      if (!name) return;
      if (unsafe(name) || names.has(name)) fail('샵목록: 중복 또는 예약된 추가 열 이름입니다.');
      names.add(name);
      const i = 15+j, vals = reg.rows.map(r => r[i]).filter(nonempty);
      const numeric = vals.length && vals.every(Number.isFinite);
      const dates = vals.length && vals.every(v => (typeof v === 'string' || v instanceof Date) && normalizeDay(v));
      columns.push({key:'custom:'+name,name,type:numeric?'number':dates?'date':'text'});
      indexes.push(i);
    });
    for (const r of reg.rows) {
      const id = text(r[0]);
      if (!id.trim() || unsafe(id) || ids.has(id) || !text(r[6]).trim()) fail('샵목록: 샵 ID 또는 추적 상태를 확인하세요.');
      ids.add(id);
      if (r[6] !== '활성') continue;
      const url = cleanShopUrl(r[2]);
      if (!url || !text(r[1]).trim()) fail('샵목록: 활성 샵의 이름 또는 URL을 확인하세요.');
      const t = ms(r[12]), custom = dict();
      columns.forEach(({key,type},j) => {
        const value = r[indexes[j]];
        custom[key] = type === 'number' ? number(value) : type === 'date' ? normalizeDay(value) || null : value ?? '';
      });
      if (Number.isFinite(t)) last = Math.max(t,last);
      active.set(id,{id,name:text(r[1]),url,country:text(r[3]),visibility:text(r[4]),total:number(r[10]),items:number(r[5]),
        favorites:number(r[11]),memo:text(r[9]),registered:stamp(ms(r[7])),source:text(r[13]),time:stamp(t),status:text(r[14]),custom});
    }
    return {ids, active, columns, last};
  }
  function readDaily(rows, active) {
    const latest = new Map(), history = dict(), sources = dict();
    for (const r of rows) {
      const id = text(r[0]), shop = active.get(id);
      if (!shop || r[9] !== '일별' || ['실패','집계제외'].includes(r[11]) || urlKey(r[10]) !== urlKey(shop.url)) continue;
      const day = normalizeDay(r[1]), t = ms(r[2]);
      if (!day || !Number.isFinite(t)) fail('일별관측: 확정 행의 날짜와 실제 관측 시각을 확인하세요.');
      const key = JSON.stringify([id,day]), old = latest.get(key);
      if (!old || t >= old.t) latest.set(key,{id,day,t,r});
    }
    for (const {id,day,r} of latest.values()) {
      const record = {sold:number(r[4]),favorites:number(r[8]),listings:number(r[6])};
      if (Object.values(record).every(v => v === null)) continue;
      (history[id] ||= dict())[day] = record;
      const memo = text(r[12]), prefix = '시간추적: ';
      (sources[id] ||= dict())[day] = memo.startsWith(prefix) ? memo.slice(prefix.length) : '기존 일별관측';
    }
    return {history,sources};
  }
  function readIntervals(rows, active) {
    const picked = new Map(), history = dict();
    let latestDay = '';
    for (const r of rows) {
      const id = text(r[0]), shop = active.get(id), t = ms(r[3]);
      if (!shop || !Number.isFinite(t) || (t + KST) % (12 * HOUR) !== 0) continue;
      const am = (t + KST) % DAY !== 0, day = stamp(am ? t : t-1).slice(0,10);
      const at = ms(r[4]), url = urlKey(r[12]), metrics = [counter(r[5]),counter(r[7]),counter(r[9])];
      const ok = normalizeDay(r[1]) === day && r[2] === (am?'00~12시':'12~24시') && url === urlKey(shop.url) &&
        okStatus.has(text(r[13]).trim()) && Math.abs(at-t) <= HOUR && metrics.some(v => v !== null);
      const seed = r[14] === '기준값' || r[15] === '기존 자정 관측 · 기준값';
      if (!picked.has(id)) picked.set(id,new Map());
      const slots = picked.get(id), old = slots.get(t), rank = Number.isFinite(at) ? at : -Infinity;
      // Pick first, validate second: an invalid latest retry must not expose stale sales.
      if (!old || rank >= old.rank) slots.set(t,{id,t,at,rank,day,side:am?'am':'pm',ok,url,seed,metrics});
    }
    for (const [id, slots] of picked) for (const x of slots.values()) {
      if (!x.ok) continue;
      const prev = slots.get(x.t - 12*HOUR), paired = !x.seed && prev?.ok && prev.url === x.url && x.at > prev.at;
      const [sold,listings,favorites] = x.metrics.map((n,j) => paired && n !== null && prev.metrics[j] !== null ? n-prev.metrics[j] : null);
      const valid = [sold,listings,favorites].filter(n => n !== null).length;
      const status = x.seed ? '기준값' : !paired ? '이전 구간 미확인' : sold < 0 ? 'SOLD 정정' : valid === 3 ? '구간확인' : valid ? '구간 부분확인' : '비교 수치 없음';
      const days = history[id] ||= dict();
      (days[x.day] ||= dict())[x.side] = {sold:sold < 0 ? null : sold, favorites,listings,observedAt:stamp(x.at),
        previousObservedAt:paired ? stamp(prev.at) : null,scheduledAt:stamp(x.t),elapsedHours:paired ? (x.at-prev.at)/HOUR : null,status};
      if (x.day > latestDay) latestDay = x.day;
    }
    return {history,latestDay};
  }
  function buildDashboardData(input, options = {}) {
    const reg = table(input?.registry, registryHeader, '샵목록', true);
    const daily = table(input?.daily, dailyHeader, '일별관측');
    const raw = table(input?.intervals, intervalHeader, '구간별관측');
    const settings = table(input?.settings, ['항목','값'], '설정안내');
    const anchors = settings.rows.filter(r => r[0] === '기준일');
    const anchor = anchors.length === 1 && normalizeDay(anchors[0][1]);
    if (!anchor) fail('설정안내: 기준일을 읽을 수 없습니다.');
    const {ids,active,columns,last} = readShops(reg);
    if (!ids.size && (daily.rows.length || raw.rows.length)) fail('샵목록: 관측 기록이 있으나 샵 목록을 읽지 못했습니다.');
    const d = readDaily(daily.rows,active), intervals = readIntervals(raw.rows,active);
    const now = options.now ?? Date.now(), fetchedAt = stamp(typeof now === 'number' ? now : ms(now));
    if (!fetchedAt) fail('읽은 시각을 확인하세요.');
    return {anchorDay:anchor > intervals.latestDay ? anchor : intervals.latestDay,collectedAt:stamp(last),fetchedAt,
      shops:Array.from(active.values()),customColumns:columns,history:d.history,intervalHistory:intervals.history,
      dailySources:d.sources,pinned:[],sheetUrl:text(options.sheetUrl)};
  }
  return {buildDashboardData,normalizeDay,cleanShopUrl};
});
