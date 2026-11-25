/* Pure write planner for header-inclusive Sheets values arrays.  Callers must
 * send plan.data with valueInputOption: 'RAW'; this module does no I/O. */
(function (root, factory) {
  'use strict';
  const api = Object.freeze(factory());
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.MagicBoardWrites = api;
})(globalThis, function () {
  'use strict';
  const CORE = ['샵 ID','샵명','샵 URL','국가','SOLD 공개','활성 리스팅','추적 상태','신규 문서 등록 시각(KST)','기존 등록 시각(KST)','메모','마지막 누적 판매','마지막 누적 즐겨찾기','마지막 관측 시각(KST)','출처','기록 구분'];
  const COL = Object.freeze({name:2,url:3,country:4,visibility:5,memo:10});
  const VISIBILITY = new Set(['공개','비공개','확인 필요','확인 불가']);
  const UNSAFE = new Set(['__proto__','prototype','constructor']);
  const fail = message => { throw new Error(message); };
  const own = (o,k) => Object.prototype.hasOwnProperty.call(o,k);
  const plain = v => v !== null && typeof v === 'object' && !Array.isArray(v) && (Object.getPrototypeOf(v) === Object.prototype || Object.getPrototypeOf(v) === null);
  const safeKey = k => typeof k === 'string' && !UNSAFE.has(k);
  const text = v => v == null ? '' : String(v);
  const nonempty = v => v != null && v !== '';
  const letter = n => { let s=''; while(n){ const x=(n-1)%26; s=String.fromCharCode(65+x)+s; n=(n-1-x)/26; } return s; };
  const range = (row,col) => "'샵목록'!" + letter(col) + row;
  function cleanShopUrl(value) {
    let s = text(value).trim().replace(/^http:\/\//i,'https://');
    if (!/^https?:\/\//i.test(s)) s='https://'+s;
    // Deliberately accept only a public shop root.  Everything after its slug
    // is discarded, so tracking URLs cannot make the stored canonical URL long.
    const m=/^https:\/\/(?:www\.|m\.)?etsy\.com\/(?:[a-z]{2}(?:-[a-z]{2})?\/)?shop\/([a-z0-9_]+)(?:[/?#]|$)/i.exec(s);
    return m ? 'https://www.etsy.com/shop/'+m[1] : '';
  }
  const urlKey = v => cleanShopUrl(v).toLowerCase();
  // Mirrors sheets-data's KST date interpretation for inferred custom columns.
  function normalizeDay(v) {
    let t;
    if (v instanceof Date) t=v.getTime();
    else if (typeof v==='string') {
      const s=v.trim(), m=/^(\d{4})([-.])\s*(\d{1,2})\2\s*(\d{1,2})\.?(?:(?:T|\s+)(?:오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*(?:AM|PM)?\s*(KST|Z|[+-]\d{2}:\d{2})?)?$/i.exec(s);
      if (!m) return '';
      const y=+m[1], mo=+m[3], d=+m[4], base=Date.UTC(y,mo-1,d), check=new Date(base);
      if(check.getUTCFullYear()!==y||check.getUTCMonth()!==mo-1||check.getUTCDate()!==d) return '';
      if (!m[5]) t=base-9*3600000;
      else { const h=+m[5], min=+m[6], sec=+(m[7]||0), z=(m[8]||'KST').toUpperCase(); if(h>23||min>59||sec>59) return ''; let offset=z==='KST'?540:0; if(/^[+-]/.test(z)){const zh=+z.slice(1,3),zm=+z.slice(4);if(zh>23||zm>59)return '';offset=(z[0]==='-'?-1:1)*(60*zh+zm);} t=base+((h*60+min-offset)*60+sec)*1000; }
    } else return '';
    return Number.isFinite(t) ? new Date(t+9*3600000).toISOString().slice(0,10) : '';
  }
  function day(v) {
    if (typeof v !== 'string' || !/^\d{4}-\d\d-\d\d$/.test(v)) return '';
    const d=new Date(v+'T00:00:00Z'); return d.toISOString().slice(0,10)===v ? v : '';
  }
  function decimal(v) {
    if (typeof v==='string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(v.trim())) v=Number(v);
    return typeof v==='number' && Number.isFinite(v) ? v : null;
  }
  function rawValue(v) { return v == null ? '' : v; }
  function customValue(v,type) {
    if (type==='number') return decimal(v);
    if (type==='date') return normalizeDay(v) || null;
    return v instanceof Date && Number.isFinite(v.getTime()) ? normalizeDay(v) : text(v);
  }
  function readRegistry(values) {
    if (!Array.isArray(values) || !Array.isArray(values[0]) || !values.every(Array.isArray)) fail('샵목록: 시트 범위를 읽을 수 없습니다.');
    const h=values[0].map(v=>text(v).trim());
    if (h.length<CORE.length || CORE.some((x,i)=>h[i]!==x) || new Set(h.slice(0,CORE.length)).size!==CORE.length) fail('샵목록: A:O 핵심 헤더를 확인하세요.');
    const columns=new Map(), headerNames=new Set();
    for(let i=15;i<h.length;i++) {
      const name=h[i]; if (!name) continue;
      if (!safeKey(name) || headerNames.has(name)) fail('샵목록: 중복 또는 예약된 추가 열 이름입니다.');
      headerNames.add(name);
      const vals=values.slice(1).map(r=>r[i]).filter(nonempty);
      const type=vals.length && vals.every(Number.isFinite) ? 'number' : vals.length && vals.every(v=>(typeof v==='string'||v instanceof Date)&&normalizeDay(v)) ? 'date' : 'text';
      columns.set('custom:'+name,{col:i+1,type,name});
    }
    const ids=new Map();
    values.slice(1).forEach((r,ix)=>{
      if (!r.some(nonempty)) return;
      const id=text(r[0]); const status=text(r[6]);
      if (!id || !safeKey(id) || ids.has(id) || !['활성','삭제'].includes(status)) fail('샵목록: 샵 ID 또는 추적 상태를 확인하세요.');
      ids.set(id,{row:ix+2,cells:r});
    });
    return {values,h,columns,ids};
  }
  function categoryColumn(m) {
    const end=Math.max(m.h.length,...m.values.map(r=>r.length));
    for(let col=16;col<=end;col++)if(!m.h[col-1]&&m.values.every(r=>!nonempty(r[col-1])))return col;
    return Math.max(16,end+1);
  }
  function inputText(v,max,trim) {
    if (v==null) return '';
    if (typeof v!=='string' || v.length>max || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(v)) fail('텍스트 형식이나 길이를 확인하세요.');
    return trim ? v.trim() : v;
  }
  function valueFor(key,v,column) {
    if (column?.type==='number') { if(v==null || (typeof v==='string'&&!v.trim())) return null; const n=decimal(v); if(n===null) fail('유한한 숫자를 입력하세요.'); return n; }
    if (column?.type==='date') { const s=inputText(v,32,true); if(!s) return null; if(!day(s)) fail('유효한 YYYY-MM-DD 날짜를 입력하세요.'); return s; }
    const s=inputText(v,key==='memo'||column ? 10000 : key==='url' ? 2048 : 200, !(key==='memo'||column));
    if (key==='url') return cleanShopUrl(s)||fail('유효한 Etsy 샵 URL을 입력하세요.');
    if (key==='name'&&!s) fail('샵명을 입력하세요.');
    if (key==='visibility'&&!VISIBILITY.has(s)) fail('SOLD 공개 상태를 확인하세요.');
    return s;
  }
  function current(entry,key,column) {
    const v=entry.cells[(column ? column.col : COL[key])-1];
    if (column) return customValue(v,column.type);
    if (key==='url') return cleanShopUrl(v);
    return text(v);
  }
  function equal(a,b) { return a===b || (Number.isNaN(a)&&Number.isNaN(b)); }
  function checkExpected(expected,key,old,column) {
    if (!own(expected,key)) fail('수정할 수 없는 항목입니다: '+key);
    const wanted=column ? customValue(expected[key],column.type) : key==='url' ? cleanShopUrl(expected[key]) : text(expected[key]);
    if (!equal(wanted,old)) fail('편집 충돌: '+key+' 값이 변경되었습니다. 새로고침 후 다시 시도하세요.');
  }
  function write(out,row,col,val) { out.push({range:range(row,col),values:[[rawValue(val)]]}); }
  function makePlan(data) { return Object.freeze({valueInputOption:'RAW',data:Object.freeze(data.map(x=>Object.freeze({range:x.range,values:Object.freeze([Object.freeze(x.values[0])])}))),expectedCells:Object.freeze(data.map(x=>Object.freeze({range:x.range,values:Object.freeze([Object.freeze(x.values[0])])}))),noOp:data.length===0}); }
  function createText(key,v) {
    // Create accepts a long public URL only so it can be reduced to its shop root.
    if (key==='url') {
      const raw=inputText(v,32768,true), canonical=cleanShopUrl(raw);
      if (!canonical || canonical.length>2048) fail('유효한 Etsy 샵 URL을 입력하세요.');
      return canonical;
    }
    return valueFor(key,v);
  }
  function registeredAt(v) {
    if (typeof v!=='string' || !/^\d{4}-\d\d-\d\d [0-2]\d:[0-5]\d:[0-5]\d KST$/.test(v)) fail('등록 시각을 확인하세요.');
    const m=/^(\d{4}-\d\d-\d\d) (\d\d):/.exec(v);
    if (+m[2]>23 || !day(m[1])) fail('등록 시각을 확인하세요.');
    return v;
  }
  function createPlan(m,request,meta) {
    const allowed=new Set(['action','shop']);
    if(Object.keys(request).some(k=>!allowed.has(k)||!safeKey(k)) || !plain(request.shop)) fail('생성 요청을 확인하세요.');
    const shopAllowed=new Set(['url','name','country','visibility','memo']);
    if(Object.keys(request.shop).some(k=>!shopAllowed.has(k)||!safeKey(k)) || !own(request.shop,'url')) fail('생성할 샵 정보를 확인하세요.');
    if(!plain(meta) || Object.keys(meta).length!==2 || Object.keys(meta).some(k=>!safeKey(k)||!['id','registeredAt'].includes(k))) fail('생성 메타데이터를 확인하세요.');
    let id;
    try { id=inputText(meta.id,256,true); } catch (_) { fail('생성 메타데이터를 확인하세요.'); }
    if(!id || !safeKey(id) || m.ids.has(id)) fail('생성 메타데이터를 확인하세요.');
    const url=createText('url',request.shop.url);
    for(const [,entry] of m.ids) if(urlKey(entry.cells[2])===urlKey(url)) fail('이미 등록된 URL입니다 (삭제된 샵 포함).');
    const name=createText('name',own(request.shop,'name') ? request.shop.name : url.split('/').pop());
    const country=createText('country',request.shop.country)||'확인 필요';
    const visibility=createText('visibility',request.shop.visibility||'확인 필요');
    const memo=createText('memo',request.shop.memo);
    const row=[id,name,url,country,visibility,'확인 필요','활성',registeredAt(meta.registeredAt),'',memo,'확인 필요','확인 필요','','','미수집'];
    const append=Object.freeze({range:"'샵목록'!A:O",values:Object.freeze([Object.freeze(row)])});
    return Object.freeze({valueInputOption:'RAW',data:Object.freeze([]),expectedCells:Object.freeze([]),noOp:false,append,createdId:id,canonicalUrl:url});
  }
  function plan(registryValues, request, meta) {
    const m=readRegistry(registryValues);
    if(request&&JSON.stringify(request).length>65536)fail('요청 크기를 확인하세요.');
    if (!plain(request) || !safeKey(request.action)) fail('요청 형식을 확인하세요.');
    if (request.action==='create') return createPlan(m,request,meta);
    if (request.action==='patch') {
      const allowed=new Set(['action','id','values','expected']);
      if(Object.keys(request).some(k=>!allowed.has(k)||!safeKey(k)) || typeof request.id!=='string' || !safeKey(request.id) || !plain(request.values)||!plain(request.expected)) fail('변경 값과 기존 값이 필요합니다.');
      const entry=m.ids.get(request.id); if(!entry) fail('존재하지 않는 샵 ID입니다. 새로고침하세요.');
      if(text(entry.cells[6])!=='활성') fail('비활성 샵은 수정할 수 없습니다.');
      const keys=Object.keys(request.values);
      if(!keys.length||keys.length>64||Object.keys(request.expected).length!==keys.length||keys.some(k=>!safeKey(k))) fail('변경 항목을 확인하세요.');
      const pending=[], category=m.columns.get('custom:카테고리');
      for(const key of keys) {
        let column=m.columns.get(key);
        if(!column && key==='custom:카테고리') column=category || {col:categoryColumn(m),type:'text',name:'카테고리',newHeader:!category};
        if ((!column&&!own(COL,key)) || !own(request.expected,key)) fail('수정할 수 없는 항목입니다: '+key);
        const old=column ? (column.newHeader ? '' : current(entry,key,column)) : current(entry,key);
        checkExpected(request.expected,key,old,column);
        const val=valueFor(key,request.values[key],column);
        if(!equal(old,val)) pending.push({key,column,col:column?.col||COL[key],val,old});
      }
      const url=pending.find(x=>x.key==='url'), identity=url && urlKey(url.old)!==urlKey(url.val);
      if(url) for(const [id,e] of m.ids) if(id!==request.id && urlKey(e.cells[2])===urlKey(url.val)) fail('이미 등록된 URL입니다 (삭제된 샵 포함).');
      const out=[];
      for(const x of pending) {
        if(x.column?.newHeader) write(out,1,x.col,'카테고리');
        if(!identity || x.col!==5) write(out,entry.row,x.col,x.val);
      }
      if(identity) {
        write(out,entry.row,5,'확인 필요'); write(out,entry.row,6,'확인 필요');
        write(out,entry.row,11,'확인 필요'); write(out,entry.row,12,'확인 필요'); write(out,entry.row,13,''); write(out,entry.row,14,''); write(out,entry.row,15,'미수집');
      }
      return makePlan(out);
    }
    if(request.action==='remove'||request.action==='restore') {
      const allowed=new Set(['action','ids']);
      if(Object.keys(request).some(k=>!allowed.has(k)||!safeKey(k)) || !Array.isArray(request.ids)||!request.ids.length||request.ids.length>200) fail('대상 샵 ID 목록을 확인하세요.');
      const seen=new Set(), out=[], wanted=request.action==='remove'?'활성':'삭제', next=request.action==='remove'?'삭제':'활성';
      for(const id of request.ids) { if(typeof id!=='string'||!safeKey(id)||seen.has(id)) fail('대상 샵 ID 목록을 확인하세요.'); seen.add(id); const entry=m.ids.get(id); if(!entry) fail('존재하지 않는 샵 ID입니다. 새로고침하세요.'); if(text(entry.cells[6])!==wanted) fail('샵 상태가 바뀌었습니다. 새로고침하세요.'); out.push({entry}); }
      return makePlan(out.map(x=>({range:range(x.entry.row,7),values:[[next]]})));
    }
    fail('요청 작업을 확인하세요.');
  }
  function verify(registryValues, result) {
    const m=readRegistry(registryValues);
    if(!plain(result)||!Array.isArray(result.expectedCells)) fail('쓰기 계획을 확인하세요.');
    if(own(result,'append') || own(result,'createdId') || own(result,'canonicalUrl')) {
      if(!result.append || !Array.isArray(result.data) || result.data.length || !Array.isArray(result.expectedCells) || result.expectedCells.length || typeof result.createdId!=='string' || !safeKey(result.createdId) || typeof result.canonicalUrl!=='string' || result.canonicalUrl!==cleanShopUrl(result.canonicalUrl) || result.append.range!=="'샵목록'!A:O" || !Array.isArray(result.append.values) || result.append.values.length!==1 || !Array.isArray(result.append.values[0]) || result.append.values[0].length!==15) fail('쓰기 계획을 확인하세요.');
      const expected=result.append.values[0], entry=m.ids.get(result.createdId);
      if(!entry || expected[0]!==result.createdId || expected[2]!==result.canonicalUrl || expected[6]!=='활성' || expected[5]!=='확인 필요' || expected[10]!=='확인 필요' || expected[11]!=='확인 필요' || expected[14]!=='미수집' || expected[8]!=='' || expected[12]!=='' || expected[13]!=='' || !expected[1] || !expected[2] || !expected[3] || !VISIBILITY.has(expected[4]) || !registeredAt(expected[7])) fail('쓰기 확인에 실패했습니다: 생성 행');
      if(entry.cells.slice(0,15).some((v,i)=>!equal(rawValue(v),rawValue(expected[i])))) fail('쓰기 확인에 실패했습니다: 생성 행');
      return true;
    }
    for(const cell of result.expectedCells) {
      if(!plain(cell)||typeof cell.range!=='string'||!Array.isArray(cell.values)||cell.values.length!==1||!Array.isArray(cell.values[0])||cell.values[0].length!==1) fail('쓰기 계획을 확인하세요.');
      const match=/^'샵목록'!([A-Z]+)([1-9]\d*)$/.exec(cell.range); if(!match) fail('쓰기 계획을 확인하세요.');
      let col=0; for(const ch of match[1]) col=col*26+ch.charCodeAt(0)-64;
      const got=registryValues[+match[2]-1]?.[col-1]; if(!equal(rawValue(got),rawValue(cell.values[0][0]))) fail('쓰기 확인에 실패했습니다: '+cell.range);
    }
    return true;
  }
  return {plan,verify,cleanShopUrl};
});
