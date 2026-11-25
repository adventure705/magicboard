/* Private Sheets connection. Tokens and sheet data stay in this tab's memory. */
(function(root){'use strict';
  const SCOPE='https://www.googleapis.com/auth/spreadsheets';
  const BASE='https://sheets.googleapis.com/v4/spreadsheets/';
  const CACHE_VERSION='1';
  let gisLoad;
  const VIEW_NAMES=['판매대시보드','즐겨찾기대시보드'];
  function creationViewPlan(viewRows,id){
    if(!/^shop-[A-Za-z0-9_-]{1,100}$/.test(id))throw new Error('새 샵 ID를 확인하세요.');
    const data=[],checks=[];
    function column(n){let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s}
    VIEW_NAMES.forEach((name,index)=>{
      const rows=viewRows[index],header=rows?.[0],idCol=header?.indexOf('샵 ID');
      if(!Array.isArray(header)||header[0]!=='샵명'||idCol<1||header.lastIndexOf('샵 ID')!==idCol)throw new Error(name+': 샵 ID 열을 확인하세요.');
      const positions=[];for(let i=1;i<=200;i++)if(String(rows[i]?.[idCol]||'')===id)positions.push(i);
      if(positions.length>1)throw new Error(name+': 샵 ID가 중복되었습니다.');
      let at=positions[0];
      if(at===undefined){for(let i=1;i<=200;i++){if((rows[i]||[]).every(v=>v===null||v===undefined||v==='')){at=i;break}}
        if(at===undefined)throw new Error(name+': 빈 대시보드 ID 공간이 없습니다(200행).');
        const row=at+2,ref='$'+column(idCol)+row;
        const formula='=IF('+ref+'="","",IF(XLOOKUP('+ref+',\'샵목록\'!$A$2:$A,\'샵목록\'!$G$2:$G,"")="활성",XLOOKUP('+ref+',\'샵목록\'!$A$2:$A,\'샵목록\'!$B$2:$B,""),""))';
        data.push({range:"'"+name+"'!"+column(idCol)+row,values:[[id]]},{range:"'"+name+"'!A"+row,values:[[formula]]});
      }
      checks.push({index,rowIndex:at,idCol,id});
    });return {data,checks};
  }

  function loadGis(){
    if(root.google?.accounts?.oauth2)return Promise.resolve();
    if(gisLoad)return gisLoad;
    gisLoad=new Promise((resolve,reject)=>{
      const script=document.createElement('script');let done=false;
      const timer=setTimeout(()=>finish(new Error('로그인 도구를 불러오지 못했습니다. 새로고침해 주세요.')),20000);
      function finish(error){if(done)return;done=true;clearTimeout(timer);error?reject(error):resolve()}
      script.src='https://accounts.google.com/gsi/client';script.async=true;script.defer=true;
      script.onload=()=>finish(root.google?.accounts?.oauth2?null:new Error('Google 로그인 도구를 확인하세요.'));
      script.onerror=()=>finish(new Error('Google 로그인 도구를 불러오지 못했습니다.'));
      document.head.appendChild(script);
    });
    return gisLoad;
  }
  function create({config={},onData=()=>{},onClear=()=>{},onState=()=>{},shouldAutoRefresh=()=>true}={}){
    const clientId=String(config.googleClientId||'').trim(),sid=String(config.spreadsheetId||'').trim();
    const interval=Math.min(3600,Math.max(30,Number(config.refreshSeconds)||60))*1000;
    // Deliberately per-tab: a private sheet token must never survive a browser restart.
    const cacheKey='magic-board.google-token.'+CACHE_VERSION+'.'+encodeURIComponent(clientId)+'.'+encodeURIComponent(sid)+'.'+encodeURIComponent(SCOPE);
    let token='',expiresAt=0,grantedScope='',tokenClient,ready=false,busy=false,verified=false,writeGranted=false,epoch=0;
    let expiryTimer,pollTimer,requestController,probeState=null,refreshCount=0;
    const fullRanges=["'샵목록'","'일별관측'!A1:M","'구간별관측'!A1:Q","'설정안내'!A1:B"];
    const sheetUrl='https://docs.google.com/spreadsheets/d/'+encodeURIComponent(sid)+'/edit';
    function valid(){return !!token&&Date.now()<expiresAt}
    function cacheDrop(){try{root.sessionStorage?.removeItem(cacheKey)}catch{}}
    function cacheRead(){
      try{
        const value=root.sessionStorage?.getItem(cacheKey);if(!value)return null;
        const saved=JSON.parse(value);
        if(!saved||typeof saved!=='object'||typeof saved.accessToken!=='string'||!saved.accessToken||typeof saved.expiresAt!=='number'||!Number.isFinite(saved.expiresAt)||typeof saved.scope!=='string'||saved.expiresAt<=Date.now()||saved.expiresAt>Date.now()+86400000){cacheDrop();return null}
        return saved;
      }catch{cacheDrop();return null}
    }
    function cacheSave(){try{root.sessionStorage?.setItem(cacheKey,JSON.stringify({accessToken:token,expiresAt,scope:grantedScope}))}catch{}}
    function armExpiry(){clearTimeout(expiryTimer);expiryTimer=setTimeout(expired,Math.max(0,expiresAt-Date.now()))}
    function autoAllowed(){try{return !!shouldAutoRefresh()}catch{return false}}
    function state(phase,text,error=false){onState({phase,text,error,authenticated:verified,canLogin:ready&&!busy,canRefresh:valid()&&!busy,hasWritePermission:verified&&valid()&&writeGranted,canWrite:verified&&valid()&&writeGranted&&!busy})}
    function clear({dropCache=true}={}){
      epoch++;requestController?.abort();requestController=null;token='';expiresAt=0;grantedScope='';verified=false;writeGranted=false;busy=false;probeState=null;refreshCount=0;
      clearTimeout(expiryTimer);if(dropCache)cacheDrop();onClear();
    }
    function expired(){clear();state('expired','로그인이 만료됐습니다. 다시 로그인해 주세요.',true)}
    function probeRanges(data){
      const ds=Math.max(2,data.daily.length-200),ts=Math.max(2,data.intervals.length-200);
      return {ranges:["'샵목록'","'설정안내'!A1:B",`'일별관측'!A${ds}:M`,`'구간별관측'!A${ts}:Q`],fingerprint:JSON.stringify([data.registry,data.settings,data.daily.slice(ds-1),data.intervals.slice(ts-1)])};
    }
    async function api(suffix,{method='GET',body,signal}={}){
      const headers={Authorization:'Bearer '+token};if(body)headers['Content-Type']='application/json';
      const response=await fetch(BASE+encodeURIComponent(sid)+suffix,{method,headers,body:body?JSON.stringify(body):undefined,credentials:'omit',cache:'no-store',signal});
      if(!response.ok){const error=new Error(response.status===403?'시트 편집 권한을 확인하고 Google로 다시 로그인해 주세요.':response.status===401?'로그인이 만료됐습니다. 다시 로그인해 주세요.':response.status===429?'조회가 일시적으로 많습니다. 잠시 후 다시 시도하세요.':'시트 연결을 확인하고 다시 시도해 주세요.');error.status=response.status;throw error}
      return response.json();
    }
    async function batch(ranges,signal){
      const params=new URLSearchParams({majorDimension:'ROWS',valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'FORMATTED_STRING'});ranges.forEach(range=>params.append('ranges',range));
      const body=await api('/values:batchGet?'+params,{signal});
      if(!Array.isArray(body.valueRanges)||body.valueRanges.length!==ranges.length)throw new Error('시트 응답이 불완전합니다. 다시 시도해 주세요.');
      return body.valueRanges.map(range=>Array.isArray(range.values)?range.values:[]);
    }
    async function readFull(signal){
      const rows=await batch(fullRanges,signal);const source={registry:rows[0],daily:rows[1],intervals:rows[2],settings:rows[3]};
      const data=root.MagicBoardSheets.buildDashboardData(source,{now:new Date(),sheetUrl});
      return {data,source,probe:probeRanges(source)};
    }
    function apply(result){onData(result.data);verified=true;probeState=result.probe;refreshCount++}
    async function refresh({auto=false,force=false}={}){
      if(busy||(auto&&!autoAllowed()))return;
      if(!valid()){if(token)expired();else state('sign-in-required','시트 권한이 있는 Google 계정으로 로그인하세요.');return}
      const activeEpoch=epoch;busy=true;requestController=new AbortController();const timeout=setTimeout(()=>{if(activeEpoch===epoch)requestController?.abort()},25000);
      state('loading',verified?'시트 확인 중':'시트 불러오는 중');
      try{
        if(auto&&!force&&verified&&probeState&&refreshCount%10!==9){const rows=await batch(probeState.ranges,requestController.signal);if(activeEpoch!==epoch)return;if(JSON.stringify(rows)===probeState.fingerprint){refreshCount++;busy=false;state('connected','시트 연결됨');return}}
        const result=await readFull(requestController.signal);if(activeEpoch!==epoch)return;if(!valid()){expired();return}apply(result);busy=false;state('connected','시트 연결됨');
      }catch(error){
        if(activeEpoch!==epoch)return;if(error.name==='AbortError')error=new Error('시트 응답이 지연되고 있습니다. 다시 시도해 주세요.');
        if(error.status===401){expired();return}
        if(error.status===403){clear();state('error',error.message||'시트에 연결하지 못했습니다.',true);return}
        // A restored token is not authenticated until its read succeeds, but a transient
        // failure must not discard a valid, per-tab token.
        if(!verified){busy=false;state('error',error.message||'시트에 연결하지 못했습니다.',true);return}
        busy=false;state('error','연결 오류 · 마지막으로 읽은 데이터입니다. 새로고침해 주세요.',true);
      }finally{clearTimeout(timeout);if(activeEpoch===epoch){requestController=null;if(busy){busy=false;state(verified?'connected':'sign-in-required',verified?'시트 연결됨':'다시 로그인해 주세요.')}}}
    }
    async function mutate(request){
      if(busy)throw new Error('다른 시트 작업이 진행 중입니다. 잠시 후 다시 시도하세요.');
      if(!valid()||!verified||!writeGranted)throw new Error('편집 권한으로 Google에 다시 로그인해 주세요.');
      if(!root.MagicBoardWrites?.plan||!root.MagicBoardWrites?.verify)throw new Error('편집 도구를 불러오지 못했습니다. 페이지를 새로고침하세요.');
      const activeEpoch=epoch;busy=true;requestController=new AbortController();let posted=false;
      const timeout=setTimeout(()=>{if(activeEpoch===epoch)requestController?.abort()},25000);state('saving','시트 저장 중');
      function current(){if(activeEpoch!==epoch||!valid())throw new Error('로그인이 만료됐습니다. 다시 로그인하고 저장 결과를 확인하세요.')}
      try{
        // Re-read before planning; only allowed registry cells can be written.
        const before=await readFull(requestController.signal);current();
        const createMeta=request.action==='create'?{id:'shop-'+root.crypto.randomUUID(),registeredAt:new Date(Date.now()+9*3600000).toISOString().slice(0,19).replace('T',' ')+' KST'}:undefined;
        const plan=root.MagicBoardWrites.plan(before.source.registry,request,createMeta);
        const viewRanges=VIEW_NAMES.map(name=>"'"+name+"'!2:202");
        if(plan.append)creationViewPlan(await batch(viewRanges,requestController.signal),plan.createdId);
        if(request.action==='patch'&&Object.hasOwn(request.values||{},'url')){
          const old=before.data.shops.find(s=>s.id===request.id);
          if(old&&root.MagicBoardSheets.cleanShopUrl(old.url).toLowerCase()!==root.MagicBoardSheets.cleanShopUrl(request.values.url).toLowerCase()){
            const headers=before.source.daily[0],idCol=headers.indexOf('샵 ID'),kindCol=headers.indexOf('기록 종류'),urlCol=headers.indexOf('출처 URL');
            if(before.source.daily.slice(1).some(row=>row[idCol]===request.id&&row[kindCol]==='일별'&&!row[urlCol]))throw new Error('출처가 없는 과거 기록이 있어 이 주소 변경은 기존 편집 페이지에서 처리해 주세요.');
          }
        }
        if(plan.append){posted=true;await api('/values/'+encodeURIComponent(plan.append.range)+':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS',{method:'POST',signal:requestController.signal,body:{majorDimension:'ROWS',values:plan.append.values}});current()}
        if(plan.data.length){posted=true;await api('/values:batchUpdate',{method:'POST',signal:requestController.signal,body:{valueInputOption:'RAW',data:plan.data}});current()}
        const after=await readFull(requestController.signal);current();root.MagicBoardWrites.verify(after.source.registry,plan);
        if(plan.append){
          const views=creationViewPlan(await batch(viewRanges,requestController.signal),plan.createdId);current();
          // Only generated UUID strings and fixed lookup formulas use USER_ENTERED. User strings stay RAW.
          if(views.data.length)await api('/values:batchUpdate',{method:'POST',signal:requestController.signal,body:{valueInputOption:'USER_ENTERED',data:views.data}});
          current();const savedViews=await batch(viewRanges,requestController.signal);current();
          if(views.checks.some(c=>savedViews[c.index]?.[c.rowIndex]?.[c.idCol]!==c.id))throw new Error('샵 등록 후 시트 뷰 연결을 확인하지 못했습니다.');
        }
        apply(after);busy=false;state('connected','시트 저장·확인 완료');return {data:after.data,createdId:plan.createdId};
      }catch(error){
        const message=posted?'저장 요청 후 결과를 확정하지 못했습니다. 새로고침으로 현재 값을 확인하세요. '+(error.status===403?'편집 권한으로 다시 로그인해 주세요.':''):(error.name==='AbortError'?'시트 응답이 지연됐습니다. 새로고침 후 다시 시도하세요.':error.message||'시트에 저장하지 못했습니다.');
        if(activeEpoch===epoch){busy=false;if(error.status===401)expired();else state('error',message,true)}
        throw new Error(message);
      }finally{clearTimeout(timeout);if(activeEpoch===epoch){requestController=null;busy=false}}
    }
    function login(){
      if(!ready||busy)return;
      clear();const loginEpoch=epoch;busy=true;state('loading','Google 로그인 중');
      tokenClient.callback=response=>{
        if(loginEpoch!==epoch)return;
        if(response.error||!response.access_token){busy=false;state('error','Google 로그인을 완료하지 못했습니다. 다시 시도해 주세요.',true);return}
        token=response.access_token;grantedScope=String(response.scope||'');writeGranted=grantedScope.split(/\s+/).includes(SCOPE);
        const rawSeconds=Number(response.expires_in),seconds=Number.isFinite(rawSeconds)&&rawSeconds>0?Math.min(86400,rawSeconds):0;expiresAt=Date.now()+seconds*1000;cacheSave();armExpiry();busy=false;refresh({force:true});
      };
      try{tokenClient.requestAccessToken({prompt:'select_account'})}catch{busy=false;state('error','로그인 창을 열지 못했습니다. 팝업을 허용하고 다시 시도하세요.',true)}
    }
    function logout(){clear();state('sign-in-required','로그아웃했습니다. 다른 계정으로 다시 로그인할 수 있습니다.')}
    async function init(){
      clear({dropCache:false});state('loading','연결 준비 중');
      if(!/^[a-zA-Z0-9_-]{20,}$/.test(sid)||!clientId.endsWith('.apps.googleusercontent.com')){state('setup-required','Google 로그인 설정이 필요합니다. config.js와 배포 안내를 확인하세요.',true);return}
      if(location.protocol!=='https:'&&!['localhost','127.0.0.1'].includes(location.hostname)){state('setup-required','HTTPS 웹 주소나 localhost에서 열어주세요. 파일을 직접 열면 로그인할 수 없습니다.',true);return}
      try{
        await loadGis();tokenClient=root.google.accounts.oauth2.initTokenClient({client_id:clientId,scope:SCOPE,include_granted_scopes:false,callback:()=>{},error_callback:()=>{busy=false;state('error','로그인 창이 닫혔거나 차단되었습니다. 다시 로그인해 주세요.',true)}});
        ready=true;const saved=cacheRead();
        if(saved){token=saved.accessToken;expiresAt=saved.expiresAt;grantedScope=saved.scope;writeGranted=grantedScope.split(/\s+/).includes(SCOPE);armExpiry();state('loading','저장된 로그인으로 시트 확인 중');refresh({force:true})}
        else state('sign-in-required','시트 접근 권한이 있는 Google 계정으로 로그인하세요.');
        clearInterval(pollTimer);pollTimer=setInterval(()=>{if(!document.hidden&&!busy&&valid()&&autoAllowed())refresh({auto:true})},interval);
        document.addEventListener('visibilitychange',()=>{if(!document.hidden&&!busy){if(valid()&&autoAllowed())refresh({auto:true,force:true});else if(token&&!valid())expired()}});
      }catch(error){ready=false;state('error',error.message||'Google 로그인 도구를 불러오지 못했습니다.',true)}
    }
    return Object.freeze({init,login,logout,refresh:()=>refresh({force:true}),mutate});
  }
  root.MagicBoardRemote=Object.freeze({create,creationViewPlan});
})(globalThis);
