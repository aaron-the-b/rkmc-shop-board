/* Shop Board live mode: Microsoft sign-in, SharePoint data, and sync.
   The board UI (app.js) keeps everything in one STATE object. This file builds that object from
   the RKMC Shop Board SharePoint site, then writes changes back whenever the UI calls save(). */
(function(){
'use strict';
const C=window.SHOPBOARD_CONFIG;
const GRAPH='https://graph.microsoft.com/v1.0';
const AUTH=`https://login.microsoftonline.com/${C.tenantId}/oauth2/v2.0`;
const SCOPES='openid profile offline_access User.Read Sites.ReadWrite.All Files.ReadWrite.All';
const TK='shopboard-token';
const gate=document.getElementById('liveGate');

/* ================= sign-in (authorization code + PKCE, no library) ================= */
const b64u=buf=>btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
const rand=n=>b64u(crypto.getRandomValues(new Uint8Array(n)));
function readTok(){try{return JSON.parse(sessionStorage.getItem(TK)||localStorage.getItem(TK)||'null')}catch(e){return null}}
function writeTok(t){const s=JSON.stringify(t);try{sessionStorage.setItem(TK,s);localStorage.setItem(TK,s)}catch(e){}}
function clearTok(){try{sessionStorage.removeItem(TK);localStorage.removeItem(TK)}catch(e){}}

async function authorizeUrl(){
  const verifier=rand(48),state=rand(16);
  const challenge=b64u(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(verifier)));
  try{localStorage.setItem('shopboard-pkce',JSON.stringify({verifier,state}))}catch(e){}
  const q=new URLSearchParams({client_id:C.clientId,response_type:'code',redirect_uri:C.redirectUri,response_mode:'query',
    scope:SCOPES,state,code_challenge:challenge,code_challenge_method:'S256',prompt:'select_account'});
  return `${AUTH}/authorize?${q}`;
}
async function tokenRequest(params){
  const r=await fetch(`${AUTH}/token`,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:new URLSearchParams({client_id:C.clientId,scope:SCOPES,...params})});
  const j=await r.json();
  if(!r.ok)throw Object.assign(new Error(j.error_description||j.error||'Sign-in failed'),{code:j.error});
  const t={access:j.access_token,refresh:j.refresh_token,exp:Date.now()+(j.expires_in-120)*1000};
  writeTok(t);return t;
}
async function redeem(code,state){
  let p=null;try{p=JSON.parse(localStorage.getItem('shopboard-pkce'))}catch(e){}
  if(!p||p.state!==state)throw new Error('Sign-in was interrupted. Try again.');
  localStorage.removeItem('shopboard-pkce');
  return tokenRequest({grant_type:'authorization_code',code,redirect_uri:C.redirectUri,code_verifier:p.verifier});
}
let tok=null;
async function token(){
  if(tok&&tok.exp>Date.now())return tok.access;
  if(tok&&tok.refresh){try{tok=await tokenRequest({grant_type:'refresh_token',refresh_token:tok.refresh});return tok.access}catch(e){}}
  clearTok();tok=null;showSignIn('Your sign-in expired. Sign in again to keep saving.');throw new Error('signed out');
}
async function signIn(){
  const url=await authorizeUrl();
  // Microsoft's sign-in page refuses to load inside a frame (a Teams website tab), so use a popup there.
  if(window.top!==window.self){
    const w=window.open(url,'shopboard-signin','width=520,height=680');
    if(!w){setGate('Allow pop-ups for this page, then click Sign in again.',true);return}
    window.addEventListener('storage',function on(e){if(e.key===TK&&e.newValue){window.removeEventListener('storage',on);location.reload()}});
    return;
  }
  location.assign(url);
}

/* ================= Graph helpers ================= */
async function g(path,opt={}){
  const url=path.startsWith('http')?path:GRAPH+path;
  for(let attempt=0;;attempt++){
    const r=await fetch(url,{...opt,headers:{Authorization:'Bearer '+await token(),...(opt.body&&!(opt.body instanceof Blob)&&!(opt.body instanceof File)?{'Content-Type':'application/json'}:{}),...(opt.headers||{})}});
    if((r.status===429||r.status===503)&&attempt<4){await new Promise(res=>setTimeout(res,(+r.headers.get('Retry-After')||2)*1000));continue}
    if(r.status===401&&attempt===0&&tok){tok.exp=0;continue}
    if(r.status===204)return null;
    const ct=r.headers.get('content-type')||'';
    const body=ct.includes('json')?await r.json():await r.text();
    if(!r.ok){const e=new Error((body&&body.error&&body.error.message)||('Request failed: '+r.status));e.status=r.status;e.body=body;throw e}
    return body;
  }
}
async function all(path){let out=[],next=path;while(next){const r=await g(next);out=out.concat(r.value||[]);next=r['@odata.nextLink']}return out}

/* ================= SharePoint lists ================= */
const LISTS=['Jobs','Skids','SerialRequirements','SerialRecords','ActivityLog','Groups','JobAccess','UsersRoles'];
const L={};   // name -> {id, cols: displayName -> {name,type,multi}, items: []}
let SITE=null,SITEURL='',DRIVE=null,PHOTODRIVE=null,UIL=null;

async function loadList(name){
  const meta=SITE.lists.find(l=>l.displayName===name);
  if(!meta){L[name]=null;return}
  const cols={};
  for(const c of await all(`/sites/${SITE.id}/lists/${meta.id}/columns`)){
    const type=c.boolean?'bool':c.number?'num':c.dateTime?'date':c.choice?'choice':c.personOrGroup?'person':'text';
    const multi=!!((c.choice&&c.choice.allowMultipleValues)||(c.personOrGroup&&c.personOrGroup.allowMultipleSelection));
    cols[c.displayName]={name:c.name,type,multi};if(!cols[c.name])cols[c.name]={name:c.name,type,multi};
  }
  const items=await all(`/sites/${SITE.id}/lists/${meta.id}/items?$expand=fields&$top=999`);
  L[name]={id:meta.id,cols,items,webUrl:meta.webUrl};
}
// read a field by the name people see in SharePoint (handles Title having been renamed on import)
function rd(list,item,disp){const c=L[list]&&L[list].cols[disp];const f=item.fields||{};const n=c?c.name:disp;
  if(c&&c.type==='person'){const ids=f[n+'LookupId'];return ids==null?[]:[].concat(ids).map(id=>(UIL&&UIL[id])||'').filter(Boolean)}
  const v=f[n];if(v==null)return c&&c.multi?[]:'';return v}
const txt=v=>Array.isArray(v)?v.join(';'):v==null?'':String(v);
const num=v=>v===''||v==null||isNaN(+v)?null:+v;
const yes=v=>v===true||/^(yes|true|1)$/i.test(String(v));
// write: convert values to what each column type expects; unknown columns are skipped
function wr(list,obj){const out={};const cols=L[list].cols;
  for(const [disp,v] of Object.entries(obj)){const c=cols[disp];if(!c)continue;
    if(c.type==='person'){const ids=[].concat(v||[]).map(e=>UIL&&UIL.byEmail[String(e).toLowerCase()]).filter(Boolean);
      if(c.multi){out[c.name+'LookupId@odata.type']='Collection(Edm.Int32)';out[c.name+'LookupId']=ids}else if(ids[0])out[c.name+'LookupId']=ids[0];continue}
    if(c.type==='bool'){out[c.name]=yes(v);continue}
    if(c.type==='num'){const n=num(v);if(n!=null)out[c.name]=n;continue}
    if(c.type==='date'){if(v)out[c.name]=new Date(v).toISOString();continue}
    if(c.type==='choice'&&c.multi){out[c.name+'@odata.type']='Collection(Edm.String)';out[c.name]=[].concat(v||[]).flatMap(x=>String(x).split(/[;,]/)).map(s=>s.trim()).filter(Boolean);continue}
    out[c.name]=v==null?'':(typeof v==='object'?JSON.stringify(v):String(v));
  }return out}
const itemPath=(list,id)=>`/sites/${SITE.id}/lists/${L[list].id}/items${id?'/'+id:''}`;
async function createItem(list,obj){const r=await g(itemPath(list),{method:'POST',body:JSON.stringify({fields:wr(list,obj)})});return r}
async function updateItem(list,id,obj){return g(itemPath(list,id)+'/fields',{method:'PATCH',body:JSON.stringify(wr(list,obj))})}
async function deleteItem(list,id){return g(itemPath(list,id),{method:'DELETE'})}

async function loadUserInfo(){
  // People columns store a site user id; this hidden list maps those ids to emails.
  try{const meta=await g(`/sites/${SITE.id}/lists/User Information List?$select=id`);
    const rows=await all(`/sites/${SITE.id}/lists/${meta.id}/items?$expand=fields($select=EMail,Title)&$top=999`);
    UIL={byEmail:{}};for(const r of rows){const e=(r.fields.EMail||'').toLowerCase();if(e){UIL[r.id]=e;UIL.byEmail[e]=+r.id}}
  }catch(e){UIL={byEmail:{}}}
}

/* ================= board configuration file (layouts, roles, teams) ================= */
const CFG_PATH=()=>`/sites/${SITE.id}/drive/root:/${C.configFolder}/board-config.json`;
async function loadConfig(){try{return await g(CFG_PATH()+':/content')}catch(e){if(e.status===404)return null;throw e}}
async function saveConfig(obj){return g(CFG_PATH()+':/content',{method:'PUT',body:JSON.stringify(obj),headers:{'Content-Type':'application/json'}})}

/* ================= build STATE from SharePoint ================= */
const PALETTE=['#0f6fe0','#9b3fd6','#e0337a','#3a9d23','#0a8fa3','#c96a00','#d0392b','#7a63c7','#2d8f9a','#77808c'];
const jid=n=>'J'+String(n).trim();
let ME=null,CONFIG=null,STATE=null,SNAP=null;
const LIVE=window.LIVE={parts:{},state:null};

function buildState(){
  const layout=CONFIG.layout;
  const S=JSON.parse(JSON.stringify(layout));
  S.jobs={};S.skids=[];S.partReqs={};S.users={};S.groups={};S.allJobs={users:[],groups:[]};
  const teams=CONFIG.teams||{};
  // people
  if(L.UsersRoles)for(const it of L.UsersRoles.items){
    const email=txt(rd('UsersRoles',it,'Email')).trim().toLowerCase();if(!email)continue;
    if(L.UsersRoles.cols.Active&&rd('UsersRoles',it,'Active')!==''&&!yes(rd('UsersRoles',it,'Active')))continue;
    const want=[].concat(rd('UsersRoles',it,'Roles')).flatMap(x=>String(x).split(/[;,\n]/)).map(s=>s.trim().toLowerCase()).filter(Boolean)[0]||'';
    const roleId=Object.keys(S.roles).find(id=>id===want||S.roles[id].label.toLowerCase()===want||(S.roles[id].short||'').toLowerCase()===want)||'tech';
    const company=txt(rd('UsersRoles',it,'Company')).trim();
    S.users[email]={name:txt(rd('UsersRoles',it,'DisplayName')).trim()||email,role:roleId,email,_sp:it.id,...(S.roles[roleId].company==='customer'?{company}:{})};
  }
  // first person to open an empty UsersRoles list becomes the admin; anyone else not listed can look but has no jobs
  if(!S.users[ME.email])S.users[ME.email]=Object.keys(S.users).length?{name:ME.name,role:'tech',email:ME.email,_guest:true}:{name:ME.name,role:'admin',email:ME.email,_new:true};
  S.defaultUser=ME.email;
  // groups
  if(L.Groups)for(const it of L.Groups.items){const name=txt(rd('Groups',it,'GroupName')).trim();if(!name)continue;
    const members=[].concat(rd('Groups',it,'MemberEmails')).flatMap(x=>String(x).split(/[;,\n]/)).map(s=>s.trim().toLowerCase()).filter(e=>S.users[e]);
    S.groups['g'+it.id]={name,members,_sp:it.id}}
  const groupByName=n=>Object.keys(S.groups).find(k=>S.groups[k].name.toLowerCase()===String(n).trim().toLowerCase());
  // jobs
  let ci=0;
  if(L.Jobs)for(const it of L.Jobs.items){const n=txt(rd('Jobs',it,'JobNumber')).trim();if(!n)continue;const id=jid(n);
    const pmEmail=txt(rd('Jobs',it,'ProjectManagerEmail')).trim().toLowerCase();const customer=txt(rd('Jobs',it,'Customer'))||'Customer';
    S.jobs[id]={num:n,customer,company:txt(rd('Jobs',it,'Company'))||customer,project:customer+' · '+n,pm:(S.users[pmEmail]||{}).name||pmEmail||'Not set',
      part:'',c:PALETTE[ci++%PALETTE.length],ship:txt(rd('Jobs',it,'ShipDate')).slice(0,10),ifs:'',acc:txt(rd('Jobs',it,'ACCFolderLink')),
      ncr:0,issues:0,teams:'',planner:[0,0],status:txt(rd('Jobs',it,'Status')),_sp:it.id,
      team:teams[id]||(S.users[pmEmail]?{pm:[pmEmail]}:{}),access:{users:[],groups:[]}}}
  // job access
  if(L.JobAccess)for(const it of L.JobAccess.items){const jn=txt(rd('JobAccess',it,'JobNumber')).trim();const type=txt(rd('JobAccess',it,'GrantType')).toLowerCase();
    const who=txt(rd('JobAccess',it,'Principal')).trim();const target=jn.toUpperCase()==='ALL'?S.allJobs:(S.jobs[jid(jn)]||{}).access;if(!target)continue;
    if(type==='group'){const gk=groupByName(who);if(gk)target.groups.push(gk)}else if(S.users[who.toLowerCase()])target.users.push(who.toLowerCase())}
  // serial requirements, grouped by part number
  if(L.SerialRequirements)for(const it of L.SerialRequirements.items){const pn=txt(rd('SerialRequirements',it,'PartNumber')).trim();if(!pn)continue;
    (S.partReqs[pn]=S.partReqs[pn]||[]).push({id:txt(rd('SerialRequirements',it,'ReqKey'))||('r'+it.id),item:txt(rd('SerialRequirements',it,'Description'))||txt(rd('SerialRequirements',it,'EquipmentTag')),
      model:txt(rd('SerialRequirements',it,'EquipmentTag'))||txt(rd('SerialRequirements',it,'ComponentPart')),qty:Math.max(1,num(rd('SerialRequirements',it,'Qty'))||1),
      photo:yes(rd('SerialRequirements',it,'PhotoRequired')),part:pn,rev:txt(rd('SerialRequirements',it,'PartRevision')),_sp:it.id})}
  // skids
  const stations=S.elements.filter(e=>e.kind==='station');
  const lineName=id=>{const l=S.elements.find(e=>e.id===id);return l?l.name:''};
  const unplaced={};
  if(L.Skids)for(const it of L.Skids.items){const tag=txt(rd('Skids',it,'SkidTag')).trim();if(!tag)continue;
    const jn=txt(rd('Skids',it,'JobNumber')).trim();const shop=(txt(rd('Skids',it,'Shop'))||'RKMC1').toLowerCase().replace(/\s/g,'');
    let bd={};try{bd=JSON.parse(rd('Skids',it,'BoardData')||'{}')||{}}catch(e){}
    const st=txt(rd('Skids',it,'Station')).trim(),ln=txt(rd('Skids',it,'Line')).trim();
    const station=stations.find(s=>s.id===st)||stations.find(s=>s.shop===shop&&(!ln||lineName(s.lineId).toLowerCase()===ln.toLowerCase())&&(s.name.toLowerCase()===st.toLowerCase()||s.name.replace(/\D/g,'')===st.replace(/\D/g,'')&&st!==''));
    const part=txt(rd('Skids',it,'PartNumber')).trim();
    const L_=(num(rd('Skids',it,'Length_in'))||240)/12,W_=(num(rd('Skids',it,'Width_in'))||96)/12;
    let x=num(rd('Skids',it,'X_ft')),y=num(rd('Skids',it,'Y_ft')),rot=num(rd('Skids',it,'Rotation_deg'))||0;
    if(x==null||y==null){
      if(station){const xs=station.pts.map(p=>p[0]),ys=station.pts.map(p=>p[1]);x=(Math.min(...xs)+Math.max(...xs))/2;y=(Math.min(...ys)+Math.max(...ys))/2;rot=station.orient==='v'?90:0}
      else{const vb=(window.PLANS[shop]||{vb:[0,0,100,100]}).vb;const n=unplaced[shop]=(unplaced[shop]||0)+1;x=vb[0]+12+L_/2;y=vb[1]+12+(n-1)*(W_+4)}}
    const sk={id:'sp'+it.id,tag,shop,jobId:jid(jn),part,so:txt(rd('Skids',it,'ShopOrderNo')),x,y,rot,L:L_,W:W_,stationId:station?station.id:null,
      pct:bd.pct||0,planPct:bd.planPct||0,est:bd.est||0,act:bd.act||0,qc:bd.qc||{},serials:{},uniqueReqs:bd.uniqueReqs||null,
      history:bd.history||[{at:new Date().toISOString().slice(0,10),txt:'Added from SharePoint'}],_sp:it.id};
    S.skids.push(sk);
    if(S.jobs[sk.jobId]&&!S.jobs[sk.jobId].part)S.jobs[sk.jobId].part=part;
    if(part&&!LIVE.parts[part])LIVE.parts[part]={desc:txt(rd('Skids',it,'Notes'))||part,L:L_,W:W_,est:bd.est||0,reqs:[]};
  }
  for(const pn of Object.keys(S.partReqs))if(!LIVE.parts[pn])LIVE.parts[pn]={desc:pn,L:20,W:8,est:0,reqs:[]};
  // serial records: RecordKey "tag|reqId|n" from the board, or SkidTag + EquipmentTag typed in SharePoint
  if(L.SerialRecords)for(const it of L.SerialRecords.items){
    const key=txt(rd('SerialRecords',it,'RecordKey'));const tag=txt(rd('SerialRecords',it,'SkidTag'))||key.split('|')[0];
    const sk=S.skids.find(k=>k.tag===tag);if(!sk)continue;
    const reqs=sk.uniqueReqs||S.partReqs[sk.part]||[];
    let slot=null;const p=key.split('|');
    if(p.length===3)slot=p[1]+':'+p[2];
    else{const et=txt(rd('SerialRecords',it,'EquipmentTag'));const r=reqs.find(r=>r.model===et);if(r){let q=0;while(sk.serials[r.id+':'+q])q++;slot=r.id+':'+q}}
    if(!slot)continue;
    const by=txt(rd('SerialRecords',it,'EnteredByEmail')).toLowerCase();
    sk.serials[slot]={sn:txt(rd('SerialRecords',it,'SerialNumber')),by:(S.users[by]||{}).name||by,at:txt(rd('SerialRecords',it,'EnteredAt')).slice(0,10),photo:txt(rd('SerialRecords',it,'PhotoLink')),_sp:it.id};
  }
  return S;
}

/* ================= sync: compare with the last saved snapshot, write the differences ================= */
const layoutPart=S=>({elements:S.elements,library:S.library,shops:S.shops,roles:S.roles,views:S.views,plan:S.plan,catalog:S.catalog,materials:S.materials});
const teamsPart=S=>Object.fromEntries(Object.entries(S.jobs).map(([k,j])=>[k,j.team||{}]));
const skidRow=(S,k)=>{const st=S.elements.find(e=>e.id===k.stationId);const ln=st&&S.elements.find(e=>e.id===st.lineId);
  return{SkidTag:k.tag,JobNumber:(S.jobs[k.jobId]||{}).num||k.jobId.replace(/^J/,''),PartNumber:k.part,ShopOrderNo:k.so,Shop:(S.shops[k.shop]||{}).name||k.shop.toUpperCase(),
    Line:ln?ln.name:'',Station:st?st.id:'',X_ft:Math.round(k.x*100)/100,Y_ft:Math.round(k.y*100)/100,Rotation_deg:k.rot||0,
    Length_in:Math.round(k.L*12*100)/100,Width_in:Math.round(k.W*12*100)/100,
    BoardData:JSON.stringify({pct:k.pct,planPct:k.planPct,est:k.est,act:k.act,qc:k.qc,uniqueReqs:k.uniqueReqs,history:(k.history||[]).slice(-50)})}};
const snapOf=S=>JSON.parse(JSON.stringify({layout:layoutPart(S),teams:teamsPart(S),skids:S.skids,partReqs:S.partReqs,users:S.users,groups:S.groups,allJobs:S.allJobs,
  access:Object.fromEntries(Object.entries(S.jobs).map(([k,j])=>[k,j.access]))}));
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
let busy=false,again=false,pending=false;
LIVE.dirty=()=>{pending=true;status('Saving…')};
LIVE.sync=async function(S){
  if(busy){again=true;return}busy=true;
  try{do{again=false;await syncOnce(S)}while(again);pending=false;status('Saved to SharePoint')}
  catch(e){console.error(e);status('Not saved: '+(e.status===403?'you do not have permission to change this':e.message),true)}
  finally{busy=false}
};
async function log(action,o){if(!L.ActivityLog)return;
  try{await createItem('ActivityLog',{EventID:crypto.randomUUID(),Timestamp:new Date().toISOString(),PersonEmail:ME.email,Source:'Board',Action:action,...o})}catch(e){console.warn('activity log',e)}}
async function syncOnce(S){
  const now=snapOf(S),was=SNAP;
  // 1. layout and job teams -> board-config.json
  if(!same(now.layout,was.layout)||!same(now.teams,was.teams)){
    CONFIG={...CONFIG,layout:now.layout,teams:now.teams,savedBy:ME.email,savedAt:new Date().toISOString()};await saveConfig(CONFIG);
    if(!same(now.layout,was.layout))log('Layout changed',{NewValue:'Board layout saved'});
  }
  // 2. skids
  const before=Object.fromEntries(was.skids.map(k=>[k.id,k]));
  for(const k of S.skids){const b=before[k.id];const row=skidRow(S,k);
    if(!b||!k._sp){if(!L.Skids)continue;const r=await createItem('Skids',{...row,Status:'In production'});k._sp=r.id;log('Skid added',{JobNumber:row.JobNumber,SkidTag:k.tag});continue}
    const old=skidRow({elements:was.layout.elements,shops:was.layout.shops,jobs:S.jobs},b);
    if(!same(row,old)){await updateItem('Skids',k._sp,row);
      if(row.Station!==old.Station||row.X_ft!==old.X_ft||row.Y_ft!==old.Y_ft)log('Skid moved',{JobNumber:row.JobNumber,SkidTag:k.tag,Field:'Station',OldValue:old.Station||'floor',NewValue:row.Station||'floor'});
      else log('Skid updated',{JobNumber:row.JobNumber,SkidTag:k.tag})}
    // serials on this skid
    const reqs=k.uniqueReqs||S.partReqs[k.part]||[];
    for(const [slot,v] of Object.entries(k.serials)){const o=(b&&b.serials[slot])||{};if(o.sn===v.sn&&o.photo===v.photo&&v._sp)continue;
      const [rid,q]=slot.split(':');const r=reqs.find(r=>r.id===rid)||{};
      const rec={RecordKey:`${k.tag}|${rid}|${q}`,SkidTag:k.tag,JobNumber:row.JobNumber,EquipmentTag:r.model||'',Description:r.item||'',SerialNumber:v.sn,PhotoLink:v.photo||'',
        EnteredByEmail:ME.email,EnteredAt:new Date().toISOString(),IsOverride:k.uniqueReqs?'Yes':'No'};
      if(!L.SerialRecords)continue;
      if(v._sp)await updateItem('SerialRecords',v._sp,rec);else{const x=await createItem('SerialRecords',rec);v._sp=x.id}
      if(o.sn!==v.sn)log('Serial entered',{JobNumber:row.JobNumber,SkidTag:k.tag,Field:r.model||rid,OldValue:o.sn||'',NewValue:v.sn})}
    if(b)for(const [slot,o] of Object.entries(b.serials))if(!k.serials[slot]&&o._sp&&L.SerialRecords){await deleteItem('SerialRecords',o._sp);log('Serial cleared',{JobNumber:row.JobNumber,SkidTag:k.tag,Field:slot,OldValue:o.sn})}
  }
  for(const b of was.skids)if(!S.skids.some(k=>k.id===b.id)&&b._sp){await deleteItem('Skids',b._sp);log('Skid removed',{SkidTag:b.tag})}
  // 3. serial requirements (shared per part number)
  if(L.SerialRequirements){
    const flat=o=>Object.entries(o).flatMap(([pn,list])=>list.map(r=>({...r,part:pn})));
    const oldR=Object.fromEntries(flat(was.partReqs).map(r=>[r.id,r]));
    for(const r of flat(S.partReqs)){const o=oldR[r.id];const row={ReqKey:r.id,PartNumber:r.part,PartRevision:r.rev||'',EquipmentTag:r.model||'',Description:r.item||'',Qty:r.qty||1};
      if(!r._sp){const x=await createItem('SerialRequirements',row);r._sp=x.id;const live=S.partReqs[r.part].find(z=>z.id===r.id);if(live)live._sp=x.id}
      else if(!o||o.item!==r.item||o.model!==r.model||o.qty!==r.qty)await updateItem('SerialRequirements',r._sp,row)}
    const nowIds=new Set(flat(S.partReqs).map(r=>r.id));
    for(const o of Object.values(oldR))if(!nowIds.has(o.id)&&o._sp)await deleteItem('SerialRequirements',o._sp);
  }
  // 4. people, groups and job access
  if(L.UsersRoles)for(const [email,u] of Object.entries(S.users)){const o=was.users[email];if(u._guest)continue;
    const row={Email:email,DisplayName:u.name,Company:u.company||'RK',Roles:(S.roles[u.role]||{}).label||u.role,Active:'Yes'};
    if(!u._sp){const x=await createItem('UsersRoles',row);u._sp=x.id;delete u._new}
    else if(!o||o.name!==u.name||o.role!==u.role||o.company!==u.company)await updateItem('UsersRoles',u._sp,row)}
  if(L.UsersRoles)for(const [email,o] of Object.entries(was.users))if(!S.users[email]&&o._sp)await deleteItem('UsersRoles',o._sp);
  if(L.Groups){for(const [gk,gr] of Object.entries(S.groups)){const o=was.groups[gk];const row={GroupName:gr.name,MemberEmails:gr.members};
      if(!gr._sp){const x=await createItem('Groups',row);gr._sp=x.id}else if(!o||!same(o.members,gr.members)||o.name!==gr.name)await updateItem('Groups',gr._sp,row)}
    for(const [gk,o] of Object.entries(was.groups))if(!S.groups[gk]&&o._sp)await deleteItem('Groups',o._sp)}
  if(L.JobAccess){
    const want=new Map();const add=(jn,type,who)=>want.set(`${jn}|${type}|${String(who).toLowerCase()}`,{JobNumber:jn,GrantType:type,Principal:who});
    for(const j of Object.values(S.jobs)){j.access.users.forEach(u=>add(j.num,'User',u));j.access.groups.forEach(gk=>S.groups[gk]&&add(j.num,'Group',S.groups[gk].name))}
    S.allJobs.users.forEach(u=>add('ALL','User',u));S.allJobs.groups.forEach(gk=>S.groups[gk]&&add('ALL','Group',S.groups[gk].name));
    const have=new Map();for(const it of L.JobAccess.items)have.set(`${txt(rd('JobAccess',it,'JobNumber')).trim()}|${txt(rd('JobAccess',it,'GrantType'))}|${txt(rd('JobAccess',it,'Principal')).trim().toLowerCase()}`,it);
    for(const [k,row] of want)if(!have.has(k)){const x=await createItem('JobAccess',{GrantID:crypto.randomUUID().slice(0,8),...row,AccessLevel:'View',GrantedByEmail:ME.email,GrantedAt:new Date().toISOString()});
      L.JobAccess.items.push({id:x.id,fields:x.fields});log('Access granted',{JobNumber:row.JobNumber,NewValue:row.GrantType+': '+row.Principal})}
    for(const [k,it] of have)if(!want.has(k)){await deleteItem('JobAccess',it.id);L.JobAccess.items=L.JobAccess.items.filter(x=>x.id!==it.id);log('Access removed',{JobNumber:k.split('|')[0],OldValue:k.split('|').slice(1).join(': ')})}
  }
  SNAP=snapOf(S);
}

/* ================= photos -> SkidPhotos library ================= */
LIVE.uploadPhoto=async function(k,slot,file){
  if(!PHOTODRIVE){status('Photo kept on this screen only: create the SkidPhotos library to save photos.',true);return}
  try{status('Uploading photo…');
    const ext=(file.name.split('.').pop()||'jpg').replace(/[^a-z0-9]/gi,'');const [rid,q]=slot.split(':');
    const reqs=k.uniqueReqs||STATE.partReqs[k.part]||[];const r=reqs.find(r=>r.id===rid)||{};
    const name=`${(r.model||rid).replace(/[^\w-]/g,'_')}-${+q+1}-${Date.now()}.${ext}`;
    const it=await g(`/drives/${PHOTODRIVE}/root:/${encodeURIComponent(k.tag)}/${encodeURIComponent(name)}:/content`,{method:'PUT',body:file,headers:{'Content-Type':file.type||'application/octet-stream'}});
    try{await g(`/drives/${PHOTODRIVE}/items/${it.id}/listItem/fields`,{method:'PATCH',body:JSON.stringify({SkidTag:k.tag,JobNumber:(STATE.jobs[k.jobId]||{}).num||'',EquipmentTag:r.model||'',SerialNumber:(k.serials[slot]||{}).sn||''})})}catch(e){}
    const v=k.serials[slot]=k.serials[slot]||{sn:'',by:ME.name,at:new Date().toISOString().slice(0,10)};v.photo=it.webUrl;
    log('Photo added',{JobNumber:(STATE.jobs[k.jobId]||{}).num||'',SkidTag:k.tag,Field:r.model||rid,NewValue:it.webUrl});
    window.save();
  }catch(e){console.error(e);status('Photo not saved: '+e.message,true)}
};

/* ================= hooks the UI calls ================= */
LIVE.isAdmin=()=>{const u=STATE&&STATE.users[ME.email];return!!u&&u.role==='admin'};
LIVE.settingsHtml=()=>`<section class="sec"><h3>Data</h3><p class="note">Everything here saves to the <a href="${SITEURL}" target="_blank" rel="noopener">RKMC Shop Board</a> SharePoint site as you work. Layouts, roles and job teams live in <span class="mono">${C.configFolder}/board-config.json</span>; jobs, skids, serials, people, groups and access live in the lists of the same names.</p><p class="note">Signed in as ${ME.name} (${ME.email}). <button class="btn" onclick="window.LIVE.signOut()" style="padding:2px 8px">Sign out</button></p></section>`;
LIVE.addPeopleHtml=()=>`<p class="note">To add someone, add a row with their work email to the <a href="${(L.UsersRoles&&L.UsersRoles.webUrl)||SITEURL}" target="_blank" rel="noopener">UsersRoles list</a>, then reload. Role and name changes here save back to that list.</p>`;
LIVE.signOut=()=>{clearTok();location.assign(`${AUTH}/logout?post_logout_redirect_uri=${encodeURIComponent(C.redirectUri)}`)};

/* ================= gate screen ================= */
function setGate(html,showBtn){gate.hidden=false;gate.querySelector('.gmsg').innerHTML=html;gate.querySelector('#gSign').hidden=!showBtn}
function showSignIn(msg){setGate(msg||'Sign in with your RK Microsoft account to open the shop board.',true)}
gate.querySelector('#gSign').onclick=signIn;
function status(msg,bad){const s=document.getElementById('liveStatus');if(!s)return;s.textContent=msg;s.classList.toggle('bad',!!bad);s.hidden=false}

async function boot(){
  if(!C.clientId||!C.tenantId){setGate('This copy of the board is not configured yet. Fill in config.js.',false);return}
  const q=new URLSearchParams(location.search);
  if(q.get('error')){showSignIn(`Microsoft sign-in said: ${q.get('error_description')||q.get('error')}`);history.replaceState(null,'',C.redirectUri);return}
  if(q.get('code')){
    try{tok=await redeem(q.get('code'),q.get('state'))}catch(e){showSignIn(e.message);return}
    history.replaceState(null,'',location.pathname);
    if(window.opener&&window.opener!==window){setGate('Signed in. You can close this window.',false);setTimeout(()=>window.close(),300);return}
  }
  tok=tok||readTok();
  if(!tok){showSignIn();return}
  try{
    setGate('Loading the shop board from SharePoint…',false);
    const me=await g('/me?$select=displayName,mail,userPrincipalName');
    ME={name:me.displayName,email:(me.mail||me.userPrincipalName).toLowerCase()};
    const site=await g(`/sites/${C.siteHost}:${C.sitePath}?$select=id,webUrl`);SITEURL=site.webUrl;
    SITE={id:site.id,lists:await all(`/sites/${site.id}/lists?$select=id,displayName,webUrl`)};
    const drives=await all(`/sites/${site.id}/drives?$select=id,name`);PHOTODRIVE=(drives.find(d=>d.name===C.photoLibrary)||{}).id||null;
    CONFIG=await loadConfig();
    if(!CONFIG){setupScreen();return}
    await loadUserInfo();
    for(const n of LISTS)await loadList(n);
    window.PLANS=CONFIG.plans;window.MC1=CONFIG.mc1||null;
    STATE=LIVE.state=buildState();SNAP=snapOf(STATE);
    if(STATE.users[ME.email]._new){SNAP.users={...SNAP.users};delete SNAP.users[ME.email]}
    const missing=LISTS.filter(n=>!L[n]);
    gate.hidden=true;
    const s=document.createElement('script');s.src='app.js?v='+C.version;
    s.onload=()=>{status(missing.length?'Missing lists (not saved): '+missing.join(', '):STATE.users[ME.email]._guest?'You are not in the UsersRoles list yet, so no jobs are shown. Ask an admin to add you.':'Connected to SharePoint',!!missing.length||!!STATE.users[ME.email]._guest);
      if(STATE.users[ME.email]._new)LIVE.sync(STATE)};
    document.body.appendChild(s);
  }catch(e){
    if(e.message==='signed out')return;
    console.error(e);
    const consent=/consent|AADSTS65001|admin/i.test(e.message);
    setGate(consent?'Your organization has not approved this app yet. An admin needs to grant consent for "RK Shop Board" in Microsoft Entra.':`Could not load the board: ${e.message}`,true);
  }
}
function setupScreen(){
  setGate(`<b>One-time setup.</b> The board's layout file isn't in SharePoint yet. Choose <span class="mono">board-config.json</span> to upload it to <span class="mono">${C.configFolder}</span> on the RKMC Shop Board site.<br><input type="file" id="cfgFile" accept=".json,application/json" style="margin-top:10px">`,false);
  document.getElementById('cfgFile').onchange=async e=>{const f=e.target.files[0];if(!f)return;
    try{const j=JSON.parse(await f.text());if(!j.plans||!j.layout)throw new Error('That file is not a board configuration.');setGate('Uploading…',false);await saveConfig(j);location.reload()}
    catch(err){setGate('Upload failed: '+err.message,false);setTimeout(setupScreen,3000)}};
}
boot();
})();
