/* Shop Board UI. Data comes from m365.js (window.LIVE); PLANS and MC1 come from board-config.json in SharePoint. */
/* ================= true-scale whiteboard PDF ================= */
// Builds a vector PDF (no library) whose drawing prints at exactly 1" = ftPerIn feet at 100%.
function boardPdf(o){
  // o: {shop, name, addr, plan:{d,f}, ftPerIn, boardW, boardH, check:{x0,x1,y,side,label}, elements:[], areaColors:{}, date}
  const IN=72;
  const pw=o.boardW*IN, ph=o.boardH*IN;                  // the sheet is the usable board area, as on the magnet-thread templates
  const k=IN/o.ftPerIn;                                  // points per foot
  const rot=o.rot||0;                                    // 0, or -90 = plan turned so its top edge runs down the left side
  const nums=o.plan.d.match(/-?\d+(?:\.\d+)?/g).map(Number);
  let x0=1e9,y0=1e9,x1=-1e9,y1=-1e9;for(let i=0;i<nums.length;i+=2){const x=nums[i],y=nums[i+1];if(x<x0)x0=x;if(x>x1)x1=x;if(y<y0)y0=y;if(y>y1)y1=y}
  // world (ft, y down) -> page (pt, y up), plan centred on the sheet
  let m;
  if(rot===-90){const ox=(pw-(y1-y0)*k)/2-y0*k, oy=(ph-(x1-x0)*k)/2-x0*k;m=[0,k,k,0,ox,oy]}
  else{const ox=(pw-(x1-x0)*k)/2-x0*k, oy=ph-(ph-(y1-y0)*k)/2+y0*k;m=[k,0,0,-k,ox,oy]}
  const P=(x,y)=>[m[0]*x+m[2]*y+m[4],m[1]*x+m[3]*y+m[5]];
  const f=v=>(Math.round(v*1000)/1000).toString();
  const out=[];
  const pathOps=d=>{const t=d.match(/[MLCZ]|-?\d+(?:\.\d+)?/g);const r=[];let i=0,cmd='';
    while(i<t.length){const tk=t[i];if(/[MLCZ]/.test(tk)){cmd=tk;i++;if(cmd==='Z'){r.push('h');continue}}
      if(cmd==='M'||cmd==='L'){r.push(f(+t[i])+' '+f(+t[i+1])+(cmd==='M'?' m':' l'));i+=2;if(cmd==='M')cmd='L'}
      else if(cmd==='C'){r.push(t.slice(i,i+6).map(v=>f(+v)).join(' ')+' c');i+=6}else i++}
    return r.join('\n')};
  const txt=(s)=>'('+String(s).replace(/[\\()]/g,m=>'\\'+m).replace(/[^\x20-\x7e]/g,'')+')';
  const text=(x,y,size,s,align,font,ang)=>{const w=String(s).length*size*.5;const dx=align==='c'?-w/2:align==='r'?-w:0;const c=Math.cos(ang||0),sn=Math.sin(ang||0);
    out.push(`BT /${font||'F1'} ${size} Tf ${f(c)} ${f(sn)} ${f(-sn)} ${f(c)} ${f(x+dx*c)} ${f(y+dx*sn)} Tm ${txt(s)} Tj ET`)};
  const rgb=h=>{const n=parseInt(h.slice(1),16);return[(n>>16&255)/255,(n>>8&255)/255,(n&255)/255].map(v=>v.toFixed(3)).join(' ')};
  // --- board edge
  out.push('q 0.7 0.7 0.7 RG 0.5 w',`0.25 0.25 ${f(pw-.5)} ${f(ph-.5)} re S`,'Q');
  // --- layout layers (optional), drawn under the building linework
  for(const e of o.elements||[]){
    if(!e.pts)continue;const pts=e.pts.map(p=>P(p[0],p[1]));
    const col=e.kind==='area'?(o.areaColors[e.type]||'#777777'):e.kind==='flow'?(e.c||'#0f6fe0'):'#555555';
    const w=e.kind==='line'?1.2:e.kind==='station'?.4:.8;const dash=e.kind==='station'||e.kind==='area'?'[3 2] 0 d':'[] 0 d';
    out.push(`q ${rgb(col)} RG ${w} w ${dash} 1 J 1 j`,pts.map((p,i)=>f(p[0])+' '+f(p[1])+(i?' l':' m')).join(' ')+(e.kind==='flow'?' S':' h S'),'Q');
    if(e.name&&(e.kind==='line'||e.kind==='area')){const xs=pts.map(p=>p[0]),ys=pts.map(p=>p[1]);out.push(`q ${rgb(col)} rg`);text(Math.min(...xs)+2,Math.max(...ys)+(e.kind==='line'?3:-8),e.kind==='line'?9:6,e.name);out.push('Q')}
  }
  // --- building linework at true scale
  const lw=0.6/k;                                         // 0.6 pt in foot units
  out.push(`q 0 0 0 RG 0 0 0 rg 1 J 1 j ${f(lw)} w`,m.map(f).join(' ')+' cm',pathOps(o.plan.d),'S');
  if(o.plan.f)out.push(pathOps(o.plan.f),'f');
  out.push('Q');
  // --- check dimension: a known inside dimension, drawn with extension lines outside the wall
  if(o.check){const c=o.check;const A=P(...c.a),B=P(...c.b),A2=P(c.a[0]+c.off[0],c.a[1]+c.off[1]),B2=P(c.b[0]+c.off[0],c.b[1]+c.off[1]);
    const ux=A2[0]-A[0],uy=A2[1]-A[1],ul=Math.hypot(ux,uy)||1,nx=ux/ul,ny=uy/ul;
    out.push('q 0.75 0 0 RG 0.75 0 0 rg 0.7 w',`${f(A[0]+nx*3)} ${f(A[1]+ny*3)} m ${f(A2[0]+nx*4)} ${f(A2[1]+ny*4)} l S ${f(B[0]+nx*3)} ${f(B[1]+ny*3)} m ${f(B2[0]+nx*4)} ${f(B2[1]+ny*4)} l S`,`${f(A2[0])} ${f(A2[1])} m ${f(B2[0])} ${f(B2[1])} l S`);
    const dx=B2[0]-A2[0],dy=B2[1]-A2[1],dl=Math.hypot(dx,dy),tx=dx/dl,ty=dy/dl;
    for(const p of [A2,B2])out.push(`${f(p[0]-(tx+nx)*4)} ${f(p[1]-(ty+ny)*4)} m ${f(p[0]+(tx+nx)*4)} ${f(p[1]+(ty+ny)*4)} l S`);
    let ang=Math.atan2(dy,dx);if(ang>Math.PI/2)ang-=Math.PI;if(ang<-Math.PI/2)ang+=Math.PI;
    const inches=Math.hypot(c.b[0]-c.a[0],c.b[1]-c.a[1])/o.ftPerIn;const mx=(A2[0]+B2[0])/2+nx*(ny<-.5?14:5),my=(A2[1]+B2[1])/2+ny*(ny<-.5?14:5);
    text(mx+(Math.abs(nx)>.5&&nx>0?3:0),my,8,`CHECK: ${c.label} inside = ${inches.toFixed(2)}" on this print`,'c',undefined,ang);out.push('Q')}
  // --- 100 ft scale bar and print note, bottom-left (matches the magnet-thread templates)
  const sx=0.5*IN, sy=0.45*IN, span=100, step=10;
  out.push('q 0 0 0 RG 0 0 0 rg 0.6 w');
  for(let i=0;i<10;i+=2)out.push(`${f(sx+i*step*k)} ${f(sy)} ${f(step*k)} 5 re f`);
  out.push(`${f(sx)} ${f(sy)} ${f(span*k)} 5 re S`);
  for(let i=0;i<=10;i+=2)text(sx+i*step*k,sy+8,6,(i*step)+"'",'c');
  text(sx,sy-9,7,`100 ft = ${(span/o.ftPerIn).toFixed(2)}" at 1" = ${o.ftPerIn}'-0". Print at 100% (actual size); do not scale to fit.`);
  out.push('Q');
  // --- title, top-left
  out.push('q 0 0 0 rg');
  text(0.5*IN,ph-0.5*IN,12,`${o.name}  -  1" = ${o.ftPerIn}'-0" (1:${o.ftPerIn*12})`,'l','F2');
  text(0.5*IN,ph-0.5*IN-11,7,`${o.verified||''}   |   Sheet ${o.boardW}" x ${o.boardH}" = usable board area   |   ${o.date}`);
  out.push('Q');
  // --- assemble
  const content=out.join('\n');
  const objs=['<< /Type /Catalog /Pages 2 0 R >>','<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${f(pw)} ${f(ph)}] /Contents 4 0 R /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> >>`,
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>',
    `<< /Title ${txt(o.name+' whiteboard 1in='+o.ftPerIn+'ft')} /Producer (RK Shop Floor) >>`];
  let pdf='%PDF-1.4\n';const offs=[];
  objs.forEach((b,i)=>{offs.push(pdf.length);pdf+=`${i+1} 0 obj\n${b}\nendobj\n`});
  const xref=pdf.length;
  pdf+=`xref\n0 ${objs.length+1}\n0000000000 65535 f \n`+offs.map(n=>String(n).padStart(10,'0')+' 00000 n \n').join('');
  pdf+=`trailer\n<< /Size ${objs.length+1} /Root 1 0 R /Info ${objs.length} 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return pdf;
}


'use strict';
/* ================= reference data ================= */
const AREA_TYPES = {
  receiving:{label:'Receiving',c:'#3b6fd1'}, shipping:{label:'Shipping / staging',c:'#7a63c7'},
  matstore:{label:'Material storage',c:'#2d8f9a'}, storage:{label:'General storage',c:'#77808c'},
  quarantine:{label:'Material quarantine',c:'#d0392b'}, breakroom:{label:'Breakroom',c:'#3f9b4f'},
  forklift:{label:'Forklift lane',c:'#e0b000'}, inspection:{label:'QC / test bay',c:'#e07b1f'},
  toolcrib:{label:'Tool crib',c:'#9a7652'}, office:{label:'Office',c:'#8d97a3'}
};
const LAYERS = [
  ['areas','Areas','#2d8f9a'],['lines','Production lines','#39434b'],['stations','Stations','#8d97a3'],
  ['flow','Layout lines','#0f6fe0'],['zones','Customer zones','#e0337a'],['equip','Equipment','#9a7652'],
  ['skids','Skids','#f2ae00'],['material','Material','#d0392b']
];
// Roles drive "Viewing as" and the job team roster. Admins edit them in Settings; these are the defaults.
// team: 1 = the role appears on every job's Team tab.
const ROLES0 = {
  admin:{label:'Administrator',short:'Admin',company:'RK',edit:1,mat:1,move:1,add:1,qc:1,sn:1,reqs:1,plan:1,teamEdit:1,grant:1,fixed:1},
  pm:{label:'Project manager',short:'PM',company:'RK',team:1,mat:1,add:1,reqs:1,plan:1,teamEdit:1,grant:1},
  pe:{label:'Project engineer',short:'PE',company:'RK',team:1,mat:1,reqs:1,plan:1},
  ame:{label:'Arch / mech engineer',short:'Arch/Mech',company:'RK',team:1,reqs:1,plan:1},
  ee:{label:'Electrical engineer',short:'EE',company:'RK',team:1,reqs:1,sn:1,plan:1},
  pa:{label:'Purchasing agent',short:'Purchasing',company:'RK',team:1,mat:1,plan:1},
  mfe:{label:'Manufacturing engineer',short:'Mfg Eng',company:'RK',team:1,mat:1,reqs:1,plan:1},
  prod:{label:'Production manager',short:'Prod Mgr',company:'RK',team:1,mat:1,move:1,add:1,plan:1,teamEdit:1},
  sup:{label:'Production supervisor',short:'Supervisor',company:'RK',team:1,mat:1,move:1,sn:1,plan:1},
  qc:{label:'QC inspector',short:'QC',company:'RK',mat:1,qc:1,sn:1},
  tech:{label:'Technician',short:'Tech',company:'RK',mat:1},
  cust:{label:'Customer',company:'customer',fixed:1}
};
const PERMS=[['edit','Edit layout'],['mat','Edit material'],['move','Move skids'],['add','Add skids'],['qc','QC checklists'],['sn','Serial numbers'],['reqs','Equipment lists'],['plan','Labor vs plan'],['teamEdit','Edit job teams'],['grant','Grant job access']];
const TEAM_NAMES=['Avery Cole','Jordan Patel','Sam Ortiz','Riley Chen','Morgan Hayes','Casey Nguyen','Taylor Brooks','Jamie Flores','Drew Kim','Quinn Walsh','Reese Turner','Rowan Price','Skyler Ward','Emerson Lane','Hayden Cruz','Parker Shaw','Logan Reyes','Kendall Moore'];
// Sample people. Job access: admins see every job; everyone else sees jobs they're on the team for,
// jobs granted to them or one of their groups, or every job if they or a group hold an explicit "All jobs" grant.
const EXTRA_PEOPLE=[['Aaron B.','admin'],['Noel Harper','qc'],['Blair Owens','qc'],['Jesse Ramos','tech'],['Kai Bennett','tech'],['Lane Foster','tech'],
  ['Northwind contact','cust','Northwind Data Centers'],['Contoso contact','cust','Contoso Power']];
function seedPeople(S){
  const users={};const byName={};const add=(name,roleId,company)=>{if(byName[name])return byName[name];const id=uid('u');users[id]={name,role:roleId,...(company?{company}:{})};byName[name]=id;return id};
  for(const [n,r,c] of EXTRA_PEOPLE)add(n,r,c);
  for(const j of Object.values(S.jobs)){if(!j.team)continue;for(const [r,list] of Object.entries(j.team))j.team[r]=list.map(x=>users[x]?x:add(x,S.roles[r]?r:'tech'))}
  S.users=users;
  const qc=Object.keys(users).filter(id=>users[id].role==='qc'),crew=[byName['Jesse Ramos'],byName['Kai Bennett']];
  S.groups={g_qc:{name:'QC inspectors',members:qc},g_l3:{name:'RKMC1 Line 3 crew',members:crew}};
  S.allJobs={users:[],groups:['g_qc']};
  for(const [jid,j] of Object.entries(S.jobs))j.access={users:[],groups:jid==='J24112'?['g_l3']:[]};
  S.defaultUser=byName['Aaron B.'];
}
function seedTeam(j,i){const t={pm:[j.pm]};let n=i*5;for(const r of ['pe','ame','ee','pa','mfe','prod','sup'])t[r]=[TEAM_NAMES[(n++)%TEAM_NAMES.length]];if(i%2)t.pe.push(TEAM_NAMES[(n+3)%TEAM_NAMES.length]);t.sup.push(TEAM_NAMES[(n+7)%TEAM_NAMES.length]);return t}
const VERIFIED_ASCII={rkmc1:"520' x 480', columns 40' o.c.",rkmc2:"298'-10 1/4\" x 418'-9 1/2\", columns 60' o.c."};
const VERIFIED = {rkmc1:"520' × 480', columns 40' o.c. (plan matches)",rkmc2:"298'-10¼\" × 418'-9½\", columns 60' o.c. (plan matches within 1\")"};
const MILESTONES = ['Received','Structure','Mechanical','Electrical','Testing','Final QC','Ship'];
const QC_TEMPLATES = [
  {id:'rcv',name:'Receiving inspection',stage:0,items:['Skid base matches drawing rev','Base welds visually inspected','Lifting lugs certified & tagged','Damage from transport documented','Material certs on file']},
  {id:'mech',name:'Mechanical installation',stage:2,items:['Equipment anchored per drawing','Pipe supports installed & level','Flange bolts torqued (record values)','Vibration isolators set','Clearances per NEC / mfr']},
  {id:'elec',name:'Electrical termination & torque',stage:3,items:['Conductors sized per schedule','Terminations torqued & marked','Grounding / bonding verified','Labels & arc-flash placards','Conduit fill & supports']},
  {id:'test',name:'Pre-energization / hi-pot',stage:4,items:['Insulation resistance test','Hi-pot results recorded','Phase rotation verified','Control wiring point-to-point','Functional test of controls']},
  {id:'final',name:'Final QC walkdown',stage:5,items:['Punch list closed','Paint touch-up complete','As-built redlines returned','Shipping bracing installed','Customer witness sign-off']}
];
const PARTS = window.LIVE.parts;
const JOBS0 = {
  J24107:{num:'24107',customer:'Northwind Data Centers',project:'NDC Phase 2 electrical houses',pm:'Dana Ruiz',part:'RK-EH-4813',c:'#0f6fe0',ship:'2026-12-18',ifs:'P-24107',acc:'NDC-PH2',ncr:2,issues:5,teams:'NDC Phase 2',planner:[14,3]},
  J24112:{num:'24112',customer:'Contoso Power',project:'Contoso substation pump skids',pm:'Marcus Lee',part:'RK-PS-4012',c:'#9b3fd6',ship:'2026-11-20',ifs:'P-24112',acc:'CON-SUB',ncr:0,issues:2,teams:'Contoso Substation',planner:[9,0]},
  J24118:{num:'24118',customer:'Fabrikam Utilities',project:'Fabrikam central plant CHW modules',pm:'Priya Shah',part:'RK-CHW-4613',c:'#e0337a',ship:'2027-01-15',ifs:'P-24118',acc:'FAB-CP',ncr:1,issues:3,teams:'Fabrikam Central Plant',planner:[21,5]},
  J24121:{num:'24121',customer:'Tailspin Biotech',project:'Tailspin lab MEP racks',pm:'Dana Ruiz',part:'RK-MEP-3610',c:'#3a9d23',ship:'2026-12-04',ifs:'P-24121',acc:'TSB-LAB',ncr:0,issues:1,teams:'Tailspin Lab',planner:[6,1]},
  J24125:{num:'24125',customer:'Northwind Data Centers',project:'NDC Campus 3 generator bases',pm:'Marcus Lee',part:'RK-GB-2611',c:'#0a8fa3',ship:'2027-02-05',ifs:'P-24125',acc:'NDC-C3',ncr:1,issues:0,teams:'NDC Campus 3',planner:[11,2]},
  J24130:{num:'24130',customer:'Woodgrove Bank',project:'Woodgrove data hall power modules',pm:'Priya Shah',part:'RK-DH-5013',c:'#c96a00',ship:'2027-01-29',ifs:'P-24130',acc:'WGB-DH',ncr:3,issues:4,teams:'Woodgrove DH',planner:[17,4]}
};
const LIB0 = [
  ['cart5s','5S material cart',4,2,'#2d8f9a'],['shelf','Rolling shelf',6,2,'#3b6fd1'],['cage','Wire cage (secure)',10,8,'#77808c'],
  ['rack','Pallet rack bay',8,4,'#2d8f9a'],['kanban','Gravity-flow kanban rack',6,3,'#3f9b4f'],['bench','Workbench',8,3,'#9a7652'],
  ['chest','Tool chest',4,2,'#d0392b'],['gang','Gang box',4,2.5,'#e0b000'],['reel','Cable reel stand',4,4,'#7a63c7'],
  ['testcart','Test / torque cart',3,2,'#e07b1f'],['weldscr','Welding screen',6,1,'#c96a00'],['flam','Flammable cabinet',3.5,1.5,'#d0392b'],
  ['spill','Spill kit',2,2,'#e0b000'],['eyewash','Eyewash station',2,2,'#3f9b4f'],['scrap','Scrap / recycling bins',4,4,'#77808c'],
  ['charge','Forklift charging',10,6,'#e0b000'],['lift','Scissor lift parking',8,5,'#e0b000']
].map(([id,name,L,W,c])=>({id,name,L,W,c}));
const MATERIALS = [
  ['100-48213','Breaker, 800A frame, 3P 65kA','rkmc1','MS1-A03',148,26,12],
  ['100-48877','Lug kit, 2-hole 600 kcmil','rkmc1','MS1-A07',236,30,240],
  ['210-10552','Unistrut P1000, 10 ft','rkmc1','MS1-B02',176,36,85],
  ['310-77120','Gasket, 6" 150# spiral wound','rkmc1','MS2-C11',214,430,64],
  ['310-77145','Butterfly valve, 6" lug','rkmc1','MS2-C14',300,436,9],
  ['410-22019','THHN 4/0 CU, 1000 ft reel','rkmc1','MS2-D04',380,424,6],
  ['500-90011','Anchor bolt, 3/4" x 8" HDG','rkmc2','R-04-2',225,190,500],
  ['500-90120','Grating panel, 3\'x10\'','rkmc2','R-07-1',240,282,22],
  ['610-33410','VFD 40 HP, 480V (quarantined)','rkmc2','Q-02',222,360,1],
  ['610-33520','Flow meter, mag 6"','rkmc2','R-11-3',230,315,3]
].map(([pn,desc,shop,bin,x,y,qty])=>({pn,desc,shop,bin,x,y,qty}));

/* ================= seed builder ================= */
function rng(seed){return function(){seed|=0;seed=seed+0x6D2B79F5|0;let t=Math.imul(seed^seed>>>15,1|seed);t=t+Math.imul(t^t>>>7,61|t)^t;return((t^t>>>14)>>>0)/4294967296}}
let SEED=null;
let _id=1; const uid=p=>p+(_id++).toString(36)+Math.random().toString(36).slice(2,5);
const rect=(x0,y0,x1,y1)=>[[x0,y0],[x1,y0],[x1,y1],[x0,y1]];

function buildSeed(){
  const R=rng(42);
  const S={elements:[],skids:[],roles:JSON.parse(JSON.stringify(ROLES0)),jobs:JSON.parse(JSON.stringify(JOBS0)),library:LIB0.map(o=>({...o})),
    partReqs:Object.fromEntries(Object.entries(PARTS).map(([k,p])=>[k,p.reqs.map(r=>({id:uid('r'),item:r[0],model:r[1],qty:r[2]}))])),
    shops:{
      rkmc1:{name:'RKMC1',addr:'',drawing:'1" = 50\'-0" (C. Pell, 2023-07-18)',board:{w:46,h:34,ftPerIn:15,rot:0}},
      rkmc2:{name:'RKMC2',addr:'Building plan, first floor',drawing:'1" = 30\'-0" (Reno James M002)',board:{w:46,h:34,ftPerIn:12,rot:-90}}
    },
    views:[], plan:{}, materials:MATERIALS.map(m=>({...m}))};
  SEED=S;
  const E=(o)=>{o.id=o.id||uid('e');S.elements.push(o);return o};
  const centroid=pts=>[pts.reduce((a,p)=>a+p[0],0)/pts.length,pts.reduce((a,p)=>a+p[1],0)/pts.length];
  // ---- production lines & stations
  const lines={};
  function mkLine(shop,id,name,cells,orient){
    const xs=cells.flatMap(c=>[c[0],c[2]]),ys=cells.flatMap(c=>[c[1],c[3]]);
    const ln=E({id,shop,kind:'line',name,pts:rect(Math.min(...xs)-1,Math.min(...ys)-1,Math.max(...xs)+1,Math.max(...ys)+1)});
    ln.stations=cells.map((c,i)=>E({id:id+'-S'+(i+1),shop,kind:'station',lineId:id,name:'S'+(i+1),pts:rect(...c),orient}).id);
    lines[id]=ln;
    // flow line through station centres
    const cs=cells.map(c=>[(c[0]+c[2])/2,(c[1]+c[3])/2]);
    const a=cs[0],b=cs[1]||cs[0],z=cs[cs.length-1],y=cs[cs.length-2]||z;
    const ext=(p,q,d)=>{const dx=p[0]-q[0],dy=p[1]-q[1],L=Math.hypot(dx,dy)||1;return[p[0]+dx/L*d,p[1]+dy/L*d]};
    E({shop,kind:'flow',name:name+' flow',lineId:id,c:'#0f6fe0',pts:[ext(a,b,16),...cs,ext(z,y,16)].map(p=>p.map(v=>Math.round(v*10)/10))});
  }
  // RKMC1: traced from "RK-MC1-BLD-Full August Flow 2026" (lines 1–9, stations #0–#6, skid flow in/out)
  for(const s of MC1.stations)E({id:s.id,shop:'rkmc1',kind:'station',lineId:s.line,name:s.name,pts:s.pts,orient:s.orient});
  for(const l of MC1.lines)lines[l.id]=E({id:l.id,shop:'rkmc1',kind:'line',name:l.name,pts:l.pts,stations:l.stations.slice()});
  for(const f of MC1.flows)E({shop:'rkmc1',kind:'flow',name:f.name,lineId:f.line||'',c:f.c,pts:f.pts});
  const c2=(x0,i)=>[x0,128+45*i,x0+36,171+45*i];
  mkLine('rkmc2','LA','Line A',[0,1,2,3,4,5,6,7].map(i=>c2(34,i)),'v');
  mkLine('rkmc2','LB','Line B',[0,1,2,3,4,5,6,7].map(i=>c2(92,i)),'v');
  mkLine('rkmc2','LC','Line C',[0,1,2,3,4,5,6,7].map(i=>c2(150,i)),'v');
  // ---- areas
  const area=(shop,type,name,pts)=>E({shop,kind:'area',type,name,pts});
  for(const a of MC1.areas)area('rkmc1',a.type,a.name,a.pts);
  area('rkmc2','forklift','North aisle',rect(22,108,252,120));
  area('rkmc2','forklift','Aisle A–B',rect(75,128,87,486));
  area('rkmc2','forklift','Aisle B–C',rect(133,128,145,486));
  area('rkmc2','forklift','Dock aisle',rect(252,108,262,505));
  area('rkmc2','matstore','Racking',rect(202,160,250,330));
  area('rkmc2','receiving','Receiving / docks',rect(264,160,311,300));
  area('rkmc2','quarantine','Material quarantine',rect(202,340,250,390));
  area('rkmc2','storage','General storage',rect(202,400,250,505));
  area('rkmc2','shipping','Ship staging',rect(264,400,311,505));
  area('rkmc2','inspection','QC / test bay',rect(264,310,311,390));
  area('rkmc2','breakroom','Breakroom',rect(276,110,296,149));
  area('rkmc2','office','Offices',rect(161,5,270,91));
  // ---- customer zones + skids (sample jobs placed on the real lines)
  const zones=[['L1',1,6,'J24107'],['L2',1,3,'J24118'],['L2',4,6,'J24130'],['L3',1,6,'J24112'],['L4',1,5,'J24107'],['L5',1,1,'J24118'],
    ['L6',1,6,'J24121'],['L7',1,4,'J24118'],['L8',1,5,'J24130'],['L9',1,3,'J24112'],
    ['LA',1,5,'J24125'],['LB',1,7,'J24121'],['LC',1,4,'J24112'],['LC',6,8,'J24125']];
  const skidCount={};
  const byId=Object.fromEntries(S.elements.map(e=>[e.id,e]));
  const stationOf={};
  for(const [lid,from,to,jid] of zones){
    const ln=lines[lid], st=ln.stations.slice(from-1,to).map(id=>byId[id]);
    // RKMC2 stations share a band; RKMC1 stations are staggered, so one zone covers the whole run
    const groups=[];
    for(const s of st){const key=ln.shop==='rkmc1'?ln.id:(s.orient==='h'?s.pts[0][1]:s.pts[0][0]);const g=groups[groups.length-1];if(g&&g.key===key)g.list.push(s);else groups.push({key,list:[s]})}
    for(const g of groups){
      const xs=g.list.flatMap(s=>s.pts.map(p=>p[0])),ys=g.list.flatMap(s=>s.pts.map(p=>p[1]));const pad=ln.shop==='rkmc1'?1.2:-.5;
      E({shop:ln.shop,kind:'zone',jobId:jid,name:'',pts:rect(Math.min(...xs)-pad,Math.min(...ys)-pad,Math.max(...xs)+pad,Math.max(...ys)+pad)});
    }
    const job=S.jobs[jid],part=PARTS[job.part];
    st.forEach((s,i)=>{
      stationOf[s.id]=jid;
      const n=(skidCount[jid]=(skidCount[jid]||0)+1);
      const cx=(s.pts[0][0]+s.pts[2][0])/2,cy=(s.pts[0][1]+s.pts[2][1])/2;
      const pct=Math.min(.97,Math.max(.04,.12+.78*(i/Math.max(1,st.length-1))+(R()-.5)*.12));
      S.skids.push(mkSkid({shop:ln.shop,jobId:jid,n,x:cx,y:cy,rot:s.orient==='v'?90:0,stationId:s.id,pct,part:job.part,L:part.L,W:part.W,R,reqs:S.partReqs[job.part]}));
    });
  }
  // ---- equipment: library items traced from the flow drawing, plus the generic ones
  const libColor=n=>/lift|jack/i.test(n)?'#e0b000':/cage|tool/i.test(n)?'#77808c':/bin|trash/i.test(n)?'#9a7652':/table|bench|board|stand/i.test(n)?'#7a63c7':'#2d8f9a';
  for(const l of MC1.library)if(!S.library.some(x=>x.id===l.id))S.library.push({id:l.id,name:l.name,L:Math.max(1,l.L),W:Math.max(1,l.W),c:libColor(l.name)});
  const eq=(shop,libId,x,y,rot=0)=>E({shop,kind:'equip',libId,x,y,rot,name:''});
  for(const q of MC1.equip)eq('rkmc1',q.libId,q.x,q.y,q.rot);
  for(let i=0;i<8;i++){eq('rkmc2','cart5s',30,150+45*i,90);eq('rkmc2','shelf',89,150+45*i,90);eq('rkmc2','gang',147,150+45*i,90)}
  eq('rkmc2','cage',220,140);eq('rkmc2','charge',285,95);eq('rkmc2','scrap',235,500);eq('rkmc2','eyewash',190,123);eq('rkmc2','flam',196,330,90);
  // ---- build material: IFS part catalogue, material zones and the pieces on the floor
  S.catalog=JSON.parse(JSON.stringify(MC1.catalog));
  for(const m of MATERIALS)S.catalog[m.pn]=S.catalog[m.pn]||{desc:m.desc,uom:'EA'};
  const mzs=MC1.matzones.map(z=>E({shop:'rkmc1',kind:'matzone',name:z.name,jobId:z.job,loc:z.loc,pts:z.pts,parts:[]}));
  const stations1=S.elements.filter(e=>e.shop==='rkmc1'&&e.kind==='station'&&e.lineId);
  for(const p of MC1.pieces){
    const z=mzs.find(z=>inPoly(p.x,p.y,z.pts));
    let job=z?z.jobId:null;
    if(!z){const near=stations1.map(s=>{const[cx,cy]=centroid(s.pts);return{s,d:Math.hypot((cx-p.x)/3,cy-p.y)}}).sort((a,b)=>a.d-b.d)[0];if(near&&near.d<16)job=stationOf[near.s.id]||null}
    const ea=S.catalog[p.pn]&&S.catalog[p.pn].uom==='EA';
    E({shop:'rkmc1',kind:'stock',pn:p.pn,x:p.x,y:p.y,L:Math.max(.8,p.L),W:Math.max(.8,p.W),rot:p.rot,jobId:job,qty:ea?1:Math.ceil(R()*12),name:''});
  }
  S.materials=S.materials.filter(m=>m.shop!=='rkmc1');
  // ---- production plan per line (earned hours)
  for(const id of Object.keys(lines)){const d=60+Math.round(R()*90);const f=.82+R()*.3;S.plan[id]={day:d,dayAct:Math.round(d*f),week:d*5,weekAct:Math.round(d*5*(.9+R()*.18))}}
  // ---- saved views
  S.views=[
    {id:'v1',name:'Line 1 close-up',shop:'rkmc1',cx:365,cy:110,s:3.4,scope:'shared',owner:'admin'},
    {id:'v2',name:'Lines 1–4',shop:'rkmc1',cx:360,cy:150,s:2.0,scope:'shared',owner:'admin'},
    {id:'v5',name:'East wall build material',shop:'rkmc1',cx:380,cy:465,s:3.2,scope:'shared',owner:'admin'},
    {id:'v3',name:'Receiving & docks',shop:'rkmc2',cx:260,cy:240,s:3.2,scope:'shared',owner:'admin'},
    {id:'v4',name:'My Woodgrove lines',shop:'rkmc1',cx:360,cy:330,s:2.2,scope:'personal',owner:'admin'}
  ];
  Object.values(S.jobs).forEach((j,i)=>j.team=seedTeam(j,i));seedPeople(S);
  return S;
}
function mkSkid({shop,jobId,n,x,y,rot,stationId,pct,part,L,W,R,reqs}){
  R=R||Math.random;
  const job=S0jobs()[jobId]||{};
  const stage=Math.min(5,Math.floor(pct*6.2));
  const qc={};
  for(const t of QC_TEMPLATES){
    if(t.stage<stage){qc[t.id]={res:t.items.map(()=>R()<.06?'f':(R()<.05?'n':'p')),done:true,by:'J. Ortega',at:daysAgo(Math.round((stage-t.stage)*4+R()*3)),acc:true}}
  }
  reqs=reqs||(STATE&&STATE.partReqs[part])||[];
  const serials={};
  let k=0;
  reqs.forEach((r,ri)=>{for(let q=0;q<r.qty;q++){k++;if(R()<pct*1.05)serials[r.id+':'+q]={sn:r.model.split('-')[0]+'-'+(100000+Math.floor(R()*899999)),by:'T. Nguyen',at:daysAgo(Math.round(R()*12))}}});
  const est=PARTS[part].est, eff=.85+R()*.25;
  const tag=(job.num||'00000')+'-M'+String(n).padStart(2,'0');
  return {id:uid('k'),tag,shop,jobId,part,so:'SO-'+(job.num||'0')+'-'+String(n).padStart(2,'0'),x,y,rot,L,W,stationId,pct:Math.round(pct*100)/100,
    planPct:Math.min(1,Math.round((pct+(R()-.45)*.14)*100)/100),est,act:Math.round(est*pct/eff),qc,serials,uniqueReqs:null,
    history:[{at:daysAgo(Math.round(20+R()*10)),txt:'Raw skid received at '+(shop==='rkmc1'?'RKMC1':'RKMC2')+' receiving'},{at:daysAgo(Math.round(R()*6)),txt:'Moved to '+stationLabel(stationId)}]};
}
function daysAgo(d){const t=new Date(Date.now()-d*864e5);return t.toISOString().slice(0,10)}
function stationLabel(id){if(!id)return 'floor (unassigned)';const src=STATE||SEED;const s=src&&src.elements.find(e=>e.id===id);if(!s)return id;const ln=s.lineId&&src.elements.find(e=>e.id===s.lineId);return (ln?ln.name+' · ':'')+s.name}
let STATE=null; const S0jobs=()=>STATE?STATE.jobs:JOBS0;

/* ================= persistence ================= */
const KEY='rk-shopfloor-v5';
function load(){return window.LIVE.state}
let saveT=0;function save(){clearTimeout(saveT);window.LIVE.dirty();saveT=setTimeout(()=>window.LIVE.sync(STATE),600)}
const PHOTOS={}; // in-memory only
STATE=load(); if(!STATE){STATE=null;STATE=buildSeed()}
if(!STATE.roles){STATE.roles=JSON.parse(JSON.stringify(ROLES0));save()}
Object.values(STATE.jobs).forEach((j,i)=>{if(!j.team){j.team=seedTeam(j,i);save()}});
if(!STATE.users){if(STATE.roles.admin){STATE.roles.admin.label='Administrator';STATE.roles.admin.short='Admin'}if(STATE.roles.tech)delete STATE.roles.tech.sn;STATE.roles.cust={...ROLES0.cust};if(STATE.roles.admin)STATE.roles.admin.grant=1;if(STATE.roles.pm)STATE.roles.pm.grant=1;seedPeople(STATE);save()}
// boards match the plotter templates: both 46" x 34" landscape, RKMC2 turned with its offices on the left
if(STATE.shops.rkmc2&&STATE.shops.rkmc2.board.rot===undefined){STATE.shops.rkmc1.board={w:46,h:34,ftPerIn:15,rot:0};STATE.shops.rkmc2.board={w:46,h:34,ftPerIn:12,rot:-90};save()}
// known inside dimensions, printed on the PDF as a tape-measure check
const CHECKS={rkmc1:{a:[600.1,8.1],b:[600.1,488.1],off:[9,0],label:"480'-0\""},rkmc2:{a:[318.2,91.5],b:[318.2,510.3],off:[38,0],label:"418'-9 1/2\""}};
const _ext={};function planExt(shop){if(_ext[shop])return _ext[shop];const n=PLANS[shop].d.match(/-?\d+(?:\.\d+)?/g).map(Number);let x0=1e9,y0=1e9,x1=-1e9,y1=-1e9;for(let i=0;i<n.length;i+=2){x0=Math.min(x0,n[i]);x1=Math.max(x1,n[i]);y0=Math.min(y0,n[i+1]);y1=Math.max(y1,n[i+1])}return _ext[shop]=[x1-x0,y1-y0]}
function boardFit(k){const bd=STATE.shops[k].board;let[w,h]=planExt(k);if(bd.rot)[w,h]=[h,w];return{pw:w/bd.ftPerIn,ph:h/bd.ftPerIn,fits:w/bd.ftPerIn<=bd.w&&h/bd.ftPerIn<=bd.h,need:Math.max(w/bd.w,h/bd.h)}}

/* ================= ui state ================= */
const UI={shop:'rkmc1',role:'admin',layers:Object.fromEntries(LAYERS.map(l=>[l[0],true])),labels:true,
  view:{s:1,tx:0,ty:0},sel:{type:'shop'},hover:null,tab:'overview',edit:false,move:false,tool:'select',drawKind:'area',drawType:'receiving',
  draft:null,editSel:null,measure:null,measuring:false,equipLib:'cart5s',matHit:null,addSkid:false};
UI.user=STATE.defaultUser;
const me=()=>STATE.users[UI.user]||Object.values(STATE.users)[0];
Object.defineProperty(UI,'role',{get(){return me().role},set(v){}});
// a role says what someone can do; a customer's company comes from their login
const role=()=>{const r=STATE.roles[me().role]||STATE.roles.admin;return r.company==='customer'?{...r,company:me().company||'none',label:me().name}:{...r,label:me().name}};
const groupsOf=uid=>Object.entries(STATE.groups).filter(([,g])=>g.members.includes(uid)).map(([id])=>id);
const onTeam=(jid,uid=UI.user)=>{const j=STATE.jobs[jid];return!!(j&&j.team&&Object.values(j.team).some(l=>l.includes(uid)))};
const hasAllJobs=(uid=UI.user)=>STATE.roles[STATE.users[uid].role]&&STATE.users[uid].role==='admin'||STATE.allJobs.users.includes(uid)||groupsOf(uid).some(g=>STATE.allJobs.groups.includes(g));
function accessWhy(jid,uid=UI.user){const j=STATE.jobs[jid];const u=STATE.users[uid];if(!j||!u)return null;const r=STATE.roles[u.role]||{};
  if(r.company==='customer')return j.customer===u.company?'customer':null;
  if(u.role==='admin')return'admin';if(onTeam(jid,uid))return'team';
  if(j.access.users.includes(uid))return'person';const g=groupsOf(uid).find(g=>j.access.groups.includes(g));if(g)return'group:'+g;
  if(hasAllJobs(uid))return'all';return null}
const canGrant=jid=>{const R=role();return!!R.grant&&(me().role==='admin'||onTeam(jid))};
const canTeamEdit=jid=>{const R=role();return!!R.teamEdit&&(me().role==='admin'||onTeam(jid))};
const uname=id=>(STATE.users[id]||{name:'Removed user'}).name;
const teamRoles=()=>Object.entries(STATE.roles).filter(([,r])=>r.team);
const roleShort=id=>{const r=STATE.roles[id];return r?(r.short||r.label):id};
const teamOf=(j,id)=>(j.team&&j.team[id])||[];
const teamNames=(j,id)=>teamOf(j,id).map(uname);
const canSeeJob=jid=>!!accessWhy(jid);
const els=()=>STATE.elements.filter(e=>e.shop===UI.shop);
const elById=id=>STATE.elements.find(e=>e.id===id);
const skidById=id=>STATE.skids.find(k=>k.id===id);
const libById=id=>STATE.library.find(l=>l.id===id)||{name:'Unknown',L:4,W:4,c:'#888'};
const undo=[];function pushUndo(){undo.push(JSON.stringify({e:STATE.elements,k:STATE.skids}));if(undo.length>40)undo.shift()}

/* ================= helpers ================= */
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function fi(ft){const t=Math.round(ft*12);return Math.floor(t/12)+"'-"+(t%12)+'"'}
function fin(inches){return (Math.round(inches*100)/100).toFixed(2)+'"'}
const ptsStr=pts=>pts.map(p=>p[0].toFixed(1)+','+p[1].toFixed(1)).join(' ');
const centroid=pts=>[pts.reduce((a,p)=>a+p[0],0)/pts.length,pts.reduce((a,p)=>a+p[1],0)/pts.length];
function inPoly(x,y,pts){let c=false;for(let i=0,j=pts.length-1;i<pts.length;j=i++){const[xi,yi]=pts[i],[xj,yj]=pts[j];if((yi>y)!==(yj>y)&&x<(xj-xi)*(y-yi)/(yj-yi)+xi)c=!c}return c}
function bbox(pts){const xs=pts.map(p=>p[0]),ys=pts.map(p=>p[1]);return[Math.min(...xs),Math.min(...ys),Math.max(...xs),Math.max(...ys)]}
function areaOf(pts){let a=0;for(let i=0,j=pts.length-1;i<pts.length;j=i++)a+=(pts[j][0]+pts[i][0])*(pts[j][1]-pts[i][1]);return Math.abs(a/2)}
const pct=v=>Math.round(v*100)+'%';
const hrs=v=>Math.round(v).toLocaleString()+' h';
function stage(k){return Math.min(5,Math.floor(k.pct*6.2))}
function skidStatus(k){return MILESTONES[stage(k)]}
function reqsFor(k){return k.uniqueReqs||STATE.partReqs[k.part]||[]}
function serialSlots(k){const out=[];reqsFor(k).forEach(r=>{for(let q=0;q<r.qty;q++)out.push({key:r.id+':'+q,r,q})});return out}
function snCount(k){const s=serialSlots(k);return[s.filter(x=>k.serials[x.key]&&k.serials[x.key].sn).length,s.length]}
function qcCount(k){return[QC_TEMPLATES.filter(t=>k.qc[t.id]&&k.qc[t.id].done).length,QC_TEMPLATES.length]}
function schedPill(act,plan){const d=act-plan;if(Math.abs(d)<.03)return'<span class="pill neutral">On plan</span>';return d>0?`<span class="pill good">Ahead ${Math.round(d*100)}%</span>`:`<span class="pill ${d<-.1?'bad':'warn'}">Behind ${Math.round(-d*100)}%</span>`}
function jobStats(jid){
  const ks=STATE.skids.filter(k=>k.jobId===jid);
  const est=ks.reduce((a,k)=>a+k.est,0),earned=ks.reduce((a,k)=>a+k.est*k.pct,0),act=ks.reduce((a,k)=>a+k.act,0),plan=ks.reduce((a,k)=>a+k.est*k.planPct,0);
  return{ks,est,earned,act,plan,pct:est?earned/est:0,planPct:est?plan/est:0,eff:act?earned/act:1,shops:[...new Set(ks.map(k=>k.shop))],lines:[...new Set(ks.map(k=>k.stationId&&k.stationId.split('-')[0]).filter(Boolean))]}
}
function toast(msg){const t=$('#toast');t.textContent=msg;t.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(()=>t.hidden=true,2400)}
const SRC=s=>`<span class="src" title="In production this comes from ${s==='IFS'?'IFS Cloud':s==='ACC'?'Autodesk Construction Cloud':'Microsoft 365'}">${s}</span>`;

/* ================= view / pan / zoom ================= */
const svg=$('#map'),world=$('#world'),stage$=$('#stage');
let fitS=1;
function vbOf(shop){return PLANS[shop].vb}
function fit(){const r=stage$.getBoundingClientRect();const[x,y,w,h]=vbOf(UI.shop);const pad=UI_PAD();fitS=Math.min((r.width-pad.l-pad.r)/w,(r.height-pad.t-pad.b)/h);UI.view.s=fitS;UI.view.tx=pad.l+(r.width-pad.l-pad.r-w*fitS)/2-x*fitS;UI.view.ty=pad.t+(r.height-pad.t-pad.b-h*fitS)/2-y*fitS;applyView()}
function UI_PAD(){return{l:16,r:16,t:16,b:stage$.clientWidth<600?16:96}}
const minS=()=>fitS*.5, maxS=()=>40;
function zoomAt(ns,sx,sy){ns=Math.max(minS(),Math.min(maxS(),ns));const v=UI.view;const wx=(sx-v.tx)/v.s,wy=(sy-v.ty)/v.s;v.s=ns;v.tx=sx-wx*ns;v.ty=sy-wy*ns;applyView()}
function centerOn(cx,cy,s){const r=stage$.getBoundingClientRect();UI.view.s=Math.max(minS(),Math.min(maxS(),s));UI.view.tx=r.width/2-cx*UI.view.s;UI.view.ty=r.height/2-cy*UI.view.s;applyView()}
function viewCenter(){const r=stage$.getBoundingClientRect();return[(r.width/2-UI.view.tx)/UI.view.s,(r.height/2-UI.view.ty)/UI.view.s]}
function applyView(){
  const v=UI.view;world.setAttribute('transform',`translate(${v.tx.toFixed(2)} ${v.ty.toFixed(2)}) scale(${v.s.toFixed(5)})`);
  svg.classList.toggle('z1',v.s<2.4);svg.classList.toggle('z2',v.s>=2.4&&v.s<5.5);svg.classList.toggle('z3',v.s>=5.5);
  const t=Math.log(v.s/minS())/Math.log(maxS()/minS());$('#zSlider').value=Math.round(t*1000);
  $('#zPct').textContent=Math.round(v.s/fitS*100)+'%';
  const target=90/v.s,nice=[1,2,5,10,20,25,50,100,200].find(n=>n>=target*.6)||200;$('#sbBar').style.width=(nice*v.s).toFixed(0)+'px';$('#sbTxt').textContent=nice+' ft';
  if(UI.edit||UI.measuring)renderOverlay();
}
function toWorld(e){const r=svg.getBoundingClientRect();return[(e.clientX-r.left-UI.view.tx)/UI.view.s,(e.clientY-r.top-UI.view.ty)/UI.view.s]}

/* ================= rendering ================= */
function renderPlan(){const p=PLANS[UI.shop];$('#plan').innerHTML=`<path class="plan-ink" d="${p.d}"/>`+(p.f?`<path class="plan-fill" d="${p.f}"/>`:'')}
function areaFill(a){if(a.type==='forklift')return'url(#pat-fork)';if(a.type==='quarantine')return'url(#pat-quar)';return AREA_TYPES[a.type]?AREA_TYPES[a.type].c:'#888'}
function renderLayers(){
  const L=UI.layers,E=els(),h=[];const R=role();
  const selJob=UI.sel.type==='job'?UI.sel.id:UI.sel.type==='skid'?(skidById(UI.sel.id)||{}).jobId:null;
  if(L.areas)for(const a of E.filter(e=>e.kind==='area')){const t=AREA_TYPES[a.type]||{c:'#888',label:a.type};const[cx,cy]=centroid(a.pts);const b=bbox(a.pts);const fs=Math.max(2.2,Math.min(5,(b[2]-b[0])/12));
    h.push(`<g data-kind="area" data-id="${a.id}"><polygon class="area${UI.editSel===a.id?' hl':''}" points="${ptsStr(a.pts)}" fill="${areaFill(a)}" fill-opacity="${a.type==='forklift'||a.type==='quarantine'?1:.14}" stroke="${t.c}"/>`+
    `<text class="lbl zt2 muted" x="${cx}" y="${cy}" font-size="${fs}" stroke-width="${fs*.25}" text-anchor="middle" dominant-baseline="middle">${esc(a.name||t.label)}</text></g>`)}
  if(L.lines)for(const l of E.filter(e=>e.kind==='line')){const b=bbox(l.pts);
    h.push(`<g data-kind="line" data-id="${l.id}"><polygon class="line-poly${UI.editSel===l.id?' hl':''}" points="${ptsStr(l.pts)}"/><text class="lbl" x="${b[0]+1}" y="${b[1]-1.6}" font-size="6" stroke-width="1.4">${esc(l.name)}</text></g>`)}
  if(L.stations)for(const s of E.filter(e=>e.kind==='station')){const[cx,cy]=centroid(s.pts);const b=bbox(s.pts);
    h.push(`<g data-kind="station" data-id="${s.id}"><polygon class="station${UI.editSel===s.id?' hl':''}" points="${ptsStr(s.pts)}"/><text class="lbl zt2 muted mono" x="${b[0]+1.2}" y="${b[1]+3.4}" font-size="2.6" stroke-width=".6">${esc(s.name)}</text></g>`)}
  if(L.equip)for(const q of E.filter(e=>e.kind==='equip')){const lb=libById(q.libId);const w=q.rot?lb.W:lb.L,hh=q.rot?lb.L:lb.W;
    h.push(`<g data-kind="equip" data-id="${q.id}"><rect class="equip${UI.editSel===q.id?' hl':''}" x="${q.x-w/2}" y="${q.y-hh/2}" width="${w}" height="${hh}" rx=".3" fill="${lb.c}" fill-opacity=".55" stroke="${lb.c}"/><text class="lbl zt3 muted" x="${q.x}" y="${q.y+hh/2+1.6}" font-size="1.3" stroke-width=".3" text-anchor="middle">${esc(q.name||lb.name)}</text></g>`)}
  if(L.material)for(const z of E.filter(e=>e.kind==='matzone')){if(z.jobId?!canSeeJob(z.jobId):R.company!=='RK')continue;const j=z.jobId&&STATE.jobs[z.jobId];const c=j?j.c:'#2d8f9a';const b=bbox(z.pts);
    h.push(`<g data-kind="matzone" data-id="${z.id}"><polygon class="matzone${UI.editSel===z.id||(UI.sel.type==='mz'&&UI.sel.id===z.id)?' hl':''}" points="${ptsStr(z.pts)}" fill="${c}" fill-opacity=".09" stroke="${c}"/><text class="lbl zt2" style="fill:${c}" x="${b[0]+.8}" y="${b[1]-.9}" font-size="2.4" stroke-width=".5">${esc(z.name)}${j?' · '+j.num:''}</text></g>`)}
  if(L.material)for(const p of E.filter(e=>e.kind==='stock')){if(p.jobId?!canSeeJob(p.jobId):R.company!=='RK')continue;const j=p.jobId&&STATE.jobs[p.jobId];const c=j?j.c:'#5d6a73';const w=p.rot?p.W:p.L,hh=p.rot?p.L:p.W;const hit=UI.matHit===p.pn;const sel=UI.editSel===p.id||(UI.sel.type==='stock'&&UI.sel.id===p.id);
    h.push(`<g class="stock" data-kind="stock" data-id="${p.id}">${hit?`<circle class="pulse" cx="${p.x}" cy="${p.y}" r="${Math.max(3,Math.max(w,hh)*.7)}"/>`:''}<rect class="stockr${hit?' hit':''}${sel?' hl':''}" x="${p.x-w/2}" y="${p.y-hh/2}" width="${w}" height="${hh}" fill="${c}" fill-opacity="${hit?.75:.25}" stroke="${c}"/><text class="lbl zt3" x="${p.x}" y="${p.y}" font-size="${Math.max(.5,Math.min(1,hh*.4))}" stroke-width=".12" text-anchor="middle" dominant-baseline="middle">${esc(catDesc(p.pn))}</text></g>`)}
  if(L.flow)for(const f of E.filter(e=>e.kind==='flow')){
    h.push(`<g data-kind="flow" data-id="${f.id}"><polyline class="flow" points="${ptsStr(f.pts)}" stroke="${f.c||'#0f6fe0'}" stroke-width="${UI.editSel===f.id?4:2.2}" marker-end="url(#arrow)" marker-mid="url(#arrow)"/></g>`)}
  if(L.zones)for(const z of E.filter(e=>e.kind==='zone')){if(!canSeeJob(z.jobId))continue;const j=STATE.jobs[z.jobId];if(!j)continue;
    const cls=(selJob===z.jobId?' sel injob':'')+(UI.editSel===z.id?' sel':'');
    h.push(`<polygon class="zone${cls}" data-kind="zone" data-id="${z.id}" data-job="${z.jobId}" points="${ptsStr(z.pts)}" fill="${j.c}" fill-opacity=".14" stroke="${j.c}"/>`)}
  if(L.skids)for(const k of STATE.skids.filter(k=>k.shop===UI.shop)){
    const vis=canSeeJob(k.jobId),j=STATE.jobs[k.jobId]||{c:'#888'};const w=k.rot?k.W:k.L,hh=k.rot?k.L:k.W;
    const st=stage(k);const cls=(UI.sel.type==='skid'&&UI.sel.id===k.id?' sel':'')+(selJob===k.jobId?' injob':'')+(vis?'':' anon');
    const col=vis?j.c:'var(--muted)';
    if(!vis){h.push(`<g class="skid anon" data-kind="skid" data-id=""><rect class="body" x="${k.x-w/2}" y="${k.y-hh/2}" width="${w}" height="${hh}" rx=".4" fill="var(--hidden)" stroke="var(--hidden-line)"/></g>`);continue}
    h.push(`<g class="skid${cls}" data-kind="skid" data-id="${k.id}" data-job="${vis?k.jobId:''}">`+
      `<rect class="body" x="${k.x-w/2}" y="${k.y-hh/2}" width="${w}" height="${hh}" rx=".4" fill="var(--panel)" stroke="${col}"/>`+
      (vis?`<rect x="${k.x-w/2}" y="${k.y+hh/2-1.1}" width="${w*k.pct}" height="1.1" fill="${col}" pointer-events="none"/>`:'')+
      (vis?`<text class="lbl mono zt2" x="${k.x}" y="${k.y-(UI.view.s>5.5?.6:0)}" font-size="1.9" stroke-width=".25" text-anchor="middle" dominant-baseline="middle">${esc(k.tag)}</text>`+
      `<text class="lbl zt3 muted" x="${k.x}" y="${k.y+1.9}" font-size="1.25" stroke-width=".2" text-anchor="middle" dominant-baseline="middle">${fi(k.L)} × ${fi(k.W)} · ${MILESTONES[st]}</text>`:
      `<text class="lbl zt2 muted" x="${k.x}" y="${k.y}" font-size="1.8" stroke-width=".25" text-anchor="middle" dominant-baseline="middle">RK work</text>`)+`</g>`)}
  if(L.material&&role().company==='RK')for(const m of STATE.materials.filter(m=>m.shop===UI.shop)){
    const hit=UI.matHit===m.pn;
    h.push(`<g class="mat" data-kind="mat" data-id="${m.pn}">${hit?`<circle class="pulse" cx="${m.x}" cy="${m.y}" r="6"/>`:''}<path d="M${m.x} ${m.y-1.8}L${m.x+1.8} ${m.y}L${m.x} ${m.y+1.8}L${m.x-1.8} ${m.y}Z" fill="${hit?'var(--bad)':'var(--panel)'}" stroke="var(--bad)" stroke-width="1.2"/><text class="lbl zt3 mono" x="${m.x+2.4}" y="${m.y+.5}" font-size="1.4" stroke-width=".3">${esc(m.bin)}</text></g>`)}
  $('#layers').innerHTML=h.join('');
  $('#layers').classList.toggle('dim',!!selJob);
  svg.classList.toggle('nolabels',!UI.labels);
  if(UI.hover)applyHover(UI.hover);
  renderOverlay();
}
function renderOverlay(){
  const h=[],s=UI.view.s;
  if(UI.edit&&UI.editSel){const e=elById(UI.editSel);if(e&&e.pts)e.pts.forEach((p,i)=>h.push(`<circle class="handle" data-kind="handle" data-i="${i}" cx="${p[0]}" cy="${p[1]}" r="${5/s}"/>`))}
  if(UI.draft){const d=UI.draft;const pts=d.cursor?[...d.pts,d.cursor]:d.pts;
    if(d.kind==='flow')h.push(`<polyline class="draft" fill="none" points="${ptsStr(pts)}"/>`);else h.push(`<polygon class="draft" points="${ptsStr(pts)}"/>`);
    d.pts.forEach(p=>h.push(`<circle class="handle" cx="${p[0]}" cy="${p[1]}" r="${4/s}"/>`));
    if(d.cursor&&d.pts.length){const a=d.pts[d.pts.length-1];const L=Math.hypot(d.cursor[0]-a[0],d.cursor[1]-a[1]);h.push(`<text class="lbl mono" x="${d.cursor[0]+8/s}" y="${d.cursor[1]-8/s}" font-size="${12/s}" stroke-width="${3/s}">${fi(L)}</text>`)}}
  if(UI.measure){const m=UI.measure;const b=m.b||m.cur;if(b){const L=Math.hypot(b[0]-m.a[0],b[1]-m.a[1]);
    h.push(`<g class="measure"><line x1="${m.a[0]}" y1="${m.a[1]}" x2="${b[0]}" y2="${b[1]}" style="vector-effect:non-scaling-stroke"/><circle cx="${m.a[0]}" cy="${m.a[1]}" r="${3/s}"/><circle cx="${b[0]}" cy="${b[1]}" r="${3/s}"/><text class="lbl mono" x="${(m.a[0]+b[0])/2+6/s}" y="${(m.a[1]+b[1])/2-6/s}" font-size="${13/s}" stroke-width="${3/s}">${fi(L)}</text></g>`)}}
  $('#overlay').innerHTML=h.join('');
}

/* hover */
function hoverKeyOf(t){const g=t&&t.closest&&t.closest('[data-kind]');if(!g||g.dataset.kind==='handle')return null;if(g.classList.contains('anon'))return null;return{kind:g.dataset.kind,id:g.dataset.id,job:g.dataset.job}}
function applyHover(h){
  document.querySelectorAll('#layers .hl').forEach(n=>{if(!(UI.edit&&n.closest('[data-id="'+UI.editSel+'"]')))n.classList.remove('hl')});
  if(!h)return;
  if((h.kind==='zone'||h.kind==='skid')&&h.job){document.querySelectorAll(`#layers [data-job="${h.job}"]`).forEach(n=>{if(n.classList.contains('zone')||n.dataset.id===h.id)n.classList.add('hl')})}
  else{const g=document.querySelector(`#layers [data-kind="${h.kind}"][data-id="${CSS.escape(h.id)}"]`);if(g){(g.querySelector('polygon,rect,polyline')||g).classList.add('hl')}}
}
function hoverText(h){
  if(!h)return'';
  if(h.kind==='zone'){const j=STATE.jobs[h.job];const s=jobStats(h.job);return`<b>${esc(j.num)} · ${esc(j.customer)}</b><br>${esc(j.project)} · ${pct(s.pct)} complete`}
  if(h.kind==='skid'){const k=skidById(h.id);if(!k)return'';return`<b>${esc(k.tag)}</b> · ${esc(k.part)}<br>${stationLabel(k.stationId)} · ${skidStatus(k)} · ${pct(k.pct)}`}
  if(h.kind==='area'){const a=elById(h.id);return`<b>${esc(a.name)}</b><br>${AREA_TYPES[a.type]?AREA_TYPES[a.type].label:''} · ${Math.round(areaOf(a.pts)).toLocaleString()} sq ft`}
  if(h.kind==='station'){const s=elById(h.id);const k=STATE.skids.find(k=>k.stationId===s.id);return`<b>${esc(s.lineId)} · ${esc(s.name)}</b><br>${k?(canSeeJob(k.jobId)?esc(k.tag):'Occupied'):'Open station'}`}
  if(h.kind==='line'){const l=elById(h.id);return`<b>${esc(l.name)}</b><br>${(l.stations||[]).length} stations`}
  if(h.kind==='equip'){const q=elById(h.id);const lb=libById(q.libId);return`<b>${esc(q.name||lb.name)}</b><br>${fi(lb.L)} × ${fi(lb.W)}`}
  if(h.kind==='flow'){const f=elById(h.id);return`<b>${esc(f.name||'Layout line')}</b>`}
  if(h.kind==='matzone'){const z=elById(h.id);const c=zoneContents(z);return`<b>${esc(z.name)}</b>${z.jobId?' · Job '+esc(STATE.jobs[z.jobId].num):''}<br>${c.length} IFS part${c.length===1?'':'s'} · ${esc(z.loc||'no IFS location yet')}`}
  if(h.kind==='stock'){const p=elById(h.id);return`<b>${esc(catDesc(p.pn))}</b><br><span class="mono">${esc(p.pn)}</span> · qty ${p.qty}${p.jobId?' · Job '+esc(STATE.jobs[p.jobId].num):''}`}
  if(h.kind==='mat'){const m=STATE.materials.find(m=>m.pn===h.id);return`<b>${esc(m.bin)}</b> · ${esc(m.pn)}<br>${esc(m.desc)} · qty ${m.qty}`}
  return'';
}

/* ================= right panel ================= */
function crumbs(){
  const parts=[`<button data-act="home">${esc(STATE.shops[UI.shop].name)}</button>`];const sel=UI.sel;
  if(sel.type==='job')parts.push(`<span>›</span><span class="cur">Job ${esc(STATE.jobs[sel.id].num)}</span>`);
  if(sel.type==='skid'){const k=skidById(sel.id);parts.push(`<span>›</span><button data-act="job" data-id="${k.jobId}">Job ${esc(STATE.jobs[k.jobId].num)}</button><span>›</span><span class="cur">${esc(k.tag)}</span>`)}
  if(sel.type==='part')parts.push(`<span>›</span><span class="cur">Part ${esc(sel.id)}</span>`);
  if(sel.type==='el'||sel.type==='mz'||sel.type==='stock'){const e=elById(sel.id);parts.push(`<span>›</span><span class="cur">${esc(e?(e.name||kindLabel(e)):'')}</span>`)}
  if(sel.type==='mat')parts.push(`<span>›</span><span class="cur">Material ${esc(sel.id)}</span>`);
  if(sel.type!=='shop')parts.push(`<button class="back" data-act="up" title="Back (Esc)">← Back</button>`);
  return`<nav class="crumbs" aria-label="Breadcrumb">${parts.join('')}</nav>`;
}
function kindLabel(e){return{area:'Area',line:'Production line',station:'Station',flow:'Layout line',zone:'Customer zone',equip:'Equipment',matzone:'Material zone',stock:'Material'}[e.kind]||e.kind}
function renderPanel(){
  const P=$('#panel');const sel=UI.sel;let h='';
  if(UI.edit&&UI.editSel)h=panelEditEl();
  else if(sel.type==='job')h=panelJob(sel.id);
  else if(sel.type==='skid')h=panelSkid(sel.id);
  else if(sel.type==='el')h=panelEl(sel.id);
  else if(sel.type==='mat')h=panelMat(sel.id);
  else if(sel.type==='mz')h=panelMz(sel.id);
  else if(sel.type==='stock')h=panelStock(sel.id);
  else if(sel.type==='part')h=panelPart(sel.id);
  else h=panelShop();
  const top=P.scrollTop;P.innerHTML=`<div class="pwrap">${h}</div>`;if(renderPanel.keep)P.scrollTop=top;renderPanel.keep=false;
}
function hoverCard(){const h=UI.hover;
  if(!h)return`<div class="hover-card empty" id="hoverCard">Move across the floor to preview a job, skid or area. Click a customer zone for job details.</div>`;
  return`<div class="hover-card" id="hoverCard">${hoverText(h)}</div>`}
function panelShop(){
  const sh=STATE.shops[UI.shop];const R=role();
  const ks=STATE.skids.filter(k=>k.shop===UI.shop&&canSeeJob(k.jobId));
  const jobs=[...new Set(ks.map(k=>k.jobId))];
  let h=`<div><p class="eyebrow">Shop overview</p><h2>${esc(sh.name)}</h2><p class="sub">${esc(sh.addr)}</p></div>`;
  h+=`<div class="sec">${hoverCard()}</div>`;
  if(R.plan){
    const lines=els().filter(e=>e.kind==='line');
    let dP=0,dA=0,wP=0,wA=0;lines.forEach(l=>{const p=STATE.plan[l.id];if(p){dP+=p.day;dA+=p.dayAct;wP+=p.week;wA+=p.weekAct}});
    const mx=Math.max(...lines.map(l=>STATE.plan[l.id]?Math.max(STATE.plan[l.id].day,STATE.plan[l.id].dayAct):1));
    h+=`<div class="sec"><h3>Production vs plan ${SRC('IFS')}</h3>
      <div class="tiles"><div class="tile"><span>Today, earned</span><b>${dA} h</b><span>plan ${dP} h</span></div>
      <div class="tile"><span>Today</span><b>${dA>=dP?'+':'−'}${Math.abs(Math.round((dA/dP-1)*100))}%</b><span>${dA>=dP?'ahead':'behind'}</span></div>
      <div class="tile"><span>This week</span><b>${wA>=wP?'+':'−'}${Math.abs(Math.round((wA/wP-1)*100))}%</b><span>${wA} / ${wP} h</span></div></div>
      <div class="list" style="gap:6px">${lines.map(l=>{const p=STATE.plan[l.id]||{day:0,dayAct:0};const ok=p.dayAct>=p.day;return`<div class="plan-row"><span class="n">${esc(l.name)}</span><div class="bar" title="Plan ${p.day} h, earned ${p.dayAct} h"><i style="width:${p.dayAct/mx*100}%;--c:${ok?'var(--good)':'var(--warn)'}"></i><s style="left:${p.day/mx*100}%"></s></div><span class="v">${p.dayAct} / ${p.day} h</span></div>`}).join('')}</div>
      <p class="note">Bars are earned labor hours today; the tick is today’s plan.</p></div>`;
  }
  h+=`<div class="sec"><h3>Jobs on this floor</h3><div class="list">${jobs.map(jid=>{const j=STATE.jobs[jid],s=jobStats(jid);
    return`<button class="li" data-act="job" data-id="${jid}"><span class="sw" style="--c:${j.c}"></span><span class="t"><div><b>${esc(j.num)}</b> · ${esc(j.customer)}</div><small>${esc(j.project)}</small></span><span class="mono">${pct(s.pct)}</span></button>`}).join('')||'<p class="note">No jobs visible to you on this floor.</p>'}</div></div>`;
  h+=`<div class="sec"><h3>Floor marking key</h3><div class="legend">${Object.values(AREA_TYPES).map(t=>`<span><span class="sw" style="--c:${t.c}"></span>${t.label}</span>`).join('')}</div>
    <p class="note">Area colors follow common 5S floor-tape practice: yellow for traffic, red for hold and defects, blue and teal for material, orange for inspection.</p></div>`;
  return crumbs()+h;
}
function panelJob(jid){
  const j=STATE.jobs[jid];const s=jobStats(jid);const R=role();
  const pm=teamNames(j,'pm')[0];
  const mz=STATE.elements.filter(e=>e.kind==='matzone'&&e.jobId===jid),pcs=STATE.elements.filter(e=>e.kind==='stock'&&e.jobId===jid);
  const nTeam=teamRoles().reduce((a,[id])=>a+teamOf(j,id).length,0);
  const tabs=[['overview','Overview'],['team',`Team ${nTeam}`],['skids',`Skids ${s.ks.length}`],['material',`Material ${mz.length}`]];
  if(R.company==='RK')tabs.push(['access','Access']),tabs.push(['links','Links']);
  if(!tabs.some(t=>t[0]===UI.jobTab))UI.jobTab='overview';
  let h=`<div><p class="eyebrow"><span class="sw" style="--c:${j.c}"></span>Job ${esc(j.num)} ${SRC('IFS')}</p><h2>${esc(j.project)}</h2><p class="sub">${esc(j.customer)}${pm?' · PM '+esc(pm):''}</p></div>`;
  h+=`<div class="tabs" role="tablist">${tabs.map(([t,l])=>`<button role="tab" aria-selected="${UI.jobTab===t}" data-act="jtab" data-id="${t}">${l}</button>`).join('')}</div>`;
  if(UI.jobTab==='overview'){
    h+=`<div class="sec"><div class="tiles">
    <div class="tile"><span>Complete</span><b>${pct(s.pct)}</b><span>plan ${pct(s.planPct)}</span></div>
    <div class="tile"><span>Earned / est.</span><b>${Math.round(s.earned).toLocaleString()}</b><span>of ${s.est.toLocaleString()} h</span></div>
    <div class="tile"><span>Efficiency</span><b>${s.eff.toFixed(2)}</b><span>${hrs(s.act)} actual</span></div>
    <div class="tile"><span>Skids</span><b>${s.ks.length}</b><span>${s.shops.map(x=>STATE.shops[x].name).join(' + ')}</span></div>
    <div class="tile"><span>Open NCRs</span><b>${j.ncr}</b><span>IFS</span></div>
    <div class="tile"><span>ACC issues</span><b>${j.issues}</b><span>open</span></div></div>
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">${schedPill(s.pct,s.planPct)}<span class="note">Ship ${esc(j.ship)} · Lines ${s.lines.join(', ')}</span></div></div>`;
    h+=`<div class="sec"><dl class="kv"><dt>IFS project</dt><dd class="mono">${esc(j.ifs)}</dd><dt>ACC project</dt><dd class="mono">${esc(j.acc)}</dd><dt>Part number</dt><dd class="mono">${esc(j.part)}</dd><dt>Material</dt><dd>${mz.length} zone${mz.length===1?'':'s'} · ${pcs.length} piece${pcs.length===1?'':'s'}</dd></dl></div>`;
    const key=teamRoles().filter(([id])=>['pm','pe','prod','sup'].includes(id)&&teamOf(j,id).length);
    if(key.length)h+=`<div class="sec"><h3>Key contacts</h3><dl class="kv">${key.map(([id,r])=>`<dt>${esc(r.label)}</dt><dd>${teamNames(j,id).map(esc).join(', ')}</dd>`).join('')}</dl><p class="note"><button class="linkish" data-act="jtab" data-id="team">See the whole team</button></p></div>`;
  }
  if(UI.jobTab==='team'){
    const can=canTeamEdit(jid);
    h+=`<div class="sec">${can?'<p class="note">Add or remove people for each role. Admins can add new roles in Settings.</p>':''}<div class="team">${teamRoles().map(([id,r])=>{const ppl=teamNames(j,id);
      return`<div class="trow"><div class="trole">${esc(r.label)}</div><div class="tppl">${ppl.map((p,i)=>`<span class="person">${esc(p)}${can?`<button data-act="tdel" data-r="${id}" data-i="${i}" aria-label="Remove ${esc(p)}">×</button>`:''}</span>`).join('')||'<span class="note">Not assigned</span>'}
</div></div>`}).join('')}</div>
      ${can?`<form class="tadd"><select aria-label="Role">${teamRoles().map(([id,r])=>`<option value="${id}" ${UI.teamRole===id?'selected':''}>${esc(r.label)}</option>`).join('')}</select><input placeholder="Name" aria-label="Name" list="peopleList"><datalist id="peopleList">${Object.values(STATE.users).filter(u=>STATE.roles[u.role]&&STATE.roles[u.role].company==='RK').map(u=>`<option value="${esc(u.name)}">`).join('')}</datalist><button class="btn pri" type="submit">Add</button></form><p class="note">Pick someone from the list, or type a new name to add them as a ${esc((STATE.roles[UI.teamRole||'pm']||{}).label||'')}. Being on the team gives them access to this job.</p>`:''}
      <p class="note">Sample names. In production, the team comes from Microsoft 365 (Entra ID groups for the job) or the IFS project's resource list.</p></div>`;
  }
  if(UI.jobTab==='skids'){
    h+=`<div class="sec"><div class="list">${s.ks.map(k=>`<button class="li" data-act="skid" data-id="${k.id}"><span class="mono" style="font-size:12px">${esc(k.tag)}</span><span class="t"><div class="bar" style="--c:${j.c}"><i style="width:${k.pct*100}%"></i><s style="left:${k.planPct*100}%"></s></div><small>${STATE.shops[k.shop].name} · ${stationLabel(k.stationId)} · ${skidStatus(k)}</small></span><span class="mono">${pct(k.pct)}</span></button>`).join('')}</div>`;
    if(R.add)h+=UI.addSkid===jid?addSkidForm(jid):`<div class="btns"><button class="btn" data-act="addskid" data-id="${jid}">Add skid</button></div>`;
    h+=`</div>`;
  }
  if(UI.jobTab==='material'){
    h+=`<div class="sec"><h3>Material zones</h3><div class="list">${mz.map(z=>{const n=zoneContents(z).length;return`<button class="li" data-act="mz" data-id="${z.id}"><span class="sw" style="--c:${j.c}"></span><span class="t"><div>${esc(z.name)}</div><small>${STATE.shops[z.shop].name} · ${esc(z.loc||'')} · ${n} IFS part${n===1?'':'s'}</small></span><span>→</span></button>`}).join('')||'<p class="note">No material zones for this job yet.</p>'}</div></div>`;
    const byPn={};pcs.forEach(p=>{byPn[p.pn]=(byPn[p.pn]||0)+(p.qty||1)});const rows=Object.entries(byPn).sort((a,b)=>catDesc(a[0]).localeCompare(catDesc(b[0])));
    if(rows.length)h+=`<div class="sec"><h3>Pieces on the floor (${pcs.length})</h3><div class="tbl-wrap"><table><thead><tr><th>IFS part</th><th>Description</th><th>Qty</th></tr></thead><tbody>${rows.map(([pn,q])=>`<tr class="prow" data-act="part" data-id="${esc(pn)}"><td class="mono">${esc(pn)}</td><td>${esc(catDesc(pn))}</td><td class="mono">${q} ${catUom(pn)}</td></tr>`).join('')}</tbody></table></div></div>`;
  }
  if(UI.jobTab==='access'){
    const can=canGrant(jid);const why=accessWhy(jid);
    const teamIds=[...new Set(Object.values(j.team||{}).flat())];
    const chip=(label,act,id)=>`<span class="person">${esc(label)}${can&&act?`<button data-act="${act}" data-id="${id}" aria-label="Remove ${esc(label)}">×</button>`:''}</span>`;
    const allU=STATE.allJobs.users.map(uname),allG=STATE.allJobs.groups.map(g=>(STATE.groups[g]||{}).name).filter(Boolean);
    const rkUsers=Object.entries(STATE.users).filter(([,u])=>u.role!=='admin'&&STATE.roles[u.role]&&STATE.roles[u.role].company==='RK');
    h+=`<div class="sec"><p class="note">${why==='admin'?'You see every job as an administrator.':why==='team'?'You see this job because you are on its team.':why==='person'?'You were given access to this job.':why&&why.startsWith('group:')?`You see this job through the group <b>${esc(STATE.groups[why.slice(6)].name)}</b>.`:why==='all'?'You have an <b>All jobs</b> grant.':''} ${can?'You can grant access to this job.':'Only administrators, and PMs on this job, can change who has access.'}</p>
      <div class="team">
      <div class="trow"><div class="trole">Team (automatic)</div><div class="tppl">${teamIds.map(id=>chip(uname(id))).join('')||'<span class="note">Nobody yet</span>'}</div></div>
      <div class="trow"><div class="trole">Groups</div><div class="tppl">${j.access.groups.map(g=>chip((STATE.groups[g]||{name:'Removed group'}).name+' ('+((STATE.groups[g]||{members:[]}).members.length)+')','gdel',g)).join('')||'<span class="note">None</span>'}</div></div>
      <div class="trow"><div class="trole">People</div><div class="tppl">${j.access.users.map(u=>chip(uname(u),'udel',u)).join('')||'<span class="note">None</span>'}</div></div>
      <div class="trow"><div class="trole">All jobs grants</div><div class="tppl">${[...allG.map(n=>chip(n+' (group)')),...allU.map(n=>chip(n))].join('')||'<span class="note">None</span>'}<span class="note" style="width:100%">Plus administrators. Set in Settings.</span></div></div></div>
      ${can?`<form class="tadd" id="grantForm"><select aria-label="Who to grant">${Object.entries(STATE.groups).filter(([g])=>!j.access.groups.includes(g)).map(([g,x])=>`<option value="g:${g}">Group · ${esc(x.name)}</option>`).join('')}${rkUsers.filter(([id])=>!j.access.users.includes(id)&&!teamIds.includes(id)).sort((a,b)=>a[1].name.localeCompare(b[1].name)).map(([id,u])=>`<option value="u:${id}">${esc(u.name)} · ${esc(STATE.roles[u.role].label)}</option>`).join('')}</select><button class="btn pri" type="submit">Grant access</button></form>`:''}
      <p class="note">Customers only ever see their own company's jobs. In production, groups map to Entra ID security groups.</p></div>`;
  }
  if(UI.jobTab==='links'){
    h+=`<div class="sec"><h3>Collaboration ${SRC('M365')}</h3><div class="ms365">
    <div>Teams channel <b>${esc(j.teams)}</b><small>not connected yet</small></div>
    <div>Planner tasks <b>${j.planner[0]}</b><small>not connected yet</small></div></div></div>`;
    h+=`<div class="sec"><dl class="kv"><dt>IFS project</dt><dd class="mono">${esc(j.ifs)}</dd><dt>ACC project</dt><dd class="mono">${esc(j.acc)}</dd></dl></div>`;
  }
  h+=`<div class="sec"><p class="note">Pilot: labor hours, NCRs, ACC issues and Planner counts aren't connected yet. Jobs, skids, serials and access come from SharePoint.</p></div>`;
  return crumbs()+h;
}
function addSkidForm(jid){
  const j=STATE.jobs[jid];const p=PARTS[j.part];
  const open=els().filter(e=>e.kind==='station'&&!STATE.skids.some(k=>k.stationId===e.id));
  return`<form class="sec" id="addSkidForm" style="background:var(--panel-2);padding:10px;border-radius:8px">
    <h3>New skid for ${esc(j.num)}</h3>
    <div class="field"><label for="as-part">IFS part number</label><select id="as-part">${Object.keys(PARTS).map(k=>`<option ${k===j.part?'selected':''}>${k}</option>`).join('')}</select></div>
    <div class="grid4"><div class="field"><label for="as-lf">Length ft</label><input id="as-lf" type="number" min="1" value="${Math.floor(p.L)}"></div><div class="field"><label for="as-li">in</label><input id="as-li" type="number" min="0" max="11" value="${Math.round((p.L%1)*12)}"></div>
    <div class="field"><label for="as-wf">Width ft</label><input id="as-wf" type="number" min="1" value="${Math.floor(p.W)}"></div><div class="field"><label for="as-wi">in</label><input id="as-wi" type="number" min="0" max="11" value="${Math.round((p.W%1)*12)}"></div></div>
    <div class="field"><label for="as-st">Place at station (${esc(STATE.shops[UI.shop].name)})</label><select id="as-st"><option value="">Centre of current view (unassigned)</option>${open.map(s=>`<option value="${s.id}">${s.lineId} · ${s.name}</option>`).join('')}</select></div>
    <div class="btns"><button class="btn pri" type="submit">Add skid</button><button class="btn" type="button" data-act="cancelskid">Cancel</button></div></form>`;
}
function panelSkid(id){
  const k=skidById(id);if(!k)return panelShop();const j=STATE.jobs[k.jobId];const R=role();const p=PARTS[k.part]||{desc:''};
  const st=stage(k);const[qd,qt]=qcCount(k);const[sd,stt]=snCount(k);const board=STATE.shops[k.shop].board;
  let h=`<div><p class="eyebrow"><span class="sw" style="--c:${j.c}"></span>${esc(j.customer)} · Job ${esc(j.num)}</p><h2 class="mono" style="font-family:var(--f-mono);font-size:22px">${esc(k.tag)}</h2><p class="sub">${esc(p.desc)} · ${stationLabel(k.stationId)}</p></div>`;
  h+=`<div class="tabs" role="tablist">${[['overview','Overview'],['qc',`QC ${qd}/${qt}`],['sn',`Serials ${sd}/${stt}`],['hist','History']].map(([t,l])=>`<button role="tab" aria-selected="${UI.tab===t}" data-act="tab" data-id="${t}">${l}</button>`).join('')}</div>`;
  if(UI.tab==='overview'){
    h+=`<div class="sec"><div class="steps">${MILESTONES.map((m,i)=>`<span class="${i<st?'d':i===st?'c':''}" title="${m}"></span>`).join('')}</div><div class="steplbl"><span>Received</span><span>Now: ${MILESTONES[st]}</span><span>Ship</span></div></div>`;
    h+=`<div class="sec"><div class="tiles"><div class="tile"><span>Complete</span><b>${pct(k.pct)}</b><span>plan today ${pct(k.planPct)}</span></div>
      <div class="tile"><span>Labor</span><b>${k.act}</b><span>of ${k.est} h est.</span></div>
      <div class="tile"><span>Schedule</span><b>${k.pct>=k.planPct?'+':'−'}${(Math.abs(k.pct-k.planPct)*k.est/40).toFixed(1)}</b><span>crew-days</span></div></div>
      <div>${schedPill(k.pct,k.planPct)}</div></div>`;
    h+=`<div class="sec"><dl class="kv"><dt>IFS part number ${SRC('IFS')}</dt><dd class="mono">${esc(k.part)}</dd><dt>Shop order ${SRC('IFS')}</dt><dd class="mono">${esc(k.so)}</dd>
      <dt>Footprint</dt><dd>${fi(k.L)} × ${fi(k.W)} (${Math.round(k.L*k.W)} sq ft)</dd><dt>Location</dt><dd>${STATE.shops[k.shop].name} · ${stationLabel(k.stationId)}</dd>
      <dt>QC checklists ${SRC('ACC')}</dt><dd>${qd} of ${qt} submitted</dd><dt>Serial numbers</dt><dd>${sd} of ${stt} recorded</dd>
      <dt>Whiteboard magnet</dt><dd>${fin(k.L*12/(board.ftPerIn*12))} × ${fin(k.W*12/(board.ftPerIn*12))} at 1" = ${board.ftPerIn}'</dd></dl></div>`;
    if(R.move)h+=`<div class="sec"><div class="btns"><button class="btn" data-act="rotskid">Rotate 90°</button><button class="btn" data-act="movemode">${UI.move?'Moving is on: drag the skid':'Turn on Move skids'}</button></div><p class="note">With Move skids on, drag a skid onto a station and it snaps in place. Its QC, serials and history move with it.</p></div>`;
  }
  if(UI.tab==='qc'){
    const can=R.qc;
    h+=`<div class="sec">${can?'':'<p class="note">Read only. QC inspectors and administrators can fill in checklists.</p>'}${QC_TEMPLATES.map(t=>{const q=k.qc[t.id]||{res:t.items.map(()=>null)};const fails=q.res.filter(r=>r==='f').length;const filled=q.res.filter(Boolean).length;
      const pill=q.done?(fails?`<span class="pill bad">${fails} fail</span>`:'<span class="pill good">Submitted</span>'):filled?`<span class="pill warn">${filled}/${t.items.length}</span>`:`<span class="pill neutral">${t.stage<=st?'Due':'Later'}</span>`;
      return`<details class="qc" ${UI.openQc===t.id?'open':''} data-qc="${t.id}"><summary><b>${esc(t.name)}</b>${pill}</summary>${t.items.map((it,i)=>`<div class="qci"><span>${esc(it)}</span><span class="seg">${[['p','Pass'],['f','Fail'],['n','N/A']].map(([v,l])=>`<button class="${v}" aria-pressed="${q.res[i]===v}" data-act="qcset" data-t="${t.id}" data-i="${i}" data-v="${v}" ${can&&!q.done?'':'disabled'}>${l}</button>`).join('')}</span></div>`).join('')}
        <div class="qci" style="grid-template-columns:1fr">${q.done?`<span class="note">Submitted by ${esc(q.by)} on ${esc(q.at)} · ${'saved to SharePoint'}${can?` · <button class="btn" data-act="qcreopen" data-t="${t.id}" style="padding:2px 8px">Reopen</button>`:''}</span>`:(can?`<div class="btns"><button class="btn pri" data-act="qcsubmit" data-t="${t.id}" ${filled===t.items.length?'':'disabled'}>Submit</button><span class="note" style="align-self:center">${filled===t.items.length?'':'Answer every item to submit.'}</span></div>`:'<span class="note">Not submitted yet.</span>')}</div></details>`}).join('')}</div>`;
  }
  if(UI.tab==='sn'){
    const can=R.sn;const slots=serialSlots(k);
    h+=`<div class="sec"><p class="note">${k.uniqueReqs?'This skid has its own equipment list.':`Equipment list comes from part number <b class="mono">${esc(k.part)}</b>, so every ${esc(k.part)} skid shares it.`} ${can?'':'Read only for your role.'}</p>
      ${slots.map(({key,r,q})=>{const v=k.serials[key]||{};const ph=PHOTOS[k.id+'|'+key];return`<div class="sn"><div class="sn-head"><b>${esc(r.item)}${r.qty>1?` #${q+1}`:''}</b><small class="mono">${esc(r.model)}</small></div>
        <div class="sn-row"><input type="text" placeholder="Serial number" value="${esc(v.sn||'')}" data-sn="${esc(key)}" aria-label="Serial number for ${esc(r.item)}" ${can?'':'disabled'}>
        <label class="photo" title="${can?'Attach nameplate photo':'Photo'}">${ph?`<img src="${ph}" alt="Nameplate photo">`:(v.photo?'✓':'＋')}${can?`<input type="file" accept="image/*" capture="environment" data-photo="${esc(key)}" hidden>`:''}</label></div>
        ${v.sn?`<small class="note">Recorded by ${esc(v.by||'you')} · ${esc(v.at||'today')}</small>`:''}</div>`}).join('')||'<p class="note">No serialized equipment on this part number.</p>'}</div>`;
    if(R.reqs)h+=reqEditor(k);
  }
  if(UI.tab==='hist'){
    h+=`<div class="sec hist">${k.history.slice().reverse().map(e=>`<div><time>${esc(e.at)}</time><span>${esc(e.txt)}</span></div>`).join('')}</div>`;
  }
  return crumbs()+h;
}
function reqEditor(k){
  const list=reqsFor(k);
  return`<div class="sec"><h3>Serialized equipment list</h3>
    <div class="field"><label for="reqMode">Applies to</label><select id="reqMode"><option value="part" ${k.uniqueReqs?'':'selected'}>Every ${esc(k.part)} skid (part number)</option><option value="unique" ${k.uniqueReqs?'selected':''}>This skid only</option></select></div>
    <div class="tbl-wrap"><table><thead><tr><th>Item</th><th>Model</th><th>Qty</th><th></th></tr></thead><tbody>${list.map((r,i)=>`<tr><td><input value="${esc(r.item)}" data-req="${i}" data-f="item" aria-label="Item"></td><td><input value="${esc(r.model)}" data-req="${i}" data-f="model" aria-label="Model"></td><td><input type="number" min="1" value="${r.qty}" data-req="${i}" data-f="qty" aria-label="Quantity" style="min-width:44px"></td><td><button class="btn danger" data-act="reqdel" data-i="${i}" aria-label="Remove" style="padding:2px 8px">×</button></td></tr>`).join('')}</tbody></table></div>
    <div class="btns"><button class="btn" data-act="reqadd">Add item</button></div></div>`;
}
function panelEl(id){
  const e=elById(id);if(!e)return panelShop();
  let h=`<div><p class="eyebrow">${kindLabel(e)}</p><h2>${esc(e.name||(e.kind==='equip'?libById(e.libId).name:kindLabel(e)))}</h2></div><div class="sec">${hoverCard()}</div>`;
  if(e.kind==='line'){const st=(e.stations||[]).map(elById).filter(Boolean);const p=STATE.plan[e.id];
    h+=`<div class="sec"><h3>Stations</h3><div class="list">${st.map(s=>{const k=STATE.skids.find(k=>k.stationId===s.id);const vis=k&&canSeeJob(k.jobId);return`<button class="li" ${vis?`data-act="skid" data-id="${k.id}"`:''}><span class="mono">${esc(s.name)}</span><span class="t"><div>${k?(vis?esc(k.tag)+' · '+esc(STATE.jobs[k.jobId].customer):'Occupied'):'<span class="note">Open</span>'}</div></span><span class="mono">${vis?pct(k.pct):''}</span></button>`}).join('')}</div></div>`;
    if(p&&role().plan)h+=`<div class="sec"><dl class="kv"><dt>Today earned / plan</dt><dd>${p.dayAct} / ${p.day} h</dd><dt>Week earned / plan</dt><dd>${p.weekAct} / ${p.week} h</dd></dl>${schedPill(p.dayAct/p.day,1)}</div>`}
  if(e.kind==='area')h+=`<div class="sec"><dl class="kv"><dt>Type</dt><dd>${AREA_TYPES[e.type]?AREA_TYPES[e.type].label:e.type}</dd><dt>Area</dt><dd>${Math.round(areaOf(e.pts)).toLocaleString()} sq ft</dd></dl></div>`;
  if(e.kind==='station'){const k=STATE.skids.find(k=>k.stationId===e.id);h+=`<div class="sec"><p class="note">${k?'Occupied.':'Open station. A production manager can drag a skid here with Move skids on.'}</p></div>`}
  return crumbs()+h;
}
function panelMat(pn){
  const m=STATE.materials.find(m=>m.pn===pn);if(!m)return panelShop();
  return crumbs()+`<div><p class="eyebrow">Material location ${SRC('IFS')}</p><h2>${esc(m.desc)}</h2><p class="sub mono">${esc(m.pn)}</p></div>
  <div class="sec"><dl class="kv"><dt>Shop</dt><dd>${STATE.shops[m.shop].name}</dd><dt>IFS location</dt><dd class="mono">${esc(m.bin)}</dd><dt>On hand</dt><dd>${m.qty}</dd></dl>
  <p class="note">The pulsing marker shows the bin. In production this reads IFS inventory locations and maps each location code to a spot on the floor plan.</p></div>`;
}
function panelEditEl(){
  const e=elById(UI.editSel);if(!e){UI.editSel=null;return panelShop()}
  let h=`<div><p class="eyebrow">Editing · ${kindLabel(e)}</p><h2>${esc(e.name||kindLabel(e))}</h2></div><form class="sec" id="elForm">`;
  if(e.kind==='matzone')h+=`<div class="field"><label for="el-job2">Job</label><select id="el-job2"><option value="">Shop stock (no job)</option>${Object.entries(STATE.jobs).map(([k,j])=>`<option value="${k}" ${k===e.jobId?'selected':''}>${j.num} · ${esc(j.customer)}</option>`).join('')}</select></div><div class="field"><label for="el-loc">IFS location code</label><input id="el-loc" class="mono" value="${esc(e.loc||'')}" placeholder="e.g. MC1-E-07"></div>`;
  if(e.kind==='stock')h+=`<div class="field"><label for="el-pn">IFS part</label><select id="el-pn">${catOptions().replace(`value="${e.pn}"`,`value="${e.pn}" selected`)}</select></div><div class="grid2"><div class="field"><label for="el-qty">Quantity</label><input id="el-qty" type="number" min="0" value="${e.qty}"></div><div class="field"><label for="el-job2">Job</label><select id="el-job2"><option value="">Shop stock</option>${Object.entries(STATE.jobs).map(([k,j])=>`<option value="${k}" ${k===e.jobId?'selected':''}>${j.num}</option>`).join('')}</select></div><div class="field"><label for="el-L">Length ft</label><input id="el-L" type="number" step=".5" min=".5" value="${e.L}"></div><div class="field"><label for="el-W">Width ft</label><input id="el-W" type="number" step=".5" min=".5" value="${e.W}"></div></div>`;
  if(e.kind!=='zone'&&e.kind!=='equip'&&e.kind!=='stock')h+=`<div class="field"><label for="el-name">Name</label><input id="el-name" value="${esc(e.name||'')}"></div>`;
  if(e.kind==='area')h+=`<div class="field"><label for="el-type">Area type</label><select id="el-type">${Object.entries(AREA_TYPES).map(([k,t])=>`<option value="${k}" ${k===e.type?'selected':''}>${t.label}</option>`).join('')}</select></div>`;
  if(e.kind==='station')h+=`<div class="field"><label for="el-line">Production line</label><select id="el-line"><option value="">None</option>${els().filter(x=>x.kind==='line').map(l=>`<option value="${l.id}" ${l.id===e.lineId?'selected':''}>${esc(l.name)}</option>`).join('')}</select></div>`;
  if(e.kind==='zone')h+=`<div class="field"><label for="el-job">Job / customer</label><select id="el-job">${Object.entries(STATE.jobs).map(([k,j])=>`<option value="${k}" ${k===e.jobId?'selected':''}>${j.num} · ${esc(j.customer)}</option>`).join('')}</select></div>`;
  if(e.kind==='flow')h+=`<div class="field"><label for="el-c">Line color</label><input id="el-c" type="color" value="${e.c||'#0f6fe0'}"></div>`;
  if(e.kind==='equip')h+=`<div class="field"><label for="el-lib">Library item</label><select id="el-lib">${STATE.library.map(l=>`<option value="${l.id}" ${l.id===e.libId?'selected':''}>${esc(l.name)}</option>`).join('')}</select></div><div class="field"><label for="el-name">Label (optional)</label><input id="el-name" value="${esc(e.name||'')}"></div>`;
  h+=`</form><div class="sec"><div class="btns">${e.kind==='equip'||e.kind==='stock'?'<button class="btn" data-act="rotel">Rotate 90°</button>':''}<button class="btn danger" data-act="delel">Delete</button><button class="btn" data-act="desel">Done</button></div>
    <p class="note">${e.pts?'Drag the round handles to reshape it, or drag the shape to move it.':'Drag it to move it.'} Delete removes it; Ctrl+Z undoes.</p></div>`;
  if(e.pts)h+=`<div class="sec"><dl class="kv"><dt>Points</dt><dd>${e.pts.length}</dd>${e.kind==='flow'?'':`<dt>Area</dt><dd>${Math.round(areaOf(e.pts)).toLocaleString()} sq ft</dd>`}</dl></div>`;
  return h;
}


/* ================= build material (IFS parts, material zones) ================= */
const catDesc=pn=>(STATE.catalog&&STATE.catalog[pn]||{}).desc||pn;
const catUom=pn=>(STATE.catalog&&STATE.catalog[pn]||{}).uom||'EA';
function zoneContents(z){
  const by={};
  for(const p of STATE.elements)if(p.kind==='stock'&&p.shop===z.shop&&inPoly(p.x,p.y,z.pts)){const r=by[p.pn]=by[p.pn]||{pn:p.pn,qty:0,count:0,manual:false};r.qty+=p.qty||0;r.count++}
  const rows=Object.values(by);
  for(const l of z.parts||[])rows.push({pn:l.pn,qty:l.qty,count:0,manual:true,lineId:l.id});
  rows.forEach(r=>{r.desc=catDesc(r.pn);r.uom=catUom(r.pn)});
  return rows.sort((a,b)=>a.desc.localeCompare(b.desc));
}
function zonesAt(shop,x,y){return STATE.elements.filter(z=>z.kind==='matzone'&&z.shop===shop&&inPoly(x,y,z.pts))}
function partLocations(pn){
  const out=[];
  for(const p of STATE.elements)if(p.kind==='stock'&&p.pn===pn&&(!p.jobId||canSeeJob(p.jobId))){const z=zonesAt(p.shop,p.x,p.y)[0];out.push({shop:p.shop,x:p.x,y:p.y,qty:p.qty,where:z?z.name+(z.loc?' ('+z.loc+')':''):nearStation(p),id:p.id,jobId:p.jobId})}
  for(const z of STATE.elements)if(z.kind==='matzone'&&(!z.jobId||canSeeJob(z.jobId)))for(const l of z.parts||[])if(l.pn===pn){const[cx,cy]=centroid(z.pts);out.push({shop:z.shop,x:cx,y:cy,qty:l.qty,where:z.name+(z.loc?' ('+z.loc+')':'')+', listed',zone:z.id,jobId:z.jobId})}
  for(const m of STATE.materials)if(m.pn===pn)out.push({shop:m.shop,x:m.x,y:m.y,qty:m.qty,where:'Bin '+m.bin,bin:m.pn});
  return out;
}
function nearStation(p){const s=STATE.elements.filter(e=>e.kind==='station'&&e.shop===p.shop).map(s=>{const[cx,cy]=centroid(s.pts);return{s,d:Math.hypot(cx-p.x,cy-p.y)}}).sort((a,b)=>a.d-b.d)[0];return s&&s.d<40?'Beside '+stationLabel(s.s.id):'Open floor'}
function catOptions(){return Object.entries(STATE.catalog).sort((a,b)=>a[1].desc.localeCompare(b[1].desc)).map(([pn,c])=>`<option value="${pn}">${esc(c.desc)}</option>`).join('')}
function jobLine(jid){const j=jid&&STATE.jobs[jid];return j?`<span class="sw" style="--c:${j.c}"></span>Job ${esc(j.num)} · ${esc(j.customer)}`:'Shop stock (no job)'}
function panelMz(id){
  const z=elById(id);if(!z)return panelShop();const rows=zoneContents(z);const R=role();
  let h=`<div><p class="eyebrow">${jobLine(z.jobId)} ${SRC('IFS')}</p><h2>${esc(z.name)}</h2><p class="sub">IFS location <span class="mono">${esc(z.loc||'not set')}</span> · ${Math.round(areaOf(z.pts)).toLocaleString()} sq ft</p></div>`;
  h+=`<div class="sec"><h3>What's here (${rows.length} parts)</h3>${rows.length?`<div class="tbl-wrap"><table><thead><tr><th>IFS part</th><th>Description</th><th>Qty</th><th></th></tr></thead><tbody>${rows.map(r=>`<tr class="prow" data-act="part" data-id="${esc(r.pn)}" title="Show every location of this part"><td class="mono">${esc(r.pn)}</td><td>${esc(r.desc)}${r.manual?' <span class="src">listed</span>':''}</td><td class="mono">${r.qty} ${r.uom}</td><td>${r.manual&&R.mat?`<button class="btn danger" data-act="mzdel" data-id="${r.lineId}" aria-label="Remove line" style="padding:1px 7px">×</button>`:''}</td></tr>`).join('')}</tbody></table></div>`:'<p class="note">Nothing recorded here yet.</p>'}
    <p class="note">Pieces drawn inside the zone count automatically. Listed lines are parts stored here that aren't drawn on the map. Click a row to see everywhere that part is.</p></div>`;
  if(R.mat)h+=`<form class="sec" id="mzAdd"><h3>Add an IFS part to this zone</h3><div class="grid2"><div class="field"><label for="mz-pn">IFS part number</label><input id="mz-pn" list="catList" placeholder="e.g. 610-…" autocomplete="off" required><datalist id="catList">${catOptions()}</datalist></div><div class="field"><label for="mz-qty">Quantity</label><input id="mz-qty" type="number" min="1" value="1"></div></div><p class="note" id="mz-desc"></p><div class="btns"><button class="btn pri" type="submit">Add to zone</button></div></form>`;
  return crumbs()+h;
}
function panelStock(id){
  const p=elById(id);if(!p)return panelShop();const z=zonesAt(p.shop,p.x,p.y)[0];const locs=partLocations(p.pn);const total=locs.reduce((a,l)=>a+(l.qty||0),0);
  return crumbs()+`<div><p class="eyebrow">${jobLine(p.jobId)} ${SRC('IFS')}</p><h2>${esc(catDesc(p.pn))}</h2><p class="sub mono">${esc(p.pn)}</p></div>
  <div class="sec"><dl class="kv"><dt>Quantity here</dt><dd>${p.qty} ${catUom(p.pn)}</dd><dt>Material zone</dt><dd>${z?`<button class="btn" data-act="mz" data-id="${z.id}" style="padding:1px 8px">${esc(z.name)}</button>`:'None'}</dd><dt>Location</dt><dd>${esc(z&&z.loc?z.loc:nearStation(p))}</dd>
  <dt>Footprint</dt><dd>${fi(p.L)} × ${fi(p.W)}</dd><dt>On hand, all locations ${SRC('IFS')}</dt><dd>${total} ${catUom(p.pn)} in ${locs.length} spot${locs.length===1?'':'s'}</dd></dl>
  <div class="btns"><button class="btn pri" data-act="part" data-id="${esc(p.pn)}">Show every location</button></div>
  <p class="note">Pilot: part numbers and quantities are typed in here. Later they come from IFS inventory.</p></div>`;
}
function panelPart(pn){
  const locs=partLocations(pn);const total=locs.reduce((a,l)=>a+(l.qty||0),0);
  return crumbs()+`<div><p class="eyebrow">IFS part ${SRC('IFS')}</p><h2>${esc(catDesc(pn))}</h2><p class="sub mono">${esc(pn)} · ${catUom(pn)}</p></div>
  <div class="sec"><div class="tiles"><div class="tile"><span>On hand</span><b>${total}</b><span>${catUom(pn)}</span></div><div class="tile"><span>Locations</span><b>${locs.length}</b><span>pulsing on the map</span></div><div class="tile"><span>Shops</span><b>${new Set(locs.map(l=>l.shop)).size}</b><span>${[...new Set(locs.map(l=>STATE.shops[l.shop].name))].join(' + ')}</span></div></div></div>
  <div class="sec"><h3>Where it is</h3><div class="list">${locs.map((l,i)=>`<button class="li" data-act="goloc" data-id="${i}"><span class="mono">${l.qty}</span><span class="t"><div>${esc(l.where)}</div><small>${STATE.shops[l.shop].name}${l.jobId?' · Job '+esc(STATE.jobs[l.jobId].num):''}</small></span><span>→</span></button>`).join('')||'<p class="note">Not on any floor yet.</p>'}</div></div>`;
}

/* ================= edit bar ================= */
function renderEditBar(){
  const b=$('#editBar');b.hidden=!UI.edit;if(!UI.edit)return;
  const T=(id,label,c,kind,type)=>`<button class="tool" aria-pressed="${UI.tool===id}" data-tool="${id}" data-kind="${kind||''}" data-type="${type||''}"><i style="--c:${c}"></i>${label}</button>`;
  const mzJob=`<select id="eb-mzjob" aria-label="Job"><option value="">Shop stock (no job)</option>${Object.entries(STATE.jobs).map(([k,j])=>`<option value="${k}" ${k===UI.drawMzJob?'selected':''}>${j.num} · ${esc(j.customer)}</option>`).join('')}</select>`;
  const matTools=`<h4>Build material</h4>${T('matzone','Material zone','#2d8f9a','matzone')}${UI.tool==='matzone'?mzJob:''}${T('stock','Material piece','#5d6a73','stock')}${UI.tool==='stock'?`<select id="eb-pn" aria-label="IFS part">${catOptions().replace(`value="${UI.drawPn}"`,`value="${UI.drawPn}" selected`)}</select>`+mzJob:''}`;
  if(!role().edit){b.innerHTML=`<h4>Edit material</h4>${T('select','Select / move','var(--ink)')}${matTools}<p class="hint">${UI.tool==='stock'?'Click the floor to place a piece. R rotates it.':UI.tool==='matzone'?'Click to add corners. Double-click or Enter finishes; Esc cancels.':'Click a material zone or piece to edit it.'}</p>`;return}
  b.innerHTML=`<h4>Edit layout</h4>${T('select','Select / move','var(--ink)')}
    ${T('area','Area polygon',AREA_TYPES[UI.drawType]?AREA_TYPES[UI.drawType].c:'#888','area')}
    ${UI.tool==='area'?`<select id="eb-area" aria-label="Area type">${Object.entries(AREA_TYPES).map(([k,t])=>`<option value="${k}" ${k===UI.drawType?'selected':''}>${t.label}</option>`).join('')}</select>`:''}
    ${T('line','Production line','#39434b','line')}${T('station','Station','#8d97a3','station')}
    ${T('zone','Customer zone','#e0337a','zone')}
    ${UI.tool==='zone'?`<select id="eb-job" aria-label="Job">${Object.entries(STATE.jobs).map(([k,j])=>`<option value="${k}" ${k===UI.drawJob?'selected':''}>${j.num} · ${esc(j.customer)}</option>`).join('')}</select>`:''}
    ${T('flow','Layout line','#0f6fe0','flow')}
    ${T('equip','Place equipment','#9a7652','equip')}
    ${UI.tool==='equip'?`<select id="eb-lib" aria-label="Equipment">${STATE.library.map(l=>`<option value="${l.id}" ${l.id===UI.equipLib?'selected':''}>${esc(l.name)}</option>`).join('')}</select>`:''}
    ${matTools}
    <p class="hint">${UI.tool==='select'?'Click a shape to edit it. Middle-drag pans.':(UI.tool==='equip'||UI.tool==='stock')?'Click the floor to place. R rotates the selected item.':'Click to add points (snaps to 6"). Hold Shift for square angles. Double-click or Enter finishes; Backspace removes a point; Esc cancels.'}</p>`;
}

/* ================= interactions ================= */
let drag=null;
svg.addEventListener('pointerdown',e=>{
  if(e.button===1||(e.button===0&&e.altKey)){e.preventDefault();drag={mode:'pan',x:e.clientX,y:e.clientY,tx:UI.view.tx,ty:UI.view.ty};svg.setPointerCapture(e.pointerId);svg.classList.add('panning');return}
  if(e.button!==0)return;
  const w=toWorld(e);const g=e.target.closest('[data-kind]');const kind=g&&g.dataset.kind;
  if(UI.measuring){if(!UI.measure||UI.measure.b)UI.measure={a:w};else UI.measure.b=w;renderOverlay();return}
  if(UI.edit&&UI.tool!=='select'){
    if(UI.tool==='stock'){pushUndo();const q={id:uid('e'),shop:UI.shop,kind:'stock',pn:UI.drawPn||Object.keys(STATE.catalog)[0],x:snap(w[0]),y:snap(w[1]),L:4,W:4,rot:0,jobId:UI.drawMzJob||null,qty:1,name:''};STATE.elements.push(q);UI.editSel=q.id;save();renderLayers();renderPanel();return}
    if(UI.tool==='equip'){pushUndo();const q={id:uid('e'),shop:UI.shop,kind:'equip',libId:UI.equipLib,x:snap(w[0]),y:snap(w[1]),rot:0,name:''};STATE.elements.push(q);UI.editSel=q.id;save();renderLayers();renderPanel();return}
    const now=Date.now(),lc=UI.lastClick;UI.lastClick={t:now,x:e.clientX,y:e.clientY};
    if(UI.draft&&lc&&now-lc.t<400&&Math.hypot(e.clientX-lc.x,e.clientY-lc.y)<6){UI.lastClick=null;finishDraft();return}
    const p=snapPt(w,e.shiftKey&&UI.draft&&UI.draft.pts.length?UI.draft.pts[UI.draft.pts.length-1]:null);if(!UI.draft)UI.draft={kind:UI.tool,pts:[]};
    UI.draft.pts.push(p);renderOverlay();return}
  if(UI.edit&&kind==='handle'){const el=elById(UI.editSel);pushUndo();drag={mode:'vertex',el,i:+g.dataset.i};svg.setPointerCapture(e.pointerId);return}
  if(UI.edit&&g&&kind!=='skid'&&kind!=='mat'){const el=elById(g.dataset.id);if(el&&!role().edit&&el.kind!=='matzone'&&el.kind!=='stock'){drag={mode:'maybe',x:e.clientX,y:e.clientY,tx:UI.view.tx,ty:UI.view.ty,target:null};svg.setPointerCapture(e.pointerId);return}if(el){if(UI.editSel!==el.id){UI.editSel=el.id;renderLayers();renderPanel()}pushUndo();drag={mode:'moveel',el,start:w,orig:JSON.stringify(el.pts||[el.x,el.y]),moved:false};svg.setPointerCapture(e.pointerId);return}}
  if(UI.move&&role().move&&kind==='skid'){const k=skidById(g.dataset.id);if(k){pushUndo();drag={mode:'skid',k,start:w,ox:k.x,oy:k.y,moved:false};svg.setPointerCapture(e.pointerId);return}}
  drag={mode:'maybe',x:e.clientX,y:e.clientY,tx:UI.view.tx,ty:UI.view.ty,target:g};svg.setPointerCapture(e.pointerId);
});
svg.addEventListener('pointermove',e=>{
  const w=toWorld(e);
  if(drag){
    if(drag.mode==='maybe'&&Math.hypot(e.clientX-drag.x,e.clientY-drag.y)>4){drag.mode='pan';svg.classList.add('panning')}
    if(drag.mode==='pan'){UI.view.tx=drag.tx+e.clientX-drag.x;UI.view.ty=drag.ty+e.clientY-drag.y;applyView();return}
    if(drag.mode==='vertex'){drag.el.pts[drag.i]=snapPt(w,e.shiftKey?drag.el.pts[(drag.i-1+drag.el.pts.length)%drag.el.pts.length]:null);raf(renderLayers);return}
    if(drag.mode==='moveel'){const dx=snap(w[0]-drag.start[0]),dy=snap(w[1]-drag.start[1]);const o=JSON.parse(drag.orig);if(dx||dy)drag.moved=true;
      if(drag.el.pts)drag.el.pts=o.map(p=>[p[0]+dx,p[1]+dy]);else{drag.el.x=o[0]+dx;drag.el.y=o[1]+dy}raf(renderLayers);return}
    if(drag.mode==='skid'){drag.k.x=drag.ox+w[0]-drag.start[0];drag.k.y=drag.oy+w[1]-drag.start[1];drag.moved=true;raf(renderLayers);return}
    return;
  }
  if(UI.draft){UI.draft.cursor=snapPt(w,e.shiftKey&&UI.draft.pts.length?UI.draft.pts[UI.draft.pts.length-1]:null);raf(renderOverlay)}
  if(UI.measuring&&UI.measure&&!UI.measure.b){UI.measure.cur=w;raf(renderOverlay)}
  const h=hoverKeyOf(e.target);const key=h?h.kind+h.id:'';
  if(key!==(UI.hover?UI.hover.kind+UI.hover.id:'')){UI.hover=h;applyHover(h);const hc=$('#hoverCard');if(hc){hc.className='hover-card'+(h?'':' empty');hc.innerHTML=h?hoverText(h):'Move across the floor to preview a job, skid or area. Click a customer zone for job details.'}}
  const tip=$('#tip');if(h&&!(UI.sel.type==='shop'&&!UI.edit)){const r=stage$.getBoundingClientRect();tip.innerHTML=hoverText(h);tip.hidden=false;tip.style.left=Math.min(e.clientX-r.left+14,r.width-270)+'px';tip.style.top=(e.clientY-r.top+16)+'px'}else tip.hidden=true;
});
svg.addEventListener('pointerleave',()=>{$('#tip').hidden=true});
svg.addEventListener('pointerup',e=>{
  const d=drag;drag=null;svg.classList.remove('panning');if(!d)return;
  if(d.mode==='maybe')onClick(d.target);
  if(d.mode==='skid'){if(d.moved)dropSkid(d.k,d.ox,d.oy);else{undo.pop();onClick(e.target.closest('[data-kind]'))}}
  if(d.mode==='moveel'||d.mode==='vertex'){if(d.mode==='moveel'&&!d.moved)undo.pop();save();renderLayers();renderPanel()}
});
svg.addEventListener('dblclick',e=>e.preventDefault());
svg.addEventListener('auxclick',e=>{if(e.button===1)e.preventDefault()});
svg.addEventListener('wheel',e=>{e.preventDefault();const r=svg.getBoundingClientRect();const f=Math.exp(-e.deltaY*(e.deltaMode===1?.05:.0015));zoomAt(UI.view.s*f,e.clientX-r.left,e.clientY-r.top)},{passive:false});
let rafQ=new Set();function raf(fn){if(rafQ.has(fn))return;rafQ.add(fn);requestAnimationFrame(()=>{rafQ.delete(fn);fn()})}
const snap=v=>Math.round(v*2)/2;
function snapPt(w,from){let p=[snap(w[0]),snap(w[1])];if(from){if(Math.abs(p[0]-from[0])>Math.abs(p[1]-from[1]))p[1]=from[1];else p[0]=from[0]}return p}

function onClick(g){
  if(UI.edit){UI.editSel=null;renderLayers();renderPanel();return}
  const kind=g&&g.dataset.kind;
  if(!g||g.classList.contains('anon')){if(UI.matHit){UI.matHit=null}select({type:'shop'});return}
  if(kind==='zone')return select({type:'job',id:g.dataset.job});
  if(kind==='skid')return select({type:'skid',id:g.dataset.id});
  if(kind==='mat'){UI.matHit=g.dataset.id;return select({type:'mat',id:g.dataset.id})}
  if(kind==='matzone')return select({type:'mz',id:g.dataset.id});
  if(kind==='stock')return select({type:'stock',id:g.dataset.id});
  if(kind==='station'){const k=STATE.skids.find(k=>k.stationId===g.dataset.id);if(k&&canSeeJob(k.jobId))return select({type:'skid',id:k.id})}
  select({type:'el',id:g.dataset.id});
}
function select(sel){if(sel.type==='skid'&&UI.sel.id!==sel.id)UI.tab='overview';if(sel.type==='job'&&!(UI.sel.type==='job'&&UI.sel.id===sel.id)&&UI.sel.type!=='skid')UI.jobTab='overview';UI.sel=sel;UI.addSkid=false;$('#tip').hidden=true;renderLayers();renderPanel()}
function up(){const s=UI.sel;if(s.type==='skid'){const k=skidById(s.id);select({type:'job',id:k.jobId})}else{UI.matHit=null;select({type:'shop'})}}
function dropSkid(k,ox,oy){
  const st=els().filter(e=>e.kind==='station').find(s=>inPoly(k.x,k.y,s.pts));
  const occupied=st&&STATE.skids.find(o=>o!==k&&o.stationId===st.id);
  const from=k.stationId;
  if(st&&occupied){k.x=ox;k.y=oy;undo.pop();toast(`${st.lineId} · ${st.name} already holds ${canSeeJob(occupied.jobId)?occupied.tag:'a skid'}. Move that one first.`);renderLayers();return}
  if(st){const[cx,cy]=centroid(st.pts);k.x=cx;k.y=cy;k.stationId=st.id;k.rot=st.orient==='v'?90:k.rot}else{k.stationId=null;k.x=snap(k.x);k.y=snap(k.y)}
  if(from!==k.stationId){k.history.push({at:new Date().toISOString().slice(0,10),txt:`Moved from ${stationLabel(from)} to ${stationLabel(k.stationId)} by ${role().label}`});toast(`${k.tag} moved to ${stationLabel(k.stationId)}`)}
  save();renderLayers();renderPanel();
}
function finishDraft(){
  const d=UI.draft;if(!d)return;UI.draft=null;
  const need=d.kind==='flow'?2:3;
  // collapse duplicate trailing points created by double-click
  const pts=d.pts.filter((p,i,a)=>i===0||p[0]!==a[i-1][0]||p[1]!==a[i-1][1]);
  if(pts.length<need){toast(`Needs at least ${need} points.`);renderOverlay();return}
  pushUndo();
  const el={id:uid('e'),shop:UI.shop,kind:d.kind,pts,name:''};
  if(d.kind==='area'){el.type=UI.drawType;el.name=AREA_TYPES[UI.drawType].label}
  if(d.kind==='line'){const n=els().filter(e=>e.kind==='line').length+1;el.name='New line '+n;el.stations=[];el.id='NL'+Date.now().toString(36)}
  if(d.kind==='station'){const[cx,cy]=centroid(pts);const ln=els().filter(e=>e.kind==='line').find(l=>inPoly(cx,cy,l.pts));el.lineId=ln?ln.id:'';el.orient='h';
    if(ln){ln.stations=ln.stations||[];ln.stations.push(el.id);el.name='S'+ln.stations.length;el.id=ln.id+'-S'+ln.stations.length+'-'+Date.now().toString(36).slice(-3)}else el.name='Station'}
  if(d.kind==='zone'){el.jobId=UI.drawJob||Object.keys(STATE.jobs)[0]}
  if(d.kind==='matzone'){el.name='Material zone';el.jobId=UI.drawMzJob||null;el.loc='';el.parts=[]}
  if(d.kind==='flow'){el.name='Layout line';el.c='#0f6fe0'}
  STATE.elements.push(el);UI.editSel=el.id;save();renderLayers();renderPanel();
}
document.addEventListener('keydown',e=>{
  const typing=/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName);
  if(e.key==='Escape'){if(!$('#modal').hidden){closeModal();return}if(!$('#viewsMenu').hidden){toggleViews(false);return}
    if(UI.draft){UI.draft=null;renderOverlay();return}if(UI.measuring){UI.measure=null;renderOverlay();return}
    if(UI.edit&&UI.editSel){UI.editSel=null;renderLayers();renderPanel();return}if(UI.sel.type!=='shop')up();return}
  if(typing)return;
  if(e.key==='Enter'&&UI.draft){finishDraft();return}
  if(e.key==='Backspace'&&UI.draft){UI.draft.pts.pop();renderOverlay();e.preventDefault();return}
  if((e.key==='Delete'||e.key==='Backspace')&&UI.edit&&UI.editSel){delEl();e.preventDefault();return}
  if((e.key==='r'||e.key==='R')){if(UI.edit&&UI.editSel){rotEl()}else if(UI.sel.type==='skid'&&role().move){rotSkid()}}
  if(e.key==='z'&&(e.ctrlKey||e.metaKey)&&undo.length){const s=JSON.parse(undo.pop());STATE.elements=s.e;STATE.skids=s.k;save();renderLayers();renderPanel();toast('Undone')}
  if(e.key==='+'||e.key==='=')zoomBy(1.25);if(e.key==='-')zoomBy(.8);
});
function zoomBy(f){const r=stage$.getBoundingClientRect();zoomAt(UI.view.s*f,r.width/2,r.height/2)}
function delEl(){const e=elById(UI.editSel);if(!e)return;pushUndo();STATE.elements=STATE.elements.filter(x=>x!==e);if(e.kind==='station'){const ln=elById(e.lineId);if(ln&&ln.stations)ln.stations=ln.stations.filter(s=>s!==e.id)}UI.editSel=null;save();renderLayers();renderPanel();toast(kindLabel(e)+' deleted. Ctrl+Z to undo.')}
function rotEl(){const e=elById(UI.editSel);if(e&&(e.kind==='equip'||e.kind==='stock')){pushUndo();e.rot=e.rot?0:90;save();renderLayers()}}
function rotSkid(){const k=skidById(UI.sel.id);if(!k)return;k.rot=k.rot?0:90;k.history.push({at:new Date().toISOString().slice(0,10),txt:'Rotated 90°'});save();renderLayers();renderPanel()}

/* panel events */
$('#panel').addEventListener('click',e=>{
  const b=e.target.closest('[data-act]');if(!b)return;const a=b.dataset.act,id=b.dataset.id;
  if(a==='home'){UI.matHit=null;select({type:'shop'})}
  if(a==='up')up();
  if(a==='job'){select({type:'job',id});focusJob(id)}
  if(a==='skid'){const k=skidById(id);if(k.shop!==UI.shop)setShop(k.shop,true);select({type:'skid',id});centerOn(k.x,k.y,Math.max(UI.view.s,5))}
  if(a==='tab'){UI.tab=id;renderPanel()}
  if(a==='jtab'){UI.jobTab=id;renderPanel()}
  if(a==='gdel'||a==='udel'){const j=STATE.jobs[UI.sel.id];const k=a==='gdel'?'groups':'users';j.access[k]=j.access[k].filter(x=>x!==id);save();renderLayers();renderPanel.keep=true;renderPanel()}
  if(a==='tdel'){const j=STATE.jobs[UI.sel.id];j.team[b.dataset.r].splice(+b.dataset.i,1);save();renderPanel.keep=true;renderPanel()}
  if(a==='addskid'){UI.addSkid=id;renderPanel.keep=true;renderPanel()}
  if(a==='cancelskid'){UI.addSkid=false;renderPanel.keep=true;renderPanel()}
  if(a==='rotskid')rotSkid();
  if(a==='movemode')setMove(!UI.move);
  if(a==='qcset'){const k=skidById(UI.sel.id);const t=QC_TEMPLATES.find(t=>t.id===b.dataset.t);k.qc[t.id]=k.qc[t.id]||{res:t.items.map(()=>null)};k.qc[t.id].res[+b.dataset.i]=b.dataset.v;UI.openQc=t.id;save();renderPanel.keep=true;renderPanel()}
  if(a==='qcsubmit'){const k=skidById(UI.sel.id);const q=k.qc[b.dataset.t];q.done=true;q.by=role().label;q.at=new Date().toISOString().slice(0,10);q.acc=true;k.history.push({at:q.at,txt:QC_TEMPLATES.find(t=>t.id===b.dataset.t).name+' submitted to ACC'});UI.openQc=b.dataset.t;save();renderPanel.keep=true;renderPanel();toast('Checklist submitted. Sample sync to ACC Forms.')}
  if(a==='qcreopen'){const k=skidById(UI.sel.id);k.qc[b.dataset.t].done=false;UI.openQc=b.dataset.t;save();renderPanel.keep=true;renderPanel()}
  if(a==='reqadd'||a==='reqdel'){const k=skidById(UI.sel.id);const list=editableReqs(k);if(a==='reqadd')list.push({id:uid('r'),item:'New item',model:'MODEL',qty:1});else list.splice(+b.dataset.i,1);save();renderPanel.keep=true;renderPanel()}
  if(a==='delel')delEl();
  if(a==='rotel')rotEl();
  if(a==='desel'){UI.editSel=null;renderLayers();renderPanel()}
  if(a==='part'){UI.matHit=id;UI.layers.material=true;renderRibbon();select({type:'part',id})}
  if(a==='mz'){const z=elById(id);if(z&&z.shop!==UI.shop)setShop(z.shop,true);select({type:'mz',id});if(z){const[cx,cy]=centroid(z.pts);centerOn(cx,cy,Math.max(UI.view.s,3))}}
  if(a==='goloc'){const l=partLocations(UI.sel.id)[+id];if(!l)return;if(l.shop!==UI.shop)setShop(l.shop,true);renderLayers();centerOn(l.x,l.y,Math.max(UI.view.s,6))}
  if(a==='mzdel'){const z=elById(UI.sel.id);z.parts=(z.parts||[]).filter(x=>x.id!==id);save();renderPanel.keep=true;renderPanel()}
});
$('#panel').addEventListener('toggle',e=>{if(e.target.matches&&e.target.matches('details.qc')&&e.target.open)UI.openQc=e.target.dataset.qc},true);
function editableReqs(k){if(k.uniqueReqs)return k.uniqueReqs;return STATE.partReqs[k.part]=STATE.partReqs[k.part]||[]}
$('#panel').addEventListener('change',e=>{
  const t=e.target;const k=UI.sel.type==='skid'?skidById(UI.sel.id):null;
  if(t.dataset.sn!==undefined&&k){const v=t.value.trim();if(v)k.serials[t.dataset.sn]={sn:v,by:role().label,at:new Date().toISOString().slice(0,10)};else delete k.serials[t.dataset.sn];save();renderPanel.keep=true;renderPanel();return}
  if(t.dataset.photo!==undefined&&k&&t.files[0]){const fr=new FileReader();fr.onload=()=>{PHOTOS[k.id+'|'+t.dataset.photo]=fr.result;renderPanel.keep=true;renderPanel();window.LIVE.uploadPhoto(k,t.dataset.photo,t.files[0])};fr.readAsDataURL(t.files[0]);return}
  if(t.id==='reqMode'&&k){if(t.value==='unique')k.uniqueReqs=JSON.parse(JSON.stringify(STATE.partReqs[k.part]||[]));else k.uniqueReqs=null;save();renderPanel.keep=true;renderPanel();return}
  if(t.dataset.req!==undefined&&k){const list=editableReqs(k);const r=list[+t.dataset.req];r[t.dataset.f]=t.dataset.f==='qty'?Math.max(1,+t.value||1):t.value;save();renderPanel.keep=true;renderPanel();return}
  if(t.id==='mz-pn'){const d=$('#mz-desc');if(d)d.textContent=STATE.catalog[t.value]?catDesc(t.value)+' · '+catUom(t.value):'';return}
  if(t.id==='as-part'){const p=PARTS[t.value];$('#as-lf').value=Math.floor(p.L);$('#as-li').value=Math.round((p.L%1)*12);$('#as-wf').value=Math.floor(p.W);$('#as-wi').value=Math.round((p.W%1)*12);return}
  if(t.closest('#elForm')){const el=elById(UI.editSel);if(!el)return;pushUndo();
    if(t.id==='el-name')el.name=t.value;if(t.id==='el-type'){el.type=t.value;if(!el.name||Object.values(AREA_TYPES).some(a=>a.label===el.name))el.name=AREA_TYPES[t.value].label}
    if(t.id==='el-line'){const old=elById(el.lineId);if(old&&old.stations)old.stations=old.stations.filter(s=>s!==el.id);el.lineId=t.value;const ln=elById(t.value);if(ln){ln.stations=ln.stations||[];ln.stations.push(el.id)}}
    if(t.id==='el-job2')el.jobId=t.value||null;if(t.id==='el-loc')el.loc=t.value.trim();if(t.id==='el-pn')el.pn=t.value;if(t.id==='el-qty')el.qty=Math.max(0,+t.value||0);if(t.id==='el-L')el.L=Math.max(.5,+t.value||1);if(t.id==='el-W')el.W=Math.max(.5,+t.value||1);
    if(t.id==='el-job')el.jobId=t.value;if(t.id==='el-c')el.c=t.value;if(t.id==='el-lib')el.libId=t.value;
    save();renderLayers();renderPanel()}
});
$('#panel').addEventListener('submit',e=>{
  e.preventDefault();
  if(e.target.id==='grantForm'){const j=STATE.jobs[UI.sel.id];const v=e.target.querySelector('select').value;if(!v)return;const[t,id]=v.split(':');j.access[t==='g'?'groups':'users'].push(id);save();renderPanel.keep=true;renderPanel();toast('Access granted');return}
  if(e.target.classList&&e.target.classList.contains('tadd')){const j=STATE.jobs[UI.sel.id];const r=e.target.querySelector('select').value;UI.teamRole=r;const v=e.target.querySelector('input').value.trim();if(!v)return;let id=Object.keys(STATE.users).find(k=>STATE.users[k].name.toLowerCase()===v.toLowerCase());if(!id){id=uid('u');STATE.users[id]={name:v,role:r}}const list=(j.team[r]=j.team[r]||[]);if(!list.includes(id))list.push(id);save();renderPanel.keep=true;renderPanel();return}
  if(e.target.id==='mzAdd'){const z=elById(UI.sel.id);const raw=$('#mz-pn').value.trim();const pn=Object.keys(STATE.catalog).find(k=>k.toLowerCase()===raw.toLowerCase())||Object.keys(STATE.catalog).find(k=>raw.toLowerCase().startsWith(k.toLowerCase()));
    if(!pn){$('#mz-desc').textContent='That part number is not in the sample IFS catalogue. Pick one from the list.';return}
    z.parts=z.parts||[];const q=Math.max(1,+$('#mz-qty').value||1);const ex=z.parts.find(x=>x.pn===pn);if(ex)ex.qty+=q;else z.parts.push({id:uid('p'),pn,qty:q});save();renderPanel.keep=true;renderPanel();toast(`Added ${q} × ${catDesc(pn)}`);return}
  if(e.target.id!=='addSkidForm')return;
  const jid=UI.addSkid;const part=$('#as-part').value;const L=(+$('#as-lf').value||0)+(+$('#as-li').value||0)/12,W=(+$('#as-wf').value||0)+(+$('#as-wi').value||0)/12;
  if(L<=0||W<=0){toast('Enter a length and width.');return}
  const stId=$('#as-st').value;const st=stId?elById(stId):null;let x,y,rot=0;
  if(st){[x,y]=centroid(st.pts);rot=st.orient==='v'?90:0}else{[x,y]=viewCenter()}
  const n=STATE.skids.filter(k=>k.jobId===jid).length+1;
  const k=mkSkid({shop:UI.shop,jobId:jid,n,x,y,rot,stationId:st?st.id:null,pct:0,part,L,W});
  k.qc={};k.serials={};k.act=0;k.planPct=0;k.history=[{at:new Date().toISOString().slice(0,10),txt:`Added by ${role().label} at ${stationLabel(k.stationId)}`}];
  pushUndo();STATE.skids.push(k);UI.addSkid=false;save();select({type:'skid',id:k.id});centerOn(x,y,Math.max(UI.view.s,5));toast(`${k.tag} added`);
});
function focusJob(jid){const zs=STATE.elements.filter(z=>z.kind==='zone'&&z.jobId===jid);const here=zs.filter(z=>z.shop===UI.shop);const use=here.length?here:zs;if(!use.length)return;if(use[0].shop!==UI.shop)setShop(use[0].shop,true);
  const b=bbox(use.flatMap(z=>z.pts));const r=stage$.getBoundingClientRect();const s=Math.min(r.width/(b[2]-b[0]+60),r.height/(b[3]-b[1]+60));centerOn((b[0]+b[2])/2,(b[1]+b[3])/2,Math.min(Math.max(s,UI.view.s),8))}

/* ================= ribbon ================= */
function renderRibbon(){
  $('#shopSel').innerHTML=Object.entries(STATE.shops).map(([k,s])=>`<option value="${k}" ${k===UI.shop?'selected':''}>${esc(s.name)}</option>`).join('');
  $('#roleSel').innerHTML=Object.entries(STATE.roles).sort((a,b)=>(a[1].company==='RK'?0:1)-(b[1].company==='RK'?0:1)).map(([rid,r])=>{const us=Object.entries(STATE.users).filter(([,u])=>u.role===rid).sort((a,b)=>a[1].name.localeCompare(b[1].name));return us.length?`<optgroup label="${esc(r.label)}">${us.map(([id,u])=>`<option value="${id}" ${id===UI.user?'selected':''}>${esc(u.name)} · ${esc(r.short||r.label)}${u.company?' · '+esc(u.company):''}</option>`).join('')}</optgroup>`:''}).join('');
  $('#layerChips').innerHTML='<span class="lbl">Layers</span>'+LAYERS.map(([k,l,c])=>`<button class="chip" style="--c:${c}" aria-pressed="${UI.layers[k]}" data-layer="${k}"><i></i>${l}</button>`).join('');
  const R=role();$('#editBtn').hidden=!(R.edit||R.mat);$('#editBtn').textContent=R.edit?'Edit layout':'Edit material';$('#libBtn').hidden=!R.edit;$('#setBtn').hidden=!R.edit;$('#pdfBtn').hidden=R.company!=='RK';$('#moveBtn').hidden=!R.move;$('#searchBox').hidden=R.company!=='RK';
  $('#editBtn').setAttribute('aria-pressed',UI.edit);$('#moveBtn').setAttribute('aria-pressed',UI.move);document.querySelector('.who').hidden=!window.LIVE.isAdmin();
}
$('#layerChips').addEventListener('click',e=>{const b=e.target.closest('[data-layer]');if(!b)return;UI.layers[b.dataset.layer]=!UI.layers[b.dataset.layer];b.setAttribute('aria-pressed',UI.layers[b.dataset.layer]);renderLayers()});
function setShop(s,keepView){UI.shop=s;UI.draft=null;UI.editSel=null;UI.measure=null;$('#shopSel').value=s;renderPlan();if(!keepView)fit();else{const r=stage$.getBoundingClientRect();const[x,y,w,h]=vbOf(s);fitS=Math.min((r.width-32)/w,(r.height-112)/h)}if(UI.sel.type!=='shop'&&UI.sel.type!=='job')UI.sel={type:'shop'};renderLayers();renderPanel()}
$('#shopSel').addEventListener('change',e=>{UI.sel={type:'shop'};setShop(e.target.value)});
$('#roleSel').addEventListener('change',e=>{UI.user=e.target.value;try{localStorage.setItem(KEY+'-user',UI.user)}catch(_){}setEdit(false);if(!role().move)setMove(false);
  if(UI.sel.type==='job'&&!canSeeJob(UI.sel.id))UI.sel={type:'shop'};if(UI.sel.type==='skid'&&!canSeeJob((skidById(UI.sel.id)||{}).jobId))UI.sel={type:'shop'};if(UI.sel.type==='mat'&&role().company!=='RK')UI.sel={type:'shop'};
  renderRibbon();renderLayers();renderPanel();const nj=Object.keys(STATE.jobs).filter(canSeeJob).length;toast(`Viewing as ${me().name}: ${nj} of ${Object.keys(STATE.jobs).length} jobs visible`)});
function setEdit(on){UI.edit=on;UI.draft=null;UI.editSel=null;UI.tool='select';if(on){setMove(false);UI.measuring=false;$('#measureBtn').setAttribute('aria-pressed','false')}$('#editBtn').setAttribute('aria-pressed',on);svg.classList.toggle('drawing',false);renderEditBar();renderLayers();renderPanel()}
function setMove(on){UI.move=on;$('#moveBtn').setAttribute('aria-pressed',on);if(on&&UI.edit)setEdit(false);renderPanel.keep=true;renderPanel();if(on)toast('Drag a skid onto a station. R rotates the selected skid.')}
$('#editBtn').addEventListener('click',()=>setEdit(!UI.edit));
$('#moveBtn').addEventListener('click',()=>setMove(!UI.move));
$('#editBar').addEventListener('click',e=>{const b=e.target.closest('[data-tool]');if(!b)return;UI.tool=b.dataset.tool;UI.draft=null;if(UI.tool==='zone'&&!UI.drawJob)UI.drawJob=Object.keys(STATE.jobs)[0];svg.classList.toggle('drawing',UI.tool!=='select');renderEditBar();renderOverlay()});
$('#editBar').addEventListener('change',e=>{if(e.target.id==='eb-mzjob')UI.drawMzJob=e.target.value;if(e.target.id==='eb-pn')UI.drawPn=e.target.value;if(e.target.id==='eb-area')UI.drawType=e.target.value;if(e.target.id==='eb-job')UI.drawJob=e.target.value;if(e.target.id==='eb-lib')UI.equipLib=e.target.value;renderEditBar()});

/* material search */
$('#matSearch').addEventListener('input',e=>{const q=e.target.value.trim().toLowerCase();const box=$('#matResults');if(q.length<2){box.hidden=true;return}
  const hits=Object.entries(STATE.catalog).filter(([pn,c])=>(pn+' '+c.desc).toLowerCase().includes(q)).slice(0,30);
  box.innerHTML=hits.length?hits.map(([pn,c])=>{const n=partLocations(pn).length;return`<button class="mi" data-pn="${esc(pn)}"><span class="mono">${esc(pn)}</span><span style="white-space:normal">${esc(c.desc)}</span><small>${n} spot${n===1?'':'s'}</small></button>`}).join(''):'<p class="hint">No IFS part matches. Try UPS, rack, 610-, or VFD.</p>';box.hidden=false});
$('#matResults').addEventListener('click',e=>{const b=e.target.closest('[data-pn]');if(!b)return;const pn=b.dataset.pn;$('#matResults').hidden=true;const l=partLocations(pn).sort((a,b)=>(a.shop===UI.shop?0:1)-(b.shop===UI.shop?0:1))[0];
  if(l&&l.shop!==UI.shop)setShop(l.shop,true);UI.layers.material=true;renderRibbon();UI.matHit=pn;select({type:'part',id:pn});if(l)centerOn(l.x,l.y,Math.max(UI.view.s,5))});
document.addEventListener('click',e=>{if(!e.target.closest('#searchBox'))$('#matResults').hidden=true;if(!e.target.closest('#viewsMenu')&&!e.target.closest('#viewsBtn'))toggleViews(false)});

/* ================= tray ================= */
$('#zIn').onclick=()=>zoomBy(1.3);$('#zOut').onclick=()=>zoomBy(1/1.3);$('#fitBtn').onclick=fit;
$('#zSlider').addEventListener('input',e=>{const t=e.target.value/1000;const s=minS()*Math.pow(maxS()/minS(),t);const r=stage$.getBoundingClientRect();zoomAt(s,r.width/2,r.height/2)});
$('#measureBtn').onclick=()=>{UI.measuring=!UI.measuring;UI.measure=null;$('#measureBtn').setAttribute('aria-pressed',UI.measuring);svg.classList.toggle('drawing',UI.measuring);renderOverlay();if(UI.measuring)toast('Click two points to measure. Esc clears.')};
$('#labelsBtn').onclick=()=>{UI.labels=!UI.labels;$('#labelsBtn').setAttribute('aria-pressed',UI.labels);renderLayers()};
$('#viewsBtn').onclick=()=>toggleViews($('#viewsMenu').hidden);
function toggleViews(open){const m=$('#viewsMenu');m.hidden=!open;$('#viewsBtn').setAttribute('aria-expanded',open);if(open)renderViews()}
function renderViews(){
  const R=role();const mine=STATE.views.filter(v=>v.scope==='personal'&&v.owner===UI.role),shared=STATE.views.filter(v=>v.scope==='shared');
  const item=v=>`<button class="mi" data-view="${v.id}"><span>${esc(v.name)}</span><small>${STATE.shops[v.shop].name}</small>${(v.owner===UI.role||(R.edit&&v.scope==='shared'))?`<span class="x" data-delview="${v.id}" role="button" aria-label="Delete view" title="Delete view">×</span>`:''}</button>`;
  $('#viewsMenu').innerHTML=`<h4>Shared views</h4>${shared.map(item).join('')||'<p class="hint">None yet.</p>'}<h4>My views</h4>${mine.map(item).join('')||'<p class="hint">Frame the floor the way you like, then save it here.</p>'}
    <form id="saveView"><label class="hint" for="vname" style="padding:0">Save current view</label><input id="vname" class="field" style="border:1px solid var(--line);border-radius:6px;padding:6px 8px;background:var(--panel)" placeholder="e.g. Line 5 morning walk" required>
    ${R.edit?'<label class="hint" style="padding:0;display:flex;gap:6px;align-items:center"><input type="checkbox" id="vshared"> Share with everyone</label>':''}<button class="btn pri" type="submit">Save view</button></form>`;
}
$('#viewsMenu').addEventListener('click',e=>{const d=e.target.closest('[data-delview]');if(d){e.stopPropagation();STATE.views=STATE.views.filter(v=>v.id!==d.dataset.delview);save();renderViews();return}
  const b=e.target.closest('[data-view]');if(!b)return;const v=STATE.views.find(v=>v.id===b.dataset.view);if(v.shop!==UI.shop){UI.sel={type:'shop'};setShop(v.shop,true)}
  if(v.layers)Object.assign(UI.layers,v.layers);renderRibbon();renderLayers();centerOn(v.cx,v.cy,v.s);toggleViews(false);toast('View: '+v.name)});
$('#viewsMenu').addEventListener('submit',e=>{e.preventDefault();const name=$('#vname').value.trim();if(!name)return;const[cx,cy]=viewCenter();
  STATE.views.push({id:uid('v'),name,shop:UI.shop,cx:Math.round(cx),cy:Math.round(cy),s:+UI.view.s.toFixed(3),layers:{...UI.layers},scope:$('#vshared')&&$('#vshared').checked?'shared':'personal',owner:UI.role});save();renderViews();toast('Saved view “'+name+'”')});

/* ================= modals ================= */
function openModal(title,body){const m=$('#modal');m.hidden=false;m.innerHTML=`<div class="scrim" data-close="1"><div class="modal" role="dialog" aria-modal="true" aria-label="${esc(title)}"><header><h3>${esc(title)}</h3><button class="btn" data-close="1">Close</button></header><div class="mbody">${body}</div></div></div>`}
function closeModal(){$('#modal').hidden=true;$('#modal').innerHTML='';renderLayers();renderPanel()}
$('#modal').addEventListener('click',e=>{if(e.target.dataset.close)closeModal();const b=e.target.closest('[data-mact]');if(!b)return;const a=b.dataset.mact;
  if(a==='libadd'){STATE.library.push({id:uid('l'),name:'New item',L:4,W:2,c:'#77808c'});save();libModal()}
  if(a==='libdel'){const id=b.dataset.id;if(STATE.elements.some(x=>x.libId===id)){b.textContent='In use';return}STATE.library=STATE.library.filter(l=>l.id!==id);save();libModal()}
  if(a==='suggest'){const sh=b.dataset.id;const bd=STATE.shops[sh].board;const need=boardFit(sh).need;const nice=[5,8,10,12,15,16,20,24,25,30,40,50].find(n=>n>=need)||Math.ceil(need);bd.ftPerIn=nice;save();setModal()}
  if(a==='flip'){const bd=STATE.shops[b.dataset.id].board;bd.rot=bd.rot?0:-90;save();setModal()}
  if(a==='pdf')exportPdf(b);
  if(a==='useradd'){const id=uid('u');STATE.users[id]={name:'New person',role:'tech'};save();renderRibbon();setModal();setTimeout(()=>{const i=document.querySelector(`[data-user="${id}"][data-f="name"]`);if(i){i.scrollIntoView({block:'center'});i.select()}},0)}
  if(a==='userdel'){const id=b.dataset.id;if(!b.dataset.armed){b.dataset.armed=1;b.textContent='Click again';return}delete STATE.users[id];Object.values(STATE.jobs).forEach(j=>{for(const r in j.team)j.team[r]=j.team[r].filter(x=>x!==id);j.access.users=j.access.users.filter(x=>x!==id)});Object.values(STATE.groups).forEach(g=>g.members=g.members.filter(x=>x!==id));STATE.allJobs.users=STATE.allJobs.users.filter(x=>x!==id);if(UI.user===id)UI.user=STATE.defaultUser;save();renderRibbon();setModal()}
  if(a==='groupadd'){const id=uid('g');STATE.groups[id]={name:'New group',members:[]};save();setModal();setTimeout(()=>{const i=document.querySelector(`[data-group="${id}"][data-f="name"]`);if(i){i.scrollIntoView({block:'center'});i.select()}},0)}
  if(a==='groupdel'){const id=b.dataset.id;if(!b.dataset.armed){b.dataset.armed=1;b.textContent='Click again';return}delete STATE.groups[id];Object.values(STATE.jobs).forEach(j=>j.access.groups=j.access.groups.filter(x=>x!==id));STATE.allJobs.groups=STATE.allJobs.groups.filter(x=>x!==id);save();setModal()}
  if(a==='gmemdel'){const g=STATE.groups[b.dataset.g];g.members=g.members.filter(x=>x!==b.dataset.id);save();setModal()}
  if(a==='roleadd'){const id=uid('role');STATE.roles[id]={label:'New role',short:'',company:'RK',team:1,plan:1};save();renderRibbon();setModal();setTimeout(()=>{const i=document.querySelector(`[data-role="${id}"][data-f="label"]`);if(i){i.scrollIntoView({block:'center'});i.select()}},0)}
  if(a==='roledel'){const id=b.dataset.id;if(!b.dataset.armed){b.dataset.armed=1;b.textContent='Click again';return}delete STATE.roles[id];Object.values(STATE.users).forEach(u=>{if(u.role===id)u.role='tech'});Object.values(STATE.jobs).forEach(j=>{if(j.team)delete j.team[id]});save();renderRibbon();setModal()}
  if(a==='reset'){if(b.dataset.armed){try{localStorage.removeItem(KEY)}catch(_){}STATE=null;STATE=buildSeed();UI.sel={type:'shop'};closeModal();setShop(UI.shop);toast('Sample data restored')}else{b.dataset.armed=1;b.textContent='Click again to reset everything'}}
});
$('#modal').addEventListener('change',e=>{const t=e.target;
  const keepScroll=()=>{const mb=document.querySelector('#modal .mbody');const y=mb?mb.scrollTop:0;setModal();const m2=document.querySelector('#modal .mbody');if(m2)m2.scrollTop=y};
  if(t.dataset.user){const u=STATE.users[t.dataset.user];const f=t.dataset.f;if(f==='all'){const L=STATE.allJobs.users;STATE.allJobs.users=t.checked?[...new Set([...L,t.dataset.user])]:L.filter(x=>x!==t.dataset.user)}else if(f==='name')u.name=t.value.trim()||'Unnamed';else if(f==='role'){u.role=t.value;if((STATE.roles[u.role]||{}).company==='customer'&&!u.company)u.company=Object.values(STATE.jobs)[0].customer}else u[f]=t.value;save();renderRibbon();keepScroll();return}
  if(t.dataset.group){const g=STATE.groups[t.dataset.group];const f=t.dataset.f;if(f==='name')g.name=t.value.trim()||'Unnamed group';else if(f==='all'){const L=STATE.allJobs.groups;STATE.allJobs.groups=t.checked?[...new Set([...L,t.dataset.group])]:L.filter(x=>x!==t.dataset.group)}else if(f==='addmem'&&t.value)g.members.push(t.value);save();keepScroll();return}
  if(t.dataset.role){const r=STATE.roles[t.dataset.role];const f=t.dataset.f;if(t.type==='checkbox'){if(t.checked)r[f]=1;else delete r[f]}else r[f]=t.value.trim()||(f==='label'?'Unnamed role':'');save();renderRibbon();if(t.dataset.role===UI.role)setEdit(false);return}
  if(t.dataset.lib){const l=STATE.library.find(l=>l.id===t.dataset.lib);const f=t.dataset.f;l[f]=f==='L'||f==='W'?Math.max(.5,+t.value||1):t.value;save()}
  if(t.dataset.shop){const s=STATE.shops[t.dataset.shop];const f=t.dataset.f;if(f==='ftPerIn')s.board.ftPerIn=Math.max(1,+t.value||1);else s[f]=t.value;save();if(f!=='ftPerIn')renderRibbon();setModal()}});
function libModal(){openModal('Equipment library',`<p class="note">Items admins can place on any layout. Sizes are footprint length × width in feet. Deleting is blocked while an item is placed on a floor.</p>
  <div class="tbl-wrap"><table><thead><tr><th>Name</th><th>Length ft</th><th>Width ft</th><th>Color</th><th>Placed</th><th></th></tr></thead><tbody>${STATE.library.map(l=>`<tr><td><input value="${esc(l.name)}" data-lib="${l.id}" data-f="name" aria-label="Name"></td><td><input type="number" step=".5" value="${l.L}" data-lib="${l.id}" data-f="L" aria-label="Length"></td><td><input type="number" step=".5" value="${l.W}" data-lib="${l.id}" data-f="W" aria-label="Width"></td><td><input type="color" value="${l.c}" data-lib="${l.id}" data-f="c" aria-label="Color"></td><td class="mono">${STATE.elements.filter(x=>x.libId===l.id).length}</td><td><button class="btn danger" data-mact="libdel" data-id="${l.id}" style="padding:2px 8px">Delete</button></td></tr>`).join('')}</tbody></table></div>
  <div class="btns"><button class="btn pri" data-mact="libadd">Add library item</button></div>`)}
function setModal(){
  openModal('Shop settings',Object.entries(STATE.shops).map(([k,s])=>{const{pw:bw,ph:bh,fits}=boardFit(k);
    return`<section class="sec"><h3>${esc(s.name)}</h3><div class="grid2"><div class="field"><label for="s-${k}-n">Display name</label><input id="s-${k}-n" value="${esc(s.name)}" data-shop="${k}" data-f="name"></div><div class="field"><label for="s-${k}-a">Address / note</label><input id="s-${k}-a" value="${esc(s.addr)}" data-shop="${k}" data-f="addr"></div></div>
    <dl class="kv"><dt>Source drawing scale</dt><dd>${esc(s.drawing)}</dd><dt>Verified inside dimensions</dt><dd>${VERIFIED[k]||'Not checked'}</dd></dl>
    <div class="grid2"><div class="field"><label for="s-${k}-b">Whiteboard scale: feet per inch</label><input id="s-${k}-b" type="number" min="1" step="1" value="${s.board.ftPerIn}" data-shop="${k}" data-f="ftPerIn"></div>
    <div class="field"><label>Board usable area</label><div style="display:flex;gap:6px;align-items:center"><span class="mono">${s.board.w}" × ${s.board.h}"</span><span class="note">plan ${s.board.rot?'turned 90°':'upright'}</span><button class="btn" data-mact="flip" data-id="${k}" style="padding:2px 8px">Turn plan</button></div></div></div>
    <p class="note">At 1" = ${s.board.ftPerIn}', the plan prints ${bw.toFixed(1)}" × ${bh.toFixed(1)}" ${fits?'and fits the board.':'and does <b>not</b> fit the board.'} Skid magnets use this scale.</p>
    <div class="btns"><button class="btn" data-mact="suggest" data-id="${k}">Use smallest scale that fits</button></div></section>`}).join('')+
    rolesSection()+peopleSection()+`${window.LIVE.settingsHtml()}`)}
function rolesSection(){
  const rows=Object.entries(STATE.roles).filter(([,r])=>r.company==='RK');
  return`<section class="sec" id="rolesSec"><h3>Roles</h3><p class="note">Roles fill the "Viewing as" list. Tick <b>Job team</b> to give every job a slot for that role on its Team tab. Permissions decide what each role can change.</p>
  <div class="tbl-wrap"><table class="roles"><thead><tr><th>Role</th><th>Short</th><th title="Shown on each job's Team tab">Job team</th>${PERMS.map(([,l])=>`<th>${l}</th>`).join('')}<th></th></tr></thead><tbody>${rows.map(([id,r])=>`<tr>
    <td><input value="${esc(r.label)}" data-role="${id}" data-f="label" aria-label="Role name"></td><td><input value="${esc(r.short||'')}" data-role="${id}" data-f="short" aria-label="Short name" style="width:80px"></td>
    <td class="c"><input type="checkbox" ${r.team?'checked':''} data-role="${id}" data-f="team" ${r.fixed?'disabled':''} aria-label="On job teams"></td>
    ${PERMS.map(([k,l])=>`<td class="c"><input type="checkbox" ${r[k]?'checked':''} data-role="${id}" data-f="${k}" ${r.fixed?'disabled':''} aria-label="${l}"></td>`).join('')}
    <td>${r.fixed?'':`<button class="btn danger" data-mact="roledel" data-id="${id}" style="padding:2px 8px">Delete</button>`}</td></tr>`).join('')}</tbody></table></div>
  <div class="btns"><button class="btn pri" data-mact="roleadd">Add role</button></div>
  <p class="note">The administrator row is locked so nobody can lock themselves out. Customer logins are set per customer company.</p></section>`}
function peopleSection(){
  const roleOpts=sel=>Object.entries(STATE.roles).map(([id,r])=>`<option value="${id}" ${id===sel?'selected':''}>${esc(r.label)}</option>`).join('');
  const users=Object.entries(STATE.users).sort((a,b)=>a[1].name.localeCompare(b[1].name));const nj=Object.keys(STATE.jobs).length;
  return`<section class="sec" id="peopleSec"><h3>People</h3><p class="note">Roles decide what someone can do. Job access decides which jobs they see: their team jobs, jobs granted to them or their groups, or every job with an <b>All jobs</b> grant. Only administrators see every job by default.</p>
  <div class="tbl-wrap"><table class="roles"><thead><tr><th>Name</th><th>Role</th><th>Customer company</th><th>All jobs</th><th>Sees</th><th></th></tr></thead><tbody>${users.map(([id,u])=>{const cust=(STATE.roles[u.role]||{}).company==='customer';const n=Object.keys(STATE.jobs).filter(j=>accessWhy(j,id)).length;return`<tr>
    <td><input value="${esc(u.name)}" data-user="${id}" data-f="name" aria-label="Name"></td><td><select data-user="${id}" data-f="role" aria-label="Role" ${id===STATE.defaultUser?'disabled':''}>${roleOpts(u.role)}</select></td>
    <td>${cust?`<select data-user="${id}" data-f="company" aria-label="Company">${[...new Set(Object.values(STATE.jobs).map(j=>j.customer))].map(c=>`<option ${c===u.company?'selected':''}>${esc(c)}</option>`).join('')}</select>`:'<span class="note">RK</span>'}</td>
    <td class="c"><input type="checkbox" data-user="${id}" data-f="all" ${STATE.allJobs.users.includes(id)||u.role==='admin'?'checked':''} ${u.role==='admin'||cust?'disabled':''} aria-label="All jobs"></td>
    <td class="mono">${n} / ${nj}</td><td>${id===STATE.defaultUser?'':`<button class="btn danger" data-mact="userdel" data-id="${id}" style="padding:2px 8px">Delete</button>`}</td></tr>`}).join('')}</tbody></table></div>
  ${window.LIVE.addPeopleHtml()}</section>
  <section class="sec" id="groupsSec"><h3>Groups</h3><p class="note">Grant a group access on a job's Access tab, or tick <b>All jobs</b> here. Nobody gets All jobs automatically.</p>
  <div class="tbl-wrap"><table class="roles"><thead><tr><th>Group</th><th>Members</th><th>All jobs</th><th>Jobs granted</th><th></th></tr></thead><tbody>${Object.entries(STATE.groups).map(([g,x])=>`<tr>
    <td><input value="${esc(x.name)}" data-group="${g}" data-f="name" aria-label="Group name"></td>
    <td><div class="gmem">${x.members.map(u=>`<span class="person">${esc(uname(u))}<button data-mact="gmemdel" data-g="${g}" data-id="${u}" aria-label="Remove">×</button></span>`).join('')}<select data-group="${g}" data-f="addmem" aria-label="Add member"><option value="">Add…</option>${users.filter(([id,u])=>!x.members.includes(id)&&(STATE.roles[u.role]||{}).company==='RK').map(([id,u])=>`<option value="${id}">${esc(u.name)}</option>`).join('')}</select></div></td>
    <td class="c"><input type="checkbox" data-group="${g}" data-f="all" ${STATE.allJobs.groups.includes(g)?'checked':''} aria-label="All jobs"></td>
    <td class="mono">${Object.values(STATE.jobs).filter(j=>j.access.groups.includes(g)).map(j=>j.num).join(', ')||'none'}</td>
    <td><button class="btn danger" data-mact="groupdel" data-id="${g}" style="padding:2px 8px">Delete</button></td></tr>`).join('')}</tbody></table></div>
  <div class="btns"><button class="btn pri" data-mact="groupadd">Add group</button></div></section>`}
$('#libBtn').onclick=libModal;$('#setBtn').onclick=setModal;
function pdfModal(){
  const opts=[['areas','Areas'],['lines','Production lines and stations'],['flow','Layout lines']];
  openModal('Export whiteboard PDF',`<p class="note">Makes a full-size vector PDF of a shop at its whiteboard scale. Plot it at 100% (actual size) and it matches the board and the skid magnets. Each sheet carries a 100 ft scale bar and a known inside dimension to check with a tape measure.</p>
    ${Object.entries(STATE.shops).map(([k,s])=>{const f=boardFit(k);return`<section class="sec"><h3>${esc(s.name)}</h3><dl class="kv"><dt>Scale</dt><dd>1" = ${s.board.ftPerIn}'-0" (1:${s.board.ftPerIn*12})</dd><dt>Sheet</dt><dd>${s.board.w}" × ${s.board.h}"${s.board.rot?', plan turned 90°':''}</dd><dt>Plan on sheet</dt><dd>${f.pw.toFixed(1)}" × ${f.ph.toFixed(1)}"${f.fits?'':' (does not fit)'}</dd></dl>
      <div class="btns"><button class="btn pri" data-mact="pdf" data-id="${k}">Download ${esc(s.name)} PDF</button></div></section>`}).join('')}
    <section class="sec"><h3>Include layout layers</h3><div class="btns">${opts.map(([k,l])=>`<label class="note" style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="pdf-${k}" ${UI.pdfLayers&&UI.pdfLayers[k]?'checked':''}> ${l}</label>`).join('')}</div>
    <p class="note">Leave these off for a permanent board graphic and draw changing layouts with markers. Customer zones, skids and equipment are never printed; the magnets stand in for them.</p></section>
    <p class="note" id="pdfMsg" role="status"></p>`)}
async function exportPdf(btn){
  const k=btn.dataset.id,s=STATE.shops[k],bd=s.board;const msg=$('#pdfMsg');
  UI.pdfLayers={areas:$('#pdf-areas').checked,lines:$('#pdf-lines').checked,flow:$('#pdf-flow').checked};
  const kinds=[UI.pdfLayers.areas&&'area',UI.pdfLayers.lines&&'line',UI.pdfLayers.lines&&'station',UI.pdfLayers.flow&&'flow'].filter(Boolean);
  const date=new Date().toISOString().slice(0,10);
  const pdf=boardPdf({name:s.name,plan:PLANS[k],ftPerIn:bd.ftPerIn,boardW:bd.w,boardH:bd.h,rot:bd.rot||0,check:CHECKS[k],verified:VERIFIED_ASCII[k]?'Inside '+VERIFIED_ASCII[k]:'',
    elements:STATE.elements.filter(e=>e.shop===k&&kinds.includes(e.kind)),areaColors:Object.fromEntries(Object.entries(AREA_TYPES).map(([t,a])=>[t,a.c])),date});
  const filename=`${s.name}_whiteboard_1in-${bd.ftPerIn}ft_${bd.w}x${bd.h}.pdf`;
  const dl=window.claude&&window.claude.use?await window.claude.use('downloads'):null;
  if(!dl){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([pdf],{type:'application/pdf'}));a.download=filename;document.body.appendChild(a);a.click();a.remove();msg.textContent='Saved '+filename+'.';return}
  btn.disabled=true;
  try{await dl.save({filename,data:new Blob([pdf])});msg.textContent=`Saved ${filename}.`}
  catch(e){msg.textContent=e&&e.code==='declined'?'Download cancelled.':'Could not save the PDF ('+(e&&e.code||'error')+').'}
  finally{btn.disabled=false}
}
$('#pdfBtn').onclick=pdfModal;

/* ================= boot ================= */
renderRibbon();renderPlan();renderEditBar();
requestAnimationFrame(()=>{fit();renderLayers();renderPanel()});
new ResizeObserver(()=>{const c=viewCenter();const old=fitS;const r=stage$.getBoundingClientRect();const[x,y,w,h]=vbOf(UI.shop);const pad=UI_PAD();fitS=Math.min((r.width-pad.l-pad.r)/w,(r.height-pad.t-pad.b)/h);if(Math.abs(UI.view.s-old)<1e-6)fit();else centerOn(c[0],c[1],UI.view.s)}).observe(stage$);

