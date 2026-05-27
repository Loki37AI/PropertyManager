
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="PropManager">
<meta name="theme-color" content="#0d1117">
<title>🏢 PropertyManager</title>
<script src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
<script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
<script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
body{font-family:system-ui,-apple-system,sans-serif;background:#0d1117;color:#e6edf3;min-height:100vh}
input,button,select,textarea{font-family:inherit}
@keyframes pop{from{transform:scale(.93);opacity:0}to{transform:scale(1);opacity:1}}
</style>
</head>
<body>
<div id="root"></div>
<script type="text/babel">

const PAL=['#58a6ff','#3fb950','#d29922','#a78bfa','#f85149','#ffd166'];
const BC=[{bg:'rgba(88,166,255,.15)',tx:'#58a6ff'},{bg:'rgba(63,185,80,.15)',tx:'#3fb950'},{bg:'rgba(210,153,34,.15)',tx:'#d29922'},{bg:'rgba(167,139,250,.15)',tx:'#a78bfa'},{bg:'rgba(248,81,73,.15)',tx:'#f85149'}];
const MOS=['January','February','March','April','May','June','July','August','September','October','November','December'];
const NOW=new Date();
const DEF_M=`${NOW.getFullYear()}-${String(NOW.getMonth()+1).padStart(2,'0')}`;
const uid=()=>'_'+Math.random().toString(36).substr(2,9);
const SCRIPT_URL='https://script.google.com/macros/s/AKfycbyQvai32XBdjl4tpJcK7vG5L3LvJrDqAMN0ZPdOKd2-wCguA7MoqVpyFeredNtPH8W8/exec';
const SHEET_ID='1LQZhC8UjRB0xX2JHD8NLyIrJefbaxebP_d-DCOaTPdY';
let syncTimer=null;

// Push all data to Google Sheets
function pushToSheets(data){
  clearTimeout(syncTimer);
  syncTimer=setTimeout(()=>{
    fetch(SCRIPT_URL,{
      method:'POST',
      body:JSON.stringify({sheetId:SHEET_ID,action:'save',payload:JSON.stringify(data)})
    }).catch(()=>{});
  },2000);
}

// Pull data from Google Sheets
function pullFromSheets(){
  return fetch(SCRIPT_URL+'?sheetId='+SHEET_ID)
    .then(r=>r.text())
    .then(t=>{
      if(!t||t==='') return null;
      try{return JSON.parse(t);}catch(e){return null;}
    })
    .catch(()=>null);
}

const fmt=m=>{if(!m)return'';const[y,mo]=m.split('-');return MOS[+mo-1]+' '+y;};
const fmtS=m=>{if(!m)return'';const[y,mo]=m.split('-');return MOS[+mo-1].slice(0,3)+' '+y.slice(2);};
const mBill=m=>Math.max(0,(+m.curr||0)-(+m.last||0))*(+m.rate||0);
const wBill=r=>r.meters.reduce((s,m)=>s+mBill(m),0);
const rTotal=r=>(+r.rent||0)+wBill(r);
const getPK=m=>{const[y,mo]=m.split('-').map(Number);const pm=mo===1?12:mo-1,py=mo===1?y-1:y;return`${py}-${String(pm).padStart(2,'0')}`;};
const lsGet=k=>{try{const v=localStorage.getItem(k);return v?JSON.parse(v):null;}catch(e){return null;}};
const lsSet=(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v));}catch(e){}};

function Modal({show,onClose,title,children}){
  if(!show)return null;
  return(
    <div onClick={onClose} style={{position:'fixed',inset:0,background:'rgba(0,0,0,.88)',zIndex:300,display:'flex',alignItems:'center',justifyContent:'center',padding:14}}>
      <div onClick={e=>e.stopPropagation()} style={{background:'#161b22',border:'1px solid #30363d',borderRadius:14,padding:20,width:'100%',maxWidth:420,maxHeight:'90vh',overflowY:'auto'}}>
        <div style={{fontWeight:700,fontSize:'.95rem',marginBottom:14}}>{title}</div>
        {children}
      </div>
    </div>
  );
}

function Btn({children,onClick,color='#1f6feb',text='#fff',ghost,danger,sm,style={}}){
  return <button onClick={onClick} style={{display:'inline-flex',alignItems:'center',justifyContent:'center',gap:4,padding:sm?'4px 9px':'7px 13px',border:ghost?'1px solid #30363d':danger?'1px solid rgba(248,81,73,.2)':'none',borderRadius:8,fontSize:sm?'.7rem':'.78rem',fontWeight:600,cursor:'pointer',background:ghost?'transparent':danger?'rgba(248,81,73,.12)':color,color:ghost?'#8b949e':danger?'#f85149':text,fontFamily:'inherit',...style}}>{children}</button>;
}

function FRow({label,children}){
  return <div style={{marginBottom:11}}><label style={{display:'block',fontSize:'.63rem',fontWeight:700,textTransform:'uppercase',letterSpacing:'.5px',color:'#8b949e',marginBottom:4}}>{label}</label>{children}</div>;
}

function Inp({value,onChange,placeholder,type='text',style={}}){
  return <input type={type} value={value||''} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={{width:'100%',background:'#0d1117',border:'1px solid #30363d',borderRadius:7,padding:'8px 11px',color:'#e6edf3',fontSize:'.86rem',outline:'none',fontFamily:'inherit',...style}}/>;
}

// ── EXPORT FUNCTIONS ──────────────────────────────
function exportExcel(S,PAY,allMs,getMS){
  const script=document.createElement('script');
  script.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  script.onload=()=>{
    const wb=window.XLSX.utils.book_new();
    const gPS=(rid,m)=>PAY[rid]?.[m]||{rent:'unpaid',water:'unpaid'};
    const ms=allMs();
    const H=(v,bg)=>({v,t:'s',s:{font:{bold:true,color:{rgb:'FFFFFF'},sz:10},fill:{fgColor:{rgb:bg||'1F3A5F'}},alignment:{horizontal:'center'}}});
    const T=v=>({v:v||'',t:'s',s:{font:{sz:10}}});
    const M=v=>({v:+v||0,t:'n',s:{numFmt:'#,##0.00',font:{color:{rgb:'166534'},sz:10},alignment:{horizontal:'right'}}});
    const rows=[[H('PROPERTY MANAGER — MONTHLY BILLS','0F2744'),'','','','','','','',''],[''],
      [H('Month'),H('Building'),H('Room'),H('Tenant'),H('Rent'),H('Water'),H('Total'),H('Rent Status'),H('Water Status')]];
    ms.forEach(m=>{
      const sv=getMS(m);if(!sv)return;
      sv.buildings.forEach(b=>b.rooms.forEach(r=>{
        const wb=r.meters.reduce((s,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s+d*(+mt.rate||0);},0);
        const rr=+r.rent||0,ps=gPS(r.id,m);
        rows.push([T(fmt(m)),T(b.name),T(r.name),T(r.tenant||''),M(rr),M(wb),M(rr+wb),T(ps.rent==='paid'?'✓ Paid':'✗ Unpaid'),T(ps.water==='paid'?'✓ Paid':'✗ Unpaid')]);
      }));
      rows.push([]);
    });
    const ws=window.XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=[{wch:18},{wch:14},{wch:10},{wch:18},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12}];
    window.XLSX.utils.book_append_sheet(wb,ws,'Monthly Bills');
    const oRows=[[H('OUTSTANDING & DEPOSITS','0F2744'),'','','',''],[''],
      [H('Building'),H('Room'),H('Tenant'),H('Monthly Rent'),H('Security Deposit'),H('Total Outstanding')]];
    S.buildings.forEach(b=>b.rooms.forEach(r=>{
      let due=0;
      ms.forEach(m=>{
        const sv=getMS(m);if(!sv)return;
        const sb=sv.buildings.find(x=>x.id===b.id);
        const sr=sb?.rooms.find(x=>x.id===r.id);if(!sr)return;
        const wb=sr.meters.reduce((s,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s+d*(+mt.rate||0);},0);
        const ps=gPS(r.id,m);
        due+=(ps.rent!=='paid'?+sr.rent||0:0)+(ps.water!=='paid'?wb:0);
      });
      oRows.push([T(b.name),T(r.name),T(r.tenant||''),M(r.rent||0),M(r.deposit||0),M(due)]);
    }));
    const ws2=window.XLSX.utils.aoa_to_sheet(oRows);
    ws2['!cols']=[{wch:14},{wch:10},{wch:18},{wch:14},{wch:16},{wch:16}];
    window.XLSX.utils.book_append_sheet(wb,ws2,'Outstanding');
    window.XLSX.writeFile(wb,'PropertyManager_'+new Date().toISOString().slice(0,10)+'.xlsx');
  };
  document.head.appendChild(script);
}

function exportPDF(S,PAY,allMs,getMS){
  const gPS=(rid,m)=>PAY[rid]?.[m]||{rent:'unpaid',water:'unpaid'};
  const ms=allMs();if(!ms.length){alert('No data to export');return;}
  const now=new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'long',year:'numeric'});
  let html=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Property Report</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:11px;color:#000;background:#fff}.cov{padding:34px;text-align:center;page-break-after:always}.cov h1{font-size:22px;color:#1e3a5f}.ms{padding:14px 18px;page-break-after:always}.mt{background:#1e3a5f;color:#fff;padding:7px 13px;font-size:12px;font-weight:bold}.sr{display:flex;gap:7px;margin:9px 0}.sb{flex:1;border:1px solid #ddd;border-radius:4px;padding:7px;text-align:center}.v{font-size:14px;font-weight:bold;color:#1e3a5f}.l{font-size:9px;color:#666}table{width:100%;border-collapse:collapse;margin-bottom:11px;font-size:10px}th{background:#1e3a5f;color:#fff;padding:5px 6px;text-align:left}td{padding:4px 6px;border-bottom:1px solid #eee}tr:nth-child(even) td{background:#f8fafc}.paid{color:#166534;font-weight:bold}.unpaid{color:#991b1b;font-weight:bold}.ra{text-align:right;font-weight:bold}.ft{font-size:9px;color:#aaa;text-align:center;padding:4px}@media print{.ms{page-break-after:always}}</style></head><body>`;
  html+=`<div class="cov"><h1>🏢 Property Manager Report</h1><p>${now} · ${ms.length} months</p></div>`;
  ms.forEach(m=>{
    const sv=getMS(m);if(!sv)return;
    let tR=0,tW=0,pc=0,uc=0;
    sv.buildings.forEach(b=>b.rooms.forEach(r=>{
      const wb=r.meters.reduce((s,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s+d*(+mt.rate||0);},0);
      tR+=(+r.rent||0);tW+=wb;
      const ps=gPS(r.id,m);
      if(ps.rent==='paid'&&ps.water==='paid')pc++;else uc++;
    }));
    html+=`<div class="ms"><div class="mt">📅 ${fmt(m)}</div><div class="sr"><div class="sb"><div class="v">₹${(tR+tW).toFixed(2)}</div><div class="l">Total</div></div><div class="sb"><div class="v">₹${tR.toFixed(2)}</div><div class="l">Rent</div></div><div class="sb"><div class="v">₹${tW.toFixed(2)}</div><div class="l">Water</div></div><div class="sb"><div class="v" style="color:#166534">${pc}</div><div class="l">Paid</div></div><div class="sb"><div class="v" style="color:#991b1b">${uc}</div><div class="l">Unpaid</div></div></div>`;
    sv.buildings.forEach(b=>{
      html+=`<div style="font-weight:bold;font-size:10px;color:#555;margin:7px 0 3px">🏢 ${b.name}</div><table><thead><tr><th>Room</th><th>Tenant</th><th class="ra">Rent</th><th class="ra">Water</th><th class="ra">Total</th><th>Rent</th><th>Water</th></tr></thead><tbody>`;
      b.rooms.forEach(r=>{
        const wb=r.meters.reduce((s,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s+d*(+mt.rate||0);},0);
        const rr=+r.rent||0,ps=gPS(r.id,m);
        html+=`<tr><td>${r.name}</td><td>${r.tenant||'—'}</td><td class="ra">₹${rr.toFixed(2)}</td><td class="ra">₹${wb.toFixed(2)}</td><td class="ra">₹${(rr+wb).toFixed(2)}</td><td class="${ps.rent==='paid'?'paid':'unpaid'}">${ps.rent==='paid'?'✓ Paid':'✗ Unpaid'}</td><td class="${ps.water==='paid'?'paid':'unpaid'}">${ps.water==='paid'?'✓ Paid':'✗ Unpaid'}</td></tr>`;
      });
      html+=`</tbody></table>`;
    });
    html+=`<div class="ft">PropertyManager · ${fmt(m)}</div></div>`;
  });
  html+=`</body></html>`;
  const w=window.open('','_blank');w.document.write(html);w.document.close();w.focus();setTimeout(()=>w.print(),700);
}

// ── BILLING PAGE ──────────────────────────────────
function BillingPage({S,setS,PAY,setPAY,month,showToast,allMs,getMS,applyReadings}){
  const[showAB,setShowAB]=useState(false);
  const[showAR,setShowAR]=useState(false);
  const[showER,setShowER]=useState(false);
  const[showRN,setShowRN]=useState(false);
  const[arBid,setArBid]=useState('');
  const[abN,setAbN]=useState('');
  const[arD,setArD]=useState({name:'',tenant:'',rent:'',deposit:'',moveIn:''});
  const[erBid,setErBid]=useState('');
  const[erRid,setErRid]=useState('');
  const[erD,setErD]=useState({name:'',tenant:'',rent:'',deposit:'',moveIn:''});
  const[rnBid,setRnBid]=useState('');
  const[rnN,setRnN]=useState('');
  const[col,setCol]=useState({});
  const[mI,setMI]=useState({});

  const gPS=rid=>PAY[rid]?.[month]||{rent:'unpaid',water:'unpaid'};
  const togR=rid=>{const ps=gPS(rid);const nv=ps.rent==='paid'?'unpaid':'paid';setPAY(p=>({...p,[rid]:{...p[rid],[month]:{...(p[rid]?.[month]||{rent:'unpaid',water:'unpaid'}),rent:nv}}}));showToast(nv==='paid'?'✅ Rent paid!':'Rent unpaid');};
  const togW=rid=>{const ps=gPS(rid);const nv=ps.water==='paid'?'unpaid':'paid';setPAY(p=>({...p,[rid]:{...p[rid],[month]:{...(p[rid]?.[month]||{rent:'unpaid',water:'unpaid'}),water:nv}}}));showToast(nv==='paid'?'✅ Water paid!':'Water unpaid');};
  const updR=(bid,rid,mid,f,v)=>setS(p=>({...p,buildings:p.buildings.map(b=>b.id===bid?{...b,rooms:b.rooms.map(r=>r.id===rid?{...r,meters:r.meters.map(m=>m.id===mid?{...m,[f]:v}:m)}:r)}:b)}));
  const doAB=()=>{if(!abN.trim()){alert('Enter name');return;}setS(p=>({...p,buildings:[...p.buildings,{id:uid(),name:abN.trim(),rooms:[]}]}));setAbN('');setShowAB(false);showToast('🏢 Building added!');};
  const doAR=()=>{if(!arD.name.trim()){alert('Enter room number');return;}setS(p=>({...p,buildings:p.buildings.map(b=>b.id===arBid?{...b,rooms:[...b.rooms,{id:uid(),name:arD.name.trim(),tenant:arD.tenant.trim(),rent:+arD.rent||0,deposit:+arD.deposit||0,moveIn:arD.moveIn,meters:[]}]}:b)}));setArD({name:'',tenant:'',rent:'',deposit:'',moveIn:''});setShowAR(false);showToast('🏠 Room added!');};
  const doSER=()=>{setS(p=>({...p,buildings:p.buildings.map(b=>b.id===erBid?{...b,rooms:b.rooms.map(r=>r.id===erRid?{...r,name:erD.name.trim(),tenant:erD.tenant.trim(),rent:+erD.rent||0,deposit:+erD.deposit||0,moveIn:erD.moveIn}:r)}:b)}));setShowER(false);showToast('✅ Updated!');};
  const doRN=()=>{if(!rnN.trim())return;setS(p=>({...p,buildings:p.buildings.map(b=>b.id===rnBid?{...b,name:rnN.trim()}:b)}));setShowRN(false);};
  const doAM=(bid,rid)=>{const k=bid+'_'+rid,v=mI[k]||{};if(!(v.n||'').trim()){alert('Enter meter name');return;}setS(p=>({...p,buildings:p.buildings.map(b=>b.id===bid?{...b,rooms:b.rooms.map(r=>r.id===rid?{...r,meters:[...r.meters,{id:uid(),name:v.n.trim(),unit:v.u||'units',rate:+(v.r)||0,last:'',curr:''}]}:r)}:b)}));setMI(p=>({...p,[k]:{n:'',u:'',r:''}}));showToast('Meter added!');};
  const delM=(bid,rid,mid)=>setS(p=>({...p,buildings:p.buildings.map(b=>b.id===bid?{...b,rooms:b.rooms.map(r=>r.id===rid?{...r,meters:r.meters.filter(m=>m.id!==mid)}:r)}:b)}));
  const delR=(bid,rid)=>{if(!confirm('Delete room?'))return;setS(p=>({...p,buildings:p.buildings.map(b=>b.id===bid?{...b,rooms:b.rooms.filter(r=>r.id!==rid)}:b)}));};
  const delB=bid=>{if(!confirm('Delete building?'))return;setS(p=>({...p,buildings:p.buildings.filter(b=>b.id!==bid)}));};

  const copyLastMonth=()=>{
    const prev_rd=lsGet('pm_r_'+getPK(month));
    if(!prev_rd){showToast('⚠️ No previous month data found');return;}
    setS(p=>({...p,buildings:p.buildings.map(b=>({...b,rooms:b.rooms.map(r=>({...r,meters:r.meters.map(mt=>{const x=prev_rd[r.id]?.[mt.id];return x?{...mt,last:x.curr||'',curr:''}:{...mt,last:'',curr:''};})}))}))}));
    showToast('✅ Last month readings copied!');
  };

  const iS={background:'#161b22',border:'1px solid #30363d',borderRadius:5,padding:'4px 5px',color:'#e6edf3',fontSize:'.73rem',textAlign:'center',outline:'none',width:68,fontFamily:'inherit'};

  return(
    <div>
      {/* Action bar */}
      <div style={{display:'flex',gap:7,flexWrap:'wrap',marginBottom:14}}>
        <Btn onClick={()=>{setAbN('');setShowAB(true);}}>＋ Building</Btn>
        <Btn color="#0d419d" ghost onClick={copyLastMonth} style={{fontSize:'.72rem'}}>⬅ Copy Last Month</Btn>
        <Btn color="#238636" onClick={()=>exportExcel(S,PAY,allMs,getMS)}>📊 Excel</Btn>
        <Btn color="#d29922" text="#fff" onClick={()=>exportPDF(S,PAY,allMs,getMS)}>📄 PDF</Btn>
      </div>

      {!S.buildings.length&&(
        <div style={{textAlign:'center',padding:'50px 20px',color:'#484f58'}}>
          <div style={{fontSize:'2.5rem',marginBottom:12}}>🏢</div>
          <div style={{fontWeight:600,fontSize:'1rem',color:'#8b949e',marginBottom:6}}>No buildings yet</div>
          <div style={{fontSize:'.82rem'}}>Tap <b style={{color:'#58a6ff'}}>＋ Building</b> above</div>
        </div>
      )}

      {S.buildings.map((b,bi)=>{
        const c=BC[bi%BC.length],bt=b.rooms.reduce((s,r)=>s+rTotal(r),0),isC=col[b.id];
        return(
          <div key={b.id} style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,marginBottom:12,overflow:'hidden'}}>
            <div onClick={()=>setCol(p=>({...p,[b.id]:!p[b.id]}))} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'10px 14px',background:'#0d1117',cursor:'pointer',userSelect:'none'}}>
              <div style={{display:'flex',alignItems:'center',gap:9}}>
                <div style={{width:28,height:28,borderRadius:7,background:c.bg,color:c.tx,display:'flex',alignItems:'center',justifyContent:'center',fontSize:'.8rem',fontWeight:700}}>{bi+1}</div>
                <div><div style={{fontWeight:700,fontSize:'.9rem'}}>{b.name}</div><div style={{fontSize:'.65rem',color:'#8b949e',marginTop:1}}>{b.rooms.length} room{b.rooms.length!==1?'s':''}</div></div>
              </div>
              <div style={{display:'flex',alignItems:'center',gap:9}}>
                <div style={{fontWeight:700,color:'#3fb950'}}>{bt>0?'₹'+bt.toFixed(2):'—'}</div>
                <span style={{color:'#484f58',fontSize:'.7rem',transform:isC?'rotate(-90deg)':'',transition:'.2s',display:'inline-block'}}>▼</span>
              </div>
            </div>
            {!isC&&(
              <div style={{padding:12}}>
                {b.rooms.map(r=>{
                  const wb=wBill(r),rr=+r.rent||0,tot=rr+wb,ps=gPS(r.id),rP=ps.rent==='paid',wP=ps.water==='paid';
                  return(
                    <div key={r.id} style={{border:'1px solid #30363d',borderRadius:9,marginBottom:10,overflow:'hidden',background:'#0d1117'}}>
                      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'9px 12px',background:'#161b22',flexWrap:'wrap',gap:6}}>
                        <div>
                          <div style={{fontWeight:700,fontSize:'.88rem'}}>🏠 Room {r.name}</div>
                          <div style={{fontSize:'.7rem',color:'#8b949e',marginTop:2}}>
                            {r.tenant?<span style={{color:'#a78bfa'}}>👤 {r.tenant}</span>:<span style={{color:'#484f58'}}>No tenant</span>}
                            {r.rent?<span style={{color:'#3fb950',fontWeight:600}}> · ₹{r.rent}/mo</span>:null}
                            {r.deposit?<span style={{color:'#d29922'}}> · 🔐 ₹{r.deposit}</span>:null}
                          </div>
                        </div>
                        <div style={{display:'flex',alignItems:'center',gap:7}}>
                          <div style={{fontWeight:700,color:'#58a6ff',fontSize:'.9rem'}}>{tot>0?'₹'+tot.toFixed(2):'—'}</div>
                          <Btn sm ghost onClick={()=>{setErBid(b.id);setErRid(r.id);setErD({name:r.name,tenant:r.tenant||'',rent:r.rent||'',deposit:r.deposit||'',moveIn:r.moveIn||''});setShowER(true);}}>✏️</Btn>
                          <Btn sm danger onClick={()=>delR(b.id,r.id)}>🗑</Btn>
                        </div>
                      </div>
                      {/* Bill summary */}
                      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:7,padding:'10px 12px',borderBottom:'1px solid #21262d'}}>
                        {[{l:'🏠 Rent',v:'₹'+(rr>0?rr.toFixed(2):'0'),c:'#3fb950'},{l:'💧 Water',v:'₹'+wb.toFixed(2),c:'#58a6ff'},{l:'💰 Total',v:'₹'+tot.toFixed(2),c:'#e6edf3'},{l:'Status',v:rP&&wP?'✅ Paid':'⚠️ Due',c:rP&&wP?'#3fb950':'#f85149'}].map((x,i)=>(
                          <div key={i} style={{background:'#0d1117',borderRadius:7,padding:'7px 9px',textAlign:'center',border:i===2?'1px solid rgba(88,166,255,.2)':i===3?`1px solid ${rP&&wP?'rgba(63,185,80,.2)':'rgba(248,81,73,.2)'}`:'none'}}>
                            <div style={{fontSize:'.58rem',color:'#484f58',fontWeight:700,textTransform:'uppercase',letterSpacing:'.4px',marginBottom:3}}>{x.l}</div>
                            <div style={{fontSize:'.88rem',fontWeight:700,color:x.c}}>{x.v}</div>
                          </div>
                        ))}
                      </div>
                      {/* Meter table */}
                      {r.meters.length>0&&(
                        <div style={{overflowX:'auto'}}>
                          <table style={{width:'100%',borderCollapse:'collapse',fontSize:'.73rem',minWidth:420}}>
                            <thead><tr style={{background:'rgba(255,255,255,.03)'}}>
                              {['Meter','Last Reading','Current Reading','Used','Rate','Amount',''].map((h,i)=>(
                                <th key={i} style={{color:'#8b949e',padding:'5px 7px',textAlign:i===0?'left':'center',fontSize:'.58rem',textTransform:'uppercase',fontWeight:700,paddingLeft:i===0?11:7}}>{h}</th>
                              ))}
                            </tr></thead>
                            <tbody>
                              {r.meters.map((m,mi)=>{
                                const diff=Math.max(0,(+m.curr||0)-(+m.last||0)),bill=diff>0?(diff*(+m.rate||0)).toFixed(2):null;
                                return(
                                  <tr key={m.id} style={{borderBottom:'1px solid rgba(255,255,255,.04)'}}>
                                    <td style={{padding:'6px 7px 6px 11px'}}><div style={{display:'flex',alignItems:'center',gap:5}}><div style={{width:5,height:5,borderRadius:'50%',background:PAL[mi%PAL.length],flexShrink:0}}></div><div style={{fontWeight:600}}>{m.name}</div></div></td>
                                    <td style={{padding:'6px 7px',textAlign:'center'}}><input type="number" value={m.last||''} onChange={e=>updR(b.id,r.id,m.id,'last',e.target.value)} placeholder="0" min="0" style={iS}/></td>
                                    <td style={{padding:'6px 7px',textAlign:'center'}}><input type="number" value={m.curr||''} onChange={e=>updR(b.id,r.id,m.id,'curr',e.target.value)} placeholder="0" min="0" style={iS}/></td>
                                    <td style={{padding:'6px 7px',textAlign:'center',fontWeight:700,color:diff>0?'#3fb950':'#484f58'}}>{m.curr===''&&m.last===''?'—':diff+' '+m.unit}</td>
                                    <td style={{padding:'6px 7px',textAlign:'center',color:'#8b949e',fontSize:'.7rem'}}>₹{m.rate}/{m.unit}</td>
                                    <td style={{padding:'6px 7px',textAlign:'center',fontWeight:700,color:bill?'#58a6ff':'#484f58'}}>{bill?'₹'+bill:'—'}</td>
                                    <td style={{padding:'6px 7px',textAlign:'center'}}><Btn sm danger onClick={()=>delM(b.id,r.id,m.id)}>✕</Btn></td>
                                  </tr>
                                );
                              })}
                              {r.meters.length>1&&(
                                <tr style={{background:'rgba(88,166,255,.06)',borderTop:'1px solid rgba(88,166,255,.15)'}}>
                                  <td colSpan={3} style={{textAlign:'right',fontSize:'.64rem',color:'#484f58',padding:'6px 9px'}}>Water Total →</td>
                                  <td colSpan={2} style={{color:'#58a6ff',textAlign:'center',fontWeight:700,padding:'6px 7px'}}>₹{wb.toFixed(2)}</td>
                                  <td colSpan={2}></td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      )}
                      {/* Pay row */}
                      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'8px 12px',borderTop:'1px solid #21262d',flexWrap:'wrap',gap:6}}>
                        <div style={{fontSize:'.7rem',color:'#8b949e'}}>Payment — {fmt(month)}</div>
                        <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                          <button onClick={()=>togR(r.id)} style={{border:'none',borderRadius:6,padding:'5px 11px',fontSize:'.7rem',fontWeight:700,cursor:'pointer',fontFamily:'inherit',background:rP?'rgba(63,185,80,.15)':'rgba(248,81,73,.1)',color:rP?'#3fb950':'#f85149',borderWidth:1,borderStyle:'solid',borderColor:rP?'rgba(63,185,80,.3)':'rgba(248,81,73,.2)'}}>🏠 Rent {rP?'✓':'✗'}</button>
                          <button onClick={()=>togW(r.id)} style={{border:'none',borderRadius:6,padding:'5px 11px',fontSize:'.7rem',fontWeight:700,cursor:'pointer',fontFamily:'inherit',background:wP?'rgba(88,166,255,.15)':'rgba(248,81,73,.1)',color:wP?'#58a6ff':'#f85149',borderWidth:1,borderStyle:'solid',borderColor:wP?'rgba(88,166,255,.3)':'rgba(248,81,73,.2)'}}>💧 Water {wP?'✓':'✗'}</button>
                        </div>
                      </div>
                      {/* Add meter */}
                      <div style={{display:'flex',gap:5,padding:'7px 12px',background:'rgba(255,255,255,.02)',borderTop:'1px dashed #21262d',flexWrap:'wrap',alignItems:'center'}}>
                        <span style={{fontSize:'.7rem',color:'#8b949e',fontWeight:600}}>＋ Meter:</span>
                        {[{ph:'Name',k:'n',w:88},{ph:'Unit',k:'u',w:72},{ph:'Rate ₹',k:'r',w:62,type:'number'}].map(f=>(
                          <input key={f.k} type={f.type||'text'} value={(mI[b.id+'_'+r.id]||{})[f.k]||''} onChange={e=>setMI(p=>({...p,[b.id+'_'+r.id]:{...(p[b.id+'_'+r.id]||{}),[f.k]:e.target.value}}))} placeholder={f.ph}
                            style={{background:'#0d1117',border:'1px solid #30363d',borderRadius:5,padding:'4px 7px',color:'#e6edf3',fontSize:'.72rem',outline:'none',width:f.w,fontFamily:'inherit'}}/>
                        ))}
                        <Btn sm onClick={()=>doAM(b.id,r.id)}>Add</Btn>
                      </div>
                    </div>
                  );
                })}
                <div style={{display:'flex',gap:6,marginTop:8,flexWrap:'wrap'}}>
                  <Btn sm onClick={()=>{setArBid(b.id);setArD({name:'',tenant:'',rent:'',deposit:'',moveIn:''});setShowAR(true);}}>＋ Add Room</Btn>
                  <Btn sm ghost onClick={()=>{setRnBid(b.id);setRnN(b.name);setShowRN(true);}}>✏️ Rename</Btn>
                  <Btn sm danger onClick={()=>delB(b.id)}>🗑 Delete</Btn>
                </div>
              </div>
            )}
          </div>
        );
      })}

      <Modal show={showAB} onClose={()=>setShowAB(false)} title="🏢 Add Building">
        <FRow label="Building Name"><Inp value={abN} onChange={setAbN} placeholder="e.g. Block A"/></FRow>
        <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:12}}><Btn ghost onClick={()=>setShowAB(false)}>Cancel</Btn><Btn onClick={doAB}>Add Building</Btn></div>
      </Modal>
      <Modal show={showAR} onClose={()=>setShowAR(false)} title="🏠 Add Room">
        <FRow label="Room Number"><Inp value={arD.name} onChange={v=>setArD(p=>({...p,name:v}))} placeholder="e.g. 101"/></FRow>
        <FRow label="Tenant Name"><Inp value={arD.tenant} onChange={v=>setArD(p=>({...p,tenant:v}))} placeholder="e.g. Ravi Kumar"/></FRow>
        <FRow label="Monthly Rent (₹)"><Inp type="number" value={arD.rent} onChange={v=>setArD(p=>({...p,rent:v}))} placeholder="5000"/></FRow>
        <FRow label="Security Deposit (₹)"><Inp type="number" value={arD.deposit} onChange={v=>setArD(p=>({...p,deposit:v}))} placeholder="10000"/></FRow>
        <FRow label="Move-in Date"><Inp type="date" value={arD.moveIn} onChange={v=>setArD(p=>({...p,moveIn:v}))}/></FRow>
        <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:12}}><Btn ghost onClick={()=>setShowAR(false)}>Cancel</Btn><Btn onClick={doAR}>Add Room</Btn></div>
      </Modal>
      <Modal show={showER} onClose={()=>setShowER(false)} title="✏️ Edit Room">
        <FRow label="Room Number"><Inp value={erD.name} onChange={v=>setErD(p=>({...p,name:v}))} placeholder="Room number"/></FRow>
        <FRow label="Tenant Name"><Inp value={erD.tenant} onChange={v=>setErD(p=>({...p,tenant:v}))} placeholder="Tenant name"/></FRow>
        <FRow label="Monthly Rent (₹)"><Inp type="number" value={erD.rent} onChange={v=>setErD(p=>({...p,rent:v}))} placeholder="Rent"/></FRow>
        <FRow label="Security Deposit (₹)"><Inp type="number" value={erD.deposit} onChange={v=>setErD(p=>({...p,deposit:v}))} placeholder="Deposit"/></FRow>
        <FRow label="Move-in Date"><Inp type="date" value={erD.moveIn} onChange={v=>setErD(p=>({...p,moveIn:v}))}/></FRow>
        <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:12}}><Btn ghost onClick={()=>setShowER(false)}>Cancel</Btn><Btn onClick={doSER}>Save</Btn></div>
      </Modal>
      <Modal show={showRN} onClose={()=>setShowRN(false)} title="✏️ Rename Building">
        <FRow label="Name"><Inp value={rnN} onChange={setRnN} placeholder="Building name"/></FRow>
        <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:12}}><Btn ghost onClick={()=>setShowRN(false)}>Cancel</Btn><Btn onClick={doRN}>Save</Btn></div>
      </Modal>
    </div>
  );
}

// ── OUTSTANDING ───────────────────────────────────
function OutstandingPage({S,PAY,setPAY,allMs,getMS,showToast}){
  const[open,setOpen]=useState({});
  const gPS=(rid,m)=>PAY[rid]?.[m]||{rent:'unpaid',water:'unpaid'};
  const togR=(rid,m)=>{const ps=gPS(rid,m);const nv=ps.rent==='paid'?'unpaid':'paid';setPAY(p=>({...p,[rid]:{...p[rid],[m]:{...(p[rid]?.[m]||{rent:'unpaid',water:'unpaid'}),rent:nv}}}));};
  const togW=(rid,m)=>{const ps=gPS(rid,m);const nv=ps.water==='paid'?'unpaid':'paid';setPAY(p=>({...p,[rid]:{...p[rid],[m]:{...(p[rid]?.[m]||{rent:'unpaid',water:'unpaid'}),water:nv}}}));};
  const ms=allMs();
  const rooms=S.buildings.flatMap(b=>b.rooms.map(r=>({b,r})));
  let totC=0,totD=0,totDep=0;
  const entries=rooms.map(({b,r})=>{
    let due=0,coll=0,hist=[];
    ms.forEach(m=>{
      const sv=getMS(m);if(!sv)return;
      const sb=sv.buildings.find(x=>x.id===b.id);
      const sr=sb?.rooms.find(x=>x.id===r.id);if(!sr)return;
      const wb=sr.meters.reduce((s,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s+d*(+mt.rate||0);},0);
      const rr=+sr.rent||0,ps=gPS(r.id,m);
      const rd=(ps.rent!=='paid'?rr:0)+(ps.water!=='paid'?wb:0);
      const rc=(ps.rent==='paid'?rr:0)+(ps.water==='paid'?wb:0);
      if(rr>0||wb>0){hist.push({m,rent:rr,water:wb,total:rr+wb,rp:ps.rent==='paid',wp:ps.water==='paid',due:rd,coll:rc});due+=rd;coll+=rc;}
    });
    totC+=coll;totD+=due;totDep+=+(r.deposit||0);
    return{b,r,hist,due,coll};
  }).filter(e=>e.hist.length>0).sort((a,b)=>b.due-a.due);

  return(
    <div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(100px,1fr))',gap:9,marginBottom:14}}>
        {[{i:'✅',v:'₹'+totC.toFixed(0),l:'Collected',c:'#3fb950'},{i:'⚠️',v:'₹'+totD.toFixed(0),l:'Outstanding',c:'#f85149'},{i:'🔐',v:'₹'+totDep.toFixed(0),l:'Deposits',c:'#a78bfa'},{i:'🏠',v:rooms.length,l:'Rooms',c:'#d29922'}].map(s=>(
          <div key={s.l} style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,padding:12}}>
            <div style={{fontSize:'1.2rem',marginBottom:6}}>{s.i}</div>
            <div style={{fontSize:'1.2rem',fontWeight:700,color:s.c,lineHeight:1}}>{s.v}</div>
            <div style={{fontSize:'.65rem',color:'#8b949e',marginTop:3}}>{s.l}</div>
          </div>
        ))}
      </div>
      {!entries.length&&<div style={{textAlign:'center',padding:30,color:'#484f58'}}>No billing history yet. Enter readings and save months.</div>}
      {entries.map((e,i)=>{
        const ac=PAL[i%PAL.length],um=e.hist.filter(h=>h.due>0).length,isO=open[e.r.id];
        return(
          <div key={e.r.id} style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,marginBottom:9,overflow:'hidden'}}>
            <div onClick={()=>setOpen(p=>({...p,[e.r.id]:!p[e.r.id]}))} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'11px 14px',cursor:'pointer',flexWrap:'wrap',gap:8}}>
              <div style={{display:'flex',alignItems:'center',gap:9}}>
                <div style={{width:34,height:34,borderRadius:'50%',background:ac+'22',color:ac,display:'flex',alignItems:'center',justifyContent:'center',fontSize:'.85rem',fontWeight:700,flexShrink:0}}>{e.r.name.slice(0,2).toUpperCase()}</div>
                <div>
                  <div style={{fontWeight:700,fontSize:'.88rem'}}>Room {e.r.name}</div>
                  <div style={{fontSize:'.66rem',color:'#8b949e',marginTop:2}}>{e.r.tenant?<span style={{color:'#a78bfa'}}>{e.r.tenant} · </span>:null}{e.b.name} · {um} month{um!==1?'s':''} unpaid</div>
                </div>
              </div>
              <div style={{textAlign:'right'}}>
                <div style={{fontWeight:700,fontSize:'.95rem',color:e.due>0?'#f85149':'#3fb950'}}>{e.due>0?'₹'+e.due.toFixed(2):'✅ Clear'}</div>
                <div style={{fontSize:'.65rem',color:'#8b949e',marginTop:1}}>₹{e.coll.toFixed(2)} collected</div>
                {e.r.deposit?<div style={{fontSize:'.65rem',color:'#a78bfa',marginTop:1}}>🔐 ₹{e.r.deposit} deposit</div>:null}
              </div>
            </div>
            {isO&&(
              <div style={{padding:'12px 14px',borderTop:'1px solid #21262d'}}>
                {e.r.deposit?<div style={{background:'rgba(110,64,201,.08)',border:'1px solid rgba(110,64,201,.25)',borderRadius:8,padding:'9px 12px',marginBottom:10,display:'flex',alignItems:'center',justifyContent:'space-between'}}><div style={{fontSize:'.7rem',color:'#a78bfa',fontWeight:700}}>🔐 Security Deposit{e.r.moveIn?' · Move-in: '+e.r.moveIn:''}</div><div style={{fontWeight:700,color:'#a78bfa'}}>₹{e.r.deposit}</div></div>:null}
                <div style={{overflowX:'auto'}}>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:'.73rem'}}>
                    <thead><tr style={{background:'#0d1117'}}>
                      {['Month','Rent','Water','Total','Rent','Water','Due'].map((h,i)=><th key={i} style={{color:'#8b949e',padding:'6px 8px',textAlign:'center',fontSize:'.6rem',textTransform:'uppercase',fontWeight:700,borderBottom:'1px solid #21262d'}}>{h}</th>)}
                    </tr></thead>
                    <tbody>
                      {e.hist.map(h=>(
                        <tr key={h.m} style={{borderBottom:'1px solid rgba(255,255,255,.04)'}}>
                          <td style={{padding:'6px 8px',textAlign:'center'}}><span style={{background:'rgba(88,166,255,.1)',color:'#58a6ff',borderRadius:4,padding:'2px 6px',fontWeight:700,fontSize:'.68rem'}}>{fmtS(h.m)}</span></td>
                          <td style={{textAlign:'right',padding:'6px 8px'}}>₹{h.rent.toFixed(2)}</td>
                          <td style={{textAlign:'right',padding:'6px 8px'}}>₹{h.water.toFixed(2)}</td>
                          <td style={{textAlign:'right',padding:'6px 8px',fontWeight:700}}>₹{h.total.toFixed(2)}</td>
                          <td style={{textAlign:'center',padding:'6px 8px'}}><button onClick={()=>togR(e.r.id,h.m)} style={{border:'none',borderRadius:5,padding:'2px 8px',fontSize:'.66rem',fontWeight:700,cursor:'pointer',fontFamily:'inherit',background:h.rp?'rgba(63,185,80,.12)':'rgba(248,81,73,.12)',color:h.rp?'#3fb950':'#f85149'}}>Rent {h.rp?'✓':'✗'}</button></td>
                          <td style={{textAlign:'center',padding:'6px 8px'}}><button onClick={()=>togW(e.r.id,h.m)} style={{border:'none',borderRadius:5,padding:'2px 8px',fontSize:'.66rem',fontWeight:700,cursor:'pointer',fontFamily:'inherit',background:h.wp?'rgba(88,166,255,.12)':'rgba(248,81,73,.12)',color:h.wp?'#58a6ff':'#f85149'}}>Water {h.wp?'✓':'✗'}</button></td>
                          <td style={{textAlign:'right',padding:'6px 8px',fontWeight:700,color:h.due>0?'#f85149':'#3fb950'}}>{h.due>0?'₹'+h.due.toFixed(2):'✓'}</td>
                        </tr>
                      ))}
                      <tr style={{background:'rgba(88,166,255,.06)',fontWeight:700}}>
                        <td style={{padding:'6px 8px',textAlign:'center'}}>Total</td>
                        <td style={{textAlign:'right',padding:'6px 8px'}}>₹{e.hist.reduce((s,h)=>s+h.rent,0).toFixed(2)}</td>
                        <td style={{textAlign:'right',padding:'6px 8px'}}>₹{e.hist.reduce((s,h)=>s+h.water,0).toFixed(2)}</td>
                        <td style={{textAlign:'right',padding:'6px 8px'}}>₹{e.hist.reduce((s,h)=>s+h.total,0).toFixed(2)}</td>
                        <td colSpan={2}></td>
                        <td style={{textAlign:'right',padding:'6px 8px',color:e.due>0?'#f85149':'#3fb950'}}>₹{e.due.toFixed(2)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ── TENANTS ───────────────────────────────────────
function TenantsPage({S,PAY,allMs,getMS}){
  const[search,setSearch]=useState('');
  const gPS=(rid,m)=>PAY[rid]?.[m]||{rent:'unpaid',water:'unpaid'};
  const ms=allMs();
  const rooms=S.buildings.flatMap(b=>b.rooms.map(r=>({b,r})));
  const filtered=rooms.filter(({b,r})=>{if(!search)return true;return r.name.toLowerCase().includes(search.toLowerCase())||(r.tenant||'').toLowerCase().includes(search.toLowerCase())||b.name.toLowerCase().includes(search.toLowerCase());});
  return(
    <div>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Search…" style={{width:'100%',background:'#161b22',border:'1px solid #30363d',borderRadius:8,padding:'8px 12px',color:'#e6edf3',fontSize:'.78rem',outline:'none',marginBottom:12,fontFamily:'inherit'}}/>
      {!filtered.length&&<div style={{textAlign:'center',padding:28,color:'#484f58'}}>No tenants found</div>}
      {filtered.map(({b,r},i)=>{
        const ac=PAL[i%PAL.length];
        let totD=0,totC=0;
        ms.forEach(m=>{const sv=getMS(m);if(!sv)return;const sb=sv.buildings.find(x=>x.id===b.id);const sr=sb?.rooms.find(x=>x.id===r.id);if(!sr)return;const wb=sr.meters.reduce((s,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s+d*(+mt.rate||0);},0);const rr=+sr.rent||0,ps=gPS(r.id,m);totD+=(ps.rent!=='paid'?rr:0)+(ps.water!=='paid'?wb:0);totC+=(ps.rent==='paid'?rr:0)+(ps.water==='paid'?wb:0);});
        return(
          <div key={r.id} style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,marginBottom:9,overflow:'hidden'}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'11px 14px',flexWrap:'wrap',gap:8}}>
              <div style={{display:'flex',alignItems:'center',gap:9}}>
                <div style={{width:34,height:34,borderRadius:'50%',background:ac+'22',color:ac,display:'flex',alignItems:'center',justifyContent:'center',fontSize:'.85rem',fontWeight:700,flexShrink:0}}>{r.name.slice(0,2).toUpperCase()}</div>
                <div><div style={{fontWeight:700,fontSize:'.88rem'}}>Room {r.name}</div><div style={{fontSize:'.66rem',color:'#8b949e',marginTop:2}}>{b.name}{r.moveIn?' · Since '+r.moveIn:''}</div></div>
              </div>
              {r.rent?<div style={{color:'#3fb950',fontWeight:700,fontSize:'.88rem'}}>₹{r.rent}/mo</div>:null}
            </div>
            <div style={{padding:'12px 14px',borderTop:'1px solid #21262d'}}>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
                {[{l:'👤 Tenant',v:r.tenant||'—',c:'#e6edf3'},{l:'🏠 Monthly Rent',v:'₹'+(r.rent||0),c:'#3fb950'},{l:'🔐 Security Deposit',v:'₹'+(r.deposit||0),c:'#a78bfa'},{l:'📅 Move-in Date',v:r.moveIn||'—',c:'#8b949e'},{l:'✅ Total Collected',v:'₹'+totC.toFixed(2),c:'#3fb950'},{l:'⚠️ Outstanding',v:'₹'+totD.toFixed(2),c:totD>0?'#f85149':'#3fb950'}].map(x=>(
                  <div key={x.l} style={{background:'#0d1117',borderRadius:7,padding:'7px 9px'}}>
                    <div style={{fontSize:'.58rem',color:'#484f58',fontWeight:700,textTransform:'uppercase',letterSpacing:'.4px',marginBottom:3}}>{x.l}</div>
                    <div style={{fontSize:'.82rem',fontWeight:700,color:x.c}}>{x.v}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ── DASHBOARD ─────────────────────────────────────
function DashboardPage({S,PAY,month,allMs,getMS}){
  const gPS=(rid,m)=>PAY[rid]?.[m]||{rent:'unpaid',water:'unpaid'};
  const rooms=S.buildings.flatMap(b=>b.rooms.map(r=>({b,r})));
  const tBill=rooms.reduce((s,{r})=>s+rTotal(r),0);
  let coll=0,due=0;
  rooms.forEach(({r})=>{const ps=gPS(r.id,month);if(ps.rent==='paid')coll+=(+r.rent||0);else due+=(+r.rent||0);if(ps.water==='paid')coll+=wBill(r);else due+=wBill(r);});
  const ms=allMs().slice(-8);
  const mdata=ms.map(m=>{const sv=getMS(m);if(!sv)return{l:fmtS(m),v:0};const tot=sv.buildings.reduce((s,b)=>s+b.rooms.reduce((s2,r)=>{const wb=r.meters.reduce((s3,mt)=>{const d=Math.max(0,(+mt.curr||0)-(+mt.last||0));return s3+d*(+mt.rate||0);},0);return s2+(+r.rent||0)+wb;},0),0);return{l:fmtS(m),v:tot};});
  const mxV=Math.max(...mdata.map(d=>d.v),1);
  const topR=rooms.map(({b,r})=>({bldg:b.name,room:r,bill:rTotal(r)})).sort((a,b)=>b.bill-a.bill).slice(0,5);
  const mxR=Math.max(...topR.map(h=>h.bill),1);
  return(
    <div>
      <div style={{fontWeight:700,marginBottom:12}}>📊 Dashboard — {fmt(month)}</div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(100px,1fr))',gap:9,marginBottom:14}}>
        {[{i:'💰',v:'₹'+tBill.toFixed(0),l:'Total Bill',c:'#58a6ff'},{i:'✅',v:'₹'+coll.toFixed(0),l:'Collected',c:'#3fb950'},{i:'⚠️',v:'₹'+due.toFixed(0),l:'Pending',c:'#f85149'},{i:'🏠',v:rooms.length,l:'Rooms',c:'#d29922'}].map(s=>(
          <div key={s.l} style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,padding:12}}>
            <div style={{fontSize:'1.2rem',marginBottom:6}}>{s.i}</div>
            <div style={{fontSize:'1.2rem',fontWeight:700,color:s.c,lineHeight:1}}>{s.v}</div>
            <div style={{fontSize:'.65rem',color:'#8b949e',marginTop:3}}>{s.l}</div>
          </div>
        ))}
      </div>
      <div style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,padding:13,marginBottom:11}}>
        <div style={{fontSize:'.8rem',fontWeight:700,marginBottom:11}}>💰 Monthly Collection Trend</div>
        {ms.length===0?<div style={{color:'#484f58',fontSize:'.76rem',textAlign:'center',padding:12}}>Save data for multiple months to see trend</div>:(
          <div style={{display:'flex',alignItems:'flex-end',gap:5,height:90,paddingBottom:16}}>
            {mdata.map((d,i)=>{const h=Math.max(3,d.v/mxV*78);return(
              <div key={d.l} style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',gap:2,height:'100%',justifyContent:'flex-end'}}>
                <div style={{fontSize:'.54rem',color:'#8b949e',fontWeight:600}}>{d.v>999?(d.v/1000).toFixed(1)+'k':d.v.toFixed(0)}</div>
                <div style={{width:'100%',height:h,background:PAL[i%PAL.length],borderRadius:'4px 4px 0 0'}}></div>
                <div style={{fontSize:'.54rem',color:'#484f58',textAlign:'center',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'100%'}}>{d.l}</div>
              </div>
            );})}
          </div>
        )}
      </div>
      {topR.length>0&&(
        <div style={{background:'#161b22',border:'1px solid #21262d',borderRadius:10,padding:13}}>
          <div style={{fontSize:'.8rem',fontWeight:700,marginBottom:11}}>🏆 Top Rooms by Bill</div>
          {topR.map((h,i)=>(
            <div key={h.room.id} style={{display:'flex',alignItems:'center',gap:7,padding:'5px 0',borderBottom:'1px solid rgba(255,255,255,.04)'}}>
              <div style={{width:19,height:19,borderRadius:'50%',background:PAL[i]+'22',color:PAL[i],display:'flex',alignItems:'center',justifyContent:'center',fontSize:'.62rem',fontWeight:700}}>{i+1}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:'.76rem',fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>Room {h.room.name} <span style={{color:'#484f58',fontSize:'.66rem'}}>{h.bldg}</span></div>
                <div style={{background:'#21262d',borderRadius:2,height:4,marginTop:3}}><div style={{width:(h.bill/mxR*100).toFixed(1)+'%',background:PAL[i],borderRadius:2,height:4}}></div></div>
              </div>
              <div style={{fontWeight:700,fontSize:'.8rem',color:PAL[i]}}>₹{h.bill.toFixed(0)}</div>
            </div>
          ))}
        </div>
      )}
      {!rooms.length&&<div style={{textAlign:'center',padding:'40px 20px',color:'#484f58'}}>Add buildings and rooms to see dashboard</div>}
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────
export default function App(){
  const[month,setMonth]=useState(DEF_M);
  const[page,setPage]=useState('b');
  const[S,setS]=useState({buildings:[]});
  const[PAY,setPAY]=useState({});
  const[toastMsg,setToastMsg]=useState('');
  const[saved,setSaved]=useState(false);

  const showToast=useCallback(msg=>{setToastMsg(msg);setTimeout(()=>setToastMsg(''),3000);},[]);

  useEffect(()=>{
    // 1. Load from localStorage immediately (instant)
    const st=lsGet('pm_s');
    const py=lsGet('pm_p');
    const m=localStorage.getItem('pm_month');
    if(st){
      let loaded=st;
      if(m){
        const rd=lsGet('pm_r_'+m);
        if(rd){
          loaded={...st,buildings:st.buildings.map(b=>({...b,rooms:b.rooms.map(r=>({...r,meters:r.meters.map(mt=>{const x=rd[r.id]?.[mt.id];return x?{...mt,last:x.last||'',curr:x.curr||''}:{...mt};})}))}))};
        }
      }
      setS(loaded);
    }
    if(py)setPAY(py);
    if(m)setMonth(m);

    // 2. Pull from Google Sheets (may have newer data from other phone)
    pullFromSheets().then(cloud=>{
      if(!cloud)return;
      const localTs=parseInt(localStorage.getItem('pm_ts')||'0');
      // Use cloud if: local is empty OR cloud is newer
      const localEmpty=!st||!st.buildings||st.buildings.length===0;
      if(localEmpty||(cloud.ts&&cloud.ts>localTs)){
        try{
          const cs=JSON.parse(cloud.s);
          const cp=cloud.p?JSON.parse(cloud.p):{};
          // Restore all monthly readings
          if(cloud.rd){Object.keys(cloud.rd).forEach(k=>localStorage.setItem(k,cloud.rd[k]));}
          lsSet('pm_s',cs);lsSet('pm_p',cp);
          if(cloud.month){localStorage.setItem('pm_month',cloud.month);setMonth(cloud.month);}
          if(cloud.ts)localStorage.setItem('pm_ts',cloud.ts);
          // Apply readings for current month
          const cm=cloud.month||m||DEF_M;
          const rd=lsGet('pm_r_'+cm);
          if(rd){
            const loaded={...cs,buildings:cs.buildings.map(b=>({...b,rooms:b.rooms.map(r=>({...r,meters:r.meters.map(mt=>{const x=rd[r.id]?.[mt.id];return x?{...mt,last:x.last||'',curr:x.curr||''}:{...mt};})}))}))};
            setS(loaded);
          }else{setS(cs);}
          setPAY(cp);
          showToast('☁️ Data loaded from cloud!');
        }catch(e){}
      }
    });
  },[]);

  // Auto-save to localStorage AND Google Sheets
  useEffect(()=>{
    const t=setTimeout(()=>{
      lsSet('pm_s',S);
      lsSet('pm_p',PAY);
      localStorage.setItem('pm_month',month);
      const rd={};
      S.buildings.forEach(b=>b.rooms.forEach(r=>r.meters.forEach(mt=>{if(!rd[r.id])rd[r.id]={};rd[r.id][mt.id]={last:mt.last||'',curr:mt.curr||''};})));
      lsSet('pm_r_'+month,rd);
      if(S.buildings.length>0){
        setSaved(true);
        setTimeout(()=>setSaved(false),1500);
        // Build full backup payload
        const allRd={};
        for(let i=0;i<localStorage.length;i++){
          const k=localStorage.key(i);
          if(k&&k.startsWith('pm_r_'))allRd[k]=localStorage.getItem(k);
        }
        const payload={
          s:JSON.stringify(S),
          p:JSON.stringify(PAY),
          month:month,
          rd:allRd,
          ts:Date.now()
        };
        pushToSheets(payload);
      }
    },500);
    return()=>clearTimeout(t);
  },[S,PAY,month]);

  const onMonthChange=m=>{
    // Save current readings first
    const rd={};
    S.buildings.forEach(b=>b.rooms.forEach(r=>r.meters.forEach(mt=>{if(!rd[r.id])rd[r.id]={};rd[r.id][mt.id]={last:mt.last||'',curr:mt.curr||''};})));
    lsSet('pm_r_'+month,rd);
    lsSet('pm_s',S);
    setMonth(m);
    localStorage.setItem('pm_month',m);
    // Load readings for new month
    const saved_rd=lsGet('pm_r_'+m);
    if(saved_rd){
      // Month has data - load it
      setS(p=>({...p,buildings:p.buildings.map(b=>({...b,rooms:b.rooms.map(r=>({...r,meters:r.meters.map(mt=>{const x=saved_rd[r.id]?.[mt.id];return x?{...mt,last:x.last||'',curr:x.curr||''}:{...mt,last:'',curr:''};})}))}))})  );
    }else{
      // New month - auto copy current → last
      const prev_rd=lsGet('pm_r_'+getPK(m));
      if(prev_rd){
        setS(p=>({...p,buildings:p.buildings.map(b=>({...b,rooms:b.rooms.map(r=>({...r,meters:r.meters.map(mt=>{const x=prev_rd[r.id]?.[mt.id];return x?{...mt,last:x.curr||'',curr:''}:{...mt,last:'',curr:''};})}))}))})  );
        showToast('✅ Last month current readings copied as new Last readings!');
      }else{
        setS(p=>({...p,buildings:p.buildings.map(b=>({...b,rooms:b.rooms.map(r=>({...r,meters:r.meters.map(mt=>({...mt,last:'',curr:''}))}))}))})  );
      }
    }
    showToast('📅 '+fmt(m));
  };

  const allMs=useCallback(()=>{const ms=[];for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i);if(k&&k.startsWith('pm_r_'))ms.push(k.replace('pm_r_',''));}return ms.sort();},[]);
  const getMS=useCallback(m=>{const st=lsGet('pm_s');if(!st)return null;const rd=lsGet('pm_r_'+m);if(rd){st.buildings.forEach(b=>b.rooms.forEach(r=>r.meters.forEach(mt=>{const x=rd[r.id]?.[mt.id];mt.last=x?x.last:'';mt.curr=x?x.curr:'';})));}return st;},[]);

  const NAVS=[{id:'b',i:'🏢',l:'Billing'},{id:'o',i:'💰',l:'Outstanding'},{id:'t',i:'👤',l:'Tenants'},{id:'d',i:'📊',l:'Dashboard'}];

  return(
    <div style={{background:'#0d1117',minHeight:'100vh',color:'#e6edf3',fontFamily:'system-ui,-apple-system,sans-serif'}}>
      {toastMsg&&<div style={{position:'fixed',bottom:80,left:'50%',transform:'translateX(-50%)',background:'#161b22',border:'1px solid #30363d',borderRadius:9,padding:'9px 18px',fontSize:'.78rem',fontWeight:600,zIndex:999,whiteSpace:'nowrap',boxShadow:'0 4px 20px rgba(0,0,0,.6)',pointerEvents:'none'}}>{toastMsg}</div>}
      <div style={{background:'#161b22',borderBottom:'1px solid #21262d',padding:'10px 14px',display:'flex',alignItems:'center',justifyContent:'space-between',position:'sticky',top:0,zIndex:50,gap:8}}>
        <div style={{fontWeight:700,fontSize:'.95rem'}}>🏢 Property<b style={{color:'#58a6ff'}}>Manager</b></div>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          {saved&&<span style={{fontSize:'.68rem',color:'#3fb950'}}>✅ Saved</span>}
          <input type="month" value={month} onChange={e=>onMonthChange(e.target.value)} style={{background:'#0d1117',border:'1px solid #30363d',borderRadius:6,padding:'5px 9px',color:'#e6edf3',fontSize:'.78rem',outline:'none',colorScheme:'dark'}}/>
        </div>
      </div>
      <div style={{background:'#161b22',borderBottom:'1px solid #21262d',display:'flex',overflowX:'auto'}}>
        {NAVS.map(n=>(
          <button key={n.id} onClick={()=>setPage(n.id)} style={{flex:1,border:'none',background:'none',color:page===n.id?'#58a6ff':'#8b949e',fontFamily:'inherit',fontSize:'.7rem',fontWeight:page===n.id?700:500,padding:'9px 4px',cursor:'pointer',borderBottom:page===n.id?'2px solid #58a6ff':'2px solid transparent',marginBottom:-1,minWidth:60}}>
            <span style={{display:'block',fontSize:'.95rem',marginBottom:2}}>{n.i}</span>{n.l}
          </button>
        ))}
      </div>
      <div style={{padding:'14px 12px 80px',maxWidth:900,margin:'0 auto'}}>
        {page==='b'&&<BillingPage S={S} setS={setS} PAY={PAY} setPAY={setPAY} month={month} showToast={showToast} allMs={allMs} getMS={getMS}/>}
        {page==='o'&&<OutstandingPage S={S} PAY={PAY} setPAY={setPAY} allMs={allMs} getMS={getMS} showToast={showToast}/>}
        {page==='t'&&<TenantsPage S={S} PAY={PAY} allMs={allMs} getMS={getMS}/>}
        {page==='d'&&<DashboardPage S={S} PAY={PAY} month={month} allMs={allMs} getMS={getMS}/>}
      </div>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(React.createElement(App));
</script>
</body>
</html>
