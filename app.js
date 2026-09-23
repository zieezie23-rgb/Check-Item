const $=id=>document.getElementById(id);

/* ============================================================
   SHARED HELPERS (identik di kedua aplikasi asal)
   ============================================================ */
function esc(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function norm(v){return String(v??'').replace(/^\uFEFF/,'').trim()}
function num(v){let n=Number(v);return Number.isFinite(n)?n:0}
function ext(n){return String(n||'').toLowerCase().split('.').pop()}
function sleep(){return new Promise(r=>setTimeout(r,0))}
function isSpreadsheet(n){return ['xlsx','xls','xlsm','xlsb'].includes(ext(n))}
function detectDelimiter(text){const lines=text.split(/\r?\n/).filter(x=>x.trim()).slice(0,5);const candidates=[',',';','\t','|'];let best=',',score=-1;for(const d of candidates){let n=0;for(const line of lines){let q=false,c=0;for(let i=0;i<line.length;i++){const ch=line[i];if(ch==='"')q=!q;else if(ch===d&&!q)c++;}n+=c;}if(n>score){score=n;best=d}}return best}
function hmap(h){const m={};h.forEach((v,i)=>{if(v)m[v]=i});return m}
function headRow(rows,i){return (rows[i]||[]).map(x=>norm(x).toUpperCase())}
function unpipe(rows){const sample=rows.slice(0,10).filter(r=>(r||[]).some(x=>norm(x)!==''));const hit=sample.filter(r=>typeof r[0]==='string'&&r[0].includes('|')&&r.slice(1).every(x=>norm(x)==='')).length;if(!sample.length||hit<Math.ceil(sample.length/2))return rows;return rows.map(r=>String((r||[])[0]??'').split('|').map(x=>x.trim()))}
function parseDate(v){if(v instanceof Date&&!isNaN(v))return new Date(v.getFullYear(),v.getMonth(),v.getDate());if(typeof v==='number'&&v>0&&v<600000){const d=XLSX.SSF.parse_date_code(v);if(d)return new Date(d.y,d.m-1,d.d)}const s=norm(v);if(!s)return null;let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(m)return new Date(+m[3],+m[1]-1,+m[2]);return null}
function sameDate(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate()}
function daysBetween(a,b){return Math.round((b-a)/86400000)}
function fmtDate(d){return d?d.toLocaleDateString('id-ID',{day:'2-digit',month:'short',year:'2-digit'}):''}
function today0(){const t=new Date();t.setHours(0,0,0,0);return t}

/* ============================================================
   BAGIAN "BB" — logic app "Check Barang Baru" (dipertahankan)
   ============================================================ */
const bb_resultCols=['PRDCD','DESC2','ZONA','ALAMAT','FRAC','ACOST','PTAG','TGL AKTIF','TGL Open TAG','X BPB','FIRST BPB','Tgl Kedatangan Terakhir','QTY BPB','STOCK','DSI','JTD','OVER','PKM 0','Est Kebutuhan','Hari','KET1','KET2'];
const bb_colKey={'PRDCD':'PRDCD','DESC2':'DESC2','ZONA':'ZONA','ALAMAT':'ALAMAT','FRAC':'FRAC','ACOST':'ACOST','PTAG':'PTAG','TGL AKTIF':'TGL_AKTIF','TGL Open TAG':'TGL_OPEN_TAG','X BPB':'X_BPB','FIRST BPB':'FIRST_BPB','Tgl Kedatangan Terakhir':'LAST_BPB','QTY BPB':'QTY_BPB','STOCK':'STOCK','DSI':'DSI','JTD':'JTD','OVER':'OVER','PKM 0':'PKM0','Est Kebutuhan':'EST','Hari':'HARI','KET1':'KET1','KET2':'KET2'};
const bb_numCols=new Set(['FRAC','ACOST','X BPB','QTY BPB','STOCK','DSI','JTD','OVER','PKM 0','Est Kebutuhan']);
const bb_dateCols=new Set(['TGL AKTIF','TGL Open TAG','FIRST BPB','Tgl Kedatangan Terakhir']);
const bb_ket2Names=['Belum Ada JTD Toko','PKM Toko 0 (Nol)','Over Stock di Toko','Stock Exist masih >50% BPB terakhir'];
const bb_kindLabel={PRODMAST:'PRODMAST',POMGG:'POMGG',HISTAG:'HISTAG',STOCK:'Stock Akhir',MASTER:'Master'};
const bbState={prodmast:null,pomgg:null,master:null,stock:null,histag:null,kandidat:[],rows:[],tgl:''};

function bb_parseCSVText(text,delim){const rows=[];let row=[],field='',inQ=false,i=0,n=text.length;while(i<n){const c=text[i];if(inQ){if(c==='"'){if(text[i+1]==='"'){field+='"';i+=2;continue}else{inQ=false;i++;continue}}field+=c;i++;continue}if(c==='"'&&field===''){inQ=true;i++;continue}if(c===delim){row.push(field);field='';i++;continue}if(c==='\r'){i++;continue}if(c==='\n'){row.push(field);rows.push(row);row=[];field='';i++;continue}field+=c;i++}if(field!==''||row.length){row.push(field);rows.push(row)}return rows}
async function bb_readAny(file){const ab=await file.arrayBuffer(),u=new Uint8Array(ab);if(isSpreadsheet(file.name)||(u[0]===80&&u[1]===75)){const wb=XLSX.read(ab,{type:'array',cellDates:true,raw:true});const ws=wb.Sheets[wb.SheetNames[0]];return ws?XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:false}):[]}let text=new TextDecoder('utf-8').decode(ab);if(text.charCodeAt(0)===0xFEFF)text=text.slice(1);const delim=detectDelimiter(text);return bb_parseCSVText(text,delim)}

function bb_detectKind(rows){for(let i=0;i<Math.min(rows.length,15);i++){const h=headRow(rows,i);if(!h.length)continue;if(h.includes('BPB_DATE')&&h.includes('BPB_QTY'))return 'POMGG';if(h.includes('PTAG_OLD')&&h.includes('PTAG_NEW'))return 'HISTAG';if(h.includes('BRG_AKTIF')&&h.includes('PRDCD'))return 'PRODMAST';if((h.includes('PLU')||h.includes('PLUKODE'))&&(h.includes('QTY_LPP')||h.includes('QTY_AKHIR')))return 'STOCK'}return null}
function bb_parseMaster(rows){let h3=-1,h4=-1;for(let i=0;i<10;i++){const r=headRow(rows,i);if(r.includes('PRDCD')&&r.includes('DESC2'))h3=i;if(r.includes('ST')&&r.includes('ZONA'))h4=i}if(h3<0)return null;const a=hmap(headRow(rows,h3)),b=h4>=0?hmap(headRow(rows,h4)):{};const map={...a};for(const k of ['ST','ZONA','LT','RAK','BAR','CELL'])if(b[k]!==undefined)map[k]=b[k];if(map.LT===undefined)return null;const start=Math.max(h3,h4)+2,out=new Map();for(let i=start;i<rows.length;i++){const r=rows[i]||[],id=norm(r[map.PRDCD]);if(!id||norm(r[map.LT])==='')continue;if(!out.has(id))out.set(id,{row:r,map})}return out}
function bb_parseStock(rows){const IDK=['PLU','PLUKODE'],QK=['QTY_LPP','QTY_AKHIR'];let row=-1,headers=headRow(rows,0);for(let i=0;i<Math.min(rows.length,10);i++){const h=headRow(rows,i);if(IDK.some(k=>h.includes(k))&&QK.some(k=>h.includes(k))){row=i;headers=h;break}}const idCol=IDK.map(k=>headers.indexOf(k)).find(x=>x>=0),qCol=QK.map(k=>headers.indexOf(k)).find(x=>x>=0);if(row<0||idCol===undefined||qCol===undefined)throw Error('Stock Akhir: kolom PLU / QTY_LPP tidak ditemukan.');const out=new Map();for(let i=row+1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[idCol]);if(id)out.set(id,num(r[qCol]))}return out}
function bb_parseProdmast(rows){const h=hmap(headRow(rows,0));for(const k of ['PRDCD','DESC2','CAT_COD','BRG_AKTIF','TGL_TAMBAH'])if(h[k]===undefined)throw Error('PRODMAST: kolom '+k+' tidak ditemukan.');const out=[];for(let i=1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[h.PRDCD]);if(!id)continue;out.push({PRDCD:id,DESC2:norm(r[h.DESC2]),CAT_COD:norm(r[h.CAT_COD]),BRG_AKTIF:norm(r[h.BRG_AKTIF]),TGL_TAMBAH:r[h.TGL_TAMBAH]})}return out}
function bb_parsePomgg(rows){const h=hmap(headRow(rows,0));for(const k of ['PRDCD','BPB_DATE','BPB_QTY'])if(h[k]===undefined)throw Error('POMGG: kolom '+k+' tidak ditemukan.');const idx=new Map();for(let i=1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[h.PRDCD]);if(!id)continue;const date=parseDate(r[h.BPB_DATE]);if(!date)continue;const qty=num(r[h.BPB_QTY]);const nopb=h.NOPB!==undefined?norm(r[h.NOPB]):'';if(!idx.has(id))idx.set(id,[]);idx.get(id).push({date,qty,nopb})}return idx}
function bb_parseHistag(rows){const h=hmap(headRow(rows,0));for(const k of ['PLU','PTAG_OLD','PTAG_NEW','TGL_AKHIR'])if(h[k]===undefined)throw Error('HISTAG: kolom '+k+' tidak ditemukan.');const byId=new Map();for(let i=1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[h.PLU]);if(!id)continue;const tglAkhir=parseDate(r[h.TGL_AKHIR]),po=norm(r[h.PTAG_OLD]),pn=norm(r[h.PTAG_NEW]).toUpperCase();if(!byId.has(id))byId.set(id,[]);byId.get(id).push({tglAkhir,po,pn})}const out=new Map();for(const [id,list] of byId){const withDate=list.filter(x=>x.tglAkhir);const best=withDate.length?withDate.reduce((a,b)=>b.tglAkhir>a.tglAkhir?b:a):list[list.length-1];const ket=(best.po!==''&&['E','D','S','L',''].includes(best.pn))?'Item Open Tag':'';out.set(id,{ket,tglAkhir:best.tglAkhir})}return out}

// Catatan penggabungan: file yang tidak dikenali BB (mis. REG/PRG milik ITB) DILEWATI, bukan error —
// karena sekarang satu bucket upload dipakai bersama oleh kedua aplikasi.
async function bb_parseBahan(files){
  const out={prodmast:null,pomgg:null,master:null,stock:null,histag:null};
  for(const f of files){
    let rows;try{rows=await bb_readAny(f)}catch(e){continue}
    let kind=bb_detectKind(rows);
    if(!kind){const m=bb_parseMaster(rows);if(m){out.master=m;continue}}
    if(kind==='POMGG')out.pomgg=bb_parsePomgg(rows);
    else if(kind==='HISTAG')out.histag=bb_parseHistag(rows);
    else if(kind==='PRODMAST')out.prodmast=bb_parseProdmast(rows);
    else if(kind==='STOCK')out.stock=bb_parseStock(rows);
    else{const m=bb_parseMaster(rows);if(m)out.master=m}
  }
  const missing=['prodmast','pomgg','master','stock','histag'].filter(k=>!out[k]);
  if(missing.length)throw Error('BAHAN belum lengkap, kurang: '+missing.join(', ').toUpperCase());
  return out;
}

function bb_findHeader(rows,required,maxScan=15){const req=required.map(x=>x.toUpperCase());for(let i=0;i<Math.min(rows.length,maxScan);i++){const h=headRow(rows,i);if(req.every(k=>h.includes(k)))return {row:i,headers:h}}return {row:-1,headers:[]}}
function bb_parseSsft(rows){rows=unpipe(rows);const h=bb_findHeader(rows,['PRDCD','STOCK','PKM','MINOR']);if(h.row<0)throw Error('SS-FT: kolom PRDCD/STOCK/PKM/MINOR tidak ditemukan. Header terbaca: '+headRow(rows,0).filter(Boolean).slice(0,14).join(', '));const m=hmap(h.headers),stats=new Map();for(let i=h.row+1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[m.PRDCD]);if(!id)continue;const stock=num(r[m.STOCK]),pkm=num(r[m.PKM]),minor=num(r[m.MINOR]);let ket2='';if(pkm===0)ket2='PKM Toko 0';else if(stock+minor>pkm)ket2='Over Stock Toko';const est=minor>0?Math.max(0,Math.floor((pkm-stock)/minor)*minor):Math.max(0,pkm-stock);if(!stats.has(id))stats.set(id,{jtd:0,over:0,pkm0:0,est:0});const x=stats.get(id);x.jtd++;if(ket2==='Over Stock Toko')x.over++;if(ket2==='PKM Toko 0')x.pkm0++;x.est+=est}return stats}

function bb_mget(mrow,key){return mrow&&mrow.map[key]!==undefined?mrow.row[mrow.map[key]]:''}
function bb_classifyItem(pm,pomggIdx,stockMap,histagMap,tdy){const catcod=norm(pm.CAT_COD);const dnd=catcod===''?'':(catcod.charAt(0)==='4'?'NON DRY':'DRY');if(dnd==='NON DRY')return null;const brgAktif=norm(pm.BRG_AKTIF).toUpperCase()==='Y';const tglTambah=parseDate(pm.TGL_TAMBAH);const hariAktif=(brgAktif&&tglTambah)?daysBetween(tglTambah,tdy):null;let CI;if(hariAktif!==null&&hariAktif<90){CI='Barang Baru'}else{const h=histagMap.get(pm.PRDCD);CI=(h&&h.ket)?h.ket:''}if(CI==='')return null;const bpbAll=pomggIdx.get(pm.PRDCD)||[];const bpbPos=bpbAll.filter(x=>x.qty>0);const CJ=bpbPos.length;let CK=null,CL=0;if(CJ===1){CK=bpbAll.reduce((mx,x)=>(!mx||x.date>mx)?x.date:mx,null)}else if(CJ>1){const noNopb=bpbAll.filter(x=>x.nopb==='');CK=noNopb.reduce((mx,x)=>(!mx||x.date>mx)?x.date:mx,null)}if(CK)CL=bpbAll.filter(x=>sameDate(x.date,CK)).reduce((s,x)=>s+x.qty,0);const CM=stockMap.has(pm.PRDCD)?stockMap.get(pm.PRDCD):'';const CN=(CL>0&&CM!=='')?(CM/CL):null;let CO='';if(CJ===1&&CN!==null&&CN>=0.999995)CO=`${CI} Belum pernah ada NPB `;else if(CJ===1&&CN!==null&&CN>0.5&&CN<1)CO=`Sisa Stock Masih ada ${(CN*100).toFixed(2)}%`;else if(CK&&sameDate(CK,tdy))CO=`${CI}Datang Hari Ini`;if(CO==='')return null;const firstBpb=bpbAll.length?bpbAll.reduce((mn,x)=>(!mn||x.date<mn)?x.date:mn,null):null;return {CI,CJ,CK,CL,CM,CO,CN,tglTambah,firstBpb}}
function bb_ket2Class(ss){if(ss.jtd===0)return 'Belum Ada JTD Toko';if(ss.pkm0===ss.jtd)return 'PKM Toko 0 (Nol)';if(ss.over===ss.jtd)return 'Over Stock di Toko';return 'Stock Exist masih >50% BPB terakhir'}
function bb_buildRekap(bahan,ssftStats){const tdy=today0(),masterById=bahan.master,histagMap=bahan.histag,pomggIdx=bahan.pomgg,stockMap=bahan.stock;const out=new Map();for(const pm of bahan.prodmast){if(out.has(pm.PRDCD))continue;const cls=bb_classifyItem(pm,pomggIdx,stockMap,histagMap,tdy);if(!cls)continue;const master=masterById.get(pm.PRDCD),histag=histagMap.get(pm.PRDCD);const ss=ssftStats.get(pm.PRDCD)||{jtd:0,over:0,pkm0:0,est:0};const daysSinceLast=cls.CK?daysBetween(cls.CK,tdy):null;let dsi=0;if(daysSinceLast&&cls.CL!==''&&cls.CM!==''){const soldQty=cls.CL-cls.CM;if(soldQty!==0){const rate=soldQty/daysSinceLast;if(rate)dsi=cls.CM/rate}}if(!isFinite(dsi)||isNaN(dsi))dsi=0;out.set(pm.PRDCD,{PRDCD:pm.PRDCD,DESC2:pm.DESC2,ZONA:master?norm(bb_mget(master,'ZONA')):'',ALAMAT:master?norm(bb_mget(master,'ALAMAT')):'',FRAC:master?num(bb_mget(master,'FRAC')):'',ACOST:master?num(bb_mget(master,'ACOST')):'',PTAG:master?norm(bb_mget(master,'PTAG')):'',TGL_AKTIF:cls.tglTambah,TGL_OPEN_TAG:histag?histag.tglAkhir:null,X_BPB:cls.CJ,FIRST_BPB:cls.firstBpb,LAST_BPB:cls.CK,QTY_BPB:cls.CL,STOCK:cls.CM,DSI:dsi,JTD:ss.jtd,OVER:ss.over,PKM0:ss.pkm0,EST:ss.est,HARI:daysSinceLast!==null?`${daysSinceLast} Hari`:'',KET1:cls.CO,KET2:bb_ket2Class(ss),PERSEN:(cls.CN!==null&&cls.CN!==undefined)?cls.CN:-1})}return [...out.values()].sort((a,b)=>b.PERSEN-a.PERSEN||String(a.PRDCD).localeCompare(String(b.PRDCD),'en',{numeric:true}))}
function bb_computeCandidates(bahan){const kandidat=bahan.prodmast.filter(pm=>bb_classifyItem(pm,bahan.pomgg,bahan.stock,bahan.histag,today0()));return [...new Set(kandidat.map(x=>x.PRDCD))]}

function bb_fmt(c,v){if(v===''||v===null||v===undefined)return '';if(bb_dateCols.has(c))return fmtDate(v);if(bb_numCols.has(c))return num(v).toLocaleString('id-ID',{maximumFractionDigits:2});return v}
function bb_render(arr){const th=$('bb_resultTable').querySelector('thead'),tb=$('bb_resultTable').querySelector('tbody');th.innerHTML='<tr>'+bb_resultCols.map(c=>`<th>${esc(c)}</th>`).join('')+'</tr>';tb.innerHTML=arr.map(r=>'<tr>'+bb_resultCols.map(c=>{const v=r[bb_colKey[c]];return `<td class="${typeof v==='number'?'num':''}">${esc(bb_fmt(c,v))}</td>`}).join('')+'</tr>').join('');$('bb_summary').textContent=`${arr.length.toLocaleString('id-ID')} item`}
function bb_renderKesimpulan(){const rows=bbState.rows;if(!rows.length){$('bb_kesimpulan').textContent='Tidak ada item.';return}const tgl=new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});bbState.tgl=tgl;const cnt=n=>rows.filter(r=>r.KET2===n).length;const lines=[`*Monitoring Stock Item Barang Baru dan Open TAG*`,tgl,''];const jtd0=cnt('Belum Ada JTD Toko');if(jtd0>0){lines.push('Belum Pernah Ada NPB ke TOKO :',`- ${jtd0} Item Belum Ada JTD Toko`,'')}const stockExist=cnt('Stock Exist masih >50% BPB terakhir'),pkm0=cnt('PKM Toko 0 (Nol)');if(stockExist>0||pkm0>0){lines.push('Stock Sisa di DC');if(stockExist>0)lines.push(`- ${stockExist} Item Stock Exist masih >50% BPB terakhir`);if(pkm0>0)lines.push(`- ${pkm0} Item PKM Toko 0 (Nol)`);lines.push('')}const over=cnt('Over Stock di Toko');if(over>0){lines.push('Over Stock di Toko',`- ${over} Item Over Stock di Toko`,'')}while(lines.length&&lines[lines.length-1]==='')lines.pop();$('bb_kesimpulan').textContent=lines.join('\n')}
function bb_exportXlsx(){const wb=XLSX.utils.book_new();const data=bbState.rows.map(r=>{const o={};bb_resultCols.forEach(c=>{const v=r[bb_colKey[c]];o[c]=bb_dateCols.has(c)?fmtDate(v):v});return o});XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),'REKAP');const d=bb_ket2Names.map(n=>({KET2:n,ITEM:bbState.rows.filter(r=>r.KET2===n).length}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d),'Rincian');XLSX.writeFile(wb,`Monitoring Stock Item Barang Baru dan Open TAG ${bbState.tgl}.xlsx`)}
function bb_exportJpg(){const rows=bbState.rows;if(!rows.length)return;const FF='Calibri, Carlito, "Segoe UI", Roboto, Arial, sans-serif',FONT='13px '+FF,BOLD='bold 13px '+FF,TFONT='22px '+FF,RH=19,PAD=5,TH=38,S=1.5,title=`Monitoring Stock Item Barang Baru dan Open TAG ${bbState.tgl}`;const text=rows.map(r=>bb_resultCols.map(c=>{const v=r[bb_colKey[c]];const s=bb_dateCols.has(c)?fmtDate(v):bb_fmt(c,v);return String(s??'')}));const m=document.createElement('canvas').getContext('2d');const widths=bb_resultCols.map((c,j)=>{m.font=BOLD;let w=m.measureText(c).width;m.font=FONT;for(const t of text)w=Math.max(w,m.measureText(t[j]).width);w=Math.ceil(w)+PAD*2;return c==='DESC2'?Math.min(w,360):w});const W=widths.reduce((a,b)=>a+b,0)+3;const per=Math.max(20,Math.floor((Math.floor(14e6/(W*S*S))-TH-RH-6)/RH)),pages=Math.ceil(rows.length/per);for(let pg=0;pg<pages;pg++){const a=pg*per,b=Math.min(rows.length,a+per),H=TH+RH*(b-a+1)+3,cv=document.createElement('canvas');cv.width=Math.ceil(W*S);cv.height=Math.ceil(H*S);const g=cv.getContext('2d');g.scale(S,S);g.fillStyle='#fff';g.fillRect(0,0,W,H);g.textBaseline='middle';g.strokeStyle='#000';g.lineWidth=1;g.fillStyle='#000';g.font=TFONT;g.textAlign='left';g.fillText(title,2,TH/2);const cell=(x,y,w,t,o={})=>{g.strokeRect(x+1.5,y+.5,w,RH);g.fillStyle='#000';g.font=o.bold?BOLD:FONT;g.save();g.beginPath();g.rect(x+2,y+1,w-2,RH-1);g.clip();if(o.al==='r'){g.textAlign='right';g.fillText(t,x+w-PAD+1,y+RH/2+1)}else if(o.al==='c'){g.textAlign='center';g.fillText(t,x+1+w/2,y+RH/2+1)}else{g.textAlign='left';g.fillText(t,x+PAD,y+RH/2+1)}g.restore()};let y=TH,x=0;bb_resultCols.forEach((c,j)=>{cell(x,y,widths[j],c,{bold:true,al:'c'});x+=widths[j]});for(let i=a;i<b;i++){y+=RH;x=0;bb_resultCols.forEach((c,j)=>{const t=text[i][j],right=bb_numCols.has(c);cell(x,y,widths[j],t,{al:right?'r':'l'});x+=widths[j]})}const name=`Monitoring Stock Item Barang Baru dan Open TAG ${bbState.tgl}${pages>1?` (${pg+1})`:''}.jpg`;setTimeout(()=>cv.toBlob(bl=>{const u=URL.createObjectURL(bl),el=document.createElement('a');el.href=u;el.download=name;document.body.appendChild(el);el.click();el.remove();setTimeout(()=>URL.revokeObjectURL(u),5000)},'image/jpeg',.92),pg*500)}}
async function bb_identifyBahanFile(f){try{const rows=await bb_readAny(f);let kind=bb_detectKind(rows);if(!kind&&bb_parseMaster(rows))kind='MASTER';return kind}catch(e){return null}}

/* ============================================================
   BAGIAN "ITB" — logic app "Check Item Tidak Bergerak" (dipertahankan)
   ============================================================ */
const itb_resultCols=['PRDCD','DESC2','ZONA','ALAMAT','DIV','FRAC','ACOST','MINOR','PTAG','QTY','CTN','RUPIAH','SPD','DSI','KEB','KET ZONA','JTD','Over','PKM 0','Hari','KETERANGAN','KET'];
const itb_categoryNames=['Non Dry','Item WHK','Over Stock di toko','PKM Toko 0 (Nol)','Stock kurang dari Minor','Tidak ada JTD Toko','Tidak ada Kebutuhan (PBRO)','Check Fisik WH'];
const itbState={master:null,stock:null,regIndex:null,kandidat:[],rows:[],filtered:[],tgl:''};

function itb_toWb(ab,file){const u=new Uint8Array(ab);if(isSpreadsheet(file.name)||(u[0]===80&&u[1]===75))return XLSX.read(ab,{type:'array',cellDates:true,raw:true});let text=new TextDecoder('utf-8').decode(ab);if(text.charCodeAt(0)===0xFEFF)text=text.slice(1);const delim=detectDelimiter(text);return XLSX.read(text,{type:'string',raw:true,cellDates:true,FS:delim,codepage:65001})}
async function itb_readAny(file){return itb_toWb(await file.arrayBuffer(),file)}
function itb_sheetRows(wb,name){const ws=wb.Sheets[name]||wb.Sheets[wb.SheetNames[0]];return ws?XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:false}):[]}
function itb_findHeader(rows,required,maxScan=15){const req=required.map(x=>x.toUpperCase());for(let i=0;i<Math.min(rows.length,maxScan);i++){const h=headRow(rows,i);if(req.every(k=>h.includes(k)))return {row:i,headers:h}}return {row:0,headers:headRow(rows,0)}}
function itb_parseMaster(rows){let h3=-1,h4=-1;for(let i=0;i<10;i++){const r=headRow(rows,i);if(r.includes('PRDCD')&&r.includes('DESC2'))h3=i;if(r.includes('ST')&&r.includes('ZONA'))h4=i}if(h3<0)throw Error('Master: header PRDCD/DESC2 tidak ditemukan.');const a=hmap(headRow(rows,h3)),b=h4>=0?hmap(headRow(rows,h4)):{};const map={...a};for(const k of ['ST','ZONA','LT','RAK','BAR','CELL'])if(b[k]!==undefined)map[k]=b[k];const data=[];if(map.LT===undefined)throw Error('Master: kolom LT tidak ditemukan.');const start=Math.max(h3,h4)+2;for(let i=start;i<rows.length;i++){const r=rows[i]||[],id=norm(r[map.PRDCD]);if(!id||norm(r[map.LT])==='')continue;data.push({id,row:r,map})}return data}
function itb_parseStock(rows){const IDK=['PLUKODE','PLU'],QK=['QTY_AKHIR','QTY_LPP'];let row=-1,headers=headRow(rows,0);for(let i=0;i<Math.min(rows.length,10);i++){const h=headRow(rows,i);if(IDK.some(k=>h.includes(k))&&QK.some(k=>h.includes(k))){row=i;headers=h;break}}const idCol=IDK.map(k=>headers.indexOf(k)).find(x=>x>=0),qCol=QK.map(k=>headers.indexOf(k)).find(x=>x>=0);if(row<0||idCol===undefined||qCol===undefined)throw Error(`Stock Akhir: kolom kode item / qty tidak ditemukan. Header terbaca: ${headers.filter(Boolean).slice(0,12).join(', ')}`);const out=new Map();for(let i=row+1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[idCol]);if(id)out.set(id,num(r[qCol]))}return out}
function itb_parseRegRows(rows){const out=[];const starts=[];for(let c=0;c<rows[0]?.length;c++)if(norm(rows[0][c]).toUpperCase()==='PRDCD')starts.push(c);if(starts.length>1){for(const start of starts){const trnh=start+97;if(trnh>=rows[0].length||norm(rows[0][trnh]).toUpperCase()!=='TRNH')continue;out.push({rows,start,trnh})}return out}const h=itb_findHeader(rows,['PRDCD','TRNH'],12),m=hmap(h.headers);if(m.PRDCD===undefined||m.TRNH===undefined){const found=(h.headers||[]).filter(Boolean).slice(0,12).join(', ');throw Error('REG/PRG: PRDCD dan TRNH tidak ditemukan. Format CSV tidak terbaca. Header: '+found)}return [{rows,start:m.PRDCD,trnh:m.TRNH,map:m,row0:h.row}]}
function itb_monthKey(v){if(v instanceof Date&&!isNaN(v))return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}`;let s=norm(v);if(/^\d{4}$/.test(s)){return `20${s.slice(0,2)}-${s.slice(2)}`}let m=s.match(/(20\d{2})[-\/]?(\d{1,2})/);if(m)return `${m[1]}-${String(m[2]).padStart(2,'0')}`;m=s.match(/(\d{1,2})[-\/]?(20\d{2})/);if(m)return `${m[2]}-${String(m[1]).padStart(2,'0')}`;return s}
function itb_buildRegIndex(regBlocks){const idx=new Map();for(const b of regBlocks){const rows=b.rows;const start=b.start;const trnh=b.trnh;const first=Math.max(1,b.row0?b.row0+1:1);for(let r=first;r<rows.length;r++){const id=norm(rows[r]?.[start]);const mo=itb_monthKey(rows[r]?.[trnh]);if(!id||!mo)continue;if(!idx.has(id))idx.set(id,new Map());idx.get(id).set(mo,{row:rows[r],start})}}return idx}
function itb_stockVal(rec,d){if(!rec)return 0;const c=rec.start+3+(d.getDate()-1)*3;return num(rec.row[c])}
function itb_dateSeries(){const now=new Date();const end=new Date(now.getFullYear(),now.getMonth(),now.getDate()-1);const a=[];for(let i=99;i>=0;i--){const d=new Date(end);d.setDate(end.getDate()-i);a.push(d)}return a}
function itb_daily(id,idx,dates){const mm=idx.get(id);return dates.map(d=>itb_stockVal(mm?.get(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`),d))}
function itb_maxRun(v){const n=v.length,out=Array(n).fill(0);out[n-1]=v[n-1]!==0?1:0;for(let i=n-2;i>=0;i--)out[i]=(out[i+1]===0||v[i]!==v[i+1])?0:out[i+1]+1;return Math.max(...out)}
function itb_ssStats(rows){rows=unpipe(rows);const h=itb_findHeader(rows,['PRDCD'],10),m=hmap(h.headers),s=new Map();if(m.PRDCD===undefined)throw Error('SS-FT: header PRDCD tidak ditemukan. Header terbaca: '+h.headers.filter(Boolean).slice(0,12).join(', '));for(let i=h.row+1;i<rows.length;i++){const r=rows[i]||[],id=norm(r[m.PRDCD]);if(!id)continue;if(!s.has(id))s.set(id,{jtd:0,over:0,pkm0:0});const x=s.get(id);x.jtd++;let ket=m.KET!==undefined?norm(r[m.KET]):'';if(!ket){const st=num(r[m.STOCK]),pk=num(r[m.PKM]),mi=num(r[m.MINOR]);ket=pk===0?'PKM 0':(st+mi>pk?'Over':'Kurang')}if(ket==='Over')x.over++;if(ket==='PKM 0')x.pkm0++}return s}
function itb_calc(item,stockRec,regIndex,ss,dates){const r=item.row,m=item.map,id=item.id,get=k=>r[m[k]];const whk=norm(r[76]).toUpperCase()==='Y';const desc=norm(get('DESC2')),zona=get('ZONA'),div=norm(get('DIV')),frac=num(get('FRAC')),acost=num(get('ACOST')),minor=num(get('MINOR')),ptag=norm(get('PTAG')),avg=num(get('AVG')),dsi=num(get('DSISTK')),keb=num(get('KEB'));const vals=itb_daily(id,regIndex,dates),qty=vals[99]||0,stockFinal=stockRec===undefined?'':num(stockRec);const hari=itb_maxRun(vals);let keterangan='';if(!['N','R','F'].includes(ptag)&&qty!==0&&hari>=5&&qty!==stockFinal){} else if(!['N','R','F'].includes(ptag)&&qty!==0&&hari>=5&&qty===stockFinal){if(hari>=90)keterangan='Tdk Bergerak >3 Bulan';else if(hari>=60)keterangan='Tdk Bergerak >2 Bulan';else if(hari>30)keterangan='Tdk Bergerak >1 Bulan';else keterangan=`Tdk Bergerak ${hari} Hari`}
if(!keterangan)return null;const st=ss.get(id)||{jtd:0,over:0,pkm0:0};let ket;if((num(zona)===0)&&div==='4')ket='Non Dry';else if(whk)ket='Item WHK';else if(st.jtd===0)ket='Tidak ada JTD Toko';else if(st.pkm0>st.jtd*.5)ket='PKM Toko 0 (Nol)';else if(st.over>st.jtd*.5)ket='Over Stock di toko';else if(qty<minor)ket='Stock kurang dari Minor';else if(num(zona)!==0&&keb<10)ket='Tidak ada Kebutuhan (PBRO)';else ket='Check Fisik WH';const ctn=frac?qty/frac:0;return {PRDCD:id,DESC2:desc,ZONA:zona,ALAMAT:norm(get('ALAMAT')),DIV:div,FRAC:frac,ACOST:acost,MINOR:minor,PTAG:ptag,QTY:qty,CTN:ctn,RUPIAH:acost*qty,SPD:avg,DSI:dsi,KEB:keb,'KET ZONA':num(zona)===0?0:(['14','15'].includes(norm(zona))?'Bulky':'Bulky Fraction'),JTD:st.jtd,Over:st.over,'PKM 0':st.pkm0,Hari:hari,KETERANGAN:keterangan,KET:ket}}
function itb_sheetType(name){const n=String(name).toLowerCase().replace(/[^a-z0-9]/g,'');if(n.startsWith('master'))return 'master';if(n.startsWith('stock'))return 'stock';if(/^(reg|prg)/.test(n))return 'reg';return null}
function itb_contentType(rows){const scan=rows.slice(0,12).map(r=>(r||[]).map(x=>norm(x).toUpperCase()));const has=(...k)=>scan.some(r=>k.every(x=>r.includes(x)));if(has('PLUKODE','QTY_AKHIR')||has('PLU','QTY_LPP'))return 'stock';if(has('PRDCD','TRNH'))return 'reg';if(has('PRDCD','DESC2'))return 'master';return null}
function itb_sortRows(a){const kr=Object.fromEntries(itb_categoryNames.map((n,i)=>[n,i])),ag=k=>String(k).includes('>3 Bulan')?0:String(k).includes('>2 Bulan')?1:String(k).includes('>1 Bulan')?2:3;return a.sort((x,y)=>(kr[x.KET]??99)-(kr[y.KET]??99)||ag(x.KETERANGAN)-ag(y.KETERANGAN)||y.Hari-x.Hari||String(x.PRDCD).localeCompare(String(y.PRDCD),'en',{numeric:true}))}

// Catatan penggabungan: file PRODMAST (milik BB) memiliki kolom PRDCD & DESC2 sehingga bisa salah
// terdeteksi sebagai "Master" oleh pengenal berbasis konten milik ITB. Penjaga di bawah ini
// mencegah itu: jika file yang sama sudah dikenali BB sebagai PRODMAST/POMGG/HISTAG, file itu
// tidak akan dianggap Master/Stock/REG oleh ITB.
async function itb_parseBahanOnly(files){
  if(!files.length)throw Error('BAHAN wajib diupload (Master, REG 4 bulan, dan Stock Akhir).');
  const found={master:[],stock:[],reg:[]},seen=[];
  for(let i=0;i<files.length;i++){
    const f=files[i],wb=await itb_readAny(f);
    seen.push(`${f.name} [${wb.SheetNames.join(', ')}]`);
    for(const n of wb.SheetNames){
      let t=itb_sheetType(n),rows=null;
      if(!t&&wb.SheetNames.length===1){
        rows=itb_sheetRows(wb,n);
        t=itb_contentType(rows);
        if(t==='master'){const bbk=bb_detectKind(rows);if(bbk==='PRODMAST'||bbk==='POMGG'||bbk==='HISTAG')t=null}
      }
      if(!t)continue;
      found[t].push({src:f.name+(wb.SheetNames.length>1?' / '+n:''),rows:rows||itb_sheetRows(wb,n)});
    }
  }
  const one=(k,label)=>{const a=found[k];if(!a.length)throw Error(`BAHAN: ${label} tidak ditemukan. File terbaca: ${seen.join('; ')}`);if(a.length>1)throw Error(`BAHAN: ${label} terdeteksi lebih dari satu (${a.map(x=>x.src).join(', ')}).`);return a[0].rows};
  const master=itb_parseMaster(one('master','Master'));
  const stock=itb_parseStock(one('stock','Stock Akhir'));
  let blocks=[];for(const r of found.reg)blocks.push(...itb_parseRegRows(r.rows));
  if(blocks.length<4)throw Error(`REG/PRG terdeteksi ${blocks.length} blok. Diperlukan 4 bulan.`);
  return {master,stock,regIndex:itb_buildRegIndex(blocks)};
}
function itb_computeResult(bahan,ss){const dates=itb_dateSeries();const arr=[];for(let i=0;i<bahan.master.length;i++){const x=itb_calc(bahan.master[i],bahan.stock.get(bahan.master[i].id),bahan.regIndex,ss,dates);if(x)arr.push(x)}itb_sortRows(arr);return arr}
function itb_computeCandidates(bahan){const arr=itb_computeResult(bahan,new Map());return [...new Set(arr.filter(r=>r.KETERANGAN).map(r=>r.PRDCD))]}

function itb_fmt(c,v){if(v===''||v==null)return '';if(['ACOST','RUPIAH','SPD','DSI'].includes(c))return num(v).toLocaleString('id-ID',{maximumFractionDigits:2});if(['QTY','CTN','FRAC','MINOR','JTD','Over','PKM 0','Hari'].includes(c))return num(v).toLocaleString('id-ID',{maximumFractionDigits:4});return v}
function itb_render(arr){itbState.filtered=arr;const th=$('itb_resultTable').querySelector('thead'),tb=$('itb_resultTable').querySelector('tbody');th.innerHTML='<tr>'+itb_resultCols.map(c=>`<th>${esc(c)}</th>`).join('')+'</tr>';tb.innerHTML=arr.map(r=>'<tr>'+itb_resultCols.map(c=>`<td class="${typeof r[c]==='number'?'num':''}">${esc(itb_fmt(c,r[c]))}</td>`).join('')+'</tr>').join('');$('itb_summary').textContent=`${arr.length.toLocaleString('id-ID')} item • RUPIAH ${arr.reduce((s,r)=>s+num(r.RUPIAH),0).toLocaleString('id-ID')}`}
function itb_renderKesimpulan(){const rows=itbState.rows;if(!rows.length){$('itb_kesimpulan').textContent='Tidak ada item.';return}const nf=v=>Math.round(v).toLocaleString('en-US');const list=itb_categoryNames.map(n=>{const a=rows.filter(r=>r.KET===n);return {n,c:a.length,rp:Math.round(a.reduce((t,r)=>t+num(r.RUPIAH),0))}}).filter(x=>x.c>0).sort((a,b)=>b.c-a.c);const tgl=new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});itbState.tgl=tgl;const totC=list.reduce((t,x)=>t+x.c,0),totRp=list.reduce((t,x)=>t+x.rp,0);$('itb_kesimpulan').textContent=[`*Checking Item tidak bergerak ${tgl}*`,'','Item  KET | RP',...list.map(x=>`${x.c} Item ${x.n} | Rp ${nf(x.rp)} ,-`),'=================================',`Total | ${totC} Item | Rp ${nf(totRp)} ,-`].join('\n')}
function itb_exportXlsx(){const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(itbState.rows),'Hasil');const d=itb_categoryNames.map(n=>{const a=itbState.rows.filter(r=>r.KET===n);return {KET:n,ITEM:a.length,RUPIAH:a.reduce((s,r)=>s+num(r.RUPIAH),0)}});XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d),'Rincian');XLSX.writeFile(wb,`Checking Item tidak bergerak ${itbState.tgl}.xlsx`)}
const itb_JPG_COLS=itb_resultCols;
const itb_JPG_NUM=new Set(['FRAC','ACOST','MINOR','QTY','CTN','RUPIAH','SPD','DSI','KEB','JTD','Over','PKM 0','Hari']);
function itb_jpgText(c,v){if(itb_JPG_NUM.has(c)){const n=num(v);return n===0?'-':Math.round(n).toLocaleString('en-US')}if(c==='ZONA')return(v===0||norm(v)===''||norm(v)==='0')?'':norm(v);if(c==='KET ZONA')return v===0?'-':norm(v);return norm(v)}
function itb_pctl(a,p){const k=(a.length-1)*p,f=Math.floor(k),c=Math.ceil(k);return a[f]+(a[c]-a[f])*(k-f)}
function itb_scaleColor(v,lo,mid,hi){const A=[99,190,123],B=[255,235,132],C=[248,105,107];let f,g,t;if(v<=mid){f=A;g=B;t=mid===lo?1:(v-lo)/(mid-lo)}else{f=B;g=C;t=hi===mid?0:(v-mid)/(hi-mid)}return `rgb(${f.map((x,i)=>Math.round(x+(g[i]-x)*t)).join(',')})`}
function itb_exportJpg(){const rows=itbState.rows;if(!rows.length)return;
const FF='Calibri, Carlito, "Segoe UI", Roboto, Arial, sans-serif',FONT='13px '+FF,BOLD='bold 13px '+FF,TFONT='22px '+FF,RH=19,PAD=5,TH=38,S=1.5,title=`Checking Item tidak bergerak ${itbState.tgl}`;
const text=rows.map(r=>itb_JPG_COLS.map(c=>itb_jpgText(c,r[c]))),m=document.createElement('canvas').getContext('2d');
const widths=itb_JPG_COLS.map((c,j)=>{m.font=BOLD;let w=m.measureText(c).width;m.font=FONT;for(const t of text)w=Math.max(w,m.measureText(t[j]).width);w=Math.ceil(w)+PAD*2;return c==='DESC2'?Math.min(w,360):w});
const W=widths.reduce((a,b)=>a+b,0)+3,rp=rows.map(r=>num(r.RUPIAH)).sort((a,b)=>a-b),lo=rp[0],hi=rp[rp.length-1],mid=itb_pctl(rp,.5);
const per=Math.max(20,Math.floor((Math.floor(14e6/(W*S*S))-TH-RH-6)/RH)),pages=Math.ceil(rows.length/per);
for(let pg=0;pg<pages;pg++){const a=pg*per,b=Math.min(rows.length,a+per),H=TH+RH*(b-a+1)+3,cv=document.createElement('canvas');cv.width=Math.ceil(W*S);cv.height=Math.ceil(H*S);const g=cv.getContext('2d');g.scale(S,S);g.fillStyle='#fff';g.fillRect(0,0,W,H);g.textBaseline='middle';g.strokeStyle='#000';g.lineWidth=1;
g.fillStyle='#000';g.font=TFONT;g.textAlign='left';g.fillText(title,2,TH/2);
const cell=(x,y,w,t,o={})=>{if(o.bg){g.fillStyle=o.bg;g.fillRect(x+1,y+1,w-1,RH-1)}g.strokeRect(x+1.5,y+.5,w,RH);g.fillStyle='#000';g.font=o.bold?BOLD:FONT;g.save();g.beginPath();g.rect(x+2,y+1,w-2,RH-1);g.clip();if(o.al==='r'){g.textAlign='right';g.fillText(t,x+w-PAD+1,y+RH/2+1)}else if(o.al==='c'){g.textAlign='center';g.fillText(t,x+1+w/2,y+RH/2+1)}else{g.textAlign='left';g.fillText(t,x+PAD,y+RH/2+1)}g.restore()};
let y=TH,x=0;itb_JPG_COLS.forEach((c,j)=>{cell(x,y,widths[j],c,{bold:true,al:'c'});x+=widths[j]});
for(let i=a;i<b;i++){y+=RH;x=0;itb_JPG_COLS.forEach((c,j)=>{const t=text[i][j],right=itb_JPG_NUM.has(c)||t==='-';cell(x,y,widths[j],t,{al:right?'r':'l',bg:c==='RUPIAH'?itb_scaleColor(num(rows[i].RUPIAH),lo,mid,hi):null});x+=widths[j]})}
const name=`Checking Item tidak bergerak ${itbState.tgl}${pages>1?` (${pg+1})`:''}.jpg`;
setTimeout(()=>cv.toBlob(bl=>{const u=URL.createObjectURL(bl),el=document.createElement('a');el.href=u;el.download=name;document.body.appendChild(el);el.click();el.remove();setTimeout(()=>URL.revokeObjectURL(u),5000)},'image/jpeg',.92),pg*500)}}
async function itb_identifyKindsInFile(f){try{const wb=await itb_readAny(f);const kinds=new Set();for(const n of wb.SheetNames){let t=itb_sheetType(n);if(!t&&wb.SheetNames.length===1){const rows=itb_sheetRows(wb,n);t=itb_contentType(rows);if(t==='master'){const bbk=bb_detectKind(rows);if(bbk==='PRODMAST'||bbk==='POMGG'||bbk==='HISTAG')t=null}}if(t)kinds.add(t)}return kinds}catch(e){return new Set()}}

/* ============================================================
   ORKESTRASI GABUNGAN — 1x upload BAHAN, 1x upload SS-FT,
   diproses sesuai aturan masing-masing aplikasi
   ============================================================ */
function setStatus(t){const k={'Siap':'ok','Selesai':'ok','Memproses':'busy','Periksa':'err'}[t]||'';$('status').textContent=t;$('status').className='badge '+k}

let bbOk=false, itbOk=false, combinedPrdcd=[];

async function renderBahanFileList(){
  const list=$('bahanFileList'),files=[...($('bahanFile').files||[])];
  list.innerHTML='';
  if(!files.length)return;
  const items=files.map(f=>{const li=document.createElement('li');li.innerHTML=`<b>${esc(f.name)}</b><span>Membaca…</span>`;list.appendChild(li);return li});
  await Promise.all(files.map(async(f,i)=>{
    const [bbKind,itbKinds]=await Promise.all([bb_identifyBahanFile(f),itb_identifyKindsInFile(f)]);
    const span=items[i].querySelector('span');
    const labels=new Set();
    if(bbKind)labels.add(bb_kindLabel[bbKind]);
    if(itbKinds.has('master'))labels.add('Master');
    if(itbKinds.has('stock'))labels.add('Stock Akhir');
    if(itbKinds.has('reg'))labels.add('REG/PRG');
    if(!labels.size){span.textContent='Tidak dikenali';span.className='kind-unknown'}
    else{span.textContent=[...labels].join(' + ');span.className='kind-ok'}
  }));
}
$('bahanFile').addEventListener('change',renderBahanFileList);

async function processBahanCombined(){
  const files=[...($('bahanFile').files||[])];
  if(!files.length){$('messageA').textContent='Upload file BAHAN dulu.';return}
  setStatus('Memproses');$('messageA').textContent='Membaca & memproses BAHAN untuk kedua aplikasi…';$('barA').style.width='10%';
  $('statusBB').textContent='';$('statusBB').className='statusline';
  $('statusITB').textContent='';$('statusITB').className='statusline';
  bbOk=false;itbOk=false;
  let bbCandidates=[],itbCandidates=[];

  try{
    const bahan=await bb_parseBahan(files);
    bbState.prodmast=bahan.prodmast;bbState.pomgg=bahan.pomgg;bbState.master=bahan.master;bbState.stock=bahan.stock;bbState.histag=bahan.histag;
    bbState.kandidat=bb_computeCandidates(bahan);
    bbCandidates=bbState.kandidat;
    bbOk=true;
    $('statusBB').textContent=`Barang Baru: OK — ${bbCandidates.length} item kandidat ditemukan.`;
    $('statusBB').className='statusline ok';
  }catch(e){
    console.error(e);
    $('statusBB').textContent='Barang Baru: '+(e.message||String(e));
    $('statusBB').className='statusline err';
  }
  $('barA').style.width='55%';await sleep();

  try{
    const bahan2=await itb_parseBahanOnly(files);
    itbState.master=bahan2.master;itbState.stock=bahan2.stock;itbState.regIndex=bahan2.regIndex;
    itbState.kandidat=itb_computeCandidates(bahan2);
    itbCandidates=itbState.kandidat;
    itbOk=true;
    $('statusITB').textContent=`Item Tidak Bergerak: OK — ${itbCandidates.length} item kandidat ditemukan.`;
    $('statusITB').className='statusline ok';
  }catch(e){
    console.error(e);
    $('statusITB').textContent='Item Tidak Bergerak: '+(e.message||String(e));
    $('statusITB').className='statusline err';
  }
  $('barA').style.width='100%';

  combinedPrdcd=[...new Set([...bbCandidates,...itbCandidates])];
  $('prdcdCount').textContent=combinedPrdcd.length?`${combinedPrdcd.length.toLocaleString('id-ID')} item PRDCD (gabungan Barang Baru + Item Tidak Bergerak)`:'Tidak ada item PRDCD.';
  $('downloadListBtn').disabled=!combinedPrdcd.length;

  $('step3').hidden=true;$('bb_panel').hidden=true;$('itb_panel').hidden=true;
  if(bbOk||itbOk){
    setStatus('Selesai');
    $('messageA').textContent='Selesai memproses BAHAN. Lanjut ke langkah 2.';
    $('step2').hidden=false;
  }else{
    setStatus('Periksa');
    $('messageA').textContent='BAHAN belum lengkap untuk kedua aplikasi. Periksa pesan di atas.';
  }
}

async function processSsftCombined(){
  if(!$('ssftFile').files[0]){$('messageB').textContent='Upload file SS-FT dulu, lalu tekan PROSES.';return}
  setStatus('Memproses');$('messageB').textContent='Membaca SS-FT & menghitung hasil…';$('barB').style.width='10%';
  $('statusBB2').textContent='';$('statusBB2').className='statusline';
  $('statusITB2').textContent='';$('statusITB2').className='statusline';
  const file=$('ssftFile').files[0];
  $('step3').hidden=false;

  if(bbOk){
    try{
      const rows=await bb_readAny(file);
      const stats=bb_parseSsft(rows);
      const arr=bb_buildRekap(bbState,stats);
      bbState.rows=arr;
      bb_render(arr);bb_renderKesimpulan();
      $('bb_exportBtn').disabled=!arr.length;$('bb_exportJpgBtn').disabled=!arr.length;
      $('bb_panel').hidden=false;
      $('statusBB2').textContent=`Barang Baru: OK — ${arr.length} item masuk Rekap.`;
      $('statusBB2').className='statusline ok';
    }catch(e){
      console.error(e);
      $('statusBB2').textContent='Barang Baru: '+(e.message||String(e));
      $('statusBB2').className='statusline err';
      $('bb_panel').hidden=true;
    }
  }else{
    $('statusBB2').textContent='Barang Baru: dilewati (BAHAN langkah 1 belum lengkap).';
  }
  $('barB').style.width='55%';await sleep();

  if(itbOk){
    try{
      const wb=await itb_readAny(file);
      const rows=itb_sheetRows(wb,'SS-FT');
      const ss=itb_ssStats(rows);
      const arr=itb_computeResult(itbState,ss);
      itbState.rows=arr;
      itb_render(arr);itb_renderKesimpulan();
      $('itb_exportBtn').disabled=false;$('itb_exportJpgBtn').disabled=false;
      $('itb_panel').hidden=false;
      $('statusITB2').textContent=`Item Tidak Bergerak: OK — ${arr.length} item masuk Hasil.`;
      $('statusITB2').className='statusline ok';
    }catch(e){
      console.error(e);
      $('statusITB2').textContent='Item Tidak Bergerak: '+(e.message||String(e));
      $('statusITB2').className='statusline err';
      $('itb_panel').hidden=true;
    }
  }else{
    $('statusITB2').textContent='Item Tidak Bergerak: dilewati (BAHAN langkah 1 belum lengkap).';
  }
  $('barB').style.width='100%';
  setStatus('Selesai');
  $('messageB').textContent='Selesai memproses SS-FT.';
}

$('downloadListBtn').onclick=()=>{const b=new Blob([combinedPrdcd.join('\n')+'\n'],{type:'text/plain'}),u=URL.createObjectURL(b),x=document.createElement('a');x.href=u;x.download='PRDCD_SS-FT.txt';x.click();URL.revokeObjectURL(u)};
$('processBtn').onclick=processBahanCombined;
$('processBtn2').onclick=processSsftCombined;
$('clearBtn').onclick=()=>location.reload();
$('bb_exportBtn').onclick=bb_exportXlsx;
$('bb_exportJpgBtn').onclick=bb_exportJpg;
$('itb_exportBtn').onclick=itb_exportXlsx;
$('itb_exportJpgBtn').onclick=itb_exportJpg;
$('panduanLink').onclick=()=>{const d=$('panduan');d.showModal?d.showModal():d.setAttribute('open','')};
$('panduanClose').onclick=()=>{const d=$('panduan');d.close?d.close():d.removeAttribute('open')};
$('panduan').addEventListener('click',e=>{if(e.target===$('panduan'))$('panduanClose').onclick()});
if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
