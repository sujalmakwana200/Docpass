const $=s=>document.querySelector(s),H=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const API_BASE=window.location.protocol==='file:'?'http://localhost:8000':'';
const MB=1048576,fmt=b=>b<1024?b+' B':b<MB?Math.round(b/1024)+' KB':(b/MB).toFixed(1)+' MB';
const G=['Convert to PDF','Convert from PDF','Organize','Optimize','Edit and protect'];
const T=[
{id:'word',g:0,n:'Word to PDF',a:'.doc,.docx,.odt,.rtf',e:'FastAPI + LibreOffice',b:1,fmt:'pdf'},
{id:'xls',g:0,n:'Excel to PDF',a:'.xls,.xlsx,.ods',e:'FastAPI + LibreOffice',b:1,fmt:'pdf'},
{id:'ppt',g:0,n:'PowerPoint to PDF',a:'.ppt,.pptx,.odp',e:'FastAPI + LibreOffice',b:1,fmt:'pdf'},
{id:'img',g:0,n:'Image to PDF',a:'image/*',e:'your browser (canvas + pdf-lib)',r:1},
{id:'html',g:0,n:'Text and HTML to PDF',a:'.txt,.md,.html,.htm',e:'FastAPI + ReportLab',b:1,fmt:'pdf'},
{id:'csv',g:0,n:'CSV to PDF',a:'.csv',e:'FastAPI + ReportLab',b:1,fmt:'pdf'},
{id:'p2i',g:1,n:'PDF to JPG or PNG',e:'FastAPI + Poppler',b:1,s:['JPG','PNG'],fmts:['jpg','png']},
{id:'p2t',g:1,n:'PDF to text',e:'PyMuPDF (Phase 1)'},
{id:'p2w',g:1,n:'PDF to Word',e:'pdf2docx (Phase 1)'},
{id:'p2x',g:1,n:'PDF to Excel',e:'Camelot + pandas (Phase 1)'},
{id:'merge',g:2,n:'Merge PDFs',e:'your browser (pdf-lib)',r:1},
{id:'split',g:2,n:'Split PDF',e:'your browser (pdf-lib)',r:1,p:'Pages per file',d:'1'},
{id:'ext',g:2,n:'Extract or delete pages',e:'your browser (pdf-lib)',r:1,p:'Pages, for example 1-3,7',d:'1',s:['Keep these pages','Delete these pages']},
{id:'rot',g:2,n:'Rotate pages',e:'your browser (pdf-lib)',r:1,p:'Pages (leave blank for all)',d:'',s:['90° clockwise','180°','90° anticlockwise']},
{id:'reo',g:2,n:'Reorder pages',e:'pypdf + PyMuPDF (Phase 1)'},
{id:'cmp',g:3,n:'Compress PDF',e:'FastAPI + pypdf / Poppler',b:1,fmt:'pdf'},
{id:'rep',g:3,n:'Repair PDF',e:'qpdf (Phase 1)'},
{id:'wm',g:4,n:'Add a watermark',e:'your browser (pdf-lib)',r:1,p:'Watermark text',d:'CONFIDENTIAL'},
{id:'num',g:4,n:'Add page numbers',e:'your browser (pdf-lib)',r:1},
{id:'pw',g:4,n:'Protect with a password',e:'pypdf (Phase 1)',p:'New password',d:''},
{id:'unl',g:4,n:'Unlock a PDF',e:'pypdf or qpdf (Phase 1)',p:'Current password',d:''},
{id:'meta',g:4,n:'View or remove metadata',e:'pypdf (Phase 1)'}];
T.forEach(t=>t.a=t.a||'.pdf');

const PR=[
{id:'none',n:'No preset'},
{id:'govt_portal_200kb',n:'Govt / Admission upload (Under 200 KB)',lim:200*1024,max:40,nm:'letters, numbers, _ or -'},
{id:'visa_500kb',n:'Visa / Immigration (Under 500 KB)',lim:500*1024,max:40,nm:'letters, numbers, _ or -'},
{id:'tender_2mb',n:'Tender / GST portal (Under 2 MB)',lim:2*MB,max:40,nm:'letters, numbers, _ or -'}];

let cur=T[0],files=[],outs=[],used=0,pend='';
const day=()=>new Date().toISOString().slice(0,10);
try{const r=JSON.parse(localStorage.getItem('dp')||'{}');if(r.d==day())used=r.n||0}catch(e){}
const bump=()=>{used++;try{localStorage.setItem('dp',JSON.stringify({d:day(),n:used}))}catch(e){}};

function renderGrid(){
 $('#tg').innerHTML=G.map((g,i)=>`<div><h3>${g}</h3>${T.filter(t=>t.g==i).map(t=>{
  const badge=t.r?'<small class="lv">in browser</small>':t.b?'<small class="lv">live API</small>':'<small>phase 1</small>';
  return `<button class="tr${t.id===cur.id?' on':''}" data-t="${t.id}"><span>${t.n}</span>${badge}</button>`;
 }).join('')}</div>`).join('');
}

const rules=p=>p.lim?`Rules: up to ${fmt(p.lim)}, A4 pages, searchable text, file name with ${p.nm}.`:'No portal rules applied.';

function wb(keep){
 const t=cur,o={se:($('#se')||{}).value,pa:($('#pa')||{}).value,pr:($('#pr')||{}).value};
 const statusText=t.r?'Works in your browser':t.b?'Runs on FastAPI backend':'Phase 1 (coming soon)';
 const statusCls=t.r||t.b?'s-ok':'s-info';
 const desc=t.r?'Your file is processed in this tab and never uploaded.':t.b?'Processed by the FastAPI backend with LibreOffice/Poppler and auto-deleted after 60 minutes.':'Planned for Phase 1. Select an "in browser" or "live API" tool to run real files.';
 $('#wb').innerHTML=`<div class="wbh"><h3>${t.n}</h3><span class="tag ${statusCls}">${statusText}</span></div>
<p class="mut">Engine: ${t.e}. ${desc}</p>
<label class="drop" id="dz"><input type="file" id="fi" multiple accept="${t.a}" hidden><b>Drop files here</b><span>or click to browse · up to 3 files, 50 MB each (${H(t.a.replace('image/*','JPG, PNG, WebP'))})</span></label>
<ul class="fl">${files.map((f,i)=>`<li><span>${H(f.name)}</span><em>${fmt(f.size)}</em><button data-rm="${i}" aria-label="Remove ${H(f.name)}">Remove</button></li>`).join('')}</ul>
<div class="opts">${t.s?`<label>Option<select id="se">${t.s.map((x,i)=>`<option value="${i}">${x}</option>`).join('')}</select></label>`:''}${t.p!=null?`<label>${t.p}<input id="pa" value="${H(t.d)}"></label>`:''}<label>Portal preset<select id="pr">${PR.map((p,i)=>`<option value="${i}">${p.n}</option>`).join('')}</select></label><p class="mut" id="rl" style="margin:0"></p></div>
<button class="btn alt go" id="go">Run</button><p id="msg" role="alert"></p><p class="mut" id="quotaTxt">Free runs used today: ${used} of 10</p>`;
 if(keep){if(o.se&&$('#se'))$('#se').value=o.se;if(o.pa!=null&&$('#pa'))$('#pa').value=o.pa;if(o.pr)$('#pr').value=o.pr}
 $('#rl').textContent=rules(PR[+$('#pr').value]);
 if(pend){err(pend);pend=''}
}

const err=m=>{const x=$('#msg');if(x)x.textContent=m};

function add(list){
 const ok=cur.a.split(',');
 for(const f of list){
  const e='.'+f.name.split('.').pop().toLowerCase();
  if(!(ok.includes(e)||(ok[0]=='image/*'&&f.type.startsWith('image/')))){pend=`“${f.name}” is not supported here. This tool accepts: ${cur.a.replace('image/*','JPG, PNG, WebP')}.`;continue}
  if(f.size>50*MB){pend=`“${f.name}” is over 50 MB, the free limit. Pro raises it.`;continue}
  if(files.length>=3){pend='Free runs take 3 files at a time. Pro handles 50+ files in one batch.';break}
  files.push(f)}
 wb(1)}

function rng(s,n){const r=new Set();for(const p of s.split(',')){const m=p.trim().match(/^(\d+)(?:-(\d+))?$/);if(!m)throw new Error('Write pages like 1-3,7.');const a=+m[1],b=+(m[2]||m[1]);if(a<1||b<a||b>n)throw new Error(`Pages must be between 1 and ${n}.`);for(let i=a;i<=b;i++)r.add(i-1)}return[...r].sort((x,y)=>x-y)}

async function real(t,fs,pr){
 if(!window.PDFLib)throw new Error('The PDF engine did not load. Check your connection and reload.');
 const L=PDFLib,v=s=>($(s)||{}).value||'',par=v('#pa').trim(),sel=+v('#se')||0,out=[];
 const ld=async f=>{try{return await L.PDFDocument.load(await f.arrayBuffer())}catch(e){throw new Error(`“${f.name}” is corrupt or password-protected. Unlock it first, then try again.`)}};
 const cp=async(s,ix)=>{const d=await L.PDFDocument.create();(await d.copyPages(s,ix)).forEach(p=>d.addPage(p));return d};
 if(t.id=='merge'){
  if(fs.length<2)throw new Error('Add at least 2 PDFs to merge.');
  const d=await L.PDFDocument.create();
  for(const f of fs){const s=await ld(f);(await d.copyPages(s,s.getPageIndices())).forEach(p=>d.addPage(p))}
  out.push({name:'merged.pdf',d})
 }else if(t.id=='img'){
  const d=await L.PDFDocument.create();
  for(const f of fs){
   let bm;try{bm=await createImageBitmap(f)}catch(e){throw new Error(`“${f.name}” could not be read. Use JPG, PNG or WebP.`)}
   const k=Math.min(1,2400/Math.max(bm.width,bm.height)),c=document.createElement('canvas');
   c.width=Math.round(bm.width*k);c.height=Math.round(bm.height*k);
   const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(bm,0,0,c.width,c.height);
   const bl=await new Promise(r=>c.toBlob(r,'image/jpeg',.85)),im=await d.embedJpg(await bl.arrayBuffer());
   const pg=d.addPage([595.28,841.89]),sc=Math.min(547.28/im.width,793.89/im.height);
   pg.drawImage(im,{x:(595.28-im.width*sc)/2,y:(841.89-im.height*sc)/2,width:im.width*sc,height:im.height*sc})}
  out.push({name:'images.pdf',d,img:1})
 }else for(const f of fs){
  const s=await ld(f),n=s.getPageCount(),b=f.name.replace(/\.[^.]+$/,'');
  if(t.id=='split'){const k=Math.max(1,parseInt(par)||1);for(let i=0,j=1;i<n;i+=k,j++)out.push({name:`${b}_part${j}.pdf`,d:await cp(s,Array.from({length:Math.min(k,n-i)},(_,x)=>i+x))})}
  else if(t.id=='ext'){let ix=rng(par,n);if(sel)ix=s.getPageIndices().filter(i=>!ix.includes(i));if(!ix.length)throw new Error('That would remove every page.');out.push({name:`${b}_${sel?'trimmed':'extract'}.pdf`,d:await cp(s,ix)})}
  else if(t.id=='rot'){const dg=[90,180,270][sel],ix=par?rng(par,n):s.getPageIndices();ix.forEach(i=>{const p=s.getPage(i);p.setRotation(L.degrees((p.getRotation().angle+dg)%360))});out.push({name:`${b}_rotated.pdf`,d:s})}
  else{
   const ft=await s.embedFont(L.StandardFonts.HelveticaBold);
   s.getPages().forEach((p,i)=>{const{width:w,height:h}=p.getSize();
    if(t.id=='num'){const x=`${i+1} / ${n}`;p.drawText(x,{x:w/2-ft.widthOfTextAtSize(x,10)/2,y:20,size:10,font:ft})}
    else{const x=par||'CONFIDENTIAL',z=Math.min(w,h)/8,k=ft.widthOfTextAtSize(x,z)/2*.7071;p.drawText(x,{x:w/2-k,y:h/2-k,size:z,font:ft,rotate:L.degrees(45),opacity:.18,color:L.rgb(.4,.4,.4)})}});
   out.push({name:`${b}_${t.id=='wm'?'watermarked':'numbered'}.pdf`,d:s})}
 }
 for(const o of out){o.bytes=await o.d.save();if(pr.max)o.name=o.name.replace(/\.pdf$/,'').replace(/[^\w-]+/g,'_').slice(0,pr.max)+'.pdf'}
 return out}

async function runBackend(t,fs,pr){
 const sel=+($('#se')||{}).value||0;
 const tgt=t.fmts?t.fmts[sel]:(t.fmt||'pdf');
 const fd=new FormData();
 fs.forEach(f=>fd.append('files',f));
 fd.append('target_format',tgt);
 fd.append('preset',pr.id||'none');
 const res=await fetch(`${API_BASE}/api/convert`,{method:'POST',body:fd});
 const data=await res.json().catch(()=>({}));
 if(!res.ok){
  const msg=typeof data.detail==='string'?data.detail:(data.detail&&data.detail.error)||JSON.stringify(data.detail||data);
  throw new Error(msg||`Server returned ${res.status}`);
 }
 if(data.quota&&typeof data.quota.used==='number'){used=data.quota.used}
 return data.files.map(f=>({
  name:f.output_filename,
  downloadUrl:`${API_BASE}${f.download_url}`,
  backendReport:f.portal_report
 }));
}

async function check(o,pr){
 const c=[],L=pr.lim,sz=o.bytes.length;
 c.push({l:'Size',v:L?`${fmt(sz)} of ${fmt(L)} allowed`:fmt(sz),s:!L||sz<=L?1:0,h:'Target-size compression (Pro) shrinks it to the limit.'});
 let w=0,h=0;try{const d=await PDFLib.PDFDocument.load(o.bytes),p=d.getPage(0).getSize();w=p.width;h=p.height}catch(e){}
 const a4=(x,y)=>Math.abs(x-595.3)<8&&Math.abs(y-841.9)<8,ok=a4(w,h)||a4(h,w);
 c.push({l:'Page size',v:ok?'A4':`${Math.round(w)} × ${Math.round(h)} pt`,s:!L||ok?1:0,h:'This portal expects A4 pages.'});
 c.push(o.img?{l:'Searchable text',v:'Image-only pages',s:L?0:2,h:'Scans need OCR to searchable PDF (Pro).'}:{l:'Searchable text',v:'Checked on the server',s:2});
 const nm=!pr.max||/^[\w-]{1,}\.pdf$/.test(o.name)&&o.name.length-4<=pr.max;
 c.push({l:'File name',v:o.name,s:nm?1:0,h:'Use letters, numbers, _ or - only.'});
 return c}
const verdict=(c,pr)=>!pr.lim?['READY','s-info']:c.some(x=>!x.s)?['REJECTED','s-bad']:c.some(x=>x.s==2)?['LIKELY OK','s-warn']:['ACCEPTED','s-ok'];

async function show(t,pr){
 const R=$('#res');
 if(!outs.length||outs[0].sim){
  R.innerHTML=`<div class="rh"><h3>Phase 1 Tool</h3><span class="stamp s-info">DEMO</span></div><p class="mut">“${H(t.n)}” (${H(t.e)}) is queued for Phase 1. Try any tool marked <b>in browser</b> or <b>live API</b> (Word, Excel, PowerPoint, HTML, CSV, PDF to JPG/PNG, Compress PDF) to process real files right now.</p>`;return}
 let h='';
 for(let i=0;i<outs.length;i++){
  const o=outs[i];
  if(o.backendReport){
   const rep=o.backendReport;
   const [vt,vc]=!pr.lim?['READY','s-info']:rep.ready_for_portal?['ACCEPTED','s-ok']:['REJECTED','s-bad'];
   const checks=rep.checks.map(x=>`<li><span>${H(x.label)}</span><b>${H(x.detail)}</b><i class="${x.passed?'p1':'p0'}">${H(x.status)}</i></li>`).join('');
   h+=`<div class="ro"><div class="rh"><h4>${H(o.name)}</h4><span class="stamp ${vc}">${vt}</span></div><ul class="ck">${checks}</ul><a class="btn" href="${H(o.downloadUrl)}" download="${H(o.name)}">Save ${H(o.name)}</a></div>`;
  }else{
   const c=await check(o,pr),[vt,vc]=verdict(c,pr);
   h+=`<div class="ro"><div class="rh"><h4>${H(o.name)}</h4><span class="stamp ${vc}">${vt}</span></div><ul class="ck">${c.map(x=>`<li><span>${x.l}</span><b>${H(x.v)}</b><i class="p${x.s}">${['Fail','Pass','Server check'][x.s]}</i>${x.s==0?`<div class="ht">${x.h}</div>`:''}</li>`).join('')}</ul><button class="btn" data-dl="${i}">Save PDF</button></div>`;
  }
 }
 R.innerHTML=h;R.scrollIntoView({block:'nearest',behavior:'smooth'})}

async function go(){
 const t=cur,b=$('#go');
 if(!files.length)return err('Add a file first.');
 if(used>=10&&(t.r||t.b))return err('You have used all 10 free runs today. Pro removes the daily cap.');
 const pr=PR[+$('#pr').value||0];
 b.disabled=true;b.textContent='Working…';err('');
 try{
  if(t.r){outs=await real(t,files,pr);bump()}
  else if(t.b){outs=await runBackend(t,files,pr)}
  else{await new Promise(r=>setTimeout(r,350));outs=[{sim:1}]}
  await show(t,pr);wb(1)
 }catch(e){err(/WinAnsi/.test(e.message)?'The demo font handles Latin text only. Hindi and Gujarati will work with the server fonts.':e.message||'Something went wrong. Try again.');b.disabled=false;b.textContent='Run';return}
}

document.addEventListener('click',async e=>{
 const q=s=>e.target.closest(s);
 if(q('[data-t]')){cur=T.find(t=>t.id==q('[data-t]').dataset.t);files=[];outs=[];$('#res').innerHTML='<p class="empty">Your result and portal check will show up here.</p>';renderGrid();wb();$('#work').scrollIntoView({behavior:'smooth'})}
 else if(q('[data-rm]')){files.splice(+q('[data-rm]').dataset.rm,1);wb(1)}
 else if(q('#go'))go();
 else if(q('[data-dl]')){
  const o=outs[+q('[data-dl]').dataset.dl];
  const bl=new Blob([o.bytes],{type:'application/pdf'});
  if(window.claude&&claude.use){
   try{const d=await claude.use('downloads');if(!d)throw{code:'unavailable'};await d.save({filename:o.name,data:bl})}
   catch(x){if(x&&x.code!='declined')alert('Saving is not available in this view.')}
  }else{const a=document.createElement('a');a.href=URL.createObjectURL(bl);a.download=o.name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}}
 else if(q('#wj')){const v=$('#em').value.trim(),m=$('#wm');m.textContent=/^\S+@\S+\.\S+$/.test(v)?'Thanks. The waitlist opens once the backend is connected, so nothing was sent from this demo.':'Enter a valid email address.'}
});
document.addEventListener('change',e=>{
 if(e.target.id=='fi'){add([...e.target.files])}
 if(e.target.id=='pr')$('#rl').textContent=rules(PR[+e.target.value])});
['dragover','dragleave','drop'].forEach(ev=>document.addEventListener(ev,e=>{
 const z=e.target.closest&&e.target.closest('#dz');if(!z)return;e.preventDefault();
 z.classList.toggle('on',ev=='dragover');if(ev=='drop')add([...e.dataTransfer.files])}));

renderGrid();
wb();
fetch(`${API_BASE}/api/limits`).then(r=>r.ok?r.json():null).then(d=>{if(d&&typeof d.used_today==='number'){used=d.used_today;const q=$('#quotaTxt');if(q)q.textContent=`Free runs used today: ${used} of ${d.daily_limit}`}}).catch(()=>{});