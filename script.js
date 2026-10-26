'use strict';
/* ---------- Lecteurs : fichier -> blocs [{t:'h'|'p'|'li', text}] ---------- */
async function fromPdf(buf){
  pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf=await pdfjsLib.getDocument({data:buf}).promise, pages=[];
  for(let n=1;n<=pdf.numPages;n++){
    const items=(await (await pdf.getPage(n)).getTextContent()).items.filter(i=>i.str.trim()||i.hasEOL);
    const lines=[];
    for(const it of items){
      const y=it.transform[5], l=lines[lines.length-1];
      if(l&&Math.abs(l.y-y)<2.5){l.text+=it.str;l.h=Math.max(l.h,it.height)}
      else lines.push({y,h:it.height||10,text:it.str});
    }
    pages.push(lines.filter(l=>l.text.trim()));
  }
  const hs=pages.flat().map(l=>l.h).sort((a,b)=>a-b), med=hs[Math.floor(hs.length/2)]||10, out=[];
  for(const lines of pages){
    let prev=null;
    for(const l of lines){
      const text=l.text.replace(/\s+/g,' ').trim(), head=l.h>med*1.25;
      const last=out[out.length-1];
      if(prev&&!head&&last&&last.t==='p'&&prev.y-l.y<prev.h*1.9) last.text+=' '+text;
      else out.push({t:head?'h':/^[•▪·\-–]\s/.test(text)?'li':'p',text:text.replace(/^[•▪·\-–]\s/,'')});
      prev=l;
    }
  }
  return out;
}
async function fromDocx(buf){
  const {value}=await mammoth.convertToHtml({arrayBuffer:buf});
  const doc=new DOMParser().parseFromString(value,'text/html'), out=[];
  for(const el of doc.body.children){
    const tag=el.tagName.toLowerCase(), txt=el.textContent.trim();
    if(/^h\d$/.test(tag)) txt&&out.push({t:'h',text:txt});
    else if(tag==='ul'||tag==='ol') el.querySelectorAll('li').forEach(li=>out.push({t:'li',text:li.textContent.trim()}));
    else if(tag==='table') el.querySelectorAll('tr').forEach(tr=>out.push({t:'p',text:[...tr.children].map(c=>c.textContent.trim()).join(' | ')}));
    else if(txt) out.push({t:'p',text:txt});
  }
  return out;
}
async function fromOdt(buf){
  const zip=await JSZip.loadAsync(buf), f=zip.file('content.xml');
  if(!f) throw new Error('Ce fichier n’est pas un document .odt valide.');
  const xml=new DOMParser().parseFromString(await f.async('string'),'application/xml'), out=[];
  for(const el of xml.getElementsByTagName('*')){
    const n=el.localName;
    if(n!=='h'&&n!=='p') continue;
    const text=el.textContent.trim(); if(!text) continue;
    let inList=false; for(let p=el.parentNode;p;p=p.parentNode) if(p.localName==='list-item'){inList=true;break}
    out.push({t:n==='h'?'h':inList?'li':'p',text});
  }
  return out;
}
/* ---------- Écrivains : blocs -> Blob ---------- */
async function toDocx(blocks){
  const {Document,Packer,Paragraph,TextRun,HeadingLevel}=docx;
  const kids=blocks.map(b=>b.t==='h'?new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(b.text)]})
    :b.t==='li'?new Paragraph({bullet:{level:0},children:[new TextRun(b.text)]})
    :new Paragraph({spacing:{after:160},children:[new TextRun(b.text)]}));
  return Packer.toBlob(new Document({sections:[{children:kids}]}));
}
async function toPdf(blocks){
  const pdf=new jspdf.jsPDF({unit:'mm',format:'a4'}), M=20, W=210-2*M, H=297-M;
  let y=M;
  for(const b of blocks){
    const size=b.t==='h'?16:11, lh=size*0.5;
    pdf.setFont('helvetica',b.t==='h'?'bold':'normal'); pdf.setFontSize(size);
    const lines=pdf.splitTextToSize((b.t==='li'?'•  ':'')+b.text,W);
    if(b.t==='h') y+=2;
    for(const line of lines){ if(y>H){pdf.addPage();y=M} pdf.text(line,M,y); y+=lh }
    y+=b.t==='li'?1.5:4;
  }
  return pdf.output('blob');
}
async function toOdt(blocks){
  const esc=s=>s.replace(/[<>&"]/g,c=>({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c]));
  const body=blocks.map(b=>b.t==='h'?`<text:h text:outline-level="1">${esc(b.text)}</text:h>`
    :`<text:p>${b.t==='li'?'• ':''}${esc(b.text)}</text:p>`).join('');
  const ns='xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
  const zip=new JSZip();
  zip.file('mimetype','application/vnd.oasis.opendocument.text',{compression:'STORE'});
  zip.file('META-INF/manifest.xml',`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`);
  zip.file('content.xml',`<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ns} office:version="1.2"><office:body><office:text>${body}</office:text></office:body></office:document-content>`);
  return zip.generateAsync({type:'blob',mimeType:'application/vnd.oasis.opendocument.text'});
}
/* ---------- Interface ---------- */
const MODES={
  'pdf-docx':{title:'PDF vers Word',accept:'.pdf',ext:'docx',read:fromPdf,write:toDocx},
  'docx-pdf':{title:'Word vers PDF',accept:'.docx',ext:'pdf',read:fromDocx,write:toPdf},
  'odt-pdf':{title:'LibreOffice vers PDF',accept:'.odt',ext:'pdf',read:fromOdt,write:toPdf},
  'pdf-odt':{title:'PDF vers LibreOffice',accept:'.pdf',ext:'odt',read:fromPdf,write:toOdt}
};
const $=id=>document.getElementById(id);
let mode=null,file=null,url=null;
const setStatus=(m,c='')=>{$('status').textContent=m;$('status').className='status '+c};
function pick(f){
  if(!f) return;
  if(!f.name.toLowerCase().endsWith(MODES[mode].accept)) return setStatus(`Choisissez un fichier ${MODES[mode].accept}.`,'err');
  file=f; $('dropText').textContent=f.name; $('dropSub').textContent=(f.size/1024).toFixed(0)+' Ko — cliquez pour changer';
  $('go').disabled=false; $('dl').hidden=true; setStatus('');
}
document.querySelectorAll('.tool').forEach(btn=>btn.addEventListener('click',()=>{
  mode=btn.dataset.mode; const m=MODES[mode]; file=null;
  document.querySelectorAll('.tool').forEach(b=>b.classList.toggle('active',b===btn));
  $('panel').hidden=false; $('panelTitle').textContent=m.title; $('file').accept=m.accept; $('file').value='';
  $('dropText').textContent='Déposez votre fichier ici ou cliquez pour le choisir'; $('dropSub').textContent='Format accepté : '+m.accept;
  $('go').disabled=true; $('dl').hidden=true; setStatus('');
  $('panel').scrollIntoView({behavior:'smooth',block:'nearest'});
}));
$('file').addEventListener('change',e=>pick(e.target.files[0]));
$('drop').addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('file').click()}});
['dragover','dragenter'].forEach(ev=>$('drop').addEventListener(ev,e=>{e.preventDefault();$('drop').classList.add('over')}));
['dragleave','drop'].forEach(ev=>$('drop').addEventListener(ev,e=>{e.preventDefault();$('drop').classList.remove('over')}));
$('drop').addEventListener('drop',e=>pick(e.dataTransfer.files[0]));
$('go').addEventListener('click',async()=>{
  const m=MODES[mode]; $('go').disabled=true; $('dl').hidden=true; setStatus('Conversion en cours…');
  try{
    const blocks=await m.read(await file.arrayBuffer());
    if(!blocks.length) throw new Error('Aucun texte trouvé. Un PDF scanné (image) n’est pas pris en charge.');
    const blob=await m.write(blocks);
    if(url) URL.revokeObjectURL(url);
    url=URL.createObjectURL(blob);
    $('dl').href=url; $('dl').download=file.name.replace(/\.[^.]+$/,'')+'.'+m.ext; $('dl').hidden=false;
    setStatus('Conversion terminée.','ok');
  }catch(err){console.error(err);setStatus(err.message||'La conversion a échoué.','err')}
  $('go').disabled=false;
});

