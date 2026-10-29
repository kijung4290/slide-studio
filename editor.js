import {normalizePptObject,powerpointMarkup,powerpointCss,isImageData} from './powerpoint-model.js';
import {createPowerpointEditor} from './powerpoint-editor.js';
const $ = id => document.getElementById(id);
const uid = () => crypto.randomUUID();
const clone = value => Array.isArray(value)?value.map(clone):value&&typeof value==='object'?Object.fromEntries(Object.entries(value).map(([key,item])=>[key,clone(item)])):value;
const escapeHtml = text => String(text ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
const safeJson = value => JSON.stringify(value).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
const clamp = (v,min,max) => Math.max(min,Math.min(max,v));
const formatTime = seconds => { const value=Math.max(0,Math.floor(seconds));return `${String(Math.floor(value/60)).padStart(2,'0')}:${String(value%60).padStart(2,'0')}`; };
const freshDeck = () => ({schemaVersion:1,id:uid(),title:'새로운 발표',slides:[],targetMinutes:5,updatedAt:Date.now()});
const current = () => deck.slides[index];
let deck=freshDeck(), index=0, mode='edit', tab='notes', noteSize=20, selected=null, stageWindow=null, isBlack=false;
let db=null, saveTimer=0, savingPromise=Promise.resolve(), toastTimer=0, dirty=false, frameVersion=0, busy=false;
let history=[], historyIndex=-1, groupTime=0, groupKey='', dragIndex=-1;
let elapsedBase=0, timerStarted=0, timerRunning=false;
let desktopUpdateState=null;
let imageIntent='add';
const stageMode=new URLSearchParams(location.search).has('stage');
const pptEditor=createPowerpointEditor({getSlide:current,getMode:()=>mode,onChange:key=>{markChanged(key);const slide=current();if(slide){updateThumbnail(slide);$('slide-title').value=slide.title;}},onSelect:o=>{if(o)selectElement(null);renderPptInspector(o);},redraw:drawFrame,toast});

const nativeCss=`
html,body{margin:0;width:100%;height:100%;overflow:hidden;background:white;color:#20374d;font-family:"Malgun Gothic","Apple SD Gothic Neo",sans-serif}*{box-sizing:border-box}#canvas{position:absolute;left:0;top:0;transform-origin:0 0;background:#fff;overflow:hidden}.native-slide{position:absolute;inset:0;background:#fff}.native-slide .topline{position:absolute;left:100px;top:72px;color:#197b92;font-size:24px;letter-spacing:2px}.native-slide .title{position:absolute;left:100px;right:100px;top:150px;margin:0;font-size:68px;line-height:1.3;font-weight:700;letter-spacing:-2px;white-space:pre-wrap;word-break:keep-all}.native-slide .body{position:absolute;left:100px;right:100px;top:320px;bottom:120px;font-size:37px;line-height:1.7;white-space:pre-wrap;word-break:keep-all}.native-slide .foot{position:absolute;left:100px;bottom:55px;font-size:19px;color:#8ba0af}.native-slide .rule{position:absolute;left:100px;top:285px;width:60px;height:5px;background:#4fa9ba}.native-slide.cover{background:#edf7fa}.native-slide.cover .title{top:280px;font-size:89px;max-width:1250px}.native-slide.cover .body{top:565px;font-size:31px;color:#6d8898;bottom:150px}.native-slide.quote .title{top:280px;left:160px;right:160px;border-left:10px solid #3b9eb1;padding-left:55px;font-size:76px;line-height:1.5}.native-slide.quote .body{top:650px;left:225px;color:#7b94a4;font-size:26px}.native-slide .column{position:absolute;top:340px;bottom:160px;width:650px;padding:45px;border-radius:16px;background:#eef7fa;white-space:pre-wrap;line-height:1.7;font-size:33px}.native-slide .column.left{left:100px}.native-slide .column.right{right:100px}.native-slide .column b{font-size:42px;color:#197b92}.slide-image{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:white}.overlay{position:absolute;z-index:20;white-space:pre-wrap;line-height:1.45;min-width:40px;min-height:30px;overflow-wrap:break-word}.overlay img{width:100%;height:100%;display:block;object-fit:contain}.overlay-text{padding:12px;background:#ffffffed;color:#20374d;font-size:42px}.page-number{position:absolute;right:70px;bottom:45px;font-size:21px;color:#91a0ad;pointer-events:none}.static,.static *,.static *::before,.static *::after{animation:none!important;transition:none!important;filter:none!important}.stage,.stage *{font-synthesis:none}.stage{--bg:#fff;--ink:#17272e;--fr:"FS-Reg","Malgun Gothic",sans-serif;--fb:"FS-Bold","Malgun Gothic",sans-serif;position:absolute;inset:0;font-size:40px;line-height:1.3}.stage .slide{position:absolute;inset:0}.editor-on [data-editable]{cursor:text;outline:2px dashed transparent;outline-offset:6px}.editor-on [data-editable]:hover{outline-color:#197b9260}.editor-on [data-editable]:focus{outline:3px solid #2197b5;background-color:#1a8ea510}.editor-on .selected{outline:3px solid #2197b5!important;outline-offset:5px}.editor-on .overlay{outline:2px dashed #84b5c5;outline-offset:3px}.editor-on .overlay.selected{outline-color:#2197b5}.editor-on .overlay:before{content:'⋮⋮';position:absolute;left:-28px;top:-4px;width:24px;height:40px;color:#197b92;background:#e8f7fa;font-size:25px;line-height:40px;cursor:move}.stage.reference{font-family:var(--fr)}
`;
const referenceSelectors='.k-cv-top .m>span,.k-cv-hero .m>span,.k-cv-sub .m>span,.who,.evt,.chip,.k-ttl .m>span,.k-sub,.k-s,.k-h,.k-p,.k-term,.k-desc,.k-box>.b,.k-bp,.k-step>span:last-child,th,td,.k-pill,.k-bi,.k-name,.k-hub-top,.k-hub-it,.k-lit,.k-vh,.k-vp,.k-tl-h,.k-tl-p,.k-text .m>span,.k-note,.k-ask>span,.k-src';

function toast(message, duration=4000){$('toast').textContent=message;$('toast').hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>{$('toast').hidden=true;},duration);}
function setBusy(value,title='',detail=''){busy=value;$('busy-overlay').hidden=!value;if(title)$('busy-title').textContent=title;if(detail)$('busy-detail').textContent=detail;}
function showDialog(id){if(busy)return;const dialog=$(id);if(!dialog.open)dialog.showModal();}
function closeDialogs(){document.querySelectorAll('dialog[open]').forEach(d=>d.close());}
function cleanHtml(html){
  const template=document.createElement('template');template.innerHTML=html;
  template.content.querySelectorAll('script,iframe,object,embed,link,meta,base,form,foreignObject').forEach(el=>el.remove());
  template.content.querySelectorAll('*').forEach(el=>{for(const attr of [...el.attributes]){if(/^on/i.test(attr.name)||['srcdoc','contenteditable','autofocus'].includes(attr.name)||(/^(href|src|xlink:href)$/i.test(attr.name)&&!/^data:image\/(png|jpeg|webp|gif|svg\+xml);base64,/i.test(attr.value)))el.removeAttribute(attr.name);}el.classList.remove('selected');if(el.hasAttribute('data-editable'))el.removeAttribute('data-editor-id');});
  return template.innerHTML;
}
function validateDeck(input){
  const value=input.deck||input;
  if(value.schemaVersion!==1||!Array.isArray(value.slides)||value.slides.length>200)throw new Error('발표 작업실에서 저장한 프로젝트 JSON을 선택하세요. 최대 200장까지 지원합니다.');
  const result={schemaVersion:1,id:typeof value.id==='string'?value.id:uid(),title:String(value.title||'가져온 발표').slice(0,500),targetMinutes:clamp(Number(value.targetMinutes)||5,1,600),updatedAt:Date.now(),slides:[],sources:{}};
  if(value.sources&&typeof value.sources==='object'){for(const[key,source]of Object.entries(value.sources)){if(typeof source==='string'&&/^[a-z0-9+/=\r\n]+$/i.test(source)&&source.length<210*1024*1024)result.sources[key]=source;}}
  if(value.referenceCss){if(typeof value.referenceCss!=='string'||value.referenceCss.length>6e6)throw new Error('프로젝트의 스타일 데이터가 너무 큽니다.');result.referenceCss=value.referenceCss.replace(/<\/style/gi,'');}
  const ids=new Set();
  for(const [i,slide] of value.slides.entries()){
    if(!['image','native','reference','powerpoint'].includes(slide.kind))throw new Error(`${i+1}번 슬라이드의 형식을 확인할 수 없습니다.`);
    if(slide.kind==='image'&&!/^data:image\/(png|jpeg|webp);base64,/i.test(slide.image||''))throw new Error(`${i+1}번 슬라이드 이미지가 올바르지 않습니다.`);
    let id=typeof slide.id==='string'?slide.id:uid();if(ids.has(id))id=uid();ids.add(id);
    const overlays=(Array.isArray(slide.overlays)?slide.overlays:[]).slice(0,100).map(o=>({...o,id:uid(),type:o.type==='image'?'image':'text',x:clamp(Number(o.x)||0,-100,5000),y:clamp(Number(o.y)||0,-100,5000),width:clamp(Number(o.width)||500,30,5000),height:clamp(Number(o.height)||150,30,5000),fontSize:clamp(Number(o.fontSize)||42,10,180),color:/^#[0-9a-f]{6}$/i.test(o.color)?o.color:'#20374d',text:String(o.text||''),image:/^data:image\/(png|jpeg|webp);base64,/i.test(o.image||'')?o.image:''}));
    result.slides.push({id,title:String(slide.title||`슬라이드 ${i+1}`).slice(0,500),kind:slide.kind,pointWidth:clamp(Number(slide.pointWidth)||960,100,10000),width:clamp(Number(slide.width)||1600,100,5000),height:clamp(Number(slide.height)||900,100,5000),html:cleanHtml(String(slide.html||'').slice(0,2e6)),image:isImageData(slide.image)?slide.image:undefined,background:isImageData(slide.background)?slide.background:'',backgroundColor:/^#[0-9a-f]{6}$/i.test(slide.backgroundColor||'')?slide.backgroundColor:'#ffffff',sourceId:typeof slide.sourceId==='string'?slide.sourceId:null,sourceSlideId:Number.isFinite(Number(slide.sourceSlideId))&&slide.sourceSlideId?Number(slide.sourceSlideId):null,objects:Array.isArray(slide.objects)?slide.objects.slice(0,2000).map(normalizePptObject):[],deletedObjects:Array.isArray(slide.deletedObjects)?slide.deletedObjects.map(Number).filter(Number.isFinite):[],reorderObjects:slide.reorderObjects===true,notes:String(slide.notes||''),cue:String(slide.cue||''),duration:clamp(Number(slide.duration)||30,0,3600),overlays});
  }
  return result;
}
async function openDatabase(){return new Promise((resolve,reject)=>{const request=indexedDB.open('slide-studio-v1',1);request.onupgradeneeded=()=>request.result.createObjectStore('projects');request.onerror=()=>reject(request.error);request.onsuccess=()=>resolve(request.result);});}
function databaseGet(key){return new Promise((resolve,reject)=>{const request=db.transaction('projects').objectStore('projects').get(key);request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);});}
async function persist(){
  clearTimeout(saveTimer);if(!db){$('save-status').textContent='자동 저장 불가 · 프로젝트를 파일로 저장하세요';return;}
  const snapshot=clone(deck);$('save-status').textContent='저장 중…';
  savingPromise=savingPromise.catch(()=>{}).then(()=>new Promise((resolve,reject)=>{const transaction=db.transaction('projects','readwrite');transaction.objectStore('projects').put(snapshot,'last');transaction.oncomplete=resolve;transaction.onerror=()=>reject(transaction.error);transaction.onabort=()=>reject(transaction.error);}));
  try{await savingPromise;if(snapshot.updatedAt===deck.updatedAt){dirty=false;$('save-status').textContent='이 브라우저에 자동 저장됨';}}catch{$('save-status').textContent='자동 저장 실패 · 프로젝트를 파일로 저장하세요';toast('저장 공간이 부족할 수 있습니다. 프로젝트 저장으로 파일을 백업하세요.',7000);}
}
function markChanged(key='change'){
  deck.updatedAt=Date.now();dirty=true;$('save-status').textContent='변경 내용 저장 대기…';clearTimeout(saveTimer);saveTimer=setTimeout(persist,650);recordHistory(key);if(!/^(notes-|cue-|duration|target)/.test(key))syncStage();
}
function resetHistory(){history=[{deck:clone(deck),index}];historyIndex=0;groupTime=0;groupKey='';updateHistoryButtons();}
function recordHistory(key){const now=Date.now();const snapshot={deck:clone(deck),index};if(historyIndex>0&&historyIndex===history.length-1&&groupKey===key&&now-groupTime<1200){history[historyIndex]=snapshot;}else{history=history.slice(0,historyIndex+1);history.push(snapshot);if(history.length>30)history.shift();historyIndex=history.length-1;}groupTime=now;groupKey=key;updateHistoryButtons();}
function updateHistoryButtons(){$('undo-button').disabled=historyIndex<=0;$('redo-button').disabled=historyIndex>=history.length-1;}
function undo(direction){const next=historyIndex+direction;if(next<0||next>=history.length)return;const activeId=current()?.id;historyIndex=next;deck=clone(history[next].deck);const restoredIndex=deck.slides.findIndex(s=>s.id===activeId);index=restoredIndex>=0?restoredIndex:clamp(history[next].index,0,Math.max(0,deck.slides.length-1));groupTime=0;render();deck.updatedAt=Date.now();dirty=true;clearTimeout(saveTimer);saveTimer=setTimeout(persist,300);syncStage();}
function replaceDeck(value){deck=value;index=0;selected=null;mode='edit';changeTab('notes');resetTimer();resetHistory();render();markChanged('open');resetHistory();}

function overlayMarkup(slide){return(slide.overlays||[]).map(o=>`<div class="overlay ${o.type==='text'?'overlay-text':''}" data-overlay="${escapeHtml(o.id)}" style="left:${o.x}px;top:${o.y}px;width:${o.width}px;${o.type==='image'?`height:${o.height}px;`:''}font-size:${o.fontSize||42}px;color:${escapeHtml(o.color||'#20374d')};font-weight:${o.bold?'700':'400'}">${o.type==='image'?`<img src="${escapeHtml(o.image)}" alt="추가한 이미지">`:escapeHtml(o.text)}</div>`).join('');}
function frameDocument(slide, editable){
  const reference=slide.kind==='reference';let content=slide.kind==='powerpoint'?powerpointMarkup(slide,editable):slide.kind==='image'?`<img class="slide-image" src="${escapeHtml(slide.image)}" alt="${escapeHtml(slide.title)}">`:slide.html;
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>${reference?(deck.referenceCss||''):''}\n${nativeCss}\n${powerpointCss}</style></head><body><div id="canvas" class="${editable?'editor-on ':''}static" style="width:${slide.width}px;height:${slide.height}px;background:${slide.backgroundColor||'#ffffff'}"><div class="stage ${reference?'deck-kj reference':''}"><div class="slide" id="slide-content">${content}</div>${overlayMarkup(slide)}${reference||slide.kind==='powerpoint'?'':`<div class="page-number">${index+1} / ${deck.slides.length}</div>`}</div></div></body></html>`;
}
function fitFrame(frame,slide){const canvas=frame.contentDocument?.getElementById('canvas');if(!canvas)return;const width=frame.clientWidth,height=frame.clientHeight,scale=Math.min(width/slide.width,height/slide.height);canvas.style.transform=`translate(${(width-slide.width*scale)/2}px,${(height-slide.height*scale)/2}px) scale(${scale})`;}
function fitCanvas(){const slide=current();if(!slide||stageMode)return;const viewport=$('canvas-viewport'),ratio=slide.width/slide.height;const width=Math.max(1,Math.min(viewport.clientWidth,viewport.clientHeight*ratio));$('canvas-shell').style.width=`${width}px`;$('canvas-shell').style.height=`${width/ratio}px`;fitFrame($('slide-frame'),slide);}
function drawFrame(){
  const slide=current();if(!slide)return;const frame=$('slide-frame'),version=++frameVersion;selected=null;
  frame.onload=()=>{if(version!==frameVersion)return;fitCanvas();bindFrame(frame,slide);if(slide.kind==='reference')fitReference(frame);frame.contentDocument.fonts.ready.then(()=>{if(version===frameVersion){fitReference(frame);fitCanvas();}});};frame.srcdoc=frameDocument(slide,mode==='edit');
}
function fitReference(frame){
  const doc=frame.contentDocument;if(!doc)return;
  for(const body of doc.querySelectorAll('.k-body')){body.style.transformOrigin='top left';let scale=1;const height=body.clientHeight,width=body.clientWidth;while(scale>.60&&(body.scrollHeight*scale>height+2||body.scrollWidth*scale>width+2))scale-=.025;if(scale<1){body.style.transform=`scale(${scale})`;body.style.width=`${width/scale}px`;}}
}
function captureFrame(slide){
  const doc=$('slide-frame').contentDocument;if(!doc||slide.id!==current()?.id)return;
  if(slide.kind!=='image'&&slide.kind!=='powerpoint'){const root=doc.getElementById('slide-content').cloneNode(true);root.querySelectorAll('[data-editable]').forEach(el=>{el.removeAttribute('contenteditable');el.removeAttribute('tabindex');el.classList.remove('selected');if(slide.kind==='reference')el.removeAttribute('data-editable');});root.querySelectorAll('.selected').forEach(el=>el.classList.remove('selected'));slide.html=root.innerHTML;}
}
function bindFrame(frame,slide){
  const doc=frame.contentDocument;if(!doc)return;
  pptEditor.bind(frame,slide);
  if(mode==='edit'){
    const candidates=slide.kind==='reference'?[...doc.querySelectorAll(referenceSelectors)].filter(el=>!el.querySelector(referenceSelectors)):[...doc.querySelectorAll('#slide-content [data-editable]')];
    candidates.forEach((el,i)=>{el.dataset.editable='true';el.dataset.editorId=i;el.contentEditable='true';el.tabIndex=0;el.addEventListener('focus',()=>selectElement(el));el.addEventListener('input',()=>{captureFrame(slide);markChanged(`html-${slide.id}`);updateThumbnail(slide);});el.addEventListener('paste',e=>{e.preventDefault();doc.execCommand('insertText',false,e.clipboardData.getData('text/plain'));});});
    doc.querySelectorAll('[data-overlay]').forEach(el=>bindOverlay(el,slide));
  }
  doc.addEventListener('click',e=>{if(mode!=='edit'||e.target.closest('[data-ppt-id]'))return;pptEditor.clear();const el=e.target.closest('[data-editable],[data-overlay]');if(el)selectElement(el);else selectElement(null);});
  doc.addEventListener('keydown',e=>handleKey(e));
}
function selectElement(el){const doc=$('slide-frame').contentDocument;doc?.querySelectorAll('.selected').forEach(node=>node.classList.remove('selected'));selected=el;el?.classList.add('selected');$('remove-element').disabled=!el?.hasAttribute('data-overlay');}
function bindOverlay(el,slide){
  const object=slide.overlays.find(o=>o.id===el.dataset.overlay);if(!object)return;
  if(object.type==='text'){el.contentEditable='true';el.setAttribute('data-editable','true');el.addEventListener('input',()=>{object.text=el.innerText;markChanged(`overlay-${object.id}`);});el.addEventListener('paste',e=>{e.preventDefault();el.ownerDocument.execCommand('insertText',false,e.clipboardData.getData('text/plain'));});}
  el.addEventListener('focus',()=>selectElement(el));
  el.addEventListener('pointerdown',e=>{
    selectElement(el);const rect=el.getBoundingClientRect(),scale=rect.width/el.offsetWidth;
    if(object.type==='text'&&e.clientX>=rect.left+8*scale&&e.clientY>=rect.top+8*scale)return;
    e.preventDefault();const startX=e.clientX,startY=e.clientY,x=object.x,y=object.y;el.setPointerCapture(e.pointerId);
    const move=event=>{object.x=clamp(x+(event.clientX-startX)/scale,0,slide.width-object.width);object.y=clamp(y+(event.clientY-startY)/scale,0,slide.height-40);el.style.left=object.x+'px';el.style.top=object.y+'px';};
    const finish=()=>{el.removeEventListener('pointermove',move);el.removeEventListener('pointerup',finish);el.removeEventListener('pointercancel',finish);markChanged(`move-${object.id}`);};el.addEventListener('pointermove',move);el.addEventListener('pointerup',finish);el.addEventListener('pointercancel',finish);
  });
}
function render(){
  const slide=current(),has=!!slide;$('project-title').value=deck.title;$('target-minutes').value=deck.targetMinutes;$('empty-state').hidden=has;$('slide-area').hidden=!has;$('filmstrip-panel').hidden=!has;$('notes-editor').disabled=!has;$('duration-label').hidden=!has;
  document.body.classList.toggle('rehearsing',mode==='rehearse');$('edit-mode').classList.toggle('active',mode==='edit');$('rehearse-mode').classList.toggle('active',mode==='rehearse');$('total-slides').textContent=deck.slides.length;updateHistoryButtons();
  if(!has){$('notes-editor').value='';$('notes-slide-label').textContent='슬라이드를 가져오면 여기에 연결됩니다';$('next-title').textContent='아직 슬라이드가 없어요';$('next-cue').textContent='자료를 가져오거나 빈 슬라이드로 시작하세요.';$('filmstrip').innerHTML='';renderNotesMeta();updateTimer();return;}
  $('current-number').textContent=`SLIDE ${String(index+1).padStart(2,'0')}`;$('slide-title').value=slide.title;$('slide-title').readOnly=mode!=='edit';$('slide-count').textContent=`${index+1} / ${deck.slides.length}`;$('canvas-size').textContent=`${slide.width} × ${slide.height}`;
  $('slide-kind').textContent=({image:'가져온 화면',reference:'기존 HTML · 직접 편집',native:'웹 슬라이드 · 직접 편집',powerpoint:'PPT · 요소 직접 편집'})[slide.kind];$('edit-hint').textContent=slide.kind==='powerpoint'?'글자 두 번 클릭: 수정 · 드래그: 이동 · 모서리: 크기 조절':slide.kind==='image'?'원본 화면 위에 텍스트·이미지를 추가할 수 있어요.':'문구를 클릭해서 편집하세요. 추가한 요소는 테두리를 잡고 이동할 수 있어요.';
  $('add-shape-button').hidden=slide.kind!=='powerpoint';$('ppt-object-panel').hidden=slide.kind!=='powerpoint'||mode!=='edit';
  $('prev-button').disabled=index===0;$('next-button').disabled=index===deck.slides.length-1;$('remove-element').disabled=true;$('notes-slide-label').textContent=`${String(index+1).padStart(2,'0')} · ${slide.title}`;$('slide-duration').value=slide.duration;$('notes-editor').value=slide[tab]||'';$('notes-editor').style.fontSize=noteSize+'px';$('notes-editor').scrollTop=0;renderNotesMeta();
  const next=deck.slides[index+1];$('next-title').textContent=next?`${String(index+2).padStart(2,'0')}  ${next.title}`:'마지막 슬라이드입니다';$('next-cue').textContent=next?(next.cue||next.notes||'다음 장의 대본을 미리 준비해 보세요.'):'준비한 이야기를 마무리하세요.';
  drawFrame();renderFilmstrip();updateTimer();
}
function renderNotesMeta(){const text=current()?.[tab]||'';const chars=[...text.replace(/\s/g,'')].length;$('notes-chars').textContent=`${chars.toLocaleString()}자`;$('notes-estimate').textContent=chars?`읽는 시간 약 ${formatTime(chars/5)}`:'대본을 자유롭게 편집하세요';}
function thumbContent(slide){return slide.image&&(!slide.objects?.some(o=>Object.keys(o.changes||{}).length)||slide.kind==='image')?`<img src="${escapeHtml(slide.image)}" alt="" loading="lazy">`:`<strong>${escapeHtml(slide.title)}</strong>`;}
function renderFilmstrip(){
  $('filmstrip').innerHTML=deck.slides.map((slide,i)=>`<button class="thumbnail ${i===index?'active':''}" data-slide-id="${escapeHtml(slide.id)}" data-index="${i}" draggable="true" role="listitem" aria-label="${i+1}번 ${escapeHtml(slide.title)}" ${i===index?'aria-current="true"':''}><div class="thumb-image">${thumbContent(slide)}</div><div class="thumb-caption"><b>${String(i+1).padStart(2,'0')}</b><span>${escapeHtml(slide.title)}</span></div></button>`).join('');
  const active=$('filmstrip').querySelector('.active');active?.scrollIntoView({block:'nearest',inline:'nearest'});
}
function updateThumbnail(slide){const button=[...$('filmstrip').children].find(el=>el.dataset.slideId===slide.id);if(button){button.querySelector('.thumb-caption span').textContent=slide.title;const strong=button.querySelector('strong');if(strong)strong.textContent=slide.title;}}
function navigate(next){if(!deck.slides.length)return;const value=clamp(next,0,deck.slides.length-1);if(value===index)return;index=value;groupTime=0;render();syncStage();}
function addSlide(template){
  if(deck.slides.length>=200){toast('한 프로젝트에 최대 200장을 사용할 수 있습니다.');return;}
  const width=current()?.width||1600,height=current()?.height||900,sx=width/1600,sy=height/900;
  const text=(name,value,x,y,w,h,size,color='#20374d',bold=false)=>normalizePptObject({id:uid(),name,type:'text',text:value,x:x*sx,y:y*sy,width:w*sx,height:h*sy,fontFamily:'맑은 고딕',fontSize:size*sx,color,bold,fillVisible:false,changes:{text:true,geometry:true}});
  let title='새 슬라이드',objects=[],backgroundColor='#ffffff';
  if(template==='cover'){title='발표 제목';backgroundColor='#edf7fa';objects=[text('주제','발표 주제',100,80,1400,50,24,'#197b92'),text('발표 제목','발표 제목을\n입력하세요',100,280,1400,250,89,'#20374d',true),text('소개 문장','발표를 소개하는 한 문장을 적어 주세요.',100,590,1400,100,32,'#6d8898'),text('발표자','발표자 · 소속 · 날짜',100,790,1400,50,22,'#7b94a4')];}
  if(template==='body'){title='슬라이드 제목';objects=[text('제목',title,100,130,1400,120,68,'#20374d',true),text('본문','전하고 싶은 내용을 입력하세요.\n\n글자는 두 번 클릭해서 수정할 수 있습니다.',100,320,1400,430,38)];}
  if(template==='columns'){title='두 가지 내용을 비교해 보세요';objects=[text('제목',title,100,130,1400,120,68,'#20374d',true),text('첫 번째 내용','첫 번째 내용\n\n설명을 입력하세요.',100,340,650,400,38,'#197b92'),text('두 번째 내용','두 번째 내용\n\n설명을 입력하세요.',850,340,650,400,38,'#197b92')];}
  if(template==='quote'){title='전하고 싶은 핵심 문장';objects=[text('핵심 문장','전하고 싶은\n핵심 문장을 적어 주세요.',160,280,1280,300,76,'#197b92',true),text('설명','문장에 대한 짧은 설명 또는 출처',160,650,1280,100,30,'#7b94a4')];}
  const slide={id:uid(),title,kind:'powerpoint',width,height,pointWidth:current()?.pointWidth||960,backgroundColor,objects,deletedObjects:[],notes:'',cue:'',duration:30,overlays:[]};
  const position=deck.slides.length?index+1:0;deck.slides.splice(position,0,slide);index=position;mode='edit';closeDialogs();markChanged('add');render();
}
function addOverlay(type,image=''){
  if(current()?.kind==='powerpoint'){pptEditor.add(type==='image'?'picture':'text','rect',image);return;}
  const slide=current();if(!slide||mode!=='edit')return;const object={id:uid(),type,x:slide.width*.1,y:slide.height*.4,width:slide.width*.5,height:slide.height*.4,text:type==='text'?'텍스트를 입력하세요':'',image,fontSize:42,color:'#20374d',bold:false};slide.overlays.push(object);markChanged('insert');drawFrame();
}
function styleSelected(kind,value){
  const pptObject=pptEditor.selected();if(pptObject&&mode==='edit'){if(kind==='bold')pptEditor.update('bold',!pptObject.bold);else if(kind==='font')pptEditor.update('fontSize',pptObject.fontSize+value*current().width/(current().pointWidth||960));else if(kind==='color')pptEditor.update('color',value);return;}
  const slide=current();if(!slide||mode!=='edit'||!selected){toast('슬라이드에서 편집할 문구를 먼저 클릭하세요.');return;}
  const object=slide.overlays.find(o=>o.id===selected.dataset.overlay);
  if(object){if(object.type!=='text')return;if(kind==='bold')object.bold=!object.bold;if(kind==='font')object.fontSize=clamp((object.fontSize||42)+value,10,180);if(kind==='color')object.color=value;selected.style.fontWeight=object.bold?'700':'400';selected.style.fontSize=object.fontSize+'px';selected.style.color=object.color;}
  else {if(kind==='bold')selected.style.fontWeight=selected.ownerDocument.defaultView.getComputedStyle(selected).fontWeight==='700'?'400':'700';if(kind==='font')selected.style.fontSize=clamp(parseFloat(selected.ownerDocument.defaultView.getComputedStyle(selected).fontSize)+value,10,200)+'px';if(kind==='color')selected.style.color=value;captureFrame(slide);}
  markChanged('style');
}
function download(content,type,name){const blob=new Blob([content],{type}),url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
function safeName(name){return name.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').slice(0,90)||'발표';}
function saveProject(){download(JSON.stringify({format:'slide-studio',deck},null,2),'application/json',safeName(deck.title)+'.slides.json');persist();toast('프로젝트 파일을 저장했습니다. 다시 가져오면 이어서 편집할 수 있어요.');}

function renderPptInspector(object){
  const panel=$('ppt-object-panel');if(!panel)return;const slide=current(),visible=slide?.kind==='powerpoint'&&mode==='edit';panel.hidden=!visible;if(!visible)return;
  const list=$('ppt-object-list');list.innerHTML='<option value="">요소를 선택하세요</option>'+slide.objects.map(o=>`<option value="${escapeHtml(o.id)}">${escapeHtml(o.name)}</option>`).join('');list.value=object?.id||'';
  $('ppt-properties').disabled=!object;$('remove-element').disabled=!object;$('ppt-replace-image').hidden=object?.type!=='picture';$('ppt-edit-text').disabled=!object||!['text','shape','table'].includes(object.type)||object.geometry==='line';
  if(!object)return;
  for(const input of panel.querySelectorAll('[data-ppt-property]')){const property=input.dataset.pptProperty,displayValue=property==='fontSize'?object.fontSize*(slide.pointWidth||960)/slide.width:object[property];if(document.activeElement!==input){if(input.type==='checkbox')input.checked=!!object[property];else input.value=typeof displayValue==='number'?Math.round(displayValue*10)/10:displayValue??'';}
    const font=['fontFamily','fontSize','color','align'].includes(property),paint=['fill','fillVisible','lineWidth','lineColor'].includes(property);input.disabled=font&&!['text','shape'].includes(object.type)||paint&&(!['text','shape'].includes(object.type)||object.geometry==='custom')||property==='rotation'&&object.canRotate===false;}
}
function overlayToPpt(o){return normalizePptObject({...o,type:o.type==='image'?'picture':'text',sourceId:null,image:o.type==='image'?o.image:'',fontFamily:'맑은 고딕',fillVisible:o.type!=='image',fill:'#ffffff',changes:{text:true,geometry:true,image:o.type==='image'}});}
async function savePptx(){
  if(busy)return;if(!deck.slides.length){toast('저장할 슬라이드가 없습니다.');return;}
  pptEditor.finishEdit();setBusy(true,'PPTX로 저장하고 있어요','수정한 요소와 대본을 PowerPoint 파일에 반영합니다.');
  try{
    const sourceIds=[...new Set(deck.slides.filter(s=>s.sourceSlideId).map(s=>s.sourceId))];if(sourceIds.length>1)throw new Error('서로 다른 원본 PPT가 섞인 프로젝트는 각각 저장해 주세요.');
    const source=sourceIds.length?deck.sources?.[sourceIds[0]]:null;if(sourceIds.length&&!source)throw new Error('원본 PPTX 정보가 없습니다. PPT 파일을 다시 가져와 주세요.');
    const slides=[];
    for(const[sidx,s]of deck.slides.entries()){
      $('busy-detail').textContent=`${sidx+1} / ${deck.slides.length}장 준비 중`;
      const item={title:s.title,width:s.width,height:s.height,notes:s.notes||'',sourceSlideId:s.sourceSlideId||null,backgroundColor:s.backgroundColor||'#ffffff',deletedObjects:s.deletedObjects||[],reorderObjects:s.reorderObjects===true,objects:[]};
      if(s.kind==='powerpoint')item.objects=s.objects.map(o=>{const{preview,backdrop,...value}=clone(o);if(value.type==='picture'&&!value.sourceId&&!value.cloneSourceId)value.image=value.image||preview;return value;});
      else if(s.kind==='image')item.backgroundImage=s.image;
      else{if(!window.slideDesktop?.renderSlide)throw new Error('기존 웹 슬라이드의 PPTX 저장은 설치형 앱에서 사용할 수 있습니다.');const bytes=await window.slideDesktop.renderSlide({html:frameDocument({...s,overlays:[]},false),width:s.width,height:s.height});const blob=new Blob([new Uint8Array(bytes)],{type:'image/png'});item.backgroundImage=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=reject;reader.readAsDataURL(blob);});}
      item.objects.push(...(s.overlays||[]).map(overlayToPpt));slides.push(item);
    }
    $('busy-detail').textContent='PowerPoint에서 PPTX 파일을 생성하고 있습니다.';
    const response=await fetch('/api/export/powerpoint',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({source,slides})});
    if(!response.ok){const error=await response.json();throw new Error(error.error||'PPTX 저장에 실패했습니다.');}
    download(await response.arrayBuffer(),'application/vnd.openxmlformats-officedocument.presentationml.presentation',safeName(deck.title)+'_수정본.pptx');await persist();toast('수정한 PPTX를 저장했습니다. PowerPoint에서도 다시 편집할 수 있어요.');
  }catch(error){console.error(error);toast(error.message,10000);}finally{setBusy(false);}
}

async function importFile(file){
  if(!file||busy)return;
  const ext=file.name.split('.').pop().toLowerCase();if(!['json','pdf','ppt','pptx'].includes(ext)){toast('PPT, PPTX, PDF 또는 프로젝트 JSON 파일을 선택하세요.');return;}
  if(file.size>(ext==='json'?512:150)*1024*1024){toast(ext==='json'?'프로젝트는 512MB 이하로 선택하세요.':'150MB 이하의 파일을 선택하세요.');return;}
  closeDialogs();setBusy(true,'자료를 가져오고 있어요',file.name);
  try {
    await persist();let next;
    if(ext==='json'){next=validateDeck(JSON.parse(await file.text()));}
    else if(ext==='pdf'){next=await importPdf(file);}
    else{
      $('busy-detail').textContent='PowerPoint에서 슬라이드 화면과 발표자 노트를 가져오고 있습니다. 장수에 따라 잠시 걸릴 수 있어요.';
      const response=await fetch('/api/import/powerpoint',{method:'POST',headers:{'Content-Type':'application/octet-stream','X-File-Name':encodeURIComponent(file.name)},body:file});
      const result=await response.json();if(!response.ok)throw new Error(result.error||'PPT 변환에 실패했습니다.');
      const sourceId=uid();next={...freshDeck(),title:file.name.replace(/\.[^.]+$/,''),sources:{[sourceId]:result.pptx},slides:result.slides.map((s,i)=>({id:uid(),title:s.title||`슬라이드 ${i+1}`,kind:'powerpoint',width:s.width,height:s.height,pointWidth:s.pointWidth||960,image:s.image,background:s.background,sourceId,sourceSlideId:s.sourceSlideId,objects:(s.objects||[]).map(normalizePptObject),deletedObjects:[],notes:s.notes||'',cue:'',duration:30,overlays:[]}))};next.targetMinutes=Math.max(1,Math.round(next.slides.length*.5));
    }
    if(!next.slides.length&&ext!=='json')throw new Error('가져올 슬라이드가 없습니다.');replaceDeck(next);await persist();toast(`${next.slides.length}장을 가져왔습니다.${ext==='pptx'||ext==='ppt'?' 발표자 노트도 대본에 연결했습니다.':''}`);
  }catch(error){console.error(error);toast(error.message||'자료를 가져오지 못했습니다. 파일 형식을 확인하세요.',10000);}finally{setBusy(false);$('file-input').value='';}
}
async function importPdf(file){
  const pdfjs=await import('./vendor/pdfjs/build/pdf.mjs');pdfjs.GlobalWorkerOptions.workerSrc=new URL('./vendor/pdfjs/build/pdf.worker.mjs',import.meta.url).href;
  const task=pdfjs.getDocument({data:new Uint8Array(await file.arrayBuffer()),cMapUrl:new URL('./vendor/pdfjs/cmaps/',import.meta.url).href,cMapPacked:true,standardFontDataUrl:new URL('./vendor/pdfjs/standard_fonts/',import.meta.url).href,wasmUrl:new URL('./vendor/pdfjs/wasm/',import.meta.url).href,isEvalSupported:false,enableXfa:false});
  task.onPassword=(_,reason)=>{task.destroy();toast(reason===1?'암호가 있는 PDF입니다. 암호를 해제한 파일을 가져오세요.':'PDF 암호를 확인할 수 없습니다.',8000);};
  const pdf=await task.promise;
  try{if(pdf.numPages>200)throw new Error('PDF는 최대 200페이지까지 가져올 수 있습니다.');const slides=[];
    for(let i=1;i<=pdf.numPages;i++){
      $('busy-title').textContent=`PDF를 슬라이드로 가져오는 중 · ${i} / ${pdf.numPages}`;
      const page=await pdf.getPage(i),base=page.getViewport({scale:1}),scale=Math.min(1800/base.width,2200/base.height),viewport=page.getViewport({scale}),canvas=document.createElement('canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
      await page.render({canvasContext:canvas.getContext('2d'),viewport,background:'rgb(255,255,255)'}).promise;
      const text=await page.getTextContent();const title=text.items.filter(item=>item.str?.trim()).slice(0,3).map(item=>item.str).join(' ').slice(0,100)||`페이지 ${i}`;
      slides.push({id:uid(),title,kind:'image',width:canvas.width,height:canvas.height,image:canvas.toDataURL('image/png'),notes:'',cue:'',duration:30,overlays:[]});canvas.width=canvas.height=1;page.cleanup();
    }
    return{...freshDeck(),title:file.name.replace(/\.pdf$/i,''),slides,targetMinutes:Math.max(1,Math.round(slides.length*.5))};
  }finally{await task.destroy();}
}
async function loadReference(){if(busy)return;setBusy(true,'예제를 불러오고 있어요','슬라이드와 대본을 가져옵니다.');try{await persist();const response=await fetch('/api/reference'),result=await response.json();if(!response.ok)throw new Error(result.error);replaceDeck(validateDeck(result));toast('예제를 불러왔습니다. 화면의 문구를 클릭해 편집해 보세요.');}catch(error){toast(error.message,8000);}finally{setBusy(false);}}

function elapsed(){return elapsedBase+(timerRunning?(performance.now()-timerStarted)/1000:0);}
function toggleTimer(){if(timerRunning){elapsedBase=elapsed();timerRunning=false;}else{timerStarted=performance.now();timerRunning=true;}updateTimer();}
function resetTimer(){timerRunning=false;elapsedBase=0;timerStarted=0;updateTimer();}
function updateTimer(){if(stageMode)return;const seconds=elapsed(),remaining=Math.ceil(deck.targetMinutes*60-seconds);$('elapsed').textContent=formatTime(seconds);$('timer-button').textContent=timerRunning?'일시정지':seconds>0?'계속':'시작';$('remaining').textContent=remaining>=0?`${Math.floor(remaining/60)}분 ${remaining%60}초 남음`:`${formatTime(-remaining)} 초과`;$('timer-progress').style.width=clamp(seconds/(deck.targetMinutes*60)*100,0,100)+'%';document.body.classList.toggle('over-time',remaining<0);$('clock').textContent=new Date().toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit',hour12:false});}
function syncStage(){if(stageWindow&&!stageWindow.closed){const {sources,...visibleDeck}=deck;visibleDeck.slides=deck.slides.map((s,i)=>{if(i!==index)return{id:s.id,title:s.title,kind:s.kind,width:s.width,height:s.height};const {notes,cue,...visual}=s;return visual;});stageWindow.postMessage({type:'slide-studio-state',deck:visibleDeck,index,black:isBlack},location.origin);}}
function openStage(){if(!current()){toast('발표할 슬라이드를 먼저 가져오거나 추가하세요.');return;}stageWindow=window.open(location.pathname+'?stage=1',`slide-studio-${deck.id}`,'popup,width=1280,height=760');if(!stageWindow){toast('발표화면을 열려면 브라우저의 팝업 차단을 해제하세요.');return;}stageWindow.focus();syncStage();}
async function fullscreen(){try{if(document.fullscreenElement)await document.exitFullscreen();else await document.documentElement.requestFullscreen();}catch{toast('브라우저 메뉴에서 전체화면을 선택해 주세요.');}}
function handleKey(event){
  if(pptEditor.key(event))return;
  const target=event.target,editing=target?.matches?.('input,textarea,select')||target?.isContentEditable;
  if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='s'){event.preventDefault();if(event.shiftKey)savePptx();else saveProject();return;}
  if(editing)return;
  if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='z'){event.preventDefault();undo(event.shiftKey?1:-1);return;}
  if(event.ctrlKey||event.metaKey||event.altKey||busy||document.querySelector('dialog[open]'))return;
  const key=event.key.toLowerCase();if(['arrowright','pagedown',' '].includes(key)){event.preventDefault();navigate(index+1);}else if(['arrowleft','pageup'].includes(key)){event.preventDefault();navigate(index-1);}else if(key==='home'){event.preventDefault();navigate(0);}else if(key==='end'){event.preventDefault();navigate(deck.slides.length-1);}else if(key==='t'){toggleTimer();}else if(key==='b'){isBlack=!isBlack;syncStage();toast(isBlack?'발표화면을 검은 화면으로 가렸습니다. B로 해제하세요.':'발표화면 가리기를 해제했습니다.');}else if(key==='f'){fullscreen();}else if(key==='escape'){selectElement(null);}
}
function exportHtml(){
  if(!current()){toast('내보낼 슬라이드가 없습니다.');return;}
  const {sources,...visible}=deck;const visual={...clone(visible),slides:deck.slides.map(({notes,cue,...slide})=>slide.kind==='powerpoint'?{id:slide.id,title:slide.title,kind:'native',width:slide.width,height:slide.height,html:powerpointMarkup(slide,false),overlays:slide.overlays||[]}:slide)};
  const css=`html,body{margin:0;width:100%;height:100%;overflow:hidden;background:#0c1622;color:white;font-family:"Malgun Gothic",sans-serif}#host{position:absolute;inset:0 0 55px;display:flex;align-items:center;justify-content:center}iframe{border:0;background:white}nav{position:absolute;bottom:0;left:0;right:0;height:55px;display:flex;align-items:center;justify-content:center;gap:25px;font-size:12px;background:#13243a}button{border:0;border-radius:5px;padding:7px 14px;background:#274457;color:white;cursor:pointer}#cover{position:fixed;inset:0;background:black;z-index:5}nav.hidden{display:none}`;
  const script=`const deck=${safeJson(visual)},nativeCss=${safeJson(nativeCss+powerpointCss)};let index=0,black=false;const esc=s=>String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));const frame=document.getElementById('frame');function fit(){const s=deck.slides[index],w=innerWidth,h=innerHeight-(document.querySelector('nav').classList.contains('hidden')?0:55),scale=Math.min(w/s.width,h/s.height);frame.style.width=s.width*scale+'px';frame.style.height=s.height*scale+'px';const c=frame.contentDocument?.getElementById('canvas');if(c)c.style.transform='scale('+scale+')';}function draw(){const s=deck.slides[index];const overlays=(s.overlays||[]).map(o=>'<div class="overlay '+(o.type==='text'?'overlay-text':'')+'" style="left:'+o.x+'px;top:'+o.y+'px;width:'+o.width+'px;'+(o.type==='image'?'height:'+o.height+'px;':'')+'font-size:'+o.fontSize+'px;color:'+esc(o.color)+';font-weight:'+(o.bold?700:400)+'">'+(o.type==='image'?'<img src="'+esc(o.image)+'" alt="">':esc(o.text))+'</div>').join('');frame.onload=()=>{fit();frame.contentDocument.addEventListener('keydown',keys);};frame.srcdoc='<!doctype html><html><head><meta charset="utf-8"><style>'+(s.kind==='reference'?deck.referenceCss||'':'')+nativeCss+'</style></head><body><div id="canvas" class="static" style="width:'+s.width+'px;height:'+s.height+'px"><div class="stage '+(s.kind==='reference'?'deck-kj reference':'')+'"><div class="slide">'+(s.kind==='image'?'<img class="slide-image" src="'+esc(s.image)+'" alt="">':s.html)+'</div>'+overlays+'</div></div></body></html>';document.getElementById('count').textContent=(index+1)+' / '+deck.slides.length;}function go(d){index=Math.max(0,Math.min(deck.slides.length-1,index+d));draw();}function keys(e){if(['ArrowRight','PageDown',' '].includes(e.key)){e.preventDefault();go(1);}if(['ArrowLeft','PageUp'].includes(e.key)){e.preventDefault();go(-1);}if(e.key.toLowerCase()==='b'){black=!black;document.getElementById('cover').hidden=!black;}if(e.key.toLowerCase()==='f')document.documentElement.requestFullscreen?.().catch(()=>{});}document.getElementById('prev').onclick=()=>go(-1);document.getElementById('next').onclick=()=>go(1);document.getElementById('full').onclick=()=>{document.querySelector('nav').classList.add('hidden');document.getElementById('host').style.bottom='0';document.documentElement.requestFullscreen?.().catch(()=>{});fit();};document.addEventListener('fullscreenchange',()=>{if(!document.fullscreenElement){document.querySelector('nav').classList.remove('hidden');document.getElementById('host').style.bottom='55px';fit();}});document.addEventListener('keydown',keys);window.addEventListener('resize',fit);draw();`;
  download(`<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(deck.title)}</title><style>${css}</style></head><body><main id="host"><iframe id="frame" title="발표 슬라이드" sandbox="allow-same-origin"></iframe></main><div id="cover" hidden></div><nav><button id="prev">← 이전</button><span id="count"></span><button id="next">다음 →</button><button id="full">전체화면</button></nav><script>${script}</script></body></html>`,'text/html',safeName(deck.title)+'_발표.html');
  toast('대본과 진행 메모를 제외한 발표용 HTML을 저장했습니다. 인터넷 없이 열 수 있어요.');
}

function startStage(){
  document.body.className='stage-mode';document.body.innerHTML='<div class="stage-wrap" id="stage-wrap"><iframe id="audience-frame" title="발표 슬라이드" sandbox="allow-same-origin"></iframe></div><div id="stage-cover" class="stage-cover" hidden></div><div class="stage-hint" id="stage-hint">발표자 창에서 슬라이드를 선택하세요.<button id="stage-full">전체화면</button></div>';
  let state=null,signature='';
  const fit=()=>{if(!state?.deck.slides.length)return;const s=state.deck.slides[state.index],scale=Math.min(innerWidth/s.width,innerHeight/s.height),frame=$('audience-frame');frame.style.width=s.width*scale+'px';frame.style.height=s.height*scale+'px';fitFrame(frame,s);};
  const request=()=>window.opener?.postMessage({type:'slide-studio-ready'},location.origin);
  const send=(action)=>window.opener?.postMessage({type:'slide-studio-action',action},location.origin);
  const key=e=>{const k=e.key.toLowerCase();if(['arrowright','pagedown',' '].includes(k)){e.preventDefault();send('next');}else if(['arrowleft','pageup'].includes(k)){e.preventDefault();send('prev');}else if(k==='b')send('black');else if(k==='f')document.documentElement.requestFullscreen().catch(()=>{});};
  window.addEventListener('message',e=>{if(e.origin!==location.origin||e.source!==window.opener||e.data?.type!=='slide-studio-state')return;state=e.data;deck=state.deck;index=state.index;$('stage-cover').hidden=!state.black;const sig=deck.id+':'+deck.updatedAt+':'+index;if(sig===signature)return;signature=sig;const s=current();if(!s)return;const frame=$('audience-frame');frame.onload=()=>{fit();frame.contentDocument.addEventListener('keydown',key);};frame.srcdoc=frameDocument(s,false);$('stage-hint').firstChild.textContent='← → 장 이동 · B 화면 가리기 · F 전체화면';});
  $('stage-full').onclick=()=>document.documentElement.requestFullscreen().catch(()=>{});window.addEventListener('resize',fit);document.addEventListener('keydown',key);let hintTimer;document.addEventListener('pointermove',()=>{$('stage-hint').style.opacity='1';clearTimeout(hintTimer);hintTimer=setTimeout(()=>{$('stage-hint').style.opacity='0';},2500);});request();
  return;
}

const actions={
  'save-pptx':savePptx,'add-shape':()=>showDialog('shape-dialog'),'ppt-edit-text':()=>pptEditor.edit(),'ppt-duplicate':()=>pptEditor.duplicate(),'ppt-front':()=>pptEditor.order('front'),'ppt-back':()=>pptEditor.order('back'),'ppt-replace-image':()=>{imageIntent='replace';$('image-input').click();},
  'desktop-update':async()=>{if(!window.slideDesktop)return;if(desktopUpdateState?.phase==='ready'){await persist();if(dirty){toast('먼저 프로젝트를 파일로 저장해 주세요.');return;}await window.slideDesktop.installUpdate();}else{const state=await window.slideDesktop.checkForUpdates();renderDesktopUpdate(state);if(state.phase==='current')toast('최신 버전을 사용하고 있습니다.');if(state.phase==='error')toast(state.message);}},
  import:()=>showDialog('import-dialog'),help:()=>showDialog('help-dialog'),'close-dialog':closeDialogs,
  new:()=>{showDialog('template-dialog');},add:()=>showDialog('template-dialog'),reference:loadReference,
  save:saveProject,export:exportHtml,present:openStage,prev:()=>navigate(index-1),next:()=>navigate(index+1),
  'edit-mode':()=>{mode='edit';render();},'rehearse-mode':()=>{mode='rehearse';render();},
  timer:toggleTimer,'timer-reset':resetTimer,undo:()=>undo(-1),redo:()=>undo(1),
  'notes-tab':()=>changeTab('notes'),'cue-tab':()=>changeTab('cue'),
  'notes-smaller':()=>{noteSize=clamp(noteSize-2,14,42);$('notes-editor').style.fontSize=noteSize+'px';},'notes-larger':()=>{noteSize=clamp(noteSize+2,14,42);$('notes-editor').style.fontSize=noteSize+'px';},
  'add-text':()=>addOverlay('text'),'add-image':()=>{if(current()&&mode==='edit'){imageIntent='add';$('image-input').click();}},bold:()=>styleSelected('bold'),'font-down':()=>styleSelected('font',-4),'font-up':()=>styleSelected('font',4),
  'remove-element':()=>{if(pptEditor.selected()){pptEditor.remove();return;}const s=current();if(!s||!selected?.dataset.overlay||mode!=='edit')return;s.overlays=s.overlays.filter(o=>o.id!==selected.dataset.overlay);markChanged('remove-element');drawFrame();},
  duplicate:()=>{if(!current())return;if(deck.slides.length>=200){toast('최대 200장까지 사용할 수 있습니다.');return;}const s=clone(current());s.id=uid();s.title+=' (복사)';s.overlays.forEach(o=>o.id=uid());deck.slides.splice(index+1,0,s);index++;markChanged('duplicate');render();},
  delete:()=>{if(!current())return;deck.slides.splice(index,1);index=clamp(index,0,Math.max(0,deck.slides.length-1));markChanged('delete');render();toast('슬라이드를 삭제했습니다. 실행 취소로 복구할 수 있어요.');},
  'move-left':()=>moveSlide(index,index-1),'move-right':()=>moveSlide(index,index+1)
};
function changeTab(value){tab=value;$('notes-tab').classList.toggle('active',tab==='notes');$('cue-tab').classList.toggle('active',tab==='cue');$('notes-tab').setAttribute('aria-selected',String(tab==='notes'));$('cue-tab').setAttribute('aria-selected',String(tab==='cue'));$('notes-editor').setAttribute('aria-label',tab==='notes'?'현재 슬라이드 발표 대본':'현재 슬라이드 진행 메모');$('notes-editor').placeholder=tab==='notes'?'이 슬라이드에서 전할 이야기를 적어 보세요.':'강조할 단어, 질문, 전환 멘트 등을 적어 보세요.';$('notes-editor').value=current()?.[tab]||'';renderNotesMeta();}
function moveSlide(from,to){if(to<0||to>=deck.slides.length||from===to)return;const active=current().id;const[s]=deck.slides.splice(from,1);deck.slides.splice(to,0,s);index=deck.slides.findIndex(s=>s.id===active);markChanged('reorder');render();}
function renderDesktopUpdate(state){
  desktopUpdateState=state;const button=$('desktop-update');button.hidden=false;button.title=state.message;
  button.disabled=['checking','downloading','development'].includes(state.phase);
  button.textContent=state.phase==='ready'?'업데이트 설치':state.phase==='downloading'?`업데이트 ${state.percent}%`:state.phase==='checking'?'업데이트 확인 중':state.phase==='development'?'개발 버전':'업데이트 확인';
  button.classList.toggle('primary',state.phase==='ready');
  $('desktop-version').title=state.message;
}
async function initDesktop(){
  if(!window.slideDesktop)return;
  window.slideDesktop.onPrepareClose(async()=>{await persist();return!dirty;});
  window.slideDesktop.onUpdate(renderDesktopUpdate);
  const info=await window.slideDesktop.info();$('desktop-version').hidden=false;$('desktop-version').textContent=`발표 작업실 v${info.version} · 설치형`;
  renderDesktopUpdate(info.updates);
}
async function main(){
  if(stageMode){startStage();return;}
  $('notes-editor').placeholder='이 슬라이드에서 전할 이야기를 적어 보세요.\n\nPPT에 발표자 노트가 있으면 여기에 자동으로 가져옵니다.';
  if(location.protocol==='file:'){toast('PPT·PDF를 가져오려면 같은 폴더의 발표작업실_시작.cmd를 실행하세요.',12000);}
  document.addEventListener('click',event=>{const button=event.target.closest('[data-action]');if(button&&!busy)actions[button.dataset.action]?.();const template=event.target.closest('[data-template]');if(template&&!busy)addSlide(template.dataset.template);const shape=event.target.closest('[data-shape]');if(shape&&!busy){closeDialogs();pptEditor.add('shape',shape.dataset.shape);}});
  $('ppt-object-list').addEventListener('change',event=>{if(event.target.value)pptEditor.select(event.target.value);else pptEditor.clear();});
  for(const input of document.querySelectorAll('[data-ppt-property]'))input.addEventListener(input.type==='color'?'input':'change',()=>{let value=input.type==='checkbox'?input.checked:input.value;if(input.dataset.pptProperty==='fontSize')value=Number(value)*current().width/(current().pointWidth||960);pptEditor.update(input.dataset.pptProperty,value);});
  document.querySelectorAll('dialog').forEach(dialog=>dialog.addEventListener('click',e=>{if(e.target===dialog)dialog.close();}));
  $('project-title').addEventListener('input',()=>{deck.title=$('project-title').value;markChanged('project-title');});
  $('slide-title').addEventListener('input',()=>{const s=current();if(!s)return;s.title=$('slide-title').value;markChanged('slide-title');updateThumbnail(s);$('notes-slide-label').textContent=`${String(index+1).padStart(2,'0')} · ${s.title}`;});
  $('notes-editor').addEventListener('input',()=>{const s=current();if(!s)return;s[tab]=$('notes-editor').value;markChanged(`${tab}-${s.id}`);renderNotesMeta();});
  $('target-minutes').addEventListener('change',()=>{deck.targetMinutes=clamp(Number($('target-minutes').value)||5,1,600);$('target-minutes').value=deck.targetMinutes;markChanged('target');updateTimer();});
  $('slide-duration').addEventListener('change',()=>{const s=current();if(s){s.duration=clamp(Number($('slide-duration').value)||0,0,3600);markChanged('duration');}});
  $('text-color').addEventListener('input',e=>styleSelected('color',e.target.value));
  $('file-input').addEventListener('change',e=>importFile(e.target.files[0]));
  $('image-input').addEventListener('change',async e=>{const file=e.target.files[0];if(!file)return;if(file.size>15e6){toast('이미지는 15MB 이하로 선택하세요.');return;}const intent=imageIntent;try{const bitmap=await createImageBitmap(file),canvas=document.createElement('canvas');canvas.width=bitmap.width;canvas.height=bitmap.height;canvas.getContext('2d').drawImage(bitmap,0,0);const image=canvas.toDataURL('image/png');bitmap.close();if(intent==='replace')pptEditor.replaceImage(image);else addOverlay('image',image);}catch{toast('이미지를 읽지 못했습니다. PNG 또는 JPG 파일을 선택하세요.');}e.target.value='';});
  $('drop-zone').addEventListener('click',()=>$('file-input').click());$('drop-zone').addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('file-input').click();}});
  for(const name of ['dragenter','dragover'])$('drop-zone').addEventListener(name,e=>{e.preventDefault();$('drop-zone').classList.add('over');});
  for(const name of ['dragleave','drop'])$('drop-zone').addEventListener(name,e=>{e.preventDefault();$('drop-zone').classList.remove('over');if(name==='drop')importFile(e.dataTransfer.files[0]);});
  document.addEventListener('dragover',e=>{if(e.dataTransfer.types.includes('Files'))e.preventDefault();});document.addEventListener('drop',e=>{if(e.dataTransfer.files.length&&!$('drop-zone').contains(e.target)){e.preventDefault();importFile(e.dataTransfer.files[0]);}});
  $('filmstrip').addEventListener('click',e=>{const button=e.target.closest('[data-index]');if(button)navigate(Number(button.dataset.index));});
  $('filmstrip').addEventListener('dragstart',e=>{const button=e.target.closest('[data-index]');if(button){dragIndex=Number(button.dataset.index);e.dataTransfer.setData('text/plain',button.dataset.slideId);e.dataTransfer.effectAllowed='move';}});
  $('filmstrip').addEventListener('dragover',e=>{if(dragIndex<0)return;e.preventDefault();$('filmstrip').querySelectorAll('.drag-over').forEach(el=>el.classList.remove('drag-over'));e.target.closest('[data-index]')?.classList.add('drag-over');});
  $('filmstrip').addEventListener('drop',e=>{e.preventDefault();const button=e.target.closest('[data-index]');if(button&&dragIndex>=0)moveSlide(dragIndex,Number(button.dataset.index));dragIndex=-1;});$('filmstrip').addEventListener('dragend',()=>{dragIndex=-1;$('filmstrip').querySelectorAll('.drag-over').forEach(el=>el.classList.remove('drag-over'));});
  document.addEventListener('keydown',handleKey);window.addEventListener('resize',fitCanvas);new ResizeObserver(fitCanvas).observe($('canvas-viewport'));
  window.addEventListener('message',e=>{if(e.origin!==location.origin||e.source!==stageWindow)return;if(e.data?.type==='slide-studio-ready')syncStage();else if(e.data?.type==='slide-studio-action'){if(e.data.action==='black'){isBlack=!isBlack;syncStage();}else actions[e.data.action]?.();}});
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='hidden'&&dirty)persist();});window.addEventListener('beforeunload',e=>{if(dirty){persist();e.preventDefault();e.returnValue='';}});
  try{db=await openDatabase();const saved=await databaseGet('last');if(saved)deck=validateDeck(saved);}catch(error){console.error(error);$('save-status').textContent='자동 저장 불가 · 프로젝트를 파일로 저장하세요';}
  resetHistory();render();setInterval(updateTimer,500);await initDesktop();
}
main();
