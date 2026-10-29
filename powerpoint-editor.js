import {normalizePptObject,objectBody,escapePpt,resizedGeometry} from './powerpoint-model.js';
const clamp=(value,min,max)=>Math.max(min,Math.min(max,value));
const clone=value=>JSON.parse(JSON.stringify(value));
export function createPowerpointEditor({getSlide,getMode,onChange,onSelect,redraw,toast}){
  let frame=null,selectedId=null,editing=false;
  const current=()=>getSlide()?.kind==='powerpoint'?getSlide():null;
  const selected=()=>current()?.objects.find(o=>o.id===selectedId)||null;
  const element=()=>frame?.contentDocument?.querySelector(`[data-ppt-id="${CSS.escape(selectedId||'')}"]`);
  function clear(){if(editing)finishEdit();selectedId=null;frame?.contentDocument?.querySelectorAll('.ppt-object.is-selected').forEach(el=>{el.classList.remove('is-selected');el.querySelectorAll('.ppt-handle,.ppt-drag-label').forEach(n=>n.remove());});onSelect(null);}
  function handles(el){el.querySelectorAll('.ppt-handle,.ppt-drag-label').forEach(n=>n.remove());if(getMode()!=='edit'||editing)return;el.insertAdjacentHTML('beforeend','<span class="ppt-drag-label">'+escapePpt(selected()?.name||'요소')+'</span>'+['nw','ne','sw','se'].map(h=>`<i class="ppt-handle" data-handle="${h}"></i>`).join(''));}
  function select(id){if(editing&&id!==selectedId)finishEdit();frame?.contentDocument?.querySelectorAll('.ppt-object.is-selected').forEach(el=>{el.classList.remove('is-selected');el.querySelectorAll('.ppt-handle,.ppt-drag-label').forEach(n=>n.remove());});selectedId=id;const el=element(),o=selected();if(!el||!o){selectedId=null;onSelect(null);return;}el.classList.add('is-selected');handles(el);onSelect(o);}
  function repaint(o=selected()){
    const el=element();if(!el||!o)return;el.style.left=o.x+'px';el.style.top=o.y+'px';el.style.width=o.width+'px';el.style.height=o.height+'px';el.style.transform=`rotate(${o.rotation}deg)`;
    el.innerHTML=objectBody(o);el.classList.toggle('is-selected',o.id===selectedId);handles(el);
  }
  function changed(key){onChange(key);onSelect(selected());}
  function finishEdit(){
    if(!editing)return;editing=false;const el=element();if(el){el.classList.remove('is-editing');repaint();}onSelect(selected());
  }
  function edit(){
    const o=selected(),el=element();if(!o||!el||getMode()!=='edit')return;
    if(!['text','table','shape'].includes(o.type)||o.geometry==='line'){toast('이 요소는 이동·크기 조절·회전할 수 있습니다.');return;}
    editing=true;el.classList.add('is-editing');el.innerHTML=objectBody({...o,type:o.type==='shape'?'text':o.type,changes:{...o.changes,[o.type==='table'?'table':'text']:true}});
    const editors=o.type==='table'?[...el.querySelectorAll('[data-cell]')]:[el.querySelector('[data-text-editor]')];
    for(const editor of editors.filter(Boolean)){
      editor.contentEditable='true';editor.addEventListener('pointerdown',event=>event.stopPropagation());
      editor.addEventListener('input',()=>{if(o.type==='table'){const[r,c]=editor.dataset.cell.split(',').map(Number);o.table[r][c].text=editor.innerText;o.changes.table=true;}else{o.text=editor.innerText;o.changes.text=true;}changed(`ppt-text-${o.id}`);});
      editor.addEventListener('paste',event=>{event.preventDefault();editor.ownerDocument.execCommand('insertText',false,event.clipboardData.getData('text/plain'));});
      editor.addEventListener('keydown',event=>{if(event.key==='Escape'){event.preventDefault();event.stopPropagation();finishEdit();}else if(event.key==='Enter'&&(event.ctrlKey||event.metaKey)){event.preventDefault();finishEdit();}});
      editor.addEventListener('keydown',event=>{if((event.ctrlKey||event.metaKey)&&['b','i'].includes(event.key.toLowerCase())){event.preventDefault();event.stopPropagation();if(o.type!=='table')update(event.key.toLowerCase()==='b'?'bold':'italic',!o[event.key.toLowerCase()==='b'?'bold':'italic']);}});
      editor.addEventListener('blur',()=>{setTimeout(()=>{if(!el.contains(el.ownerDocument.activeElement))finishEdit();},0);});
    }
    editors[0]?.focus();
  }
  function bind(nextFrame,slide){
    frame=nextFrame;editing=false;if(slide.kind!=='powerpoint'){selectedId=null;onSelect(null);return;}
    const doc=frame.contentDocument;
    if(getMode()!=='edit'){selectedId=null;onSelect(null);return;}
    for(const el of doc.querySelectorAll('[data-ppt-id]')){
      el.tabIndex=0;el.setAttribute('role','group');el.setAttribute('aria-label',slide.objects.find(o=>o.id===el.dataset.pptId)?.name||'PPT 요소');
      el.addEventListener('focus',()=>{if(selectedId!==el.dataset.pptId)select(el.dataset.pptId);});
      el.addEventListener('dblclick',event=>{event.preventDefault();event.stopPropagation();select(el.dataset.pptId);edit();});
      el.addEventListener('pointerdown',event=>{
        if(event.button!==0||editing&&el.dataset.pptId===selectedId)return;
        event.preventDefault();event.stopPropagation();select(el.dataset.pptId);el.focus({preventScroll:true});const o=selected();if(!o)return;
        const handle=event.target.closest('[data-handle]')?.dataset.handle;
        const canvas=doc.getElementById('canvas'),scale=canvas.getBoundingClientRect().width/slide.width;
        const start={x:event.clientX,y:event.clientY,ox:o.x,oy:o.y,w:o.width,h:o.height};let moved=false;
        el.setPointerCapture(event.pointerId);
        const move=e=>{const dx=(e.clientX-start.x)/scale,dy=(e.clientY-start.y)/scale;if(!moved&&Math.abs(dx)+Math.abs(dy)<3)return;moved=true;
          if(handle){Object.assign(o,resizedGeometry({x:start.ox,y:start.oy,width:start.w,height:start.h,rotation:o.rotation},handle,dx,dy,e.shiftKey));
          }else{o.x=start.ox+dx;o.y=start.oy+dy;}
          el.style.left=o.x+'px';el.style.top=o.y+'px';el.style.width=o.width+'px';el.style.height=o.height+'px';onSelect(o);
        };
        const done=()=>{el.removeEventListener('pointermove',move);el.removeEventListener('pointerup',done);el.removeEventListener('pointercancel',done);if(moved){o.changes.geometry=true;changed(`ppt-geometry-${o.id}`);}};
        el.addEventListener('pointermove',move);el.addEventListener('pointerup',done);el.addEventListener('pointercancel',done);
      });
    }
    if(selectedId&&slide.objects.some(o=>o.id===selectedId))select(selectedId);else{selectedId=null;onSelect(null);}
  }
  function update(property,value){
    const o=selected();if(!o||getMode()!=='edit')return;
    if(property==='rotation'&&o.canRotate===false)return;
    if(['x','y','width','height','rotation','fontSize','lineWidth'].includes(property)){value=Number(value);if(!Number.isFinite(value))return;value=clamp(value,property==='width'||property==='height'?1:property==='fontSize'?4:property==='lineWidth'?0:-20000,property==='fontSize'?500:20000);}
    o[property]=value;
    if(['x','y','width','height','rotation'].includes(property))o.changes.geometry=true;
    else if(['fontFamily','fontSize','bold','italic','align','verticalAlign'].includes(property))o.changes.font=true;
    else if(property==='color')o.changes.color=true;
    else if(['fill','fillVisible'].includes(property))o.changes.fill=true;
    else if(['lineColor','lineWidth'].includes(property))o.changes.line=true;
    if(property==='fill'){o.fillVisible=true;}
    finishEdit();repaint(o);changed(`ppt-style-${o.id}`);
  }
  function add(type,geometry='rect',image=''){
    const slide=current();if(!slide||getMode()!=='edit')return;
    const o=normalizePptObject({id:crypto.randomUUID(),name:type==='text'?'새 텍스트':type==='picture'?'새 이미지':'새 도형',type,x:slide.width*.12,y:slide.height*.3,width:slide.width*(type==='picture'?.4:.3),height:geometry==='line'?1:slide.height*(type==='text'?.18:.25),geometry,text:type==='text'?'텍스트를 입력하세요':'',fontSize:42,fontFamily:'맑은 고딕',fill:'#d7eef4',fillVisible:type==='shape',lineColor:'#197b92',lineWidth:type==='shape'?2:0,image,changes:{geometry:true,text:type==='text',image:type==='picture'}});
    slide.objects.push(o);selectedId=o.id;changed('ppt-add');redraw();
  }
  function remove(){const o=selected(),slide=current();if(!o||!slide||getMode()!=='edit')return;if(o.sourceId){slide.deletedObjects??=[];slide.deletedObjects.push(o.sourceId);}slide.objects=slide.objects.filter(item=>item.id!==o.id);selectedId=null;changed('ppt-delete');redraw();}
  function duplicate(){const o=selected(),slide=current();if(!o||!slide)return;const copy=clone(o);copy.id=crypto.randomUUID();copy.cloneSourceId=o.sourceId||o.cloneSourceId||null;copy.sourceId=null;copy.x+=24;copy.y+=24;copy.name+=' 복사';copy.changes.geometry=true;if(!copy.cloneSourceId&&copy.type==='graphic')copy.image=copy.preview;slide.objects.push(copy);selectedId=copy.id;changed('ppt-duplicate');redraw();}
  function order(direction){const o=selected(),slide=current();if(!o||!slide)return;slide.objects=slide.objects.filter(item=>item!==o);if(direction==='front')slide.objects.push(o);else slide.objects.unshift(o);slide.reorderObjects=true;changed('ppt-order');redraw();}
  function replaceImage(image){const o=selected();if(!o||o.type!=='picture')return;o.image=image;o.changes.image=true;repaint(o);changed('ppt-image');}
  function key(event){if(!selected()||getMode()!=='edit')return false;if(event.target?.isContentEditable||event.target?.matches?.('input,textarea,select'))return false;
    if(event.key==='Delete'||event.key==='Backspace'){event.preventDefault();remove();return true;}
    if(event.key==='Enter'){event.preventDefault();edit();return true;}
    if(event.key==='Escape'){event.preventDefault();clear();return true;}
    if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='d'){event.preventDefault();duplicate();return true;}
    if(!event.ctrlKey&&!event.metaKey&&['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)){event.preventDefault();const o=selected(),amount=event.shiftKey?10:1;update(event.key==='ArrowLeft'||event.key==='ArrowRight'?'x':'y',(event.key==='ArrowLeft'||event.key==='ArrowRight'?o.x:o.y)+(event.key==='ArrowLeft'||event.key==='ArrowUp'?-amount:amount));return true;}
    return false;
  }
  return{bind,selected,select,clear,edit,update,add,remove,duplicate,order,replaceImage,key,finishEdit};
}
