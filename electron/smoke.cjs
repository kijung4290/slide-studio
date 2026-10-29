const fs=require('node:fs/promises');
const path=require('node:path');
const assert=require('node:assert/strict');
const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms));
module.exports=async({app,mainWindow,service,qaDir})=>{
  await fs.mkdir(qaDir,{recursive:true});const results=[];
  const js=source=>mainWindow.webContents.executeJavaScript(source,true);
  const until=async(expression)=>{const end=Date.now()+60000;while(Date.now()<end){if(await js(expression))return;await wait(100);}throw new Error('Timed out: '+expression);};
  const pass=(name,value)=>{assert.ok(value,name);results.push({name,passed:true});console.log('PASS '+name);};
  try{
    await until("document.getElementById('desktop-version')?.textContent.includes('설치형')");
    pass('sandboxed renderer with desktop bridge',await js("typeof require==='undefined'&&typeof window.slideDesktop.info==='function'"));
    const info=await js('window.slideDesktop.info()');pass('desktop version and persistent data location',info.version===app.getVersion()&&info.userData.length>0);
    await js("document.querySelector('[data-action=reference]').click()");
    await until("document.querySelectorAll('.thumbnail').length===3&&document.getElementById('busy-overlay').hidden");
    pass('public demo excludes personal presentation',await js("document.getElementById('project-title').value==='발표 작업실 사용 예제'"));
    await until("document.getElementById('slide-frame').contentDocument?.querySelector('.title')?.isContentEditable");
    await js("{const el=document.getElementById('slide-frame').contentDocument.querySelector('.title');el.textContent='설치형 슬라이드 편집 테스트';el.dispatchEvent(new Event('input',{bubbles:true}));const notes=document.getElementById('notes-editor');notes.value='설치 후에도 유지할 대본입니다.';notes.dispatchEvent(new Event('input',{bubbles:true}));}");
    await until("document.getElementById('save-status').textContent==='이 브라우저에 자동 저장됨'");
    pass('native slide and speaker notes save',await js("document.getElementById('slide-frame').contentDocument.querySelector('.title').textContent.includes('설치형')"));
    const rendered=await js("window.slideDesktop.renderSlide({html:document.getElementById('slide-frame').srcdoc,width:1600,height:900}).then(bytes=>Array.from(bytes))");pass('legacy HTML slides render for PPTX export',rendered.length>5000);await fs.writeFile(path.join(qaDir,'legacy-render.png'),Buffer.from(rendered));
    await mainWindow.reload();await until("document.querySelectorAll('.thumbnail').length===3&&document.getElementById('notes-editor').value.includes('유지할 대본')");
    pass('saved notes restore after renderer reload',true);
    const newWindow=new Promise(resolve=>mainWindow.webContents.once('did-create-window',resolve));await js("document.querySelector('[data-action=present]').click()");const stage=await newWindow;
    const deadline=Date.now()+15000;let stageReady=false;while(Date.now()<deadline){stageReady=await stage.webContents.executeJavaScript("!!document.getElementById('audience-frame')?.contentDocument?.querySelector('.title')");if(stageReady)break;await wait(100);}pass('separate audience window receives current slide',stageReady);
    pass('audience window excludes notes from visible UI',!(await stage.webContents.executeJavaScript('document.body.innerText')).includes('유지할 대본'));
    await js("document.querySelector('[data-action=next]').click()");await wait(250);pass('audience follows presenter navigation',await stage.webContents.executeJavaScript("document.getElementById('audience-frame').contentDocument.querySelector('.title').textContent.includes('내 자료')"));stage.destroy();
    const fixture=path.join(qaDir,'sample.pdf');
    try{const data=await fs.readFile(fixture);await js(`{const file=new File([new Uint8Array(${JSON.stringify([...data])})],'sample.pdf',{type:'application/pdf'});const dt=new DataTransfer();dt.items.add(file);const input=document.getElementById('file-input');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));}`);await until("document.getElementById('busy-overlay').hidden&&document.getElementById('slide-frame').contentDocument?.querySelector('.slide-image')?.naturalWidth>0");pass('PDF imports and renders inside packaged Chromium',await js("document.getElementById('notes-editor').value===''") );}catch(error){if(error.code!=='ENOENT')throw error;results.push({name:'PDF fixture',skipped:'No fixture in CI'});}
    const pptFixture=path.join(qaDir,'sample.pptx');
    try{
      const data=await fs.readFile(pptFixture);await js(`{const file=new File([new Uint8Array(${JSON.stringify([...data])})],'sample.pptx');const dt=new DataTransfer();dt.items.add(file);const input=document.getElementById('file-input');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));}`);
      await until("document.getElementById('busy-overlay').hidden&&document.getElementById('slide-frame').contentDocument?.querySelectorAll('.ppt-object').length>=2");
      pass('packaged editable PowerPoint importer and Korean notes',await js("document.getElementById('notes-editor').value.includes('첫 번째 대본')"));
      await js("{const select=document.getElementById('ppt-object-list');select.value=select.options[1].value;select.dispatchEvent(new Event('change',{bubbles:true}));document.getElementById('ppt-edit-text').click();const editor=document.getElementById('slide-frame').contentDocument.querySelector('.is-editing [data-text-editor]');editor.innerText='설치형 앱에서 수정한 PPT 제목';editor.dispatchEvent(new Event('input',{bubbles:true}));editor.blur();}");
      await until("document.getElementById('save-status').textContent==='이 브라우저에 자동 저장됨'");
      const output=path.join(qaDir,'desktop-edited.pptx');const downloaded=new Promise((resolve,reject)=>{const timeout=setTimeout(()=>reject(new Error('PPTX download timeout')),60000);mainWindow.webContents.session.once('will-download',(_event,item)=>{item.setSavePath(output);item.once('done',(_e,state)=>{clearTimeout(timeout);state==='completed'?resolve():reject(new Error('PPTX download '+state));});});});
      await js("document.querySelector('[data-action=save-pptx]').click()");await downloaded;
      const response=await fetch(service.url+'/api/import/powerpoint',{method:'POST',headers:{'X-File-Name':'edited.pptx'},body:await fs.readFile(output)});const saved=await response.json();pass('desktop PPTX save produces editable text in PowerPoint',response.ok&&saved.slides[0].objects.some(o=>o.text==='설치형 앱에서 수정한 PPT 제목'));
    }catch(error){if(error.code!=='ENOENT')throw error;results.push({name:'PowerPoint fixture',skipped:'Requires local PowerPoint'});}
    await fs.writeFile(path.join(qaDir,'desktop-smoke.png'),(await mainWindow.webContents.capturePage()).toPNG());
    await fs.writeFile(path.join(qaDir,'smoke-results.json'),JSON.stringify({version:app.getVersion(),packaged:app.isPackaged,results},null,2));app.exit(0);
  }catch(error){await fs.writeFile(path.join(qaDir,'smoke-results.json'),JSON.stringify({version:app.getVersion(),packaged:app.isPackaged,results,error:error.stack},null,2));console.error(error);app.exit(1);}
};
