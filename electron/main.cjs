const {app,BrowserWindow,Menu,dialog,ipcMain,shell,session}=require('electron');
const path=require('node:path');
const os=require('node:os');
const fs=require('node:fs');
const {pathToFileURL}=require('node:url');
const {randomUUID}=require('node:crypto');
const {autoUpdater}=require('electron-updater');
const {UpdateController}=require('./updates.cjs');

const smoke=process.argv.includes('--smoke-test');
const verifyFeed=process.argv.includes('--verify-update-feed');
const qa=smoke||verifyFeed;
if(qa){app.setPath('userData',path.join(os.tmpdir(),'slide-studio-qa-'+randomUUID()));}
else{app.setPath('userData',path.join(app.getPath('appData'),'SlideStudio'));}
app.setAppUserModelId('kr.parkjuim.slidestudio');
const hasLock=app.requestSingleInstanceLock();
if(!hasLock){app.quit();}
let mainWindow,service,updates,allowClose=false,closeRequest=null,updateInterval;
const qaDir=process.env.SLIDE_STUDIO_QA_DIR||path.join(process.cwd(),'.qa');
function log(error){try{fs.mkdirSync(app.getPath('userData'),{recursive:true});fs.appendFileSync(path.join(app.getPath('userData'),'app.log'),new Date().toISOString()+' '+String(error?.stack||error)+'\n');}catch{}}
function trusted(event){return mainWindow&&!mainWindow.isDestroyed()&&event.sender===mainWindow.webContents&&event.senderFrame===mainWindow.webContents.mainFrame&&new URL(event.senderFrame.url).origin===service.url;}
function securePreferences(){return{nodeIntegration:false,contextIsolation:true,sandbox:true,webSecurity:true,preload:path.join(__dirname,'preload.cjs'),devTools:!app.isPackaged||qa};}
function guardWindow(window){
  window.webContents.on('will-navigate',(event,url)=>{try{const dest=new URL(url);if(dest.origin!==service.url||!['/','/슬라이드편집기.html'].includes(decodeURIComponent(dest.pathname)))event.preventDefault();}catch{event.preventDefault();}});
  window.webContents.setWindowOpenHandler(({url})=>{try{const target=new URL(url);if(target.origin===service.url&&target.searchParams.get('stage')==='1'&&['/','/슬라이드편집기.html'].includes(decodeURIComponent(target.pathname)))return{action:'allow',overrideBrowserWindowOptions:{title:'발표화면 · 발표 작업실',autoHideMenuBar:true,backgroundColor:'#000000',icon:path.join(app.getAppPath(),'assets','icon.ico'),webPreferences:securePreferences()}};}catch{}return{action:'deny'};});
  window.webContents.on('did-create-window',child=>{guardWindow(child);child.setMenuBarVisibility(false);});
  window.webContents.on('render-process-gone',(_event,details)=>log('Renderer stopped: '+JSON.stringify(details)));
}
function createMenu(){Menu.setApplicationMenu(Menu.buildFromTemplate([
  {label:'파일',submenu:[{label:'프로젝트 저장',accelerator:'CmdOrCtrl+S',click:()=>mainWindow?.webContents.executeJavaScript("document.querySelector('[data-action=save]').click()")},{label:'PPTX로 저장',accelerator:'CmdOrCtrl+Shift+S',click:()=>mainWindow?.webContents.executeJavaScript("document.querySelector('[data-action=save-pptx]').click()")},{type:'separator'},{label:'종료',accelerator:'Alt+F4',click:()=>mainWindow?.close()}]},
  {label:'편집',submenu:[{role:'undo',label:'실행 취소'},{role:'redo',label:'다시 실행'},{type:'separator'},{role:'cut',label:'잘라내기'},{role:'copy',label:'복사'},{role:'paste',label:'붙여넣기'},{role:'selectAll',label:'모두 선택'}]},
  {label:'보기',submenu:[{role:'togglefullscreen',label:'전체화면'},{role:'resetZoom',label:'기본 배율'},{role:'zoomIn',label:'확대'},{role:'zoomOut',label:'축소'}]},
  {label:'도움말',submenu:[{label:'사용 방법',click:()=>mainWindow?.webContents.executeJavaScript("document.querySelector('[data-action=help]').click()")},{label:'업데이트 확인',click:()=>updates?.check()},{label:'GitHub 릴리스',click:()=>shell.openExternal('https://github.com/kijung4290/slide-studio/releases')},{type:'separator'},{label:'발표 작업실 정보',click:()=>dialog.showMessageBox(mainWindow,{type:'info',title:'발표 작업실',message:`발표 작업실 v${app.getVersion()}`,detail:'PPT·PDF 슬라이드와 발표 대본을 함께 준비하는 Windows 프로그램\n제작: kijung4290'})}]}
]));}
function installUpdate(){if(updates.getState().phase!=='ready')return false;allowClose=true;return updates.install();}
function wireIpc(){
  ipcMain.handle('slide:render-image',async(event,value)=>{
    if(!trusted(event))throw new Error('허용하지 않는 요청입니다.');
    const width=Math.round(Number(value?.width)),height=Math.round(Number(value?.height)),html=value?.html;
    if(width<100||height<100||width>5000||height>5000||!Number.isFinite(width)||!Number.isFinite(height)||typeof html!=='string'||html.length>25e6)throw new Error('슬라이드 크기 또는 내용이 올바르지 않습니다.');
    const captureSession=session.fromPartition('slide-capture-'+randomUUID());
    captureSession.setPermissionRequestHandler((_contents,_permission,callback)=>callback(false));
    captureSession.webRequest.onBeforeRequest((details,callback)=>callback({cancel:!details.url.startsWith('data:')&&!details.url.startsWith('about:')&&!details.url.startsWith('blob:')}));
    const captureWindow=new BrowserWindow({width,height,useContentSize:true,frame:false,show:false,backgroundColor:'#ffffff',webPreferences:{session:captureSession,nodeIntegration:false,contextIsolation:true,sandbox:true,webSecurity:true,backgroundThrottling:false}});
    captureWindow.webContents.setWindowOpenHandler(()=>({action:'deny'}));
    try{
      await captureWindow.loadURL('about:blank');
      const meta='<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'none\'; style-src \'unsafe-inline\'; img-src data: blob:; font-src data:">';
      const safeHtml=/<head[\s>]/i.test(html)?html.replace(/<head([^>]*)>/i,'<head$1>'+meta):'<!doctype html><html><head>'+meta+'</head><body>'+html+'</body></html>';
      await captureWindow.webContents.executeJavaScript('document.open();document.write('+JSON.stringify(safeHtml)+');document.close();');
      let timeout;try{await Promise.race([captureWindow.webContents.executeJavaScript('document.fonts.ready.then(()=>Promise.all([...document.images].map(img=>img.complete?Promise.resolve():new Promise(resolve=>{img.onload=resolve;img.onerror=resolve}))))'),new Promise((_resolve,reject)=>{timeout=setTimeout(()=>reject(new Error('슬라이드 이미지 생성 시간이 초과되었습니다.')),15000);})]);}finally{clearTimeout(timeout);}
      return(await captureWindow.webContents.capturePage()).toPNG();
    }finally{captureWindow.destroy();}
  });
  ipcMain.handle('slide:info',event=>{if(!trusted(event))throw new Error('허용하지 않는 요청입니다.');return{version:app.getVersion(),updates:updates.getState(),userData:app.getPath('userData')};});
  ipcMain.handle('slide:check-update',event=>{if(!trusted(event))throw new Error('허용하지 않는 요청입니다.');return updates.check();});
  ipcMain.handle('slide:install-update',event=>{if(!trusted(event))throw new Error('허용하지 않는 요청입니다.');return installUpdate();});
  ipcMain.on('slide:close-ready',async(event,value)=>{
    if(!trusted(event)||!closeRequest||value?.requestId!==closeRequest)return;closeRequest=null;
    if(value.saved===true){allowClose=true;mainWindow.close();}
    else await dialog.showMessageBox(mainWindow,{type:'error',title:'작업을 저장해 주세요',message:'자동 저장이 완료되지 않았습니다.',detail:'프로젝트 저장으로 파일을 백업한 뒤 다시 종료해 주세요.'});
  });
}
async function createMainWindow(){
  mainWindow=new BrowserWindow({title:'발표 작업실',width:1440,height:940,minWidth:960,minHeight:700,show:!qa,backgroundColor:'#edf3f7',autoHideMenuBar:false,icon:path.join(app.getAppPath(),'assets','icon.ico'),webPreferences:securePreferences()});
  guardWindow(mainWindow);createMenu();
  mainWindow.on('close',event=>{if(allowClose||qa)return;event.preventDefault();if(!closeRequest){closeRequest=randomUUID();mainWindow.webContents.send('slide:prepare-close',closeRequest);}});
  mainWindow.on('closed',()=>{mainWindow=null;for(const window of BrowserWindow.getAllWindows())window.destroy();});
  await mainWindow.loadURL(service.url);
  mainWindow.webContents.on('did-finish-load',()=>mainWindow?.webContents.send('slide:update',updates.getState()));
  if(!qa){mainWindow.show();setTimeout(()=>updates.check(),8000).unref();updateInterval=setInterval(()=>updates.check(),4*60*60*1000);updateInterval.unref();}
}
app.on('second-instance',()=>{if(mainWindow){if(mainWindow.isMinimized())mainWindow.restore();mainWindow.show();mainWindow.focus();}});
app.on('window-all-closed',()=>app.quit());
app.on('will-quit',()=>{clearInterval(updateInterval);service?.server.close();});
if(hasLock)app.whenReady().then(async()=>{
  session.defaultSession.setPermissionRequestHandler((_contents,permission,callback)=>callback(permission==='fullscreen'));
  session.defaultSession.setPermissionCheckHandler((_contents,permission)=>permission==='fullscreen');
  const {createSlideServer}=await import(pathToFileURL(path.join(app.getAppPath(),'server.mjs')).href);
  service=createSlideServer({port:qa?14319:14318,root:app.getAppPath(),publicDemo:true,conversionScript:app.isPackaged?path.join(process.resourcesPath,'convert-powerpoint.ps1'):path.join(app.getAppPath(),'convert-powerpoint.ps1'),exportScript:app.isPackaged?path.join(process.resourcesPath,'export-powerpoint.ps1'):path.join(app.getAppPath(),'export-powerpoint.ps1')});
  await service.listen();
  updates=new UpdateController(autoUpdater,{version:app.getVersion(),enabled:app.isPackaged&&!qa,log});
  updates.on('state',state=>{if(mainWindow&&!mainWindow.isDestroyed())mainWindow.webContents.send('slide:update',state);});
  wireIpc();await createMainWindow();
  if(smoke)await require('./smoke.cjs')({app,mainWindow,service,qaDir,log});
  else if(verifyFeed)await require('./verify-feed.cjs')({app,qaDir,log});
}).catch(error=>{log(error);if(qa){fs.mkdirSync(qaDir,{recursive:true});fs.writeFileSync(path.join(qaDir,'startup-error.json'),JSON.stringify({error:error.message,stack:error.stack},null,2));console.error(error);app.exit(1);}else{dialog.showErrorBox('발표 작업실을 시작하지 못했습니다.',error.code==='EADDRINUSE'?'앱 전용 포트(14318)가 사용 중입니다. 다른 발표 작업실이 실행 중인지 확인하세요.':error.message);app.quit();}});
