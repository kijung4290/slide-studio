const {contextBridge,ipcRenderer}=require('electron');

if(process.isMainFrame){
  contextBridge.exposeInMainWorld('slideDesktop',Object.freeze({
    info:()=>ipcRenderer.invoke('slide:info'),
    checkForUpdates:()=>ipcRenderer.invoke('slide:check-update'),
    installUpdate:()=>ipcRenderer.invoke('slide:install-update'),
    renderSlide:options=>ipcRenderer.invoke('slide:render-image',options),
    onUpdate:callback=>{if(typeof callback!=='function')return;const handler=(_event,state)=>callback(state);ipcRenderer.on('slide:update',handler);return()=>ipcRenderer.removeListener('slide:update',handler);},
    onPrepareClose:callback=>{if(typeof callback!=='function')return;const handler=async(_event,requestId)=>{let saved=false;try{saved=(await callback())===true;}catch{}ipcRenderer.send('slide:close-ready',{requestId,saved});};ipcRenderer.on('slide:prepare-close',handler);return()=>ipcRenderer.removeListener('slide:prepare-close',handler);}
  }));
}
