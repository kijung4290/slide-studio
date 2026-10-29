const {EventEmitter}=require('node:events');

class UpdateController extends EventEmitter {
  constructor(updater,{version,enabled=true,log=()=>{}}={}){
    super();this.updater=updater;this.enabled=enabled;this.log=log;this.pending=null;
    this.state={phase:enabled?'idle':'development',currentVersion:version,version:null,percent:0,message:enabled?'시작할 때 새 버전을 자동으로 확인합니다.':'개발 실행에서는 자동 업데이트를 사용하지 않습니다.'};
    updater.autoDownload=true;updater.autoInstallOnAppQuit=true;updater.allowDowngrade=false;updater.allowPrerelease=false;
    updater.on('checking-for-update',()=>this.set({phase:'checking',message:'새 버전을 확인하고 있습니다.'}));
    updater.on('update-available',info=>this.set({phase:'downloading',version:info.version,percent:0,message:`v${info.version} 업데이트를 다운로드하고 있습니다.`}));
    updater.on('download-progress',progress=>this.set({phase:'downloading',percent:Math.round(progress.percent),message:`업데이트 다운로드 ${Math.round(progress.percent)}%`}));
    updater.on('update-not-available',()=>this.set({phase:'current',version:null,message:'최신 버전을 사용하고 있습니다.'}));
    updater.on('update-downloaded',info=>this.set({phase:'ready',version:info.version,percent:100,message:`v${info.version} 준비 완료 · 앱을 종료하면 업데이트합니다.`}));
    updater.on('error',error=>{this.log(error);this.set({phase:'error',message:'업데이트를 확인하지 못했습니다. 인터넷 연결 후 다시 시도하세요.'});});
  }
  set(patch){this.state={...this.state,...patch};this.emit('state',this.state);}
  getState(){return{...this.state};}
  async check(){
    if(!this.enabled||this.state.phase==='ready'||this.state.phase==='downloading')return this.getState();
    if(!this.pending){this.pending=this.updater.checkForUpdates().catch(error=>{this.log(error);this.set({phase:'error',message:'업데이트를 확인하지 못했습니다. 인터넷 연결 후 다시 시도하세요.'});}).finally(()=>{this.pending=null;});}
    await this.pending;return this.getState();
  }
  install(){if(!this.enabled||this.state.phase!=='ready')return false;this.updater.quitAndInstall(false,true);return true;}
}
module.exports={UpdateController};
