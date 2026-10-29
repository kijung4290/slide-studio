const fs=require('node:fs/promises');
const path=require('node:path');
const {NsisUpdater}=require('electron-updater');
const {SemVer}=require('semver');
module.exports=async({app,qaDir})=>{
  await fs.mkdir(qaDir,{recursive:true});
  const updater=new NsisUpdater({provider:'github',owner:'kijung4290',repo:'slide-studio'});
  updater.autoDownload=false;updater.autoInstallOnAppQuit=false;updater.forceDevUpdateConfig=true;updater.currentVersion=new SemVer('0.0.1');
  updater.on('error',()=>{});
  try{const result=await updater.checkForUpdates();if(!result?.updateInfo?.files?.length)throw new Error('No update installer found');
    const downloaded=await updater.downloadUpdate();
    const report={releaseVersion:result.updateInfo.version,files:result.updateInfo.files.map(file=>({url:file.url,sha512:file.sha512,size:file.size})),downloaded,verified:true};
    await fs.writeFile(path.join(qaDir,'update-feed.json'),JSON.stringify(report,null,2));console.log('Verified GitHub update metadata and downloaded installer: '+report.releaseVersion);app.exit(0);
  }catch(error){await fs.writeFile(path.join(qaDir,'update-feed.json'),JSON.stringify({verified:false,error:error.message},null,2));console.error(error);app.exit(1);}
};
