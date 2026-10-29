const fs=require('node:fs');
const path=require('node:path');
const {spawnSync}=require('node:child_process');
const root=path.join(__dirname,'..');
const files=['server.mjs','editor.js','electron/main.cjs','electron/preload.cjs','electron/updates.cjs','electron/smoke.cjs','electron/verify-feed.cjs','scripts/publish-release.cjs'];
for(const file of files){const result=spawnSync(process.execPath,['--check',path.join(root,file)],{stdio:'inherit',windowsHide:true});if(result.status!==0)process.exit(result.status||1);}
const packageJson=JSON.parse(fs.readFileSync(path.join(root,'package.json'),'utf8'));
if(packageJson.build.publish[0].repo!=='slide-studio')throw new Error('Wrong update repository');
for(const file of ['vendor/pdfjs/build/pdf.mjs','vendor/pdfjs/build/pdf.worker.mjs','assets/icon.ico','examples/demo.json','convert-powerpoint.ps1'])if(!fs.existsSync(path.join(root,file)))throw new Error('Missing '+file);
console.log('Source and release configuration checks passed.');
