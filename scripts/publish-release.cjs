const fs=require('node:fs');
const path=require('node:path');
const {spawnSync}=require('node:child_process');
const root=path.join(__dirname,'..');
const config=JSON.parse(fs.readFileSync(path.join(root,'package.json'),'utf8'));
const repo='kijung4290/slide-studio',tag='v'+config.version;
function run(command,args,options={}){const result=spawnSync(command,args,{cwd:root,encoding:'utf8',windowsHide:true,...options});if(result.status!==0)throw new Error((result.stderr||result.stdout||command+' failed').trim());return result.stdout?.trim();}
const head=run('git',['rev-parse','HEAD']);
if(run('git',['rev-parse',tag+'^{commit}'])!==head)throw new Error('Release tag must point to the current commit.');
run(process.execPath,[require.resolve('electron-builder/cli.js'),'--win','nsis','--x64','--publish','never'],{stdio:'inherit'});
const query=spawnSync('gh',['release','view',tag,'--repo',repo,'--json','id'],{cwd:root,encoding:'utf8',windowsHide:true});
if(query.status!==0)run('gh',['release','create',tag,'--repo',repo,'--verify-tag','--title',config.version,'--notes-file','RELEASE_NOTES.md','--draft']);
const installer=`SlideStudio-Setup-${config.version}-x64.exe`;
const assets=[installer,installer+'.blockmap','latest.yml'].map(name=>path.join(root,'release',name));
for(const file of assets)if(!fs.existsSync(file))throw new Error('Missing release asset: '+file);
run('gh',['release','upload',tag,...assets,'--repo',repo,'--clobber']);
run('gh',['release','edit',tag,'--repo',repo,'--draft=false','--latest']);
console.log('Published installer and update feed together: https://github.com/'+repo+'/releases/tag/'+tag);
