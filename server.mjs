import http from 'node:http';
import { readFile, writeFile, mkdtemp, rm, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import os from 'node:os';
import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import vm from 'node:vm';
import {normalizePptObject,isImageData} from './powerpoint-model.js';

export function createSlideServer(options = {}) {
const root = options.root || path.dirname(fileURLToPath(import.meta.url));
const port = Number(options.port ?? process.env.SLIDE_STUDIO_PORT ?? 4318);
const conversionScript = options.conversionScript || path.join(root,'convert-powerpoint.ps1');
const exportScript = options.exportScript || path.join(root,'export-powerpoint.ps1');
const host = '127.0.0.1';
const maxUpload = 150 * 1024 * 1024;
const mime = {'.html':'text/html; charset=utf-8','.css':'text/css; charset=utf-8','.js':'text/javascript; charset=utf-8','.mjs':'text/javascript; charset=utf-8','.json':'application/json; charset=utf-8','.png':'image/png','.bcmap':'application/octet-stream','.ttf':'font/ttf','.wasm':'application/wasm'};
let converting = false;
function json(res, status, value) { res.writeHead(status, {'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'}); res.end(JSON.stringify(value)); }
async function body(req,limit=maxUpload) {
  if (Number(req.headers['content-length']) > limit) throw Object.assign(new Error('파일이 허용된 크기를 초과했습니다.'),{status:413});
  const chunks=[]; let size=0;
  for await (const chunk of req) { size+=chunk.length; if(size>limit) throw Object.assign(new Error('파일이 허용된 크기를 초과했습니다.'),{status:413}); chunks.push(chunk); }
  return Buffer.concat(chunks);
}
function runPowerPoint(script,args) {
  return new Promise((resolve,reject)=>{
    const child=spawn('powershell.exe',['-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',script,...args],{windowsHide:true,stdio:['ignore','ignore','pipe']});
    let errors=''; const timeout=setTimeout(()=>{child.kill();reject(new Error('PowerPoint 처리 시간이 초과되었습니다. 파일을 확인한 뒤 다시 시도하세요.'));},360000);
    child.stderr.on('data',d=>{if(errors.length<4000) errors+=d.toString();});
    child.on('error',e=>{clearTimeout(timeout);reject(e);});
    child.on('close',code=>{clearTimeout(timeout);code===0?resolve():reject(new Error('PowerPoint 처리에 실패했습니다. 파일이 암호로 보호되어 있지 않은지 확인하세요. '+errors.trim()));});
  });
}
async function reference() {
  if(options.publicDemo)return JSON.parse(await readFile(path.join(root,'examples','demo.json'),'utf8'));
  let html;
  try { html=await readFile(path.join(root,'ppt슬라이드.html'),'utf8'); }
  catch(error){if(error.code!=='ENOENT')throw error;return JSON.parse(await readFile(path.join(root,'examples','demo.json'),'utf8'));}
  const scripts=[...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(m=>m[1]);
  const context={MutationObserver:class{observe(){}},document:{body:{},addEventListener(){}}};context.window=context;
  vm.createContext(context);
  for(const code of scripts.filter(s=>s.startsWith('window.GYH=')&&s.includes('GYH.data=') || s.includes('kit.js — 공통') || s.includes('GYH.scenes.kj = ['))) vm.runInContext(code,context,{timeout:5000});
  const g=context.window.GYH, scenes=g.scenes.kj, notes=g.data.kj.slides;
  return {schemaVersion:1,id:randomUUID(),title:'한 사람의 하루 · 장애인복지',updatedAt:Date.now(),targetMinutes:5,referenceCss:[...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map(m=>m[1]).join('\n'),slides:scenes.map((s,i)=>({id:randomUUID(),title:notes[i]?.title||s.title,kind:'reference',width:1920,height:1080,html:s.html,notes:notes[i]?.blocks.map(b=>(b.segs||[]).map(x=>x.v||'').join('')).join('\n\n')||'',cue:'',duration:g.decks.kj.planHint[i]||30,overlays:[]}))};
}
const server=http.createServer(async(req,res)=>{
  try {
    const activePort=server.address()?.port||port,hostHeader=req.headers.host;
    if(![`${host}:${activePort}`,`localhost:${activePort}`].includes(hostHeader)) return json(res,403,{error:'로컬 주소에서만 사용할 수 있습니다.'});
    const url=new URL(req.url,`http://${host}:${activePort}`);
    if(req.headers.origin&&!['http://'+host+':'+activePort,'http://localhost:'+activePort].includes(req.headers.origin)) return json(res,403,{error:'다른 사이트의 요청은 허용하지 않습니다.'});
    if(req.method==='GET'&&url.pathname==='/api/health') return json(res,200,{ok:true,app:'slide-studio',platform:process.platform,powerPoint:process.platform==='win32'});
    if(req.method==='GET'&&url.pathname==='/api/reference') return json(res,200,await reference());
    if(req.method==='POST'&&url.pathname==='/api/import/powerpoint') {
      if(process.platform!=='win32') return json(res,400,{error:'PPT 변환에는 Windows의 PowerPoint가 필요합니다. PDF로 저장하여 가져오세요.'});
      if(converting) return json(res,409,{error:'다른 PPT를 변환 중입니다. 완료 후 다시 시도하세요.'});
      const extension=path.extname(decodeURIComponent(req.headers['x-file-name']||'')).toLowerCase();
      if(!['.ppt','.pptx'].includes(extension)) return json(res,400,{error:'PPT 또는 PPTX 파일을 선택하세요.'});
      converting=true; let temp;
      try {
        const data=await body(req); if(data.length<8) throw new Error('파일 내용이 비어 있습니다.');
        const isZip=data[0]===0x50&&data[1]===0x4b, isOle=data.subarray(0,4).equals(Buffer.from([0xd0,0xcf,0x11,0xe0]));
        if(!(isZip||isOle)) throw new Error('PowerPoint 파일 형식이 아닙니다.');
        temp=await mkdtemp(path.join(os.tmpdir(),'slide-studio-')); const source=path.join(temp,'source'+extension);await writeFile(source,data);await runPowerPoint(conversionScript,['-SourcePath',source,'-OutputPath',temp]);
        const result=JSON.parse(await readFile(path.join(temp,'result.json'),'utf8'));
        const imageData=async file=>'data:image/png;base64,'+(await readFile(path.join(temp,path.basename(file)))).toString('base64');
        for(const slide of result.slides){slide.image=await imageData(slide.image);slide.background=await imageData(slide.background);for(const object of slide.objects||[]){if(object.preview)object.preview=await imageData(object.preview);if(object.backdrop)object.backdrop=await imageData(object.backdrop);}}
        result.pptx=(await readFile(path.join(temp,'original.pptx'))).toString('base64');
        json(res,200,result);
      } finally {converting=false;if(temp){const resolved=path.resolve(temp),prefix=path.resolve(os.tmpdir())+path.sep+'slide-studio-';if(!resolved.startsWith(prefix))throw new Error('임시 파일 경로를 확인할 수 없습니다.');await rm(resolved,{recursive:true,force:true});}}
      return;
    }
    if(req.method==='POST'&&url.pathname==='/api/export/powerpoint'){
      if(process.platform!=='win32')return json(res,400,{error:'PPTX 저장에는 Windows의 PowerPoint가 필요합니다.'});
      if(converting)return json(res,409,{error:'PowerPoint에서 다른 자료를 처리 중입니다.'});
      converting=true;let temp;
      try{
        const input=JSON.parse((await body(req,512*1024*1024)).toString('utf8'));
        if(!Array.isArray(input.slides)||!input.slides.length||input.slides.length>200)throw new Error('저장할 슬라이드를 확인하세요.');
        temp=await mkdtemp(path.join(os.tmpdir(),'slide-studio-'));
        let sourcePath=null;
        if(input.source){const data=Buffer.from(input.source,'base64');if(data.length>maxUpload||data[0]!==0x50||data[1]!==0x4b)throw new Error('원본 PPTX 데이터가 올바르지 않습니다.');sourcePath=path.join(temp,'original.pptx');await writeFile(sourcePath,data);}
        const slides=input.slides.map(slide=>{if(!Number.isFinite(slide.width)||!Number.isFinite(slide.height)||slide.width<100||slide.height<100||slide.width>10000||slide.height>10000)throw new Error('슬라이드 크기가 올바르지 않습니다.');return{sourceSlideId:slide.sourceSlideId?Number(slide.sourceSlideId):null,title:String(slide.title||''),width:slide.width,height:slide.height,notes:String(slide.notes||''),backgroundColor:/^#[0-9a-f]{6}$/i.test(slide.backgroundColor||'')?slide.backgroundColor:'#ffffff',reorderObjects:slide.reorderObjects===true,deletedObjects:(Array.isArray(slide.deletedObjects)?slide.deletedObjects:[]).map(Number).filter(Number.isFinite),objects:(Array.isArray(slide.objects)?slide.objects:[]).slice(0,2000).map(normalizePptObject),backgroundImage:isImageData(slide.backgroundImage)?slide.backgroundImage:null};});
        const planPath=path.join(temp,'plan.json'),output=path.join(temp,'edited.pptx');await writeFile(planPath,JSON.stringify({sourcePath,slides}));
        await runPowerPoint(exportScript,['-PlanPath',planPath,'-OutputPath',output]);
        res.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.presentationml.presentation','Content-Disposition':'attachment; filename="edited.pptx"','Cache-Control':'no-store'});res.end(await readFile(output));
      }finally{converting=false;if(temp){const resolved=path.resolve(temp);if(!resolved.startsWith(path.resolve(os.tmpdir())+path.sep+'slide-studio-'))throw new Error('임시 파일 경로 오류');await rm(resolved,{recursive:true,force:true});}}
      return;
    }
    if(req.method!=='GET'&&req.method!=='HEAD') return json(res,405,{error:'허용하지 않는 요청입니다.'});
    let relative=decodeURIComponent(url.pathname);if(relative==='/')relative='/슬라이드편집기.html';
    const target=path.resolve(root,'.'+relative);
    if(!target.startsWith(root+path.sep)) return json(res,403,{error:'접근할 수 없습니다.'});
    const allowed=['슬라이드편집기.html','editor.js','editor.css','launch-hint.js','powerpoint-model.js','powerpoint-editor.js'];
    if(!allowed.includes(path.basename(target))&&!target.startsWith(path.join(root,'vendor')+path.sep)) return json(res,404,{error:'파일이 없습니다.'});
    const info=await stat(target);if(!info.isFile())return json(res,404,{error:'파일이 없습니다.'});
    res.writeHead(200,{'Content-Type':mime[path.extname(target)]||'application/octet-stream','Cache-Control':'no-cache','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer','Content-Security-Policy':"default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; worker-src 'self' blob:; frame-src 'self' about:; object-src 'none'; base-uri 'none'"});res.end(req.method==='HEAD'?undefined:await readFile(target));
  } catch(error){if(!res.headersSent)json(res,error.status||500,{error:error.code==='ENOENT'?'필요한 파일을 찾을 수 없습니다.':error.message});else res.end();}
});
return {server,get url(){return`http://${host}:${server.address()?.port||port}`;},listen:()=>new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,host,()=>{server.removeListener('error',reject);resolve({server,url:`http://${host}:${server.address().port}`});});}),close:()=>new Promise(resolve=>server.close(resolve))};
}

if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  const service=createSlideServer();
  service.listen().then(({url})=>{console.log(`발표 작업실: ${url}`);console.log('이 창을 열어 두세요. 종료하려면 Ctrl+C를 누르세요.');if(process.argv.includes('--open'))spawn('cmd.exe',['/c','start','',url],{windowsHide:true,stdio:'ignore'});}).catch(async error=>{if(error.code==='EADDRINUSE'&&process.argv.includes('--open')){try{const state=await(await fetch(service.url+'/api/health')).json();if(state.app==='slide-studio'){spawn('cmd.exe',['/c','start','',service.url],{windowsHide:true,stdio:'ignore'});return;}}catch{}}console.error(error.code==='EADDRINUSE'?'이미 실행 중이거나 포트가 사용 중입니다.':error.message);process.exitCode=1;});
}
