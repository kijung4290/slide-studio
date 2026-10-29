const {test}=require('node:test');
const assert=require('node:assert/strict');
const path=require('node:path');
const {pathToFileURL}=require('node:url');
test('local server serves packaged assets and rejects cross-origin conversion requests',async()=>{
  const {createSlideServer}=await import(pathToFileURL(path.join(__dirname,'..','server.mjs')).href);
  const service=createSlideServer({port:14329});await service.listen();
  try{const page=await fetch(service.url);assert.equal(page.status,200);assert.ok(page.headers.get('content-security-policy').includes("script-src 'self'"));assert.ok((await page.text()).includes('발표 작업실'));
    const blocked=await fetch(service.url+'/api/import/powerpoint',{method:'POST',headers:{Origin:'https://example.com'}});assert.equal(blocked.status,403);
    const secret=await fetch(service.url+'/package.json');assert.equal(secret.status,404);
  }finally{await service.close();}
});
