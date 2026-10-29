const {test}=require('node:test');const assert=require('node:assert/strict');const path=require('node:path');const {pathToFileURL}=require('node:url');
const model=()=>import(pathToFileURL(path.join(__dirname,'..','powerpoint-model.js')).href);
test('resizing a rotated object preserves its opposite corner and optional aspect ratio',async()=>{
 const{resizedGeometry}=await model(),start={x:120,y:180,width:300,height:120,rotation:35};
 const corner=o=>{const a=o.rotation*Math.PI/180;return{x:o.x+o.width/2-o.width/2*Math.cos(a)+o.height/2*Math.sin(a),y:o.y+o.height/2-o.width/2*Math.sin(a)-o.height/2*Math.cos(a)}};
 const end={...resizedGeometry(start,'se',80,40,true),rotation:start.rotation},a=corner(start),b=corner(end);
 assert.ok(Math.abs(a.x-b.x)<1e-7&&Math.abs(a.y-b.y)<1e-7);assert.ok(Math.abs(end.width/end.height-start.width/start.height)<1e-7);
});
test('slide markup escapes text and omits notes and original PPT bytes',async()=>{
 const{normalizePptObject,powerpointMarkup}=await model();const object=normalizePptObject({type:'text',text:'<script>alert(1)</script>',preview:'javascript:alert(1)',changes:{text:true}});
 const markup=powerpointMarkup({objects:[object],notes:'private speaker notes',source:'original PPT bytes'});
 assert.ok(markup.includes('&lt;script&gt;'));assert.ok(!markup.includes('<script>'));assert.ok(!markup.includes('javascript:'));assert.ok(!markup.includes('private speaker notes'));assert.ok(!markup.includes('original PPT bytes'));
});
