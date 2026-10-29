const fs=require('node:fs');
const path=require('node:path');
const {Resvg}=require('@resvg/resvg-js');
const root=path.join(__dirname,'..');
const png=new Resvg(fs.readFileSync(path.join(root,'assets','icon.svg'))).render().asPng();
fs.writeFileSync(path.join(root,'assets','icon.png'),png);
const header=Buffer.alloc(22);header.writeUInt16LE(1,2);header.writeUInt16LE(1,4);header.writeUInt16LE(1,10);header.writeUInt16LE(32,12);header.writeUInt32LE(png.length,14);header.writeUInt32LE(22,18);
fs.writeFileSync(path.join(root,'assets','icon.ico'),Buffer.concat([header,png]));
console.log('Application icon generated.');
