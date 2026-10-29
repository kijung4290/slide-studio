const {test}=require('node:test');
const assert=require('node:assert/strict');
const {EventEmitter}=require('node:events');
const {UpdateController}=require('../electron/updates.cjs');
test('concurrent checks share one request and failures can be retried',async()=>{
  const updater=new EventEmitter();let checks=0,rejectCheck;
  updater.checkForUpdates=()=>{checks++;return new Promise((_resolve,reject)=>{rejectCheck=reject;});};
  const updates=new UpdateController(updater,{version:'1.0.0'});
  const first=updates.check(),second=updates.check();assert.equal(checks,1);
  rejectCheck(new Error('offline'));await Promise.all([first,second]);assert.equal(updates.getState().phase,'error');
  updater.checkForUpdates=async()=>{checks++;updater.emit('update-not-available');};await updates.check();assert.equal(checks,2);assert.equal(updates.getState().phase,'current');
});
test('a downloaded update remains pending until the user exits or chooses installation',()=>{
  const updater=new EventEmitter();let installations=0;updater.quitAndInstall=()=>installations++;
  const updates=new UpdateController(updater,{version:'1.0.0'});assert.equal(updates.install(),false);
  updater.emit('update-available',{version:'1.1.0'});updater.emit('download-progress',{percent:50.7});assert.equal(updates.getState().percent,51);
  updater.emit('update-downloaded',{version:'1.1.0'});assert.equal(installations,0);assert.equal(updater.autoInstallOnAppQuit,true);assert.equal(updater.allowDowngrade,false);assert.equal(updates.getState().phase,'ready');assert.equal(updates.install(),true);assert.equal(installations,1);
});
