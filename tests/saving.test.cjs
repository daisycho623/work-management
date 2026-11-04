const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
function createPage(api){
  const source=fs.readFileSync(path.join(__dirname,'../app.js'),'utf8');
  const context=vm.createContext({api,JSON,Promise,clearTimeout,showSync(){},writeCache(){},tasks:[{id:'a',hours:{day:1}}],LEAVE_CALENDAR_ID:'calendar',leaveDays:{},customHolidays:{},saveTimer:null});
  vm.runInContext(source.slice(source.indexOf('let saveChain='),source.indexOf('function syncFixedColumns')),context);
  return context;
}
test('input and change with identical data share one pending and completed save',async()=>{
  let finish,calls=0;const page=createPage(async()=>{calls++;await new Promise(resolve=>finish=resolve)});
  const first=page.saveAll();await Promise.resolve();assert.equal(page.saveAll(),first);assert.equal(calls,1);
  finish();assert.equal(await first,true);assert.equal(await page.saveAll(),true);assert.equal(calls,1);
});
test('distinct edits are sent serially with their original snapshots',async()=>{
  let finish;const calls=[];const page=createPage(async(_,payload)=>{calls.push(payload);if(calls.length===1)await new Promise(resolve=>finish=resolve)});
  const first=page.saveAll();await Promise.resolve();page.tasks[0].hours.day=2;const second=page.saveAll();await Promise.resolve();
  assert.equal(calls.length,1);assert.equal(calls[0].tasks[0].hours.day,1);finish();await first;await second;
  assert.equal(calls[1].tasks[0].hours.day,2);
});
test('failed saves allow retrying identical data',async()=>{
  let calls=0;const page=createPage(async()=>{if(++calls===1)throw Error('busy')});
  assert.equal(await page.saveAll(),false);assert.equal(await page.saveAll(),true);assert.equal(calls,2);
});
test('saves keep a pending local copy until the sheet confirms the latest snapshot',async()=>{
  let calls=0;const cacheWrites=[];const page=createPage(async()=>{if(++calls===1)throw Error('blocked')});
  page.writeCache=(list,pending)=>cacheWrites.push({day:list[0].hours.day,pending});
  assert.equal(await page.saveAll(),false);assert.deepEqual(cacheWrites,[{day:1,pending:true}]);
  assert.equal(await page.saveAll(),true);assert.deepEqual(cacheWrites.slice(1),[{day:1,pending:true},{day:1,pending:false}]);
});
