const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'../app.js'),'utf8');
function load(start,end,globals={}){
  const context=vm.createContext({Date,Math,...globals});
  vm.runInContext(`const pad=n=>String(n).padStart(2,'0'),dateKey=(y,m,d)=>\`\${y}-\${pad(m+1)}-\${pad(d)}\`,monthKey=(y,m)=>\`\${y}-\${pad(m+1)}\`;`+source.slice(source.indexOf(start),source.indexOf(end)),context);
  return context;
}
const ROLLOVER='const ROLLOVER_EXCLUDED_STATUSES=new Set([\'취소\',\'완료\']);';
test('2nd and 3rd Wednesdays count as 0.875 day, required is days × 0.9',()=>{
  const page=load('function isWorkDay','function selectTodayWeek',{holidays:new Set(),customHolidays:{}});
  // 2026-09: 30 days, weekdays 22, Wednesdays 2·9·16·23·30, 2nd/3rd = 9·16
  const result=page.cumulativeWorkDays(2026,8,30);
  assert.equal(result.normalDays,20);assert.equal(result.shortDays,2);assert.equal(result.workDays,21.75);assert.equal(result.required,19.575);
  assert.equal(page.requiredWorkLabel(result),'21.75일 × 0.9 = 19.575');
  assert.equal(page.isShortWednesday(2026,8,2),false);assert.equal(page.isShortWednesday(2026,8,23),false);
});
test('holiday on a short Wednesday is not counted',()=>{
  const page=load('function isWorkDay','function selectTodayWeek',{holidays:new Set(['2026-09-09']),customHolidays:{}});
  const result=page.cumulativeWorkDays(2026,8,30);
  assert.equal(result.normalDays,20);assert.equal(result.shortDays,1);assert.equal(result.workDays,20.875);
});
test('tasks appear from their registration month, not release month',()=>{
  const page=load('function isTaskVisibleInMonth','async function api');
  vm.runInContext(ROLLOVER,page);
  const task={createdDate:'2026-10-05',dates:['2026-12-10'],status:'진행'};
  assert.equal(page.isTaskVisibleInMonth(task,2026,8),false);
  assert.equal(page.isTaskVisibleInMonth(task,2026,9),true);
  assert.equal(page.isTaskVisibleInMonth({dates:['2026-12-10'],status:'진행'},2026,9),false);
  const closed={createdDate:'2026-10-05',dates:['2026-09-10'],status:'완료'};
  assert.equal(page.isTaskVisibleInMonth(closed,2026,9),true);
  assert.equal(page.isTaskVisibleInMonth(closed,2026,10),false);
});
