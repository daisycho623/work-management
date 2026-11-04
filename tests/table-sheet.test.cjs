const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
function loadScript(globals={}){
  const context=vm.createContext({JSON,Number,Object,Set,Array,Math,Date,String,Boolean,...globals});
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../apps-script/Code.gs'),'utf8'),context);
  return context;
}
const plain=value=>JSON.parse(JSON.stringify(value));
const sample=[
  {id:'a',createdDate:'2026-10-20',rms:'00123',task:'KV 교체',dates:['2026-11-05'],status:'진행',category:'플러스탭',note:'메모',hours:{'2026-08-31':5,'2026-09-01':2,'2026-10-30':4,'2026-11-02':4,'2026-11-03':0}},
  {id:'b',task:'운영 대응',dates:[],status:'보류',category:'전사 운영',hours:{'2026-09-26':8}},
  {id:'old',createdDate:'2026-06-01',task:'지난 업무',dates:['2026-06-30'],status:'완료',hours:{'2026-06-30':8}},
  {id:'__leave_calendar__',_type:'leave-calendar',leaves:{'2026-11-04':{type:'half',substitute:true}},customHolidays:{'2026-09-02':'창립기념일'}}
];
const options={updatedAt:'2026-11-04T21:00:00+09:00',today:'2026-11-04',holidays:['2026-09-24','2026-09-25','2026-09-26','2026-10-05']};
function build(tasks=sample,opts=options){return plain(loadScript().buildTableRows(tasks,opts))}
function col(rows,label){return rows[1].indexOf(label)}

test('columns run from Sep 1 to the end of the current month as weekdays without holidays',()=>{
  const {rows,dateCount}=build();
  const days=rows[1].slice(6);
  assert.equal(days[0],'1(화)');
  assert.ok(!days.includes('2(수)')||days.indexOf('2(수)')>days.indexOf('9월 합계(일)'),'custom holiday 9/2 is skipped');
  assert.ok(!days.includes('5(토)'),'weekend skipped');
  assert.ok(!days.includes('24(목)')&&!days.includes('25(금)'),'official holidays skipped');
  assert.ok(days.includes('26(토)'),'a holiday/weekend with hours is kept so data is not lost');
  assert.deepEqual(days.filter(day=>day.includes('합계')),['9월 합계(일)','10월 합계(일)','11월 합계(일)','기간 합계(시간)','기간 합계(일)']);
  assert.equal(days[days.indexOf('11월 합계(일)')-1],'30(월)','November runs to month end');
  assert.equal(dateCount,days.length-5);
  assert.equal(rows[0][0],'최종 저장 2026-11-04T21:00:00+09:00');
  assert.equal(rows[0][6],'2026년 9월');
assert.deepEqual(rows[1].slice(0,6),['등록일','RMS, Jira','업무명','반영일','진행현황','비고'],'구분 열 없음');
  assert.equal(rows[0][col(rows,'9월 합계(일)')+1],'2026년 10월');
});
test('task rows hold daily hours, monthly day totals and period totals only from Sep',()=>{
  const {rows,taskCount}=build();
  assert.equal(taskCount,2,'old completed task outside the period is left out');
  const a=rows[2];
  assert.deepEqual(a.slice(0,6),['2026-10-20','00123','KV 교체','2026-11-05','진행','메모']);
  assert.equal(a[col(rows,'1(화)')],2);
  assert.equal(a[col(rows,'9월 합계(일)')],0.25);
  assert.equal(a[col(rows,'10월 합계(일)')],0.5);
  assert.equal(a[col(rows,'11월 합계(일)')],0.5);
  assert.equal(a[col(rows,'기간 합계(시간)')],10,'Aug 31 hours are outside the period');
  assert.equal(a[col(rows,'기간 합계(일)')],1.25);
  assert.equal(rows[3][4],'취소');
  const leave=rows.at(-1);
  assert.deepEqual(leave.slice(0,6),['휴가','','','','',''],'label sits in A so it survives the A:F merge');assert.equal(leave[col(rows,'4(수)')],'반차 · 대체휴가');
  assert.ok(rows.every(row=>row.length===rows[0].length));
});
test('sync key changes when data is saved again or a new month starts',()=>{
  const ctx=loadScript();
  const key=ctx.tableSyncKey(sample,'t1','2026-11-04');
  assert.equal(ctx.tableSyncKey(sample,'t1','2026-11-30'),key);
  assert.notEqual(ctx.tableSyncKey(sample,'t2','2026-11-04'),key);
  assert.notEqual(ctx.tableSyncKey(sample,'t1','2026-12-01'),key);
});
function fakeSheet(note=''){
  const sheet={calls:[],note,maxC:3,maxR:2,getMaxColumns(){return this.maxC},getMaxRows(){return this.maxR},insertColumnsAfter(_,n){this.maxC+=n},insertRowsAfter(_,n){this.maxR+=n},clear(){this.calls.push('clear')},setColumnWidths(start,count,width){this.calls.push(['widths',start,count,width])},setFrozenRows(n){this.calls.push(['frozenRows',n])},setFrozenColumns(n){this.calls.push(['frozenColumns',n])},
    getRange(r,c,h=1,w=1){if(typeof r==='string')return {getNote:()=>sheet.note,setNote:v=>{sheet.note=v}};if(r+h-1>this.maxR||c+w-1>this.maxC)throw new Error('out of range');const range={breakApart:()=>{sheet.calls.push(['breakApart',r,c,h,w]);return range},merge:()=>{sheet.calls.push(['merge',r,c,h,w]);return range},setFontFamily:f=>{sheet.calls.push(['font',r,c,h,w,f]);return range},setHorizontalAlignment:a=>{sheet.calls.push(['align',r,c,h,w,a]);return range},setNumberFormat:()=>range,setValues:v=>{sheet.calls.push(['values',v.length,v[0].length]);return range},setFontWeight:()=>range,setBackground:()=>range};return range}};
  return sheet;
}
function scriptWith(sheet,{existing=true,lockOk=true}={}){
  return loadScript({
    Utilities:{formatDate:()=>'2026-11-04'},
    PropertiesService:{getScriptProperties:()=>({getProperty:()=>''})},
    CacheService:{getScriptCache:()=>({get:()=>null})},
    LockService:{getScriptLock:()=>({tryLock:()=>lockOk,waitLock(){},releaseLock(){}})},
    SpreadsheetApp:{flush(){},getActiveSpreadsheet:()=>({getSheetByName:()=>existing?sheet:null,insertSheet:()=>sheet})}
  });
}
test('load syncs an empty or outdated tab and skips an up-to-date one',()=>{
  const sheet=fakeSheet();const ctx=scriptWith(sheet);
  assert.equal(ctx.syncTableSheetIfStale(sample,'t1'),true);
  assert.equal(sheet.calls[0][0],'breakApart');assert.equal(sheet.calls[1],'clear');
  const values=sheet.calls.find(call=>call[0]==='values'),widths=sheet.calls.find(call=>call[0]==='widths');
  const aligns=sheet.calls.filter(call=>call[0]==='align');
  assert.deepEqual(aligns,[['align',1,1,values[1],values[2],'left'],['align',3,7,values[1]-3,values[2]-6,'right']],'numbers in task rows right, everything else left');
  const [,height,width]=values;
  assert.deepEqual(sheet.calls.find(call=>call[0]==='font'),['font',1,1,height,width,'Roboto']);
  assert.deepEqual(sheet.calls.find(call=>call[0]==='frozenColumns'),['frozenColumns',6],'freeze through F');
  assert.deepEqual(sheet.calls.find(call=>call[0]==='merge'),['merge',height,1,1,6],'merge A:F on the leave row');
  assert.deepEqual(widths,['widths',7,values[2]-6,40],'only time/total columns after 비고 are narrowed');assert.equal(sheet.note,ctx.tableSyncKey(sample,'t1','2026-11-04'));
  sheet.calls=[];
  assert.equal(ctx.syncTableSheetIfStale(sample,'t1'),false);assert.equal(sheet.calls.length,0);
  assert.equal(ctx.syncTableSheetIfStale(sample,'t2'),true);
  assert.equal(scriptWith(fakeSheet(),{existing:false}).syncTableSheetIfStale(sample,'t1'),true,'creates a missing tab');
  assert.equal(scriptWith(fakeSheet(),{lockOk:false}).syncTableSheetIfStale(sample,'t1'),false,'skips while a save holds the lock');
});
test('save keeps JSON even if the table tab write fails',()=>{
  const written=[];
  const range={setValues:v=>written.push(v)};
  const ctx=loadScript({
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},
    Utilities:{formatDate:()=>'2026-11-04T21:00:00+09:00'},
    SpreadsheetApp:{flush(){},getActiveSpreadsheet:()=>({getSheetByName:name=>{if(name==='웹앱_업무대장')throw new Error('표 탭 오류');return {getRange:()=>range}}})}
  });
  const result=ctx.savePayload(sample);
  assert.equal(result.ok,true);assert.equal(result.tableError,'표 탭 오류');
  assert.equal(JSON.parse(written[0][0][0]).length,4);
});
