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
const sample=[
  {id:'a',createdDate:'2026-10-20',rms:'00123',task:'KV 교체',dates:['2026-11-05'],status:'진행',category:'플러스탭',note:'메모',hours:{'2026-10-31':2,'2026-11-02':4,'2026-11-03':0}},
  {id:'b',task:'운영 대응',dates:[],status:'보류',category:'전사 운영',hours:{'2026-11-02':8}},
  {id:'__leave_calendar__',_type:'leave-calendar',leaves:{'2026-11-04':{type:'half',substitute:true}},customHolidays:{}}
];
test('builds table rows with month/day headers, totals and a leave row',()=>{
  const rows=JSON.parse(JSON.stringify(loadScript().buildTableRows(sample,'2026-11-04T21:00:00+09:00')));
  assert.equal(rows[0][0],'최종 저장 2026-11-04T21:00:00+09:00');
  assert.deepEqual(rows[0].slice(7),['2026년 10월','2026년 11월','','합계(시간)','합계(일)']);
  assert.deepEqual(rows[1].slice(0,7),['등록일','RMS, Jira','업무명','반영일','진행현황','구분','비고']);
  assert.deepEqual(rows[1].slice(7,10),['31(토)','2(월)','4(수)']);
  assert.deepEqual(rows[2],['2026-10-20','00123','KV 교체','2026-11-05','진행','플러스탭','메모',2,4,'',6,0.75]);
  assert.deepEqual(rows[3].slice(4,12),['취소','전사 운영','','',8,'',8,1]);
  assert.deepEqual(rows[4].slice(2,3).concat(rows[4].slice(7,10)),['휴가','','','반차 · 대체휴가']);
  assert.equal(rows.length,5);
  assert.ok(rows.every(row=>row.length===rows[0].length));
});
test('save keeps JSON even if the table tab write fails',()=>{
  const written=[];
  const range={setValues:v=>written.push(v)};
  const dataSheet={getRange:()=>range};
  const ctx=loadScript({
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},
    Utilities:{formatDate:()=>'2026-11-04T21:00:00+09:00'},
    SpreadsheetApp:{flush(){},getActiveSpreadsheet:()=>({getSheetByName:name=>{if(name==='웹앱_업무대장')throw new Error('표 탭 오류');return dataSheet}})}
  });
  const result=ctx.savePayload(sample);
  assert.equal(result.ok,true);assert.equal(result.tableError,'표 탭 오류');
  assert.equal(JSON.parse(written[0][0][0]).length,3);
});
test('writes the table into a sheet that is grown to fit',()=>{
  const calls=[];
  const sheet={maxC:3,maxR:2,getMaxColumns(){return this.maxC},getMaxRows(){return this.maxR},insertColumnsAfter(_,n){this.maxC+=n},insertRowsAfter(_,n){this.maxR+=n},clear(){calls.push('clear')},
    getRange(r,c,h,w){if(r+h-1>this.maxR||c+w-1>this.maxC)throw new Error('out of range');const range={setNumberFormat:()=>range,setValues:v=>{calls.push(['values',v.length,v[0].length]);return range},setFontWeight:()=>range,setBackground:()=>range};return range},
    setFrozenRows(){},setFrozenColumns(){}};
  const ctx=loadScript({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>null,insertSheet:()=>sheet})}});
  ctx.writeTableSheet(sample,'t');
  assert.deepEqual(calls,['clear',['values',5,12]]);
});
