const DATA_SHEET = '웹앱_업무데이터';
const BACKUP_SHEET = '웹앱_백업데이터';
const TABLE_SHEET = '웹앱_업무대장';
const BACKUP_TIMEZONE = 'Asia/Seoul';
const HOLIDAY_API_URL = 'https://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo';
function doGet(){return jsonResponse(loadPayload())}
function doPost(e){try{const r=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');if(r.action==='load')return jsonResponse(loadPayload(r.year));if(r.action==='holidays')return jsonResponse({ok:true,holidays:getHolidays(r.year)});if(r.action==='save')return jsonResponse(savePayload(r.tasks));if(r.action==='syncTable')return jsonResponse({ok:true,...rebuildTableSheet()});return jsonResponse({ok:false,error:'지원하지 않는 요청입니다.'})}catch(error){return jsonResponse({ok:false,error:error.message})}}
function loadPayload(year){
  const s=getDataSheet(),json=s.getRange('A2').getValue();
  const holidayYear=year==null?new Date().getFullYear():Number(year);
  const payload={ok:true,tasks:json?JSON.parse(json):[],holidays:[],holidayYear,updatedAt:s.getRange('B2').getValue()||''};
  try{payload.holidays=getHolidays(holidayYear)}catch(error){payload.holidayError=error.message}
  try{syncTableSheetIfStale(payload.tasks,payload.updatedAt)}catch(error){payload.tableError=error.message}
  return payload;
}
function getHolidays(year){
  const targetYear=year==null?new Date().getFullYear():Number(year);
  if(!Number.isInteger(targetYear)||targetYear<2026||targetYear>2100)throw new Error('조회할 연도가 올바르지 않습니다.');
  const key=(PropertiesService.getScriptProperties().getProperty('DATA_GO_KR_SERVICE_KEY')||'').trim();
  if(!key)throw new Error('Apps Script 스크립트 속성에 DATA_GO_KR_SERVICE_KEY를 설정해 주세요.');
  let decodedKey=key;
  try{decodedKey=decodeURIComponent(key)}catch(error){}
  const cache=CacheService.getScriptCache(),cacheKey='data-go-kr-holidays-v1-'+targetYear,cached=cache.get(cacheKey);
  if(cached)return JSON.parse(cached);
  const dates=new Set();
  for(let page=1;page<=100;page++){
    const url=HOLIDAY_API_URL+'?ServiceKey='+encodeURIComponent(decodedKey)+'&solYear='+targetYear+'&numOfRows=100&pageNo='+page;
    let response;
    try{response=UrlFetchApp.fetch(url,{muteHttpExceptions:true})}catch(error){throw new Error('공공데이터 공휴일 API에 연결하지 못했습니다. 잠시 후 다시 시도해 주세요.')}
    if(response.getResponseCode()!==200)throw new Error('공공데이터 공휴일 API 조회 실패 (HTTP '+response.getResponseCode()+'). 인증키와 활용신청 상태를 확인해 주세요.');
    let root;
    try{root=XmlService.parse(response.getContentText()).getRootElement()}catch(error){throw new Error('공공데이터 공휴일 API 응답 형식이 올바르지 않습니다.')}
    const header=root.getChild('header'),code=header&&header.getChildText('resultCode');
    if(code!=='00')throw new Error('공공데이터 공휴일 API 오류. 인증키와 활용신청 상태를 확인해 주세요.');
    const body=root.getChild('body'),items=body&&body.getChild('items'),rows=items?items.getChildren('item'):[],total=Number(body&&body.getChildText('totalCount'));
    if(!body||!Number.isInteger(total)||total<1||!rows.length)throw new Error('해당 연도의 공휴일 데이터가 없거나 불완전합니다.');
    rows.forEach(item=>{
      if(item.getChildText('isHoliday')!=='Y')return;
      const date=item.getChildText('locdate')||'';
      if(!/^\d{8}$/.test(date)||date.slice(0,4)!==String(targetYear))throw new Error('공휴일 날짜 형식이 올바르지 않습니다.');
      dates.add(date.slice(0,4)+'-'+date.slice(4,6)+'-'+date.slice(6,8));
    });
    if(page*100>=total){
      const result=[...dates].sort();
      if(!result.length)throw new Error('해당 연도의 공휴일 데이터가 없습니다.');
      cache.put(cacheKey,JSON.stringify(result),21600);
      return result;
    }
  }
  throw new Error('공휴일 조회 페이지 한도를 초과했습니다.');
}
function formatKoreanTimestamp(date){return Utilities.formatDate(date||new Date(),BACKUP_TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX")}
function savePayload(tasks){if(!Array.isArray(tasks))throw new Error('업무 데이터 형식이 올바르지 않습니다.');const json=JSON.stringify(tasks),lock=LockService.getScriptLock();if(!lock.tryLock(30000))throw new Error('다른 저장 또는 백업 작업이 진행 중입니다. 잠시 후 저장을 다시 시도해 주세요.');try{const s=getDataSheet(),updatedAt=formatKoreanTimestamp();s.getRange('A2:B2').setValues([[json,updatedAt]]);SpreadsheetApp.flush();const result={ok:true,updatedAt};try{writeTableSheet(tasks,updatedAt);SpreadsheetApp.flush()}catch(error){result.tableError=error.message}return result}finally{lock.releaseLock()}}
function getDataSheet(){const book=SpreadsheetApp.getActiveSpreadsheet();let sheet=book.getSheetByName(DATA_SHEET);if(!sheet){sheet=book.insertSheet(DATA_SHEET);sheet.getRange('A1:B1').setValues([['업무 데이터(JSON)','최종 저장 시각']]);sheet.hideSheet()}return sheet}
function getBackupSheet(){const book=SpreadsheetApp.getActiveSpreadsheet();let sheet=book.getSheetByName(BACKUP_SHEET);if(!sheet){sheet=book.insertSheet(BACKUP_SHEET);sheet.getRange('A1:D1').setValues([['백업 월','백업 시각','원본 최종 저장 시각','업무 데이터(JSON)']]);sheet.setFrozenRows(1)}return sheet}
function backupMonthlyData(){const lock=LockService.getScriptLock();lock.waitLock(10000);try{const now=new Date(),backupMonth=Utilities.formatDate(now,BACKUP_TIMEZONE,'yyyy-MM'),backupAt=formatKoreanTimestamp(now),source=getDataSheet(),values=source.getRange('A2:B2').getValues()[0],json=values[0]||'[]',sourceUpdatedAt=values[1]||'',backup=getBackupSheet(),lastRow=backup.getLastRow(),months=lastRow>1?backup.getRange(2,1,lastRow-1,1).getDisplayValues().flat():[],existingIndex=months.indexOf(backupMonth),row=existingIndex>=0?existingIndex+2:lastRow+1;backup.getRange(row,1,1,4).setValues([[backupMonth,backupAt,sourceUpdatedAt,json]]);SpreadsheetApp.flush();return {ok:true,backupMonth,row}}finally{lock.releaseLock()}}
function setupMonthlyBackupTrigger(){getBackupSheet();ScriptApp.getProjectTriggers().filter(trigger=>trigger.getHandlerFunction()==='backupMonthlyData').forEach(trigger=>ScriptApp.deleteTrigger(trigger));ScriptApp.newTrigger('backupMonthlyData').timeBased().onMonthDay(1).atHour(1).inTimezone(BACKUP_TIMEZONE).create();return '매월 1일 01시대 월간 백업 트리거가 설정되었습니다.'}
// 화면의 업무 표(table-scroll)와 같은 형태로 사람이 읽을 수 있는 업무대장 탭을 만든다.
// 원본은 웹앱_업무데이터의 JSON이며, 이 탭은 저장·불러오기 때 최신 상태로 다시 쓴다(웹 → 시트 단방향).
// 이 탭 기록이 실패해도 JSON 저장과 불러오기는 유지된다.
const TABLE_START_DATE='2026-09-01';
const TABLE_VERSION='v5';
const TABLE_TIME_COLUMN_WIDTH=40;
const TABLE_INFO_HEADERS=['등록일','RMS, Jira','업무명','반영일','진행현황','비고'];
const TABLE_LEAVE_LABELS={annual:'연차',half:'반차',quarter:'반반차'};
const TABLE_CLOSED_STATUSES=['완료','취소','보류'];
function tableLeaveLabel(value){let type='annual',substitute=false;if(value&&typeof value==='object'){type=TABLE_LEAVE_LABELS[value.type]?value.type:'annual';substitute=Boolean(value.substitute)}else if(value==='substitute'){substitute=true}else if(TABLE_LEAVE_LABELS[value]){type=value}return TABLE_LEAVE_LABELS[type]+(substitute?' · 대체휴가':'')}
function isTableDate(value){return /^\d{4}-\d{2}-\d{2}$/.test(String(value||''))}
function tableDateAdd(date,days){const d=new Date(date+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+days);return d.toISOString().slice(0,10)}
function tableMonthEnd(date){const d=new Date(Date.UTC(Number(date.slice(0,4)),Number(date.slice(5,7)),0));return d.toISOString().slice(0,10)}
function splitTableData(tasks){
  const calendar=tasks.find(task=>task&&(task.id==='__leave_calendar__'||task._type==='leave-calendar'))||{};
  const list=tasks.filter(task=>task&&task!==calendar&&task.id!=='__leave_calendar__'&&task._type!=='leave-calendar');
  return {list,leaves:calendar.leaves||{},customHolidays:calendar.customHolidays||{}};
}
// 9월 1일부터 (오늘·마지막 입력일 중 늦은 날)이 속한 달의 말일까지.
function tableRangeEnd(tasks,today){
  const {list,leaves}=splitTableData(tasks);let last=today;
  list.forEach(task=>Object.keys(task.hours||{}).forEach(date=>{if(isTableDate(date)&&Number(task.hours[date])&&date>last)last=date}));
  Object.keys(leaves).forEach(date=>{if(isTableDate(date)&&date>last&&date.slice(0,7)===today.slice(0,7))last=date});
  return tableMonthEnd(last<TABLE_START_DATE?TABLE_START_DATE:last);
}
function tableSyncKey(tasks,updatedAt,today){return [updatedAt||'',tableRangeEnd(tasks,today),TABLE_VERSION].join('|')}
function buildTableRows(tasks,options){
  const opts=options||{},today=opts.today||TABLE_START_DATE,holidaySet=new Set(opts.holidays||[]),weekdays='일월화수목금토';
  const {list,leaves,customHolidays}=splitTableData(tasks),end=tableRangeEnd(tasks,today);
  const hasData=date=>Boolean(leaves[date])||list.some(task=>Number((task.hours||{})[date]));
  // 화면과 같이 평일만 표시하되, 주말·휴일이라도 입력값이 있으면 데이터가 빠지지 않도록 포함한다.
  const dates=[];
  for(let date=TABLE_START_DATE;date<=end;date=tableDateAdd(date,1)){
    const dow=new Date(date+'T00:00:00Z').getUTCDay(),off=dow===0||dow===6||holidaySet.has(date)||Boolean(customHolidays[date])||date.slice(5)==='05-01';
    if(!off||hasData(date))dates.push(date);
  }
  // 월별 날짜 묶음 뒤에 월 합계(일) 열을 둔다.
  const columns=[];
  dates.forEach((date,index)=>{columns.push({date});const next=dates[index+1];if(!next||next.slice(0,7)!==date.slice(0,7))columns.push({month:date.slice(0,7)})});
  const monthLabel=month=>month.slice(0,4)+'년 '+Number(month.slice(5,7))+'월';
  const monthRow=[...TABLE_INFO_HEADERS.map((_,index)=>index===0?'최종 저장 '+(opts.updatedAt||''):''),...columns.map((column,index)=>column.date&&(index===0||!columns[index-1].date||columns[index-1].date.slice(0,7)!==column.date.slice(0,7))?monthLabel(column.date.slice(0,7)):''),'',''];
  const dayRow=[...TABLE_INFO_HEADERS,...columns.map(column=>column.date?Number(column.date.slice(8))+'('+weekdays[new Date(column.date+'T00:00:00Z').getUTCDay()]+')':Number(column.month.slice(5,7))+'월 합계(일)'),'기간 합계(시간)','기간 합계(일)'];
  const toDays=hours=>Math.round(hours/8*1000)/1000;
  const visible=list.filter(task=>{
    const releaseDate=(Array.isArray(task.dates)?task.dates:(task.date?[task.date]:[])).filter(Boolean).sort()[0]||'';
    return dates.some(date=>Number((task.hours||{})[date]))||(task.createdDate||'')>=TABLE_START_DATE||releaseDate>=TABLE_START_DATE||TABLE_CLOSED_STATUSES.indexOf(task.status)<0;
  });
  const rows=visible.map(task=>{
    let total=0,monthTotal=0;
    const cells=columns.map(column=>{if(column.month){const value=monthTotal;monthTotal=0;return value?toDays(value):''}const value=Number((task.hours||{})[column.date])||0;total+=value;monthTotal+=value;return value||''});
    const releaseDate=(Array.isArray(task.dates)?task.dates:(task.date?[task.date]:[])).filter(Boolean).sort()[0]||'';
    return [task.createdDate||'',task.rms||'',task.task||'',releaseDate,task.status==='보류'?'취소':(task.status||''),task.note||'',...cells,total||'',total?toDays(total):''];
  });
  const leaveRow=[...TABLE_INFO_HEADERS.map(header=>header==='업무명'?'휴가':''),...columns.map(column=>column.date&&leaves[column.date]?tableLeaveLabel(leaves[column.date]):''),'',''];
  return {rows:[monthRow,dayRow,...rows,leaveRow],monthTotalColumns:columns.map((column,index)=>column.month?TABLE_INFO_HEADERS.length+index+1:0).filter(Boolean),dateCount:dates.length,taskCount:visible.length};
}
function tableToday(){return Utilities.formatDate(new Date(),BACKUP_TIMEZONE,'yyyy-MM-dd')}
function tableHolidays(today){const holidays=[];for(let year=Number(TABLE_START_DATE.slice(0,4));year<=Number(today.slice(0,4))+1;year++){try{getHolidays(year).forEach(date=>holidays.push(date))}catch(error){}}return holidays}
function getTableSheet(){const book=SpreadsheetApp.getActiveSpreadsheet();return book.getSheetByName(TABLE_SHEET)||book.insertSheet(TABLE_SHEET)}
function writeTableSheet(tasks,updatedAt){
  const today=tableToday(),built=buildTableRows(tasks,{updatedAt,today,holidays:tableHolidays(today)}),rows=built.rows,width=rows[0].length,height=rows.length,sheet=getTableSheet();
  if(sheet.getMaxColumns()<width)sheet.insertColumnsAfter(sheet.getMaxColumns(),width-sheet.getMaxColumns());
  if(sheet.getMaxRows()<height)sheet.insertRowsAfter(sheet.getMaxRows(),height-sheet.getMaxRows());
  sheet.clear();
  sheet.getRange(1,1,height,TABLE_INFO_HEADERS.length).setNumberFormat('@');
  sheet.getRange(1,1,height,width).setValues(rows);
  sheet.getRange(1,1,2,width).setFontWeight('bold').setBackground('#f1f3f4');
  built.monthTotalColumns.forEach(column=>sheet.getRange(1,column,height,1).setBackground('#eef3fd'));
  sheet.getRange(height,1,1,width).setBackground('#fff5f8');
  // 전체는 좌측 정렬, 업무 행의 시간·합계 숫자(예: G3부터, 머리글·휴가 행 제외)만 우측 정렬한다.
  sheet.getRange(1,1,height,width).setHorizontalAlignment('left');
  if(height>3)sheet.getRange(3,TABLE_INFO_HEADERS.length+1,height-3,width-TABLE_INFO_HEADERS.length).setHorizontalAlignment('right');
  // 업무 정보 열(등록일~비고) 너비는 그대로 두고, 시간·합계 열만 좁게 고정한다.
  sheet.setColumnWidths(TABLE_INFO_HEADERS.length+1,width-TABLE_INFO_HEADERS.length,TABLE_TIME_COLUMN_WIDTH);
  sheet.setFrozenRows(2);sheet.setFrozenColumns(3);
  sheet.getRange('A1').setNote(tableSyncKey(tasks,updatedAt,today));
  return {table:TABLE_SHEET,rows:built.taskCount,dates:built.dateCount,from:TABLE_START_DATE};
}
// 불러오기 때 탭이 없거나 마지막 저장·표시 기간과 다르면 다시 쓴다. 저장과 겹치면 이번에는 건너뛴다.
function syncTableSheetIfStale(tasks,updatedAt){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLE_SHEET);
  if(sheet&&sheet.getRange('A1').getNote()===tableSyncKey(tasks,updatedAt,tableToday()))return false;
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return false;
  try{writeTableSheet(tasks,updatedAt);SpreadsheetApp.flush();return true}finally{lock.releaseLock()}
}
function readStoredTasks(){const values=getDataSheet().getRange('A2:B2').getValues()[0];return {tasks:values[0]?JSON.parse(values[0]):[],updatedAt:values[1]||''}}
// Apps Script 편집기에서 실행하거나 웹 앱 action=syncTable 요청으로 즉시 다시 만든다.
function rebuildTableSheet(){const lock=LockService.getScriptLock();lock.waitLock(30000);try{const stored=readStoredTasks(),result=writeTableSheet(stored.tasks,stored.updatedAt||formatKoreanTimestamp());SpreadsheetApp.flush();return result}finally{lock.releaseLock()}}

function jsonResponse(payload){return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON)}
