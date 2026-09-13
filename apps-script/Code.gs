const DATA_SHEET = '웹앱_업무데이터';
const HOLIDAY_API_URL = 'https://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo';
function doGet(){return jsonResponse(loadPayload())}
function doPost(e){try{const r=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');if(r.action==='load')return jsonResponse(loadPayload(r.year));if(r.action==='holidays')return jsonResponse({ok:true,holidays:getHolidays(r.year)});if(r.action==='save')return jsonResponse(savePayload(r.tasks));return jsonResponse({ok:false,error:'지원하지 않는 요청입니다.'})}catch(error){return jsonResponse({ok:false,error:error.message})}}
function loadPayload(year){
  const s=getDataSheet(),json=s.getRange('A2').getValue();
  const holidayYear=year==null?new Date().getFullYear():Number(year);
  const payload={ok:true,tasks:json?JSON.parse(json):[],holidays:[],holidayYear,updatedAt:s.getRange('B2').getValue()||''};
  try{payload.holidays=getHolidays(holidayYear)}catch(error){payload.holidayError=error.message}
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
function savePayload(tasks){if(!Array.isArray(tasks))throw new Error('업무 데이터 형식이 올바르지 않습니다.');const lock=LockService.getScriptLock();lock.waitLock(10000);try{const s=getDataSheet(),updatedAt=new Date().toISOString();s.getRange('A2:B2').setValues([[JSON.stringify(tasks),updatedAt]]);SpreadsheetApp.flush();return {ok:true,updatedAt}}finally{lock.releaseLock()}}
function getDataSheet(){const book=SpreadsheetApp.getActiveSpreadsheet();let sheet=book.getSheetByName(DATA_SHEET);if(!sheet){sheet=book.insertSheet(DATA_SHEET);sheet.getRange('A1:B1').setValues([['업무 데이터(JSON)','최종 저장 시각']]);sheet.hideSheet()}return sheet}
function jsonResponse(payload){return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON)}
