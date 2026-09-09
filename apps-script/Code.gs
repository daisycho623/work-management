const DATA_SHEET = '웹앱_업무데이터';
function doGet(){return jsonResponse(loadPayload())}
function doPost(e){try{const r=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');if(r.action==='load')return jsonResponse(loadPayload());if(r.action==='save')return jsonResponse(savePayload(r.tasks));return jsonResponse({ok:false,error:'지원하지 않는 요청입니다.'})}catch(error){return jsonResponse({ok:false,error:error.message})}}
function loadPayload(){const s=getDataSheet(),json=s.getRange('A2').getValue();return {ok:true,tasks:json?JSON.parse(json):[],updatedAt:s.getRange('B2').getValue()||''}}
function savePayload(tasks){if(!Array.isArray(tasks))throw new Error('업무 데이터 형식이 올바르지 않습니다.');const lock=LockService.getScriptLock();lock.waitLock(10000);try{const s=getDataSheet(),updatedAt=new Date().toISOString();s.getRange('A2:B2').setValues([[JSON.stringify(tasks),updatedAt]]);SpreadsheetApp.flush();return {ok:true,updatedAt}}finally{lock.releaseLock()}}
function getDataSheet(){const book=SpreadsheetApp.getActiveSpreadsheet();let sheet=book.getSheetByName(DATA_SHEET);if(!sheet){sheet=book.insertSheet(DATA_SHEET);sheet.getRange('A1:B1').setValues([['업무 데이터(JSON)','최종 저장 시각']]);sheet.hideSheet()}return sheet}
function jsonResponse(payload){return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON)}
