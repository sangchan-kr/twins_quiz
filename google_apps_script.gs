var SHEET_NAME='results';
var VALID_SETS=['history1','history2','tech'];
function sheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 var s=ss.getSheetByName(SHEET_NAME)||ss.insertSheet(SHEET_NAME);
 if(s.getLastRow()===0)s.appendRow(['서버 기록 시각','응시자 이름','문제셋','문항 수','정답 수','점수','오답 번호','응시 시각']);
 return s;
}
function safe_(v){var t=String(v==null?'':v);return /^[=+@\-\t\r]/.test(t)?"'"+t:t;}
function doPost(e){
 try{
  var d=JSON.parse(e.postData.contents);
  var name=String(d.name||'').trim().slice(0,30),set=String(d.set||'');
  var total=Number(d.total),correct=Number(d.correct);
  if(!name||VALID_SETS.indexOf(set)<0||!Number.isInteger(total)||total<1||total>50||!Number.isInteger(correct)||correct<0||correct>total)throw Error('Invalid score');
  var score=Math.round(correct/total*100),lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{sheet_().appendRow([new Date(),safe_(name),set,total,correct,score,Array.isArray(d.wrongIds)?d.wrongIds.join(','):'',safe_(d.date||'')]);}
  finally{lock.releaseLock();}
  return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
 }catch(err){return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);}
}
function mask_(s){
 s=String(s||'').replace(/^'/,'');
 if(s.length<=1)return s.slice(0,1)+'*';
 if(s.length===2)return s.slice(0,1)+'*';
 return s.slice(0,1)+'*'.repeat(Math.min(s.length-2,3))+s.slice(-1);
}
function doGet(e){
 var result={ok:true,sets:{}};
 try{
  VALID_SETS.forEach(function(k){result.sets[k]={first:[],best:[],participants:0};});
  var rows=sheet_().getDataRange().getValues();
  var people={};
  // 최초/최고는 각 학생·문제셋의 50문항 전체 시험 기준.
  for(var i=1;i<rows.length;i++){
   var row=rows[i],set=String(row[2]||''),name=String(row[1]||'').trim().replace(/^'/,'');
   if(VALID_SETS.indexOf(set)<0||!name||Number(row[3])!==50)continue;
   var correct=Number(row[4]),score=Math.round(correct/50*100);
   if(!Number.isFinite(correct)||correct<0||correct>50)continue;
   var key=set+'\u0000'+name;
   if(!people[key])people[key]={set:set,name:name,first:{score:score,correct:correct,date:row[0],index:i},best:{score:score,correct:correct,date:row[0],index:i},tries:0};
   var p=people[key];p.tries++;
   if(score>p.best.score)p.best={score:score,correct:correct,date:row[0],index:i};
  }
  Object.keys(people).forEach(function(key){
   var p=people[key],dest=result.sets[p.set];
   dest.first.push({name:mask_(p.name),score:p.first.score,correct:p.first.correct,attempts:p.tries,date:Utilities.formatDate(new Date(p.first.date),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm')});
   dest.best.push({name:mask_(p.name),score:p.best.score,correct:p.best.correct,attempts:p.tries,date:Utilities.formatDate(new Date(p.best.date),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm')});
  });
  VALID_SETS.forEach(function(set){
   var a=result.sets[set];a.first.sort(function(x,y){return y.score-x.score||x.date.localeCompare(y.date);});
   a.best.sort(function(x,y){return y.score-x.score||x.date.localeCompare(y.date);});
   a.participants=a.first.length;a.first=a.first.slice(0,100);a.best=a.best.slice(0,100);
  });
 }catch(err){result={ok:false,error:String(err)};}
 var cb=String(e&&e.parameter&&e.parameter.callback||'');
 if(/^[A-Za-z_$][\w$]{0,80}$/.test(cb)){
  return ContentService.createTextOutput(cb+'('+JSON.stringify(result).replace(/</g,'\\u003c')+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
 }
 return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}