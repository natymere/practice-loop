import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const dir=new URL('../src/',import.meta.url);
const code=fs.readFileSync(new URL('Code.gs',dir),'utf8');
const html=fs.readFileSync(new URL('Client.html',dir),'utf8');
new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1]);
const headers={Questions:['ID','URL','Canonical URL','Title','Topic','Created at','Initial help used','Estimated minutes','Stage','Next review','Last result','Last reviewed','Status','Snoozed until','Revision','Notes','Collection','Restore location'],Reviews:['Request ID','Question ID','Reviewed at','Result','Minutes','Mistake note','Previous stage','New stage','Next review','Previous revision','New revision'],Deliveries:['Digest key','Attempt ID','Started at','Channel','Status','Finished at','Detail']};
let database,owner,locked,failCommit,sends,fetchFailure,triggerCount,props;
function reset(){database=Object.fromEntries(Object.entries(headers).map(([n,h])=>[n,[h.slice()]]));database.Settings=[['Setting','Value','Meaning'],['timezone','Asia/Singapore',''],['reminderHour',0,''],['dailyBudget',30,''],['channel','off','']];owner='fearlessjem17@gmail.com';locked=false;failCommit=false;sends=0;fetchFailure=false;triggerCount=0;props={};}
reset();
const names=[...Object.keys(database),'Import'];let uuid=0;
const sheet=name=>({getSheetId:()=>names.indexOf(name),getDataRange:()=>({getValues:()=>{const rows=structuredClone(database[name]);const width=Math.max(...rows.map(r=>r.length));return rows.map(r=>[...r,...Array(width-r.length).fill('')]);}})});
const book={getSheetByName:name=>database[name]?sheet(name):null,insertSheet:name=>{database[name]=[['']];return sheet(name);},getUrl:()=> 'https://docs.google.com/spreadsheets/d/test/edit'};
const ctx=vm.createContext({console,Date,JSON,Math,Number,String,Error,RegExp,Object,Array,Session:{getActiveUser:()=>({getEmail:()=>owner}),getEffectiveUser:()=>({getEmail:()=>owner})},SpreadsheetApp:{openById:()=>book},LockService:{getScriptLock:()=>({waitLock:()=>{if(locked)throw Error('Concurrent writer');locked=true;},releaseLock:()=>{locked=false;}})},Utilities:{getUuid:()=> 'request-'+String(++uuid).padStart(20,'0'),formatDate:(d,tz,fmt)=>{const p=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'numeric',hourCycle:'h23'}).formatToParts(d).map(p=>[p.type,p.value]));return fmt==='H'?p.hour:`${p.year}-${p.month}-${p.day}`;}},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null})},ScriptApp:{getService:()=>({getUrl:()=> 'https://script.google.com/macros/s/test/exec'}),getProjectTriggers:()=>triggerCount?[{getHandlerFunction:()=> 'sendDailyDigest_'}]:[],newTrigger:()=>({timeBased:()=>({everyHours:()=>({create:()=>{triggerCount++;}})})})},MailApp:{sendEmail:()=>{sends++;if(fetchFailure)throw Error('Lost acknowledgement');}},UrlFetchApp:{fetch:()=>{sends++;if(fetchFailure)throw Error('https://api.telegram.org/botSECRET');return {getResponseCode:()=>200,getContentText:()=>'{"ok":true}'};}},Sheets:{Spreadsheets:{batchUpdate:({requests})=>{if(failCommit)throw Error('Injected commit failure');const draft=structuredClone(database);for(const r of requests){const x=r.appendCells||r.updateCells;const id=x.sheetId??x.start.sheetId;x.rows.forEach((entry,index)=>{const row=entry.values.map(c=>Object.values(c.userEnteredValue)[0]);if(r.appendCells)draft[names[id]].push(row);else {const target=draft[names[id]][x.start.rowIndex+index]||[];row.forEach((value,col)=>target[(x.start.columnIndex||0)+col]=value);draft[names[id]][x.start.rowIndex+index]=target;}});}database=draft;}}}});
vm.runInContext(code,ctx);
let checks=0;
const focused=process.argv.includes('--focused');
function test(name,fn){if(focused&&!name.startsWith('Focused:'))return;reset();fn();checks++;console.log('PASS '+name);}
function add(){return ctx.addQuestion({url:'https://leetcode.com/problems/two-sum/description/?utm_source=x',title:'Two Sum',topic:'Arrays',minutes:20,initialHelp:true});}
function makeDue(){database.Questions[1][9]='2020-01-01';}
function review(id,requestId='review-request-00001',revision=0,result='independent'){return ctx.recordReview({id,requestId,revision,result,minutes:20,note:'Boundary case'});}
test('Known LeetCode subpages and tracking URLs deduplicate',()=>{const a=add();const b=ctx.addQuestion({url:'http://www.leetcode.com/problems/TWO-SUM/solutions/123',title:'Duplicate',topic:'',minutes:10,initialHelp:false});assert.equal(b.duplicate,true);assert.equal(a.id,b.id);assert.equal(database.Questions.length,2);});
test('Generic question query IDs and hash routes remain distinct',()=>{assert.notEqual(ctx.canonicalUrl_('https://fastprep.io/q?id=1'),ctx.canonicalUrl_('https://fastprep.io/q?id=2'));assert.notEqual(ctx.canonicalUrl_('https://x.com/#/q/1'),ctx.canonicalUrl_('https://x.com/#/q/2'));assert.equal(ctx.canonicalUrl_('https://x.com/q?b=2&utm_source=a&a=1'),'https://x.com/q?a=1&b=2');});
test('Unsafe URLs are rejected',()=>{for(const url of ['javascript:alert(1)','https://a@evil.com/','https://x.com/\\@evil.com','https://x.com/\nfoo'])assert.throws(()=>ctx.canonicalUrl_(url));});
test('Calendar arithmetic handles month ends and leap years',()=>{assert.equal(ctx.plusDays_('2026-01-31',3),'2026-02-03');assert.equal(ctx.plusDays_('2028-02-28',1),'2028-02-29');assert.equal(ctx.transition_(6,'independent','2026-10-10').stage,6);assert.equal(ctx.transition_(4,'help','2026-10-10').stage,0);assert.equal(ctx.transition_(4,'hard','2026-10-10').due,'2026-10-11');});
test('Review updates schedule and history once; replay is harmless',()=>{const a=add();makeDue();const first=review(a.id);const again=review(a.id);assert.equal(again.duplicate,true);assert.equal(first.due,again.due);assert.equal(database.Reviews.length,2);assert.equal(database.Questions[1][14],1);assert.equal(database.Questions[1][8],1);});
test('A second tab cannot advance a stale review',()=>{const a=add();makeDue();review(a.id);assert.throws(()=>review(a.id,'review-request-00002',0),/another tab/);assert.equal(database.Reviews.length,2);});
test('Retry identifiers cannot be reused with different content',()=>{const a=add();makeDue();review(a.id);assert.throws(()=>review(a.id,'review-request-00001',0,'help'),/different review/);});
test('Same-day repeat is not counted as spaced recall',()=>{const a=add();makeDue();review(a.id);assert.throws(()=>review(a.id,'review-request-00002',1),/not due/);});
test('Atomic storage failure leaves neither a review nor schedule change',()=>{const a=add();makeDue();const before=JSON.stringify(database);failCommit=true;assert.throws(()=>review(a.id),/commit failure/);assert.equal(JSON.stringify(database),before);assert.equal(locked,false);});
test('Owner check rejects other accounts and missing identity',()=>{for(const email of ['someone@example.com','']){owner=email;assert.throws(()=>add(),/Sign in/);assert.throws(()=>ctx.getDashboard(),/Sign in/);assert.throws(()=>ctx.saveSettings({}),/Sign in/);}});
test('Formula-like titles are stored as plain text',()=>{ctx.addQuestion({url:'https://x.com/q',title:'=IMPORTXML("https://evil.com", "x")',topic:'',minutes:10,initialHelp:true});assert.equal(database.Questions[1][3].startsWith('='),true);assert.ok(ctx.cell_('=1+1').userEnteredValue.stringValue);});
test('Snooze preserves stage and due date',()=>{const a=add();makeDue();const due=database.Questions[1][9];ctx.updateQuestion({id:a.id,revision:0,action:'snooze'});assert.equal(database.Questions[1][8],0);assert.equal(database.Questions[1][9],due);assert.equal(ctx.getDashboard().dueCount,0);});
test('Pause and Resume actions are removed',()=>{const a=add();for(const action of ['pause','resume'])assert.throws(()=>ctx.updateQuestion({id:a.id,revision:0,action}),/Unknown/);});
test('One digest per day, including repeated trigger and channel switch',()=>{add();makeDue();database.Settings[4][1]='email';ctx.sendDailyDigest_();ctx.sendDailyDigest_();database.Settings[4][1]='telegram';props={TELEGRAM_BOT_TOKEN:'SECRET',TELEGRAM_CHAT_ID:'1'};ctx.sendDailyDigest_();assert.equal(sends,1);assert.equal(database.Deliveries.length,2);assert.equal(database.Deliveries[1][4],'sent');});
test('Uncertain delivery is logged, not retried, and contains no credential',()=>{add();makeDue();database.Settings[4][1]='telegram';props={TELEGRAM_BOT_TOKEN:'SECRET',TELEGRAM_CHAT_ID:'1'};fetchFailure=true;ctx.sendDailyDigest_();ctx.sendDailyDigest_();assert.equal(sends,1);assert.equal(database.Deliveries[1][4],'uncertain');assert.equal(JSON.stringify(database).includes('SECRET'),false);});
test('A crash after claiming delivery prevents another send',()=>{add();makeDue();database.Settings[4][1]='email';database.Deliveries.push([ctx.day_(new Date(),'Asia/Singapore'),'a','now','email','attempting','','']);ctx.sendDailyDigest_();assert.equal(sends,0);});
test('No reminder when disabled or when nothing is due',()=>{add();makeDue();ctx.sendDailyDigest_();assert.equal(sends,0);database.Settings[4][1]='email';database.Questions[1][9]='2099-01-01';ctx.sendDailyDigest_();assert.equal(sends,0);});
test('Settings installs one trigger across retries',()=>{ctx.saveSettings({channel:'off',reminderHour:9,dailyBudget:30});ctx.saveSettings({channel:'off',reminderHour:9,dailyBudget:30});assert.equal(triggerCount,1);});
test('Dashboard never exposes Telegram credentials',()=>{props={TELEGRAM_BOT_TOKEN:'TOPSECRET',TELEGRAM_CHAT_ID:'123'};const result=ctx.getDashboard();assert.equal(ctx.getSettings().telegramConfigured,true);assert.equal(JSON.stringify(result).includes('TOPSECRET'),false);});
test('Notes persist, edit independently, and preserve the review schedule',()=>{
  const a=ctx.addQuestion({url:'https://x.com/notes',title:'Notes test',minutes:15,initialHelp:false,notes:'Line one\n=1+1 <script>'});
  let q=ctx.getQuestionDetails({id:database.Questions[1][0]}).question;
  assert.equal(q.notes,'Line one\n=1+1 <script>');
  const due=q.due;
  ctx.updateQuestion({id:a.id,revision:0,action:'notes',notes:'Revised\nSecond line'});
  q=ctx.getQuestionDetails({id:database.Questions[1][0]}).question;
  assert.equal(q.notes,'Revised\nSecond line');assert.equal(q.due,due);assert.equal(q.stage,0);assert.equal(database.Reviews.length,1);
  assert.throws(()=>ctx.updateQuestion({id:a.id,revision:0,action:'notes',notes:'stale'}),/changed/);
  ctx.updateQuestion({id:a.id,revision:1,action:'notes',notes:''});assert.equal(ctx.getQuestionDetails({id:database.Questions[1][0]}).question.notes,'');
});
test('Original questions upgrade safely and migration is repeatable',()=>{
  add();database.Questions=database.Questions.map(r=>r.slice(0,15));
  const original=database.Questions[1].slice();
  assert.equal(ctx.getQuestionDetails({id:database.Questions[1][0]}).question.notes,'');
  assert.equal(database.Questions[0][15],'Notes');
  assert.equal(JSON.stringify(database.Questions[1].slice(0,15)),JSON.stringify(original));
  assert.equal(ctx.getQuestions({view:'active'}).questions.length,1);
});
test('Notes validation and duplicate protection preserve existing content',()=>{
  const input={url:'https://x.com/notes',title:'Notes',minutes:15,initialHelp:true,notes:'Keep this'};
  ctx.addQuestion(input);ctx.addQuestion({...input,notes:'Replacement'});
  assert.equal(ctx.getQuestionDetails({id:database.Questions[1][0]}).question.notes,'Keep this');
  assert.throws(()=>ctx.updateQuestion({id:database.Questions[1][0],revision:0,action:'notes',notes:'x'.repeat(5001)}));
  owner='';assert.throws(()=>ctx.updateQuestion({}),/Sign in/);
});
test('Delete removes question from due reviews and reminders; restore keeps history and stage',()=>{
  const a=add();makeDue();review(a.id);makeDue();
  const due=database.Questions[1][9], stage=database.Questions[1][8], history=JSON.stringify(database.Reviews);
  ctx.updateQuestion({id:a.id,revision:1,action:'delete'});
  assert.equal(database.Questions[1][12],'deleted');assert.equal(ctx.getDashboard().dueCount,0);
  database.Settings[4][1]='email';ctx.sendDailyDigest_();assert.equal(sends,0);
  assert.throws(()=>review(a.id,'deleted-review-00001',2),/not due/);
  assert.throws(()=>ctx.updateQuestion({id:a.id,revision:2,action:'resume'}),/Unknown/);
  ctx.updateQuestion({id:a.id,revision:2,action:'restore'});
  assert.equal(database.Questions[1][12],'active');assert.equal(database.Questions[1][9],due);assert.equal(database.Questions[1][8],stage);
  assert.equal(JSON.stringify(database.Reviews),history);assert.equal(ctx.getDashboard().dueCount,1);
});
test('Deleted URLs remain deduplicated and stale deletes cannot overwrite a review',()=>{
  const a=add();makeDue();review(a.id);
  assert.throws(()=>ctx.updateQuestion({id:a.id,revision:0,action:'delete'}),/changed/);
  assert.equal(database.Questions[1][12],'active');
  ctx.updateQuestion({id:a.id,revision:1,action:'delete'});
  const again=add();assert.equal(again.duplicate,true);assert.equal(again.deleted,true);assert.equal(database.Questions.length,2);
  owner='';assert.throws(()=>ctx.updateQuestion({id:a.id,revision:2,action:'restore'}),/Sign in/);
});
test('Reminder channel can be enabled independently of the review schedule',()=>{
  add();const before=JSON.stringify(database.Questions);
  ctx.saveSettings({dailyBudget:30,reminderHour:9,channel:'email'});assert.equal(ctx.settings_().channel,'email');
  assert.throws(()=>ctx.saveSettings({dailyBudget:30,reminderHour:9,channel:'telegram'}),/credentials/);
  props={TELEGRAM_BOT_TOKEN:'SECRET',TELEGRAM_CHAT_ID:'1'};
  ctx.saveSettings({dailyBudget:30,reminderHour:9,channel:'telegram'});assert.equal(ctx.settings_().channel,'telegram');
  assert.equal(JSON.stringify(database.Questions),before);assert.equal(triggerCount,1);
});
test('Focused: stages 1 and 7 use account dates, clear snooze, preserve history and reject stale writes',()=>{
  const a=add();makeDue();review(a.id);
  const history=JSON.stringify(database.Reviews);database.Questions[1][15]='Keep notes';database.Questions[1][13]='2099-01-01';
  const today=ctx.day_(new Date(),'Asia/Singapore');
  for(const [stage,days,revision] of [[1,1,1],[7,90,2]]){
    const result=ctx.updateQuestion({id:a.id,revision,action:'stage',stage}).question;
    assert.equal(result.stage,stage-1);assert.equal(result.due,ctx.plusDays_(today,days));assert.equal(result.snoozedUntil,'');
    assert.equal(ctx.scheduleOptions_(today)[stage-1].due,result.due);
  }
  assert.equal(database.Questions[1][15],'Keep notes');assert.equal(JSON.stringify(database.Reviews),history);
  assert.throws(()=>ctx.updateQuestion({id:a.id,revision:1,action:'stage',stage:1}),/changed/);
  assert.equal(ctx.day_(new Date('2026-10-12T17:00:00Z'),'Asia/Singapore'),'2026-10-13');
  assert.equal(ctx.plusDays_('2026-10-13',90),'2027-01-11');
});
test('Focused: repeatable migration preserves identifiers, stage, notes and history',()=>{
  const a=add();makeDue();review(a.id);database.Questions[1][15]='Keep notes';database.Questions[1][12]='paused';
  const history=JSON.stringify(database.Reviews), original=database.Questions[1].slice();
  database.Questions=database.Questions.map(row=>row.slice(0,16));
  ctx.upgrade();const q=ctx.getQuestionDetails({id:a.id}).question;
  assert.equal(q.id,a.id);assert.equal(q.status,'bank');assert.equal(q.due,'');assert.equal(q.stage,original[8]);assert.equal(q.notes,'Keep notes');assert.equal(q.revision,original[14]+1);
  assert.equal(JSON.stringify(database.Reviews),history);assert.equal(database.Questions[0].length,18);assert.equal(database.Import[0][4],'Result');
  const migrated=JSON.stringify(database);ctx.upgrade();assert.equal(JSON.stringify(database),migrated);assert.equal(sends,0);assert.equal(triggerCount,0);
  database.Questions=database.Questions.map(row=>row.slice(0,15));database.Questions[1][12]='deleted';ctx.upgrade();
  assert.equal(database.Questions[1][17],'active');assert.equal(database.Questions[1][0],a.id);assert.equal(database.Questions[1][8],original[8]);
  ctx.updateQuestion({id:a.id,revision:q.revision,action:'restore'});assert.equal(database.Questions[1][12],'active');
});
test('Focused: imports retry safely, preserve originals and exclude bank/deleted from scheduling',()=>{
  const a=add();database.Questions[1][15]='Original notes';const original=JSON.stringify(database.Questions[1]);ctx.upgrade();
  database.Import.push(['Two sum duplicate','https://leetcode.com/problems/two-sum/solutions/123','Arrays','NeetCode 150',''],['=New bank question','https://x.com/bank?utm_source=a','Graphs','NeetCode 150',''],['Duplicate in batch','https://x.com/bank','Graphs','NeetCode 150',''],['Bad URL','javascript:alert(1)','','',''],['','https://x.com/missing','','',''],['','','','','']);
  let result=ctx.importQuestions();assert.equal(result.imported,1);assert.equal(result.duplicates,2);assert.equal(result.errors,2);
  assert.equal(database.Import[2][4],'Imported');assert.equal(database.Import[3][4],'Already saved');assert.match(database.Import[4][4],/http/);assert.match(database.Import[5][4],/Title/);
  assert.equal(database.Import[2][0],'=New bank question');assert.equal(JSON.stringify(database.Questions[1]),original);
  result=ctx.importQuestions();assert.equal(result.imported,0);assert.equal(result.duplicates,3);assert.equal(database.Questions.length,3);
  const bank=database.Questions[2];assert.equal(bank[9],'');assert.equal(bank[12],'bank');assert.equal(bank[16],'NeetCode 150');
  ctx.updateQuestion({id:a.id,revision:0,action:'delete'});database.Questions[1][9]='2020-01-01';database.Questions[2][9]='2020-01-01';
  database.Settings[4][1]='email';assert.equal(ctx.getDashboard().dueCount,0);ctx.sendDailyDigest_();assert.equal(sends,0);
  assert.throws(()=>review(bank[0],'bank-review-000001'),/not due/);
  const active=ctx.updateQuestion({id:bank[0],revision:0,action:'activate',stage:7}).question;
  assert.equal(active.status,'active');assert.equal(active.stage,6);assert.equal(active.due,ctx.plusDays_(ctx.getDashboard().today,90));
});
test('Focused: delete and restore return to active practice or bank',()=>{
  const a=add();database.Questions[1][8]=6;database.Questions[1][15]='Retain';const due=database.Questions[1][9];
  for(const [location,revision] of [['active',0],['bank',2]]){
    database.Questions[1][12]=location;if(location==='bank')database.Questions[1][9]='';
    ctx.updateQuestion({id:a.id,revision,action:'delete'});assert.equal(database.Questions[1][17],location);
    ctx.updateQuestion({id:a.id,revision:revision+1,action:'restore'});assert.equal(database.Questions[1][12],location);
    assert.equal(database.Questions[1][8],6);assert.equal(database.Questions[1][15],'Retain');assert.equal(database.Questions[1][9],location==='active'?due:'');
  }
  for(const action of ['pause','resume'])assert.throws(()=>ctx.updateQuestion({id:a.id,revision:4,action}),/Unknown/);
});
test('Focused: library paginates and searches all rows; Today and summaries omit notes',()=>{
  for(let i=0;i<30;i++)ctx.addQuestion({url:'https://x.com/q/'+i,title:'Question '+String(i).padStart(2,'0'),topic:i===29?'Special':'Arrays',minutes:20,initialHelp:false,notes:'Private note'});
  database.Questions.slice(1).forEach(row=>row[9]='2020-01-01');
  const first=ctx.getQuestions({view:'active'}),second=ctx.getQuestions({view:'active',offset:first.nextOffset});
  assert.equal(first.questions.length,25);assert.equal(second.questions.length,5);assert.equal(second.nextOffset,null);
  assert.equal(ctx.getQuestions({view:'active',search:'29',topic:'Special'}).total,1);
  assert.equal(new Set([...first.questions,...second.questions].map(q=>q.id)).size,30);
  assert.equal(JSON.stringify(first).includes('Private note'),false);
  const today=ctx.getDashboard();assert.equal(today.dueCount,30);assert.equal(today.questions.length,1);assert.equal(JSON.stringify(today).includes('Private note'),false);
  assert.equal(ctx.getQuestionDetails({id:first.questions[0].id}).question.notes,'Private note');
});
// Lightweight client boundary check; no browser or live Google integration is exercised.
const frontend=vm.createContext({console,URL,Date,Number,Map,Promise,Error,document:{querySelectorAll:()=>[],querySelector:()=>({hidden:false}),getElementById:()=>({replaceChildren(){},append(){}}),createElement:()=>({})}});
vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],frontend);
assert.doesNotThrow(()=>frontend.navigate('library'));
assert.equal(frontend.hostLabel('javascript:alert(1)'),'Question link needs repair');
assert.equal(frontend.hostLabel('https://leetcode.com/problems/a'),'leetcode.com');
await assert.rejects(frontend.api('getDashboard'),/Google-hosted/);
const index=fs.readFileSync(new URL('Index.html',dir),'utf8');
assert.match(index,/include_\('Styles'\)/);assert.match(index,/include_\('Client'\)/);
assert.equal(index.includes('<script>'),false);assert.equal(index.includes('<style>'),false);
assert.throws(()=>ctx.include_('Code'),/Unknown page file/);
console.log('PASS Client syntax, early navigation, URL safety, missing bridge, and template boundary');
const manifest=JSON.parse(fs.readFileSync(new URL('appsscript.json',dir),'utf8'));
assert.equal(manifest.webapp.access,'MYSELF');assert.equal(manifest.webapp.executeAs,'USER_DEPLOYING');
console.log('PASS Private deployment manifest');
console.log(`${checks} scenario groups plus syntax/client boundary and manifest checks passed. Live Google integration remains unverified.`);
