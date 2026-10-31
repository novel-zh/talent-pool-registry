const fs=require('fs'),vm=require('vm'),crypto=require('crypto');
const fixtures=require('./fixtures.cjs');
const assert=(v,m)=>{if(!v)throw Error(m);};
const expectError=(fn,part)=>{let error;try{fn();}catch(e){error=e;}assert(error&&(!part||error.message.includes(part)),'Expected error: '+part);};
function display(v){if(v&&v.date)return v.date.slice(0,10).split('-').reverse().join('.');if(v instanceof Date)return v.toISOString().slice(0,10).split('-').reverse().join('.');if(typeof v==='string'&&v.startsWith("'="))return v.slice(1);return v==null?'':String(v);}
class Range{
 constructor(s,r,c,n=1,m=1){Object.assign(this,{s,r,c,n,m});}
 map(fn){return Array.from({length:this.n},(_,i)=>Array.from({length:this.m},(_,j)=>fn(this.s.rows[this.r+i-1]?.[this.c+j-1])));}
 getDisplayValues(){return this.map(display);}getValues(){return this.map(v=>v&&v.date?new Date(v.date):typeof v==='string'&&v.startsWith("'=")?v.slice(1):v??'');}
 getFormulas(){return this.map(v=>typeof v==='string'&&v.startsWith('=')?v:'');}
 getValue(){return this.getValues()[0][0];}getFormula(){return this.getFormulas()[0][0];}
 canEdit(){return this.s.writable&&!Array.from({length:this.n},(_,i)=>Array.from({length:this.m},(_,j)=>`${this.r+i},${this.c+j}`)).flat().some(k=>this.s.protected.has(k));}
 setValue(v){if(!this.canEdit())throw Error('write denied');if(this.s.failOnce===`${this.r},${this.c}`){this.s.failOnce=null;throw Error('simulated write failure');}while(this.s.rows.length<this.r)this.s.rows.push([]);this.s.rows[this.r-1][this.c-1]=v;this.s.writes++;return this;}
 setValues(values){if(!this.canEdit())throw Error('write denied');if(this.s.failBatch==='before'){this.s.failBatch=null;throw Error('batch failed');}for(let i=0;i<this.n;i++){while(this.s.rows.length<this.r+i)this.s.rows.push([]);for(let j=0;j<this.m;j++)this.s.rows[this.r+i-1][this.c+j-1]=values[i][j];}this.s.writes++;if(this.s.failBatch==='after'){this.s.failBatch=null;throw Error('batch response lost');}return this;}
 clearContent(){if(this.s.failClear)throw Error('clear failed');return this.setValues(Array.from({length:this.n},()=>Array(this.m).fill('')));}
}
class Sheet{
 constructor(x){this.name=x.name;this.rows=x.rows;this.writable=true;this.denied=false;this.protected=new Set();this.writes=0;this.maxColumns=Math.max(...x.rows.map(r=>r.length));this.maxRows=x.rows.length;}
 getName(){return this.name;}getRange(...args){return new Range(this,...args);}
 getLastRow(){let n=this.rows.length;while(n&&!this.rows[n-1].some(v=>v!==null&&v!==''&&v!==undefined))n--;return n;}
 getLastColumn(){return Math.max(0,...this.rows.map(row=>{let n=row.length;while(n&&(row[n-1]===null||row[n-1]===''||row[n-1]===undefined))n--;return n;}));}
 getMaxColumns(){return this.maxColumns;}insertColumnAfter(){this.maxColumns++;}
 getMaxRows(){return this.maxRows;}insertRowsAfter(after,n){this.maxRows+=n;while(this.rows.length<this.maxRows)this.rows.push([]);}
}
const sheets=fixtures.map(x=>new Sheet(x));
const ids=['MAIN_FAKE_ID_0123456789','BKR_FAKE_ID_0123456789'];
const props={MAIN_SPREADSHEET_URL:'https://docs.google.com/spreadsheets/d/'+ids[0]+'/edit',BKR_SPREADSHEET_URL:'https://docs.google.com/spreadsheets/d/'+ids[1]+'/edit'};
const ctx=vm.createContext({Date,console:{log(){}},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null})},
 SpreadsheetApp:{openById:id=>{const i=ids.indexOf(id);if(i<0||sheets[i].denied)throw Error('read denied');return {getSheets:()=>[sheets[i]],getSheetByName:n=>sheets[i].name===n?sheets[i]:null,getSpreadsheetTimeZone:()=> 'Asia/Qyzylorda'};},flush(){}},
 LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},Utilities:{getUuid:()=>crypto.randomUUID(),DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(alg,str)=>[...crypto.createHash('sha256').update(str).digest()],base64EncodeWebSafe:bytes=>Buffer.from(bytes).toString('base64url'),formatDate:date=>date.toISOString().slice(0,10).split('-').reverse().join('.'),parseDate:str=>{const [d,m,y]=str.split('.').map(Number);return new Date(Date.UTC(y,m-1,d));}}});
vm.runInContext(fs.readFileSync(require('path').join(__dirname,'../Code.gs'),'utf8'),ctx);
expectError(()=>ctx.loadPersonnel(),'initializeEmployeeIds');
const setup=ctx.initializeEmployeeIds();assert(setup.people===79&&setup.assigned===160,'ID setup and main-table list');
assert(ctx.initializeEmployeeIds().assigned===0,'ID setup is idempotent');
let result=ctx.loadPersonnel();assert(result.people.length===79,'Main-table records loaded');assert(result.people.filter(p=>p.hasBase&&p.hasAnalysis).length===79,'Shared IDs join 79 people');
assert(result.unlinkedAnalysisCount===2,'BKR-only records counted without inventing main-table names');
assert(result.people.every(p=>!('analysisName'in p)&&!('start2026'in p)&&!('end2026'in p)),'BKR name and level fields not returned');
const person=result.people[0],id=person.id;
const save=(source,changes,p=ctx.loadPersonnel().people.find(x=>x.id===id))=>ctx.savePersonnel({source,id,revision:source==='main'?p.revisionMain:p.revisionBkr,changes});
sheets.forEach(s=>s.writable=false);result=ctx.loadPersonnel();assert(result.people.every(p=>!p.canEditMain&&!p.canEditBkr),'Reader sees no edit controls');expectError(()=>save('main',{email:'test@example.org'}),'Нет права');sheets.forEach(s=>s.writable=true);
const bkrBefore=JSON.stringify(sheets[1].rows);
save('main',{email:'test@example.org'},person);assert(ctx.loadPersonnel().people.find(p=>p.id===id).email==='test@example.org','Save writes source');assert(JSON.stringify(sheets[1].rows)===bkrBefore,'Other source preserved');
expectError(()=>save('main',{email:'stale@example.org'},person),'уже изменена');
const beforeInvalid=sheets[0].writes;expectError(()=>save('main',{email:'should-not-save@example.org',birth:'31.02.2001'}),'Некорректная дата');assert(sheets[0].writes===beforeInvalid,'Validate before any write');
save('main',{birth:'01.02.1990'});assert(ctx.loadPersonnel().people.find(p=>p.id===id).birth==='01.02.1990','Typed birth date');
expectError(()=>save('main',{unknownField:'value'}),'недоступно');
// Sorting the source preserves identity and permits a safe save by stable ID.
const prefix=sheets[0].rows.slice(0,4),data=sheets[0].rows.slice(4).reverse();sheets[0].rows=prefix.concat(data);save('main',{email:'sorted@example.org'});assert(ctx.loadPersonnel().people.find(p=>p.id===id).email==='sorted@example.org','Save after sort');
let table=ctx.readTable_('main'),record=table.rows.find(r=>r.id===id);
const noteCell=table.sheet.getRange(record.row,table.cols.mainNote+1);const oldNote=noteCell.getValue();noteCell.setValue('="computed"');expectError(()=>save('main',{mainNote:'overwrite'}),'формула');noteCell.setValue(oldNote);
table=ctx.readTable_('main');record=table.rows.find(r=>r.id===id);const emailBefore=record.values.email;
sheets[0].failOnce=`${record.row},${table.cols.mainNote+1}`;expectError(()=>save('main',{email:'partial@example.org',mainNote:'new note'}),'восстановлены');assert(ctx.loadPersonnel().people.find(p=>p.id===id).email===emailBefore,'Rollback restores earlier writes');
save('main',{mainNote:'=SUM(1,2)'});table=ctx.readTable_('main');record=table.rows.find(r=>r.id===id);assert(record.values.mainNote==='=SUM(1,2)'&&!record.formulas.mainNote,'Input stays literal text');
sheets[0].failOnce=`${record.row},${table.cols.email+1}`;expectError(()=>save('main',{mainNote:'replacement',email:'rollback@example.org'}),'восстановлены');table=ctx.readTable_('main');record=table.rows.find(r=>r.id===id);assert(record.values.mainNote==='=SUM(1,2)'&&!record.formulas.mainNote,'Rollback preserves literal text without formula execution');
const bkrVersion=ctx.loadPersonnel().people.find(p=>p.id===id);sheets[0].denied=true;expectError(()=>save('bkr',{plan2027:'forbidden'},bkrVersion),'Не удалось открыть');sheets[0].denied=false;
table=ctx.readTable_('main');const newRow=Array.from(sheets[0].rows[table.rows[0].row-1]);newRow[table.idCol]='';sheets[0].rows.push(newRow);const writes=sheets.reduce((n,s)=>n+s.writes,0);expectError(()=>ctx.initializeEmployeeIds(),'Неоднозначное ФИО');assert(sheets.reduce((n,s)=>n+s.writes,0)===writes,'Ambiguous names cause no writes');sheets[0].rows.pop();
const cleanResult=ctx.loadPersonnel();

// Unified editor validates both sources before writing and rolls back across files.
const snapshots=()=>sheets.map(s=>JSON.stringify(s.rows));
let p=ctx.loadPersonnel().people.find(p=>p.id===id);
const update=sections=>ctx.updatePersonnel({id,sections});
let versions={main:{revision:p.revisionMain,changes:{email:'unified@example.org'}},bkr:{revision:p.revisionBkr,changes:{plan2027:'Единая форма'}}};
update(versions);p=ctx.loadPersonnel().people.find(p=>p.id===id);assert(p.email==='unified@example.org'&&p.plan2027==='Единая форма','Unified form saves both sources');
let bt=ctx.readTable_('bkr'),br=bt.rows.find(r=>r.id===id);bt.sheet.getRange(br.row,bt.cols.goal+1).setValue('Changed externally');
let before=snapshots();expectError(()=>update({main:{revision:p.revisionMain,changes:{email:'must-not-write@example.org'}},bkr:{revision:p.revisionBkr,changes:{plan2027:'stale'}}}),'уже изменена');assert(JSON.stringify(snapshots())===JSON.stringify(before),'Cross-source stale preflight prevents all writes');
p=ctx.loadPersonnel().people.find(p=>p.id===id);sheets[1].failOnce=`${br.row},${bt.cols.plan2027+1}`;before=snapshots();
expectError(()=>update({main:{revision:p.revisionMain,changes:{email:'rollback-new@example.org'}},bkr:{revision:p.revisionBkr,changes:{plan2027:'failure'}}}),'восстановлены');assert(JSON.stringify(snapshots())===JSON.stringify(before),'Unified save rollback spans both tables');
p=ctx.loadPersonnel().people.find(p=>p.id===id);update({main:{revision:p.revisionMain,changes:{name:'ФИО из основного списка'}}});
assert(ctx.loadPersonnel().people.find(p=>p.id===id).name==='ФИО из основного списка','Main name remains authoritative when BKR spelling differs');
expectError(()=>save('bkr',{analysisName:'Wrong source'}),'недоступно');expectError(()=>save('bkr',{start2026:'9'}),'недоступно');
// Creation needs only the name and writes one generated identity into both sources.
const request=values=>({requestId:crypto.randomUUID(),values});
const countWrites=()=>sheets.reduce((n,s)=>n+s.writes,0);
let req=request({name:'Новый участник'}),created=ctx.createPersonnel(req);
assert(created.created&&/^EMP-/.test(created.id),'Automatically assigned ID');
const firstRows=['main','bkr'].map(k=>ctx.readTable_(k).rows.find(r=>r.id===created.id));
assert(firstRows.every(Boolean)&&firstRows[0].values.name===firstRows[1].values.analysisName,'Name and same ID written to both sources');
assert(firstRows[0].values.email===''&&firstRows[1].values.goal==='','Only name is mandatory');
let writesBefore=countWrites();assert(!ctx.createPersonnel(req).created&&countWrites()===writesBefore,'Retry cannot create duplicate records');
for(const invalid of [{name:' '},{name:'Bad',birth:'31.02.2000'},{name:'Bad',start2026:'4'},{name:'Bad',analysisName:'ignored?'}]){writesBefore=countWrites();expectError(()=>ctx.createPersonnel(request(invalid)));assert(countWrites()===writesBefore,'Invalid creation causes no writes');}
sheets[1].writable=false;writesBefore=countWrites();expectError(()=>ctx.createPersonnel(request({name:'Forbidden'})),'Редактор');assert(countWrites()===writesBefore,'Both editor permissions required');sheets[1].writable=true;
req=request({name:'Полная форма',region:'Новый город',email:'test@example.org',birth:'01.03.1995',level:'3',dec2025:'2',mainNote:'=SUM(1,2)',goal:'Цель',plan2027:'План'});created=ctx.createPersonnel(req);
let added=ctx.loadPersonnel().people.find(p=>p.id===created.id);assert(added.region===added.analysisRegion&&added.birth==='01.03.1995'&&added.goal==='Цель','Optional fields routed to their source; initial region mirrored');
assert(ctx.readTable_('main').rows.find(r=>r.id===created.id).formulas.mainNote==='','New text beginning equals cannot run as formula');
for(const failure of ['before','after']){
 req=request({name:'Rollback '+failure});const counts=['main','bkr'].map(k=>ctx.readTable_(k).rows.length);sheets[1].failBatch=failure;
 expectError(()=>ctx.createPersonnel(req),'отменены');assert(['main','bkr'].every((k,i)=>ctx.readTable_(k).rows.length===counts[i]),'Creation rollback removes only new data on '+failure+' failure');
 assert(ctx.createPersonnel(req).created,'Same request safely retries after successful rollback');
}
req=request({name:'Partial write'});sheets[1].failBatch='before';sheets[0].failClear=true;expectError(()=>ctx.createPersonnel(req),'часть данных');sheets[0].failClear=false;
expectError(()=>ctx.createPersonnel(req),'незавершённая');
const partial=ctx.readTable_('main'),partialRow=partial.rows.find(r=>r.id==='EMP-'+req.requestId);partial.sheet.getRange(partialRow.row,1,1,partial.width).clearContent();
console.log('Passed: main-only identity/levels, unchanged legacy fields, stable IDs, permissions, unified editor, stale preflight and cross-file rollback; name-only/full creation, mirrored identity, safe types, retries without duplicates and partial-failure handling.');
