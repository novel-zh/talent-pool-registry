const fs=require('fs'),vm=require('vm');
const context=vm.createContext({});
vm.runInContext(fs.readFileSync(require('path').join(__dirname,'../Code.gs'),'utf8'),context);
const schema=vm.runInContext('SCHEMA_',context);
module.exports=['main','bkr'].map((kind,index)=>{
 const fields=Object.keys(schema[kind].fields);
 const headers=fields.map(k=>schema[kind].fields[k][0]);
 const prefix=index===0?[['Synthetic test register'],[],[],headers]:[headers];
 const rows=Array.from({length:index===0?79:81},(_,i)=>fields.map(k=>{
  if(k==='name'||k==='analysisName')return 'Synthetic Person '+String(i+1).padStart(3,'0');
  if(k==='region'||k==='analysisRegion')return 'Example Region';
  if(k==='email')return 'person'+(i+1)+'@example.invalid';
  if(k==='number'||k==='analysisNumber')return String(i+1);
  return '';
 }));
 return {name:index===0?'Main':'Analysis',rows:[...prefix,...rows]};
});
