import * as XLSX from 'xlsx';
import solver from 'javascript-lp-solver';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
export const DAYS=['星期一','星期二','星期三','星期四','星期五','星期六','星期日'];
export const DEPTS=['办公室','运营采编部','技术部','设计部'];
export function mergeFreeTables(files){
 const merged=new Map();let mode=null;
 for(const file of files){const book=XLSX.read(file.buffer,{type:'array'}),campuses=book.SheetNames.filter(n=>['海甸','观澜湖'].includes(n));if(!campuses.length)throw Error(file.name+'：缺少校区工作表');
 for(const campus of campuses){const rows=XLSX.utils.sheet_to_json(book.Sheets[campus],{header:1,defval:''});for(let s=0;s<5;s++)for(let dep=0;dep<4;dep++){const row=1+s*4+dep;if(String(rows[row]?.[1]).trim()!==DEPTS[dep])throw Error(file.name+'：不符合当前四部门模板');for(let d=0;d<7;d++)for(const raw of String(rows[row]?.[d+2]||'').split(/\r?\n/)){const line=raw.trim();if(!line||['无','—','-'].includes(line))continue;const match=line.match(/^(.+?)\s*[（(]([\d\s,，、\-－—]+)[）)]$/),type=match?'total':'weekly';if(mode&&mode!==type)throw Error('不能混合周空课表和带周次的总空课表');mode=type;const name=match?match[1].trim():line,key=[campus,dep,s,d].join('|');if(!merged.has(key))merged.set(key,new Map());const names=merged.get(key);if(!names.has(name))names.set(name,new Set());if(match)weeks(match[2]).forEach(w=>names.get(name).add(w))}}}}
 if(!merged.size)throw Error('未读取到名单');return Object.fromEntries([...merged].map(([key,names])=>[key,[...names].sort(([a],[b])=>a.localeCompare(b,'zh')).map(([name,ws])=>mode==='total'?name+' ('+compact([...ws])+')':name).join('\n')]));
}
export const SLOTS=[['1、2节','07:40-09:20'],['3、4节','09:45-11:25'],['5、6节','14:30-16:10'],['7、8节','16:35-18:15'],['9、10、11节','19:20-21:55'],['中午节次(一)','11:35-13:05'],['中午节次(二)','13:05-14:20']];
export function weeks(text){const out=new Set();for(const p of String(text).replace(/[，、]/g,',').replace(/[－—]/g,'-').split(',')){const m=p.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);if(!m)throw Error('周次格式无效');const a=+m[1],b=+(m[2]||m[1]);if(a<1||b>60||a>b)throw Error('周次必须在 1～60 范围内');for(let w=a;w<=b;w++)out.add(w)}return [...out].sort((a,b)=>a-b)}
export function compact(ws){const a=[...new Set(ws)].sort((a,b)=>a-b),g=[];for(const w of a){if(g.length&&g.at(-1)[1]+1===w)g.at(-1)[1]=w;else g.push([w,w])}return g.map(([a,b])=>a===b?`${a}`:`${a}-${b}`).join(', ')}
export function time(t){const m=t.match(/^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/);if(!m||+m[1]>23||+m[3]>23||+m[2]>59||+m[4]>59)throw Error('时间格式为 HH:MM-HH:MM');const a=+m[1]*60 + +m[2],b=+m[3]*60 + +m[4];if(a>=b)throw Error('结束时间必须晚于开始时间');return[a,b]}
export function clock(c){return `${String(Math.floor(c.start/60)).padStart(2,'0')}:${String(c.start%60).padStart(2,'0')}-${String(Math.floor(c.end/60)).padStart(2,'0')}:${String(c.end%60).padStart(2,'0')}`}
export function parseCourseCell(raw,day,start,end,cell){
 const courses=[],errors=[],warnings=[];
 if(!Number.isInteger(day)||day<0||day>6||!Number.isFinite(start)||!Number.isFinite(end)||start<0||end>1440||start>=end)return{courses,errors:['无法确定有效的星期或上课时间'],warnings};
 const normalized=String(raw).replace(/\r\n?/g,'\n').trim();
 const suspiciousMarker=/[（(]\s*(?:单|双)?周\s*[）)]|^\s*\d[^\n]*[（(][^\n]*周|[\[【][^\n]*节/m;
 // Blank lines delimit records; the week/period marker delimits each record's fields.
 for(const block of normalized.split(/\n[ \t]*\n+/).filter(s=>s.trim())){
  const lines=block.split('\n').map(s=>s.trim()).filter(Boolean);
  const markers=lines.map((line,index)=>{
   const canonical=line.replace(/（/g,'(').replace(/）/g,')').replace(/【/g,'[').replace(/】/g,']');
   return {index,match:canonical.match(/^([\d\s,，、\-－—]+)\s*\(([^)]*周[^)]*)\)\s*(?:\[([^\]]*节)\])?\s*$/)};
  }).filter(m=>m.match);
  if(markers.length)for(let i=0;i<lines.length;i++)if(suspiciousMarker.test(lines[i])&&!markers.some(m=>m.index===i))errors.push('存在无法识别的周次与节次行，需要核实');
  if(!markers.length){
   if(suspiciousMarker.test(block)||!suspiciousMarker.test(normalized))errors.push('无法识别周次与节次行');
   else warnings.push('课程附加信息缺少完整结构，已忽略；已识别的上课时间仍参与统计');
   continue;
  }
  for(const {index,match} of markers)try{
   const info=lines.slice(0,index);
   if(match[3]!==undefined&&!/^\s*\d+(?:[\s,，、\-－—]+\d+)*\s*节$/.test(match[3]))throw Error('节次格式无效，需要核实');
   let ws=weeks(match[1]);if(match[2].includes('单'))ws=ws.filter(w=>w%2);if(match[2].includes('双'))ws=ws.filter(w=>!(w%2));
   if(!ws.length)throw Error('没有有效教学周');
   const incomplete=match[3]===undefined||markers.length>1||info.length<2||index===lines.length-1;
   if(match[3]===undefined)warnings.push('周次行缺少节次标记，已使用左侧行标题的起止时间；仍正常参与空课表与排班');
   if(incomplete)warnings.push('课程信息格式不完整或多门课程边界不清，名称、教师和地点留空；上课时间已保留，无需确认即可参与统计与排班');
   const teacher=incomplete?'':info.at(-1);
   const name=incomplete?'':info.slice(0,-1).join('\n');
   courses.push({name,teacher,information:incomplete?'':info.join('\n'),weeks:ws,day,start,end,location:incomplete?'':lines.slice(index+1).join('\n'),raw:block,cell,time_only:incomplete});
  }catch(e){errors.push(e.message)}
 }
 return {courses,errors,warnings};
}
export function parse(buffer,filename,department){
 const b=XLSX.read(buffer,{type:'array'}),rows=XLSX.utils.sheet_to_json(b.Sheets[b.SheetNames[0]],{header:1,defval:''}).map(r=>r.map(v=>String(v).trim()));
 const h=rows.findIndex(r=>r.includes('星期一'));if(h<0||!DAYS.every(d=>rows[h].includes(d)))throw Error('未找到完整星期表头');
 const meta=rows[1].join(' '),p={id:filename.match(/\d{8,}/)?.[0]||rows[0][0].match(/\s+(\S+)\s+学生个人课表/)?.[1]||filename,name:rows[0][0].match(/\s+(\S+)\s+学生个人课表/)?.[1]||filename,term:meta.match(/学年学期[：:]\s*(\S+)/)?.[1]||'',group:meta.match(/班级[：:]\s*(.*?)\s+(?:专业|院系)[：:]/)?.[1]||'',campus:'海甸',department,enabled:true,errors:[],errors_confirmed:false,courses:[],source:filename};
 for(let r=h+1;r<rows.length;r++){
  const m=rows[r][0].match(/\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/);if(!m)continue;
  const[start,end]=time(m[0].replace(/\s/g,''));
  DAYS.forEach((d,day)=>{
   const col=rows[h].indexOf(d),raw=rows[r][col]||'';if(!raw)return;
   const result=parseCourseCell(raw,day,start,end,XLSX.utils.encode_cell({r,c:col}));
   p.courses.push(...result.courses);p.errors.push(...result.errors.map(message=>d+' 第'+(r+1)+'行：'+message));
   p.warnings??=[];p.warnings.push(...result.warnings.map(message=>d+' 第'+(r+1)+'行：'+message));
  });
 }
 if(!p.courses.length)throw Error('没有识别到课程');return p;
}
export function eligible(p){return p.enabled&&(!p.errors?.length||p.errors_confirmed)}
export function free(p,w,d,s,slots=SLOTS){const[a,b]=time(slots[s][1]);return eligible(p)&&!p.courses.some(c=>c.day===d&&c.weeks.includes(w)&&a<c.end&&c.start<b)}
export function members(r){return r.people||[r.person]}
export function matches(r,w,d,s){return r.enabled!==false&&(!r.weeks?.length||r.weeks.includes(w))&&(r.day==null||r.day===d)&&(r.slot==null||r.slot===s)}
export function candidates(state,w,d,s){const excluded=new Set(state.rules.filter(r=>r.kind==='exclude'&&matches(r,w,d,s)).flatMap(members));return state.people.filter(p=>free(p,w,d,s,state.slots)&&!excluded.has(p.id)).map(p=>p.id)}
export function fixed(state,w,d,s){return [...new Set(state.rules.filter(r=>r.kind==='fixed'&&matches(r,w,d,s)).flatMap(members))]}
export function conflict(state,w,key,ids,cells){const[d,s]=key.split(':').map(Number),pool=candidates(state,w,d,s),[a,b]=time(state.slots[s][1]);return new Set(ids).size!==ids.length||ids.some(id=>!pool.includes(id))||fixed(state,w,d,s).some(id=>!ids.includes(id))||Object.entries(cells).some(([k,other])=>{if(k===key)return false;const[od,os]=k.split(':').map(Number),[x,y]=time(state.slots[os][1]);return od===d&&a<y&&x<b&&ids.some(id=>other.includes(id))})}
export function improveFairness(state,w,original,budgetMs=2000){
 const cells=structuredClone(original),keys=Object.keys(cells),pools=Object.fromEntries(keys.map(k=>[k,candidates(state,w,...k.split(':').map(Number))]));
 const ids=state.people.filter(p=>keys.some(k=>pools[k].includes(p.id))).map(p=>p.id),preferences=state.rules.filter(r=>r.kind==='prefer'&&r.enabled!==false);
 const score=()=>{const loads=Object.fromEntries(ids.map(id=>[id,[0,0,0,0,0,0,0,0,0,0]]));const priority=preferences.map(()=>0);for(const k of keys){const[d,s]=k.split(':').map(Number),[a,b]=time(state.slots[s][1]);for(const id of cells[k]){loads[id][0]+=b-a;loads[id][1]+=a>=1080?1:0;loads[id][2]+=d>=5?1:0;loads[id][3+d]++;preferences.forEach((r,i)=>{if(members(r).includes(id)&&matches(r,w,d,s))priority[i]--})}}const differences=[0,0,0];for(let i=0;i<ids.length;i++)for(let j=i+1;j<ids.length;j++)for(let n=0;n<3;n++)differences[n]+=Math.abs(loads[ids[i]][n]-loads[ids[j]][n]);return [...priority,...differences,ids.reduce((sum,id)=>sum+loads[id].slice(3).reduce((n,v)=>n+Math.max(0,v-1),0),0)]};
 const better=(a,b)=>{for(let i=0;i<a.length;i++){if(a[i]!==b[i])return a[i]<b[i]}return false};
 let current=score();const deadline=Date.now()+budgetMs;
 while(Date.now()<deadline){let changed=false;
  for(const k of keys){for(let i=0;i<cells[k].length;i++){const old=cells[k][i];for(const id of pools[k]){if(Date.now()>=deadline)return cells;if(id===old||cells[k].includes(id))continue;cells[k][i]=id;if(!conflict(state,w,k,cells[k],cells)){const next=score();if(better(next,current)){current=next;changed=true;break}}cells[k][i]=old}}}
  if(changed)continue;
  outer:for(let a=0;a<keys.length;a++)for(let b=a+1;b<keys.length;b++){const x=keys[a],y=keys[b];for(let i=0;i<cells[x].length;i++)for(let j=0;j<cells[y].length;j++){if(Date.now()>=deadline)return cells;const p=cells[x][i],q=cells[y][j];if(p===q||!pools[x].includes(q)||!pools[y].includes(p))continue;cells[x][i]=q;cells[y][j]=p;if(!conflict(state,w,x,cells[x],cells)&&!conflict(state,w,y,cells[y],cells)){const next=score();if(better(next,current)){current=next;changed=true;break outer}}cells[x][i]=p;cells[y][j]=q}}
  if(!changed)break;
 }
 return cells;
}
export function solve(state,w,timeoutMs=60000,onProgress=()=>{}){const model={optimize:'obj',opType:'max',constraints:{},variables:{},ints:{},options:{timeout:timeoutMs}},keys=Object.keys(state.demands).filter(k=>state.demands[k]>0),vars={},pools={};if(!keys.length)return{cells:{},notes:[]};let serial=0;for(const k of keys){const[d,s]=k.split(':').map(Number),pool=candidates(state,w,d,s);pools[k]=pool;model.constraints['slot'+k]={max:state.demands[k]};const forced=fixed(state,w,d,s);if(forced.length>state.demands[k]||forced.some(id=>!pool.includes(id)))throw Error(`${DAYS[d]} ${state.slots[s][0]}：固定人员不可用或超过人数`);for(const id of pool){const n='x'+serial++;vars[n]={key:k,id};model.variables[n]={obj:1,['slot'+k]:1};model.constraints[n]={min:forced.includes(id)?1:0,max:1};model.variables[n][n]=1;model.ints[n]=1}}
for(const p of state.people){for(let i=0;i<keys.length;i++)for(let j=i+1;j<keys.length;j++){const a=keys[i],b=keys[j],[ad,as]=a.split(':').map(Number),[bd,bs]=b.split(':').map(Number),[x,y]=time(state.slots[as][1]),[u,v]=time(state.slots[bs][1]);if(ad===bd&&x<v&&u<y){const ns=Object.keys(vars).filter(n=>vars[n].id===p.id&&(vars[n].key===a||vars[n].key===b));if(ns.length===2){const c='overlap'+serial++;model.constraints[c]={max:1};ns.forEach(n=>model.variables[n][c]=1)}}}}
const objectives=[];for(const r of state.rules)if(r.kind==='prefer'&&r.enabled!==false)objectives.push(Object.fromEntries(Object.entries(vars).map(([n,v])=>[n,members(r).includes(v.id)&&matches(r,w,...v.key.split(':').map(Number))?1:0])));objectives.push(Object.fromEntries(Object.keys(vars).map(n=>[n,1])));
const bound=keys.reduce((a,k)=>a+state.demands[k]*(time(state.slots[+k.split(':')[1]][1])[1]-time(state.slots[+k.split(':')[1]][1])[0]),0);
const active=state.people.filter(p=>Object.values(vars).some(v=>v.id===p.id));const fairness={};for(let i=0;i<active.length;i++)for(let j=i+1;j<active.length;j++){const n='diff'+serial++;model.variables[n]={obj:0};fairness[n]=-1;for(const sign of [1,-1]){const c='balance'+serial++;model.constraints[c]={max:0};model.variables[n][c]=-1;for(const[x,v]of Object.entries(vars)){const[a,b]=time(state.slots[+v.key.split(':')[1]][1]);if(v.id===active[i].id)model.variables[x][c]=sign*(b-a);if(v.id===active[j].id)model.variables[x][c]=-sign*(b-a)}}}objectives.push(fairness);
// Each subsequent objective preserves the result of the preceding priority.
for(const kind of ['evening','weekend']){const objective={};for(let i=0;i<active.length;i++)for(let j=i+1;j<active.length;j++){const n='special'+serial++;model.variables[n]={obj:0};objective[n]=-1;for(const sign of [1,-1]){const c='specialBalance'+serial++;model.constraints[c]={max:0};model.variables[n][c]=-1;for(const[x,v]of Object.entries(vars)){const[d,s]=v.key.split(':').map(Number);if(kind==='evening'?time(state.slots[s][1])[0]>=18*60:d>=5){if(v.id===active[i].id)model.variables[x][c]=sign;if(v.id===active[j].id)model.variables[x][c]=-sign}}}}objectives.push(objective)}
const repeated={};for(const p of active)for(let d=0;d<7;d++){const ns=Object.keys(vars).filter(n=>vars[n].id===p.id&&+vars[n].key.split(':')[0]===d);if(ns.length>1){const n='repeat'+serial++,c='repeatBound'+serial++;model.variables[n]={obj:0,[c]:-1};model.constraints[c]={max:1};ns.forEach(n=>model.variables[n][c]=1);repeated[n]=-1}}objectives.push(repeated);
let answer={},notes=[];const usedAux=new Set();for(let i=0;i<objectives.length;i++){const custom=state.rules.filter(r=>r.kind==='prefer'&&r.enabled!==false).length;onProgress(i<custom?'优化自定义优先规则':['满足班次人数','均衡总工时','均衡晚班','均衡周末','减少同日多班'][i-custom]);for(const[n,v]of Object.entries(model.variables))v.obj=objectives[i][n]||0;for(const n of Object.keys(objectives[i]))if(!vars[n])usedAux.add(n);const runModel=structuredClone(model),unused=Object.keys(model.variables).filter(n=>!vars[n]&&!usedAux.has(n));const removed=new Set(unused.flatMap(n=>Object.keys(model.variables[n]).filter(c=>c!=='obj'&&model.variables[n][c])));unused.forEach(n=>delete runModel.variables[n]);removed.forEach(c=>delete runModel.constraints[c]);for(const v of Object.values(runModel.variables))for(const c of removed)delete v[c];const started=Date.now(),result=solver.Solve(runModel);if(!result.feasible)throw Error('强制规则或重叠班次冲突，无法生成');if(!result.bounded)throw Error('排班模型异常');const next=Object.fromEntries(keys.map(k=>[k,Object.entries(vars).filter(([n,v])=>v.key===k&&(result[n]||0)>.5).map(([,v])=>v.id)]));const valid=Object.keys(vars).every(n=>Math.abs((result[n]||0)-Math.round(result[n]||0))<1e-6)&&keys.every(k=>next[k].length<=state.demands[k]&&!conflict(state,w,k,next[k],next));if(!valid){if(!Object.keys(answer).length)throw Error('计算未得到有效整数方案，请减少人数或规则后重试');notes.push('计算未完成，保留上一阶段有效方案');break}answer=next;if(Date.now()-started>=timeoutMs){notes.push('达到计算时间限制，当前方案可用，但未保证最优');break}const c='priority'+i;model.constraints[c]={equal:Math.round(result.result*100000)/100000};for(const[n,v]of Object.entries(model.variables))v[c]=objectives[i][n]||0}
if(notes.length){answer=improveFairness(state,w,answer);notes=['精确优化未完成，已进行受规则约束的公平性调整；当前为近似方案，未保证全局最优']}return{cells:answer,notes}}
export async function templateExport(state,ws,total,roster=null){const people=state.people.filter(eligible);if(!people.length&&!roster)throw Error('没有可导出的人员');if(people.some(p=>!['海甸','观澜湖'].includes(p.campus)||!DEPTS.includes(p.department)))throw Error('请先修正人员校区和部门');const response=await fetch('assets/free_template.xlsx');if(!response.ok)throw Error('无法读取模板');const zip=unzipSync(new Uint8Array(await response.arrayBuffer())),counts={};for(const[index,campus]of ['海甸','观澜湖'].entries()){const name=`xl/worksheets/sheet${index+1}.xml`;let xml=strFromU8(zip[name]);counts[campus]=0;for(let s=0;s<5;s++)for(let dep=0;dep<4;dep++)for(let d=0;d<7;d++){const coord=String.fromCharCode(67+d)+(2+s*4+dep),lines=roster?(roster[[campus,dep,s,d].join('|')]||'').split('\n').filter(Boolean):people.filter(p=>p.campus===campus&&p.department===DEPTS[dep]).map(p=>{const w=ws.filter(w=>free(p,w,d,s,SLOTS));return w.length?(total?`${p.name} (${compact(w)})`:p.name):''}).filter(Boolean);if(lines.length)counts[campus]++;const regex=new RegExp(`<c\\b[^>]*\\br="${coord}"[^>]*?(?:/>|>.*?</c>)`,'s');const old=xml.match(regex);if(!old)throw Error('模板单元格缺失');const style=old[0].match(/\bs="(\d+)"/)?.[1]||'0';const value=lines.join('\n').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');xml=xml.replace(regex,()=>`<c r="${coord}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${value}</t></is></c>`)}zip[name]=strToU8(xml)}if(!Object.values(counts).some(Boolean))throw Error('所选周次没有空闲人员');return{bytes:zipSync(zip),counts}}
export function newState(){return{people:[],rules:[],slots:structuredClone(SLOTS),custom_slots:structuredClone(SLOTS),slot_mode:'default',default_department:'办公室',demands:Object.fromEntries(Array.from({length:5},(_,d)=>Array.from({length:5},(_,s)=>[`${d}:${s}`,1])).flat()),schedules:{},threshold:100}}
