import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {newState,parse,solve,conflict,templateExport} from '../src/engine.js';
import {unzipSync} from 'fflate';
const sample=process.env.SAMPLE;
if(sample){const p=parse(readFileSync(sample),'学生个人课表_20243006951.xls','运营采编部');assert.equal(p.name,'陈锦瀚');assert.equal(p.courses.length,27);assert.equal(p.errors.length,0)}
const s=newState();s.people=['a','b','c'].map(id=>({id,name:id,campus:'海甸',department:'办公室',enabled:true,courses:[]}));s.demands={'0:0':1,'1:0':1,'2:0':1};
let r=solve(s,1);assert.equal(new Set(Object.values(r.cells).flat()).size,3);
s.rules=[{kind:'fixed',people:['a','b'],day:0,slot:0}];s.demands['0:0']=2;r=solve(s,1);assert.deepEqual(new Set(r.cells['0:0']),new Set(['a','b']));
s.rules.push({kind:'exclude',people:['a'],day:0,slot:0});assert.throws(()=>solve(s,1),/固定人员/);
s.rules=[];s.slots=[['甲','08:00-10:00'],['乙','09:00-11:00']];s.demands={'0:0':1,'0:1':1};r=solve(s,1);assert.notEqual(r.cells['0:0'][0],r.cells['0:1'][0]);for(const k in r.cells)assert.equal(conflict(s,1,k,r.cells[k],r.cells),false);
s.demands={};assert.deepEqual(solve(s,1).cells,{});
const original=new Uint8Array(readFileSync('public/assets/free_template.xlsx'));globalThis.fetch=async()=>({ok:true,arrayBuffer:async()=>original.buffer.slice(original.byteOffset,original.byteOffset+original.byteLength)});
const out=await templateExport(s,[1,2,3],true),a=unzipSync(original),b=unzipSync(out.bytes);for(const k in a)if(!/^xl\/worksheets\/sheet[12]\.xml$/.test(k))assert.deepEqual(b[k],a[k],k);
assert.match(new TextDecoder().decode(b['xl/worksheets/sheet1.xml']),/a \(1-3\)/);console.log('PASS: parser, balance, multi-person fixed rules, exclusions, overlap, zero demand, template preservation');
