import assert from 'node:assert/strict';
import {readFileSync,readdirSync} from 'node:fs';
import {newState,parse,solve,improveFairness,conflict} from '../src/engine.js';
const s=newState();s.people=['a','b','c'].map(id=>({id,enabled:true,courses:[]}));
s.demands={'0:0':1,'1:0':1,'2:0':1};let cells=improveFairness(s,1,{'0:0':['a'],'1:0':['a'],'2:0':['a']});assert.equal(new Set(Object.values(cells).flat()).size,3);
s.rules=[{kind:'prefer',people:['a'],day:0,slot:0,enabled:true},{kind:'fixed',people:['a'],day:1,slot:0,enabled:true},{kind:'exclude',people:['c'],day:2,slot:0,enabled:true}];cells=improveFairness(s,1,{'0:0':['a'],'1:0':['a'],'2:0':['a']});assert.deepEqual(cells['0:0'],['a']);assert.deepEqual(cells['1:0'],['a']);assert.deepEqual(cells['2:0'],['b']);
if(process.env.CASE_DIR){const state=newState(),dir=process.env.CASE_DIR;state.people=readdirSync(dir).filter(f=>/\.xls$/i.test(f)).map(f=>parse(readFileSync(dir+'/'+f),f,'办公室'));
for(const week of [1,3]){const result=solve(state,week,500),assignments=Object.values(result.cells).flat();assert.equal(assignments.length,25);for(const k in result.cells)assert.equal(conflict(state,week,k,result.cells[k],result.cells),false);const counts=state.people.map(p=>assignments.filter(id=>id===p.id).length);assert.ok(counts.every(n=>n>=2&&n<=4),JSON.stringify(counts));console.log(`PASS: week ${week}, full coverage, valid assignments, balanced shifts ${counts.join(', ')}`)}}
console.log('PASS: timeout fairness fallback preserves fixed/exclude/preference constraints');
