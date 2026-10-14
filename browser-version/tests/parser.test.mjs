import assert from 'node:assert/strict';
import {readFileSync,readdirSync} from 'node:fs';
import XLSX from 'xlsx';
import {parse,parseCourseCell,DAYS,free,candidates,newState} from '../src/engine.js';
const raw='\n课程第一行\n（课程第二行）\n教师甲\n1-5,7（周）【01-02节】\n地点第一行\n地点第二行\n\n另一门课\n教师乙\n2-8(双周)[01-02节]\n网络课堂\n';
const result=parseCourseCell(raw,0,460,560,'B4');
assert.deepEqual(result.errors,[]);assert.equal(result.courses.length,2);
assert.equal(result.courses[0].name,'课程第一行\n（课程第二行）');assert.equal(result.courses[0].teacher,'教师甲');assert.equal(result.courses[0].location,'地点第一行\n地点第二行');assert.deepEqual(result.courses[1].weeks,[2,4,6,8]);
assert.equal(parseCourseCell('未知课程\n教师\n坏周次(周)[01-02节]\n地点',0,460,560,'B4').errors.length,1);
const ambiguous=parseCourseCell('课程\n教师\n1-2(周)[01-02节]\n地点\n下一门\n教师\n3-4(周)[01-02节]\n地点',0,460,560,'B4');
assert.deepEqual(ambiguous.errors,[]);assert.equal(ambiguous.courses.length,2);assert.equal(ambiguous.warnings.length,2);assert.ok(ambiguous.courses.every(c=>c.name===''&&c.teacher===''&&c.location===''));
const fallback=parseCourseCell('1-5(周)[01-02节]',0,460,560,'B4');assert.deepEqual(fallback.errors,[]);assert.equal(fallback.warnings.length,1);
const person={id:'test',enabled:true,courses:fallback.courses,errors:[],warnings:fallback.warnings};assert.equal(free(person,1,0,0),false);assert.equal(free(person,1,1,0),true);const state=newState();state.people=[person];assert.deepEqual(candidates(state,1,0,0),[]);assert.deepEqual(candidates(state,1,1,0),['test']);
assert.equal(parseCourseCell('1-5(周)[01-02节]',0,560,460,'B4').errors.length,1);
if(process.env.CASE_DIR){let total=0;for(const f of readdirSync(process.env.CASE_DIR).filter(f=>/\.xlsx?$/i.test(f))){const data=readFileSync(process.env.CASE_DIR+'/'+f),p=parse(data,f,'办公室'),book=XLSX.read(data,{type:'array'}),rows=XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1,defval:''}),header=rows.findIndex(r=>r.includes('星期一'));let expected=0;for(const row of rows.slice(header+1)){if(!/\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/.test(String(row[0])))continue;for(const d of DAYS){const cell=String(row[rows[header].indexOf(d)]||'');expected+=(cell.match(/\(.*?周.*?\)\[/g)||[]).length}}
assert.deepEqual(p.errors,[],f);assert.equal(p.courses.length,expected,f);total+=p.courses.length;console.log(`${f}: ${p.courses.length} course records, zero errors`)}console.log(`PASS: all special files, ${total} records`)}
console.log('PASS: multiline names and locations, multiple records, full-width markers, even weeks, malformed and ambiguous records');
