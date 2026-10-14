import {chromium} from '@playwright/test';
import XLSX from 'xlsx';
import assert from 'node:assert/strict';
import {DAYS} from '../src/engine.js';
const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['海南大学 测试人员 学生个人课表'],['学年学期：2026-2027-1'],['',...DAYS],['1、2节\n07:40-09:20','1-5(周)[01-02节]']]),'课表');
const browser=await chromium.launch({channel:'msedge',headless:true});
try{
 const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto('http://127.0.0.1:8765');await page.locator('#enter').click();
 await page.locator('#files').setInputFiles({name:'学生个人课表_99999999.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:XLSX.write(book,{type:'buffer',bookType:'xlsx'})});
 await page.getByText('参与 · 信息提醒 1 条',{exact:true}).waitFor();assert.match(await page.locator('#status').innerText(),/上课时间已保留/);
 await page.locator('.detail').click();await page.getByText('课程信息提醒（不影响空课表与排班）',{exact:true}).waitFor();assert.equal(await page.locator('[name=confirmed]').count(),0);await page.locator('#cancel').click();
 await page.locator('[data-view=free]').click();assert.equal(await page.locator('.cell[data-key="0:0"]').innerText(),'无');assert.match(await page.locator('.cell[data-key="1:0"]').innerText(),/测试人员/);
 await page.locator('[data-view=schedule]').click();await page.locator('#generate').click();await page.getByText('排班完成',{exact:true}).waitFor();assert.equal(await page.locator('.cell[data-key="0:0"] .tag').count(),0);assert.equal(await page.locator('.cell[data-key="1:0"] .tag').count(),1);
 assert.deepEqual(errors,[]);console.log('PASS: partial metadata import, persistent warning, no confirmation, correct availability and schedule');
}finally{await browser.close()}
