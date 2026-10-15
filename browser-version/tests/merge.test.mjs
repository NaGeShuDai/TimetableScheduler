import {readFileSync} from 'node:fs';
import assert from 'node:assert/strict';
import {mergeFreeTables,templateExport} from '../src/engine.js';
import {chromium} from '@playwright/test';
const template=new Uint8Array(readFileSync('public/assets/free_template.xlsx'));
globalThis.fetch=async()=>({ok:true,arrayBuffer:async()=>template.buffer.slice(template.byteOffset,template.byteOffset+template.byteLength)});
const a=await templateExport({people:[]},[],false,{'海甸|1|0|0':'甲 (1-3)\n乙 (2)'}),b=await templateExport({people:[]},[],false,{'海甸|1|0|0':'甲 (3-5)','观澜湖|0|2|4':'丙 (1)'});
const merged=mergeFreeTables([{name:'a.xlsx',buffer:a.bytes},{name:'b.xlsx',buffer:b.bytes}]);assert.match(merged['海甸|1|0|0'],/甲 \(1-5\)/);assert.match(merged['海甸|1|0|0'],/乙 \(2\)/);
const browser=await chromium.launch({channel:'msedge',headless:true});try{const page=await browser.newPage();await page.goto('http://127.0.0.1:8765');await page.locator('[data-view=free]').click();const chooser=page.waitForEvent('filechooser');await page.locator('#mergefree').click();await(await chooser).setFiles([{name:'a.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(a.bytes)},{name:'b.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(b.bytes)}]);await page.getByText('汇总空课表 · 2 份',{exact:true}).waitFor();assert.match(await page.locator('#dialog-body').innerText(),/甲 \(1-5\)/);const downloaded=page.waitForEvent('download');await page.getByRole('button',{name:'导出汇总空课表',exact:true}).click();assert.equal((await downloaded).suggestedFilename(),'汇总空课表.xlsx');console.log('PASS: multiple-file import, deduplication, union of weeks, preview and template export')}finally{await browser.close()}
