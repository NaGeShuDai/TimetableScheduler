import {chromium} from '@playwright/test';
import assert from 'node:assert/strict';
import {newState} from '../src/engine.js';
const state=newState();state.people=[{id:'test',name:'测试人员',enabled:true,courses:[],errors:[],campus:'海甸',department:'办公室'}];state.demands={'0:0':1};
const browser=await chromium.launch({channel:'msedge',headless:true});
try{
 const page=await browser.newPage();page.on('dialog',d=>d.accept());
 await page.addInitScript(data=>{localStorage.setItem('hnu-browser-scheduler-v1',JSON.stringify(data));window.Worker=class{postMessage(){this.timer=setTimeout(()=>{this.onmessage({data:{type:'progress',stage:'均衡总工时'}})},50);this.end=setTimeout(()=>this.onmessage({data:window.testError?{error:'测试错误'}:{cells:{'0:0':['test']},notes:[]}}),1800)}terminate(){clearTimeout(this.timer);clearTimeout(this.end)}}},state);
 await page.goto('http://127.0.0.1:8765');await page.locator('[data-view=schedule]').click();await page.locator('#generate').click();await page.locator('#loading').waitFor({state:'visible'});await page.locator('#loading-stage').getByText('均衡总工时').waitFor();
 await page.waitForTimeout(1100);assert.match(await page.locator('#loading-time').innerText(),/1 秒/);assert.equal(await page.locator('#workspace').getAttribute('aria-busy'),'true');await page.screenshot({path:'tests/loading-desktop.png'});
 await page.locator('#cancel-computation').click();await page.locator('#loading').waitFor({state:'hidden'});assert.match(await page.locator('#status').innerText(),/已取消/);assert.equal(await page.evaluate(()=>JSON.parse(localStorage.getItem('hnu-browser-scheduler-v1')).schedules[1]),undefined);
 await page.setViewportSize({width:390,height:844});await page.locator('#generate').click();await page.locator('#loading').waitFor({state:'visible'});await page.screenshot({path:'tests/loading-mobile.png'});await page.getByText('排班完成',{exact:true}).waitFor();await page.locator('#loading').waitFor({state:'hidden'});assert.equal(await page.locator('#workspace').getAttribute('aria-busy'),null);
 await page.evaluate(()=>window.testError=true);await page.locator('#generate').click();await page.locator('#loading').waitFor({state:'visible'});await page.locator('#loading').waitFor({state:'hidden'});assert.equal(await page.locator('#status').innerText(),'测试错误');
 console.log('PASS: stage, elapsed time, desktop/mobile loader, cancellation, success and error cleanup');
}finally{await browser.close()}
