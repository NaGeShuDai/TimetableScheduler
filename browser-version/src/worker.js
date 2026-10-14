import {solve} from './engine.js';
self.onmessage=e=>{try{self.postMessage(solve(e.data.state,e.data.week,60000,stage=>self.postMessage({type:'progress',stage})))}catch(error){self.postMessage({error:error.message})}};
