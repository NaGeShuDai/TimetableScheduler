const http=require('http'),fs=require('fs'),path=require('path');
const root=path.join(__dirname,'public');
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript','.css':'text/css','.png':'image/png','.ico':'image/x-icon','.ttf':'font/ttf','.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'};
const port=Number(process.env.PORT||8765);
const server=http.createServer((req,res)=>{try{const pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname);const file=path.resolve(root,'.'+(pathname==='/'?'/index.html':pathname));if(!file.startsWith(root+path.sep)){res.writeHead(403);return res.end();}fs.readFile(file,(err,data)=>{if(err){res.writeHead(404);return res.end('Not found');}res.setHeader('Content-Type',types[path.extname(file)]||'application/octet-stream');res.end(data);});}catch{res.writeHead(400);res.end();}});
server.on('error',e=>{console.error('启动失败：'+e.message+'。可设置 PORT 使用其他端口。');process.exit(1);});
server.listen(port,'127.0.0.1',()=>console.log('本地浏览器版：http://127.0.0.1:'+port));
