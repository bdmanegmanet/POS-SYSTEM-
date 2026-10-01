const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 10000;
const GAS_WEB_APP_URL = String(process.env.GAS_WEB_APP_URL || '').trim();
const indexPath = path.join(__dirname, 'Index.html');

function send(res, status, body, type='text/plain; charset=utf-8'){
  res.writeHead(status, {'Content-Type': type, 'Cache-Control':'no-store'});
  res.end(body);
}

async function proxyToGas(body){
  if(!GAS_WEB_APP_URL) throw new Error('GAS_WEB_APP_URL is not configured on Render.');
  const response = await fetch(GAS_WEB_APP_URL, {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify(body),
    redirect:'follow'
  });
  const text = await response.text();
  return {status:response.status,text};
}

const server = http.createServer(async (req,res)=>{
  try{
    if(req.method==='GET' && req.url==='/healthz') return send(res,200,JSON.stringify({ok:true,service:'POS System',gasConfigured:!!GAS_WEB_APP_URL}),'application/json; charset=utf-8');
    if(req.method==='POST' && req.url==='/api'){
      let raw=''; req.on('data',c=>raw+=c); req.on('end',async()=>{
        try{
          const body=JSON.parse(raw||'{}'), out=await proxyToGas(body);
          send(res,out.status,out.text,'application/json; charset=utf-8');
        }catch(e){send(res,503,JSON.stringify({success:false,error:e.message}),'application/json; charset=utf-8');}
      }); return;
    }
    if(req.method==='GET' && (req.url==='/'||req.url==='/index.html')){
      return send(res,200,fs.readFileSync(indexPath,'utf8'),'text/html; charset=utf-8');
    }
    return send(res,404,'Not found');
  }catch(e){return send(res,500,e.message)}
});
server.listen(PORT,'0.0.0.0',()=>console.log('POS System listening on '+PORT));
