const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const crypto=require('node:crypto');
const path=require('node:path');
const root=path.join(__dirname,'..');
function frontend(api){
  const nodes={},storage=new Map([['dge_token','session']]);
  const context={state:{token:'session'},api,shown:0,$:selector=>nodes[selector]||=( {hidden:false,classList:{add(){nodes[selector].hidden=true},remove(){nodes[selector].hidden=false}}}),safeStorage:{setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)},setTimeout:fn=>{fn()},showApp:()=>{context.shown++}};
  const s=fs.readFileSync(path.join(root,'admin.js'),'utf8');
  vm.runInNewContext(s.slice(s.indexOf('function mostrarLoginExpirado'),s.indexOf('function hoje()')),context);
  return {context,nodes,storage};
}
test('temporary startup failures preserve session and recover',async()=>{
  let calls=0;const f=frontend(async()=>{if(++calls<3)throw new Error('network');return {usuario:{nome:'Ana'}}});
  await f.context.restaurarSessao();assert.equal(calls,3);assert.equal(f.context.shown,1);assert.equal(f.storage.get('dge_token'),'session');
});
test('persistent network failure offers retry without clearing saved login',async()=>{
  const f=frontend(async()=>{throw new Error('network')});await f.context.restaurarSessao();
  assert.equal(f.storage.get('dge_token'),'session');assert.equal(f.nodes['#sessionRetry'].hidden,false);assert.equal(f.nodes['#login'].hidden,true);
});
test('cookie restoration works without localStorage marker',async()=>{
  const f=frontend(async()=>({usuario:{nome:'Ana'}}));f.storage.clear();f.context.state.token='';await f.context.restaurarSessao();assert.equal(f.context.shown,1);
});
test('actual expiration clears login state and displays sign-in',()=>{
  const f=frontend(()=>{});f.context.mostrarLoginExpirado();assert.equal(f.storage.has('dge_token'),false);assert.equal(f.context.state.token,'');assert.equal(f.nodes['#login'].hidden,false);
});
function backend(){
  const sessions=new Map(),u={ID:'USR-1',_key:'USR-1',Nome:'Ana',Perfil:'Administrador',Status:'Ativo','Usuário':'ana','Senha Hash':'hashed'};
  const M={hash:()=> 'hashed',records:async()=>[u],one:async()=>u,query:async(sql,p)=>{
    if(sql.startsWith('INSERT INTO dge_sessions'))sessions.set(p[0],Date.now()+28800000);
    if(sql.startsWith('UPDATE dge_sessions')){if((sessions.get(p[0])||0)>Date.now()){sessions.set(p[0],Date.now()+28800000);return [{user_key:u._key}]}return []}
    if(sql.startsWith('DELETE FROM dge_sessions'))sessions.delete(p[0]);return [];
  }};
  const c={module:{exports:{}},Buffer,console,require:n=>n==='node:crypto'?crypto:n==='../lib/model'?M:n==='../lib/db'?{transaction:()=>{}}:require(path.join(root,'lib/commission-payment'))};
  vm.runInNewContext(fs.readFileSync(path.join(root,'api/interna.js'),'utf8'),c);
  async function call(body,raw){const response={headers:{},statusCode:200,setHeader(k,v){this.headers[k]=v},status(code){this.statusCode=code;return this},json(data){this.data=data;return this}};await c.module.exports({method:'POST',headers:{cookie:raw?'dge_session='+raw:''},body},response);return response}
  return {sessions,call,M};
}
test('two logins coexist; activity renews expiry; logout affects only its session',async()=>{
  const b=backend(),login={acao:'login',usuario:'ana',senha:'password'};
  const first=await b.call(login),second=await b.call(login);
  const token=r=>r.headers['Set-Cookie'].match(/^dge_session=([^;]+)/)[1];
  const a=token(first),other=token(second);assert.notEqual(a,other);assert.equal(b.sessions.size,2);
  const key=crypto.createHash('sha256').update(a).digest('hex');b.sessions.set(key,Date.now()+1000);
  const check=await b.call({acao:'verificarToken'},a);assert.equal(check.data.ok,true);assert.ok(b.sessions.get(key)>Date.now()+28000000);assert.match(check.headers['Set-Cookie'],/HttpOnly/);
  await b.call({acao:'sair'},a);assert.equal((await b.call({acao:'verificarToken'},a)).statusCode,401);assert.equal((await b.call({acao:'verificarToken'},other)).data.ok,true);
});
