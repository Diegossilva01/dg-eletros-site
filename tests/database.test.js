const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const {paymentStatements}=require('../lib/commission-payment');
const dependency=process.env.PGLITE_MODULE;
test('database validates commission payments and preserves concurrent stock changes',{skip:!dependency},async()=>{
  const {PGlite}=require(dependency);const db=new PGlite();
  try{
    await db.exec(fs.readFileSync(path.join(__dirname,'../schema.sql'),'utf8'));
    const put=async(collection,key,data)=>db.query("INSERT INTO dge_records VALUES($1,$2,nextval('dge_row_seq'),$3)",[collection,key,JSON.stringify(data)]);
    await put('Vendas','sale1',{'Vendedora':' Ana  Thaynaa ','Status comissão':'Aprovada','Valor da venda':'23.646,50'});
    const pay=async(key,value,name='ANA THAYNAA')=>db.transaction(async tx=>{let results;for(const [sql,params] of paymentStatements(key,{ID:key,Vendedora:name,'Valor pago':value}))results=await tx.query(sql,params);return results.rows});
    assert.equal((await pay('pay1',472.93)).length,1);
    assert.equal((await pay('pay2',472.93)).length,0);
    assert.equal((await pay('over',0.01)).length,0);
    await put('Vendas','sale2',{'Vendedora':'Ana Thaynaa','Status comissão':'Aprovada','Valor da venda':1000});
    assert.equal((await pay('partial',10)).length,1);
    assert.equal((await pay('remainder',10)).length,1);
    assert.equal((await pay('tooMuch',1)).length,0);
    // Exercise the same transaction statements through overlapping callers.
    await put('Vendas','sale3',{'Vendedora':'Ana Thaynaa','Status comissão':'Aprovada','Valor da venda':500});
    const together=await Promise.all([pay('parallel1',10),pay('parallel2',10)]);
    assert.equal(together.reduce((sum,r)=>sum+r.length,0),1);
    const old={ID:'P1',Quantidade:1,Status:'Publicado','Promoção':'Não'};
    await put('Produtos','P1',old);
    await db.query("UPDATE dge_records SET data=jsonb_set(data,'{Quantidade}','0') WHERE collection='Produtos' AND record_key='P1'");
    await db.query("UPDATE dge_records SET data=jsonb_set(data,$3::text[],$4::jsonb,true) WHERE collection=$1 AND record_key=$2",['Produtos','P1',['Promoção'],JSON.stringify('Sim')]);
    const stale=await db.query("UPDATE dge_records SET data=$3::jsonb WHERE collection=$1 AND record_key=$2 AND ($4::jsonb IS NULL OR (data - '_key' - '_linha')=($4::jsonb - '_key' - '_linha')) RETURNING source_row",['Produtos','P1',JSON.stringify({...old,Produto:'Edited'}),JSON.stringify(old)]);
    assert.equal(stale.rows.length,0);
    const row=(await db.query("SELECT data FROM dge_records WHERE collection='Produtos' AND record_key='P1'")).rows[0].data;
    assert.equal(row.Quantidade,0);assert.equal(row['Promoção'],'Sim');
  }finally{await db.close()}
});
