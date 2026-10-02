// The lock and balance check are separate statements: Read Committed sees
// payments committed by another administrator while this transaction waited.
const vendorKey = value => String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
const sqlVendorKey = `lower(trim(regexp_replace(COALESCE(data->>'Vendedora',''),'[[:space:]]+',' ','g')))`;
function sqlMoney(field) {
  const text = `COALESCE(data->>'${field}','0')`;
  return `COALESCE(NULLIF(regexp_replace(CASE WHEN strpos(${text},',')>0 THEN replace(replace(${text},'.',''),',','.') ELSE ${text} END,'[^0-9.-]','','g'),'')::numeric,0)`;
}
function paymentStatements(key, payment) {
  const name = vendorKey(payment.Vendedora);
  return [
    [`SELECT pg_advisory_xact_lock(hashtext('dge-commission'),hashtext($1))`, [name]],
    [`WITH balance AS (
       SELECT GREATEST(0,
         COALESCE(SUM(ROUND(${sqlMoney('Valor da venda')}*0.02,2)) FILTER
           (WHERE collection='Vendas' AND data->>'Status comissão'='Aprovada'),0)
         - COALESCE(SUM(${sqlMoney('Valor pago')}) FILTER
           (WHERE collection='Pagamentos Comissoes'),0)) AS available
       FROM dge_records
       WHERE collection IN ('Vendas','Pagamentos Comissoes') AND ${sqlVendorKey}=$1
     )
     INSERT INTO dge_records(collection,record_key,source_row,data)
     SELECT 'Pagamentos Comissoes',$2,nextval('dge_row_seq'),$3::jsonb
     FROM balance WHERE available >= $4::numeric
     RETURNING record_key`, [name, key, JSON.stringify(payment), payment['Valor pago']]]
  ];
}
module.exports = { paymentStatements };
