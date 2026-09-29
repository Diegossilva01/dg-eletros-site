const { neon } = require('@neondatabase/serverless');

let connection;

function db() {
  if (!process.env.DATABASE_URL) {
    throw new Error('DATABASE_URL não configurada na Vercel.');
  }
  return connection ||= neon(process.env.DATABASE_URL);
}

const query = (text, params = []) => db().query(text, params);

async function records(collection) {
  const rows = await query(
    'SELECT record_key, source_row, data FROM dge_records WHERE collection=$1 ORDER BY source_row',
    [collection]
  );
  return rows.map(r => ({
    ...r.data,
    _key: r.record_key,
    _linha: r.source_row
  }));
}

async function one(collection, id, row) {
  const rows = await query(
    'SELECT record_key, source_row, data FROM dge_records WHERE collection=$1 AND (record_key=$2 OR source_row=$3) LIMIT 1',
    [collection, String(id || ''), Number(row) || -1]
  );
  return rows[0] && {
    ...rows[0].data,
    _key: rows[0].record_key,
    _linha: rows[0].source_row
  };
}

async function save(collection, key, data, row) {
  if (row) {
    const result = await query(
      'UPDATE dge_records SET data=$3::jsonb WHERE collection=$1 AND record_key=$2 RETURNING source_row',
      [collection, key, JSON.stringify(data)]
    );
    if (!result.length) {
      throw new Error('Registro não encontrado. Atualize a página.');
    }
    return result[0].source_row;
  }

  const result = await query(
    "INSERT INTO dge_records(collection,record_key,source_row,data) VALUES($1,$2,nextval('dge_row_seq'),$3::jsonb) RETURNING source_row",
    [collection, key, JSON.stringify(data)]
  );
  return result[0].source_row;
}

async function remove(collection, key) {
  await query(
    'DELETE FROM dge_records WHERE collection=$1 AND record_key=$2',
    [collection, key]
  );
}

module.exports = { query, records, one, save, remove };
