-- Execute no Editor SQL do projeto Neon usado exclusivamente pela DG Eletros.
CREATE TABLE IF NOT EXISTS dge_records (
  collection TEXT NOT NULL CHECK (collection IN ('Produtos','Vendas','Usuarios','LogExclusoes','Pagamentos Comissoes','Horas Extras')),
  record_key TEXT NOT NULL,
  source_row INTEGER NOT NULL,
  data JSONB NOT NULL,
  PRIMARY KEY (collection, record_key),
  UNIQUE (collection, source_row)
);
CREATE INDEX IF NOT EXISTS dge_records_collection_row ON dge_records (collection, source_row);
CREATE SEQUENCE IF NOT EXISTS dge_row_seq START WITH 1000000;

CREATE TABLE IF NOT EXISTS dge_sessions (
  token_hash TEXT PRIMARY KEY,
  user_key TEXT NOT NULL,
  expires_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS dge_photos (
  photo_id TEXT PRIMARY KEY,
  mime_type TEXT NOT NULL CHECK (mime_type IN ('image/jpeg','image/png','image/webp')),
  content BYTEA NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
