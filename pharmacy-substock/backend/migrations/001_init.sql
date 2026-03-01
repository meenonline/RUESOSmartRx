-- ===== Pharmacy Sub-stock Management Schema =====

-- Transactions table: all stock movements
CREATE TABLE IF NOT EXISTS transactions (
  id            SERIAL PRIMARY KEY,
  disp_no       VARCHAR(20),
  drug_code     VARCHAR(50),
  drug_name     VARCHAR(255) NOT NULL,
  pack_size     VARCHAR(50),
  price_per_unit NUMERIC(12,2) DEFAULT 0,
  lot_no        VARCHAR(100),
  exp_date      DATE,
  qty           INTEGER NOT NULL,       -- positive = รับเข้า, negative = เบิกยา
  txn_type      VARCHAR(50) NOT NULL,   -- 'receive','dispense','adjust','opening'
  dept          VARCHAR(100) DEFAULT 'substock',
  recorder      VARCHAR(100),
  remark        TEXT,
  created_at    TIMESTAMPTZ DEFAULT NOW()
);

-- Formulary: drug master list with min-stock and location
CREATE TABLE IF NOT EXISTS formulary (
  id              SERIAL PRIMARY KEY,
  drug_name       VARCHAR(255) NOT NULL UNIQUE,
  drug_code       VARCHAR(50),
  pack_size       VARCHAR(50),
  min_stock       INTEGER DEFAULT 0,
  cabinet         VARCHAR(100),
  reorder_point   INTEGER DEFAULT 0,
  shelf_position  VARCHAR(100),
  is_inactive     BOOLEAN DEFAULT FALSE,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- Ignore expiry: track lots to hide from dashboard alerts
CREATE TABLE IF NOT EXISTS ignore_expiry (
  id        SERIAL PRIMARY KEY,
  drug_name VARCHAR(255) NOT NULL,
  lot_no    VARCHAR(100) NOT NULL,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  UNIQUE(drug_name, lot_no)
);

-- Audit trail: random count audit records
CREATE TABLE IF NOT EXISTS audit_trail (
  id            SERIAL PRIMARY KEY,
  drug_name     VARCHAR(255) NOT NULL,
  system_qty    INTEGER,
  counted_qty   INTEGER,
  diff          INTEGER,
  cabinet       VARCHAR(100),
  shelf_position VARCHAR(100),
  lot_no        VARCHAR(100),
  recorder      VARCHAR(100),
  created_at    TIMESTAMPTZ DEFAULT NOW()
);

-- Settings: key-value store
CREATE TABLE IF NOT EXISTS settings (
  key   VARCHAR(100) PRIMARY KEY,
  value TEXT
);

-- ===== Views =====

-- Stock summary by lot
CREATE OR REPLACE VIEW stock_summary_lot AS
SELECT
  drug_name,
  drug_code,
  pack_size,
  lot_no,
  exp_date,
  price_per_unit,
  SUM(qty) AS balance,
  SUM(qty * COALESCE(price_per_unit, 0)) AS total_value
FROM transactions
GROUP BY drug_name, drug_code, pack_size, lot_no, exp_date, price_per_unit;

-- Stock summary without lot (aggregated)
CREATE OR REPLACE VIEW stock_summary_nolot AS
SELECT
  drug_name,
  drug_code,
  pack_size,
  MAX(price_per_unit) AS price_per_unit,
  SUM(qty) AS balance,
  SUM(qty * COALESCE(price_per_unit, 0)) AS total_value
FROM transactions
GROUP BY drug_name, drug_code, pack_size;

-- ===== Indexes =====
CREATE INDEX IF NOT EXISTS idx_transactions_drug_name ON transactions(drug_name);
CREATE INDEX IF NOT EXISTS idx_transactions_created_at ON transactions(created_at);
CREATE INDEX IF NOT EXISTS idx_transactions_disp_no ON transactions(disp_no);
CREATE INDEX IF NOT EXISTS idx_formulary_drug_name ON formulary(drug_name);
CREATE INDEX IF NOT EXISTS idx_ignore_expiry_drug_lot ON ignore_expiry(drug_name, lot_no);
