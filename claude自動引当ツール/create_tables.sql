-- 在庫自動引当・紐付けツール テーブル作成SQL
-- Ver 2.0

-- orders（受注）
CREATE TABLE IF NOT EXISTS orders (
  id INT AUTO_INCREMENT PRIMARY KEY,
  yahoo_order_id VARCHAR(50) NOT NULL UNIQUE,
  ordered_at DATETIME NOT NULL,
  desired_delivery_date DATE,
  customer_name VARCHAR(100),
  priority_ship_flag TINYINT(1) DEFAULT 0,
  yahoo_ship_status VARCHAR(30),
  status VARCHAR(30) DEFAULT 'pending',
  delay_memo TEXT,
  customer_contacted_flag TINYINT(1) DEFAULT 0,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- order_items（受注明細）
CREATE TABLE IF NOT EXISTS order_items (
  id INT AUTO_INCREMENT PRIMARY KEY,
  order_id INT NOT NULL,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  product_name VARCHAR(200),
  quantity INT NOT NULL,
  inventory_type VARCHAR(20) NOT NULL,
  status VARCHAR(30) DEFAULT 'pending',
  allocated_qty INT DEFAULT 0,
  shipped_flag TINYINT(1) DEFAULT 0,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW(),
  FOREIGN KEY (order_id) REFERENCES orders(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- purchases（発注）
CREATE TABLE IF NOT EXISTS purchases (
  id INT AUTO_INCREMENT PRIMARY KEY,
  order_item_id INT NOT NULL,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  product_name VARCHAR(200),
  quantity INT NOT NULL,
  shop_name VARCHAR(100),
  ordered_at DATE NOT NULL,
  status VARCHAR(30) DEFAULT 'ordered',
  memo TEXT,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW(),
  FOREIGN KEY (order_item_id) REFERENCES order_items(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- ems（EMS便）
CREATE TABLE IF NOT EXISTS ems (
  id INT AUTO_INCREMENT PRIMARY KEY,
  ems_number VARCHAR(50) NOT NULL UNIQUE,
  shipped_at DATE NOT NULL,
  estimated_arrival DATE NOT NULL,
  arrived_at DATE,
  status VARCHAR(30) DEFAULT 'in_transit',
  memo TEXT,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- ems_items（EMS明細）
CREATE TABLE IF NOT EXISTS ems_items (
  id INT AUTO_INCREMENT PRIMARY KEY,
  ems_id INT NOT NULL,
  order_item_id INT NOT NULL,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  quantity INT NOT NULL,
  created_at DATETIME DEFAULT NOW(),
  FOREIGN KEY (ems_id) REFERENCES ems(id),
  FOREIGN KEY (order_item_id) REFERENCES order_items(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- inventory（在庫）
CREATE TABLE IF NOT EXISTS inventory (
  id INT AUTO_INCREMENT PRIMARY KEY,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  inventory_type VARCHAR(20) NOT NULL,
  quantity INT DEFAULT 0,
  reserved_qty INT DEFAULT 0,
  available_qty INT DEFAULT 0,
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- allocations（引当記録）
CREATE TABLE IF NOT EXISTS allocations (
  id INT AUTO_INCREMENT PRIMARY KEY,
  order_item_id INT NOT NULL,
  inventory_type VARCHAR(20) NOT NULL,
  allocation_type VARCHAR(20) NOT NULL,
  quantity INT NOT NULL,
  ems_item_id INT,
  allocated_by VARCHAR(50),
  allocated_at DATETIME DEFAULT NOW(),
  FOREIGN KEY (order_item_id) REFERENCES order_items(id),
  FOREIGN KEY (ems_item_id) REFERENCES ems_items(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- alerts（アラート）
CREATE TABLE IF NOT EXISTS alerts (
  id INT AUTO_INCREMENT PRIMARY KEY,
  alert_type VARCHAR(50) NOT NULL,
  order_id INT,
  order_item_id INT,
  product_code VARCHAR(100),
  message TEXT NOT NULL,
  resolved_flag TINYINT(1) DEFAULT 0,
  resolved_at DATETIME,
  created_at DATETIME DEFAULT NOW(),
  FOREIGN KEY (order_id) REFERENCES orders(id),
  FOREIGN KEY (order_item_id) REFERENCES order_items(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- japan_inventory_staging（日本在庫仕分け）
CREATE TABLE IF NOT EXISTS japan_inventory_staging (
  id INT AUTO_INCREMENT PRIMARY KEY,
  ems_item_id INT NOT NULL,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  quantity INT NOT NULL,
  status VARCHAR(30) DEFAULT 'waiting',
  assigned_order_item_id INT,
  reflected_at DATETIME,
  excluded_reason TEXT,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW(),
  FOREIGN KEY (ems_item_id) REFERENCES ems_items(id),
  FOREIGN KEY (assigned_order_item_id) REFERENCES order_items(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- インデックス
CREATE INDEX idx_orders_ordered_at ON orders(ordered_at);
CREATE INDEX idx_orders_status ON orders(status);
CREATE INDEX idx_order_items_order_id ON order_items(order_id);
CREATE INDEX idx_order_items_product_code ON order_items(product_code);
CREATE INDEX idx_order_items_status ON order_items(status);
CREATE INDEX idx_purchases_order_item_id ON purchases(order_item_id);
CREATE INDEX idx_purchases_product_code ON purchases(product_code);
CREATE INDEX idx_ems_items_ems_id ON ems_items(ems_id);
CREATE INDEX idx_ems_items_order_item_id ON ems_items(order_item_id);
CREATE INDEX idx_inventory_product ON inventory(product_code, inventory_type);
CREATE INDEX idx_allocations_order_item_id ON allocations(order_item_id);
CREATE INDEX idx_alerts_resolved ON alerts(resolved_flag);
CREATE INDEX idx_japan_staging_ems_item ON japan_inventory_staging(ems_item_id);
CREATE INDEX idx_japan_staging_status ON japan_inventory_staging(status);
