CREATE TABLE orders (
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

CREATE TABLE order_items (
  id INT AUTO_INCREMENT PRIMARY KEY,
  order_id INT NOT NULL,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  quantity INT NOT NULL,
  inventory_type VARCHAR(20) NOT NULL,
  status VARCHAR(30) DEFAULT 'pending',
  allocated_qty INT DEFAULT 0,
  shipped_flag TINYINT(1) DEFAULT 0,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW(),
  FOREIGN KEY (order_id) REFERENCES orders(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE purchases (
  id INT AUTO_INCREMENT PRIMARY KEY,
  order_item_id INT NOT NULL,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  quantity INT NOT NULL,
  shop_name VARCHAR(100),
  ordered_at DATE NOT NULL,
  status VARCHAR(30) DEFAULT 'ordered',
  memo TEXT,
  created_at DATETIME DEFAULT NOW(),
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW(),
  FOREIGN KEY (order_item_id) REFERENCES order_items(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE ems (
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

CREATE TABLE ems_items (
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

CREATE TABLE inventory (
  id INT AUTO_INCREMENT PRIMARY KEY,
  product_code VARCHAR(100) NOT NULL,
  product_sub_code VARCHAR(100),
  inventory_type VARCHAR(20) NOT NULL,
  quantity INT DEFAULT 0,
  reserved_qty INT DEFAULT 0,
  available_qty INT DEFAULT 0,
  updated_at DATETIME DEFAULT NOW() ON UPDATE NOW()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE allocations (
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

CREATE TABLE alerts (
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

CREATE TABLE japan_inventory_staging (
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
