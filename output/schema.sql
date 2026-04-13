-- Generated SQLite Schema for Excel2App

CREATE TABLE IF NOT EXISTS Products (
    row_id INTEGER PRIMARY KEY,
    A TEXT,
    B TEXT,
    C TEXT,
    D TEXT,
    E TEXT
);

INSERT INTO Products (row_id, A, B, C, D, E) VALUES (1, 'ID', 'Name', 'Category', 'Price', 'Stock');
INSERT INTO Products (row_id, A, B, C, D, E) VALUES (2, 101, 'Laptop', 'Electronics', 1200, 15);
INSERT INTO Products (row_id, A, B, C, D, E) VALUES (3, 102, 'Chair', 'Furniture', 150, 40);
INSERT INTO Products (row_id, A, B, C, D, E) VALUES (4, 103, 'Desk', 'Furniture', 300, 10);
INSERT INTO Products (row_id, A, B, C, D, E) VALUES (5, 104, 'Monitor', 'Electronics', 250, 25);

CREATE TABLE IF NOT EXISTS Categories (
    row_id INTEGER PRIMARY KEY,
    A TEXT,
    B TEXT
);

INSERT INTO Categories (row_id, A, B) VALUES (1, 'Category Name', 'Description');
INSERT INTO Categories (row_id, A, B) VALUES (2, 'Electronics', 'Gadgets and hardware');
INSERT INTO Categories (row_id, A, B) VALUES (3, 'Furniture', 'Office and home furniture');

CREATE TABLE IF NOT EXISTS Summary (
    row_id INTEGER PRIMARY KEY,
    A TEXT,
    B TEXT
);

INSERT INTO Summary (row_id, A, B) VALUES (1, 'Metric', 'Value');
INSERT INTO Summary (row_id, A) VALUES (2, 'Total Stock');
INSERT INTO Summary (row_id, A) VALUES (3, 'Total Value');
INSERT INTO Summary (row_id, A) VALUES (4, 'Unknown Function Test');
INSERT INTO Summary (row_id, A) VALUES (5, 'Undefined Ref Test');
INSERT INTO Summary (row_id, A) VALUES (6, 'Total Stock (Again)');

CREATE TABLE IF NOT EXISTS Cycles (
    row_id INTEGER PRIMARY KEY,
    A TEXT,
    B TEXT
);


