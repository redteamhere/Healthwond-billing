PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS Users (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Username TEXT NOT NULL UNIQUE,
    PasswordHash TEXT NOT NULL,
    PasswordSalt TEXT NOT NULL,
    FullName TEXT NOT NULL,
    Role TEXT NOT NULL,
    IsActive INTEGER NOT NULL DEFAULT 1,
    LastLoginAt TEXT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Products (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ProductName TEXT NOT NULL,
    Packing TEXT NULL,
    HsnCode TEXT NULL,
    BatchNumber TEXT NOT NULL,
    ExpiryDate TEXT NOT NULL,
    GstPercentage NUMERIC NOT NULL DEFAULT 0,
    MRP NUMERIC NOT NULL DEFAULT 0,
    PTR NUMERIC NOT NULL DEFAULT 0,
    PTS NUMERIC NOT NULL DEFAULT 0,
    CompanyName TEXT NULL,
    Composition TEXT NULL,
    CurrentStock INTEGER NOT NULL DEFAULT 0,
    Barcode TEXT NULL,
    IsDeleted INTEGER NOT NULL DEFAULT 0,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Customers (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    CustomerName TEXT NOT NULL,
    Gstin TEXT NULL,
    DrugLicenseNumber TEXT NULL,
    Address TEXT NULL,
    Phone TEXT NULL,
    Email TEXT NULL,
    OutstandingBalance NUMERIC NOT NULL DEFAULT 0,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Suppliers (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    SupplierName TEXT NOT NULL,
    Gstin TEXT NULL,
    DrugLicenseNumber TEXT NULL,
    Address TEXT NULL,
    Phone TEXT NULL,
    Email TEXT NULL,
    OutstandingBalance NUMERIC NOT NULL DEFAULT 0,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Purchases (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    PurchaseNumber TEXT NOT NULL UNIQUE,
    SupplierId INTEGER NOT NULL,
    PurchaseDate TEXT NOT NULL,
    SupplierInvoiceNumber TEXT NULL,
    SupplierInvoiceDate TEXT NULL,
    PurchaseOrderNumber TEXT NULL,
    PurchaseOrderDate TEXT NULL,
    PlaceOfSupply TEXT NULL,
    CaseCount INTEGER NOT NULL DEFAULT 0,
    TransportName TEXT NULL,
    EwayBillNumber TEXT NULL,
    SubTotal NUMERIC NOT NULL DEFAULT 0,
    DiscountAmount NUMERIC NOT NULL DEFAULT 0,
    GstAmount NUMERIC NOT NULL DEFAULT 0,
    RoundOffAmount NUMERIC NOT NULL DEFAULT 0,
    NetAmount NUMERIC NOT NULL DEFAULT 0,
    Notes TEXT NULL,
    CreatedBy INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL,
    FOREIGN KEY (SupplierId) REFERENCES Suppliers(Id)
);

CREATE TABLE IF NOT EXISTS PurchaseItems (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    PurchaseId INTEGER NOT NULL,
    ProductId INTEGER NOT NULL,
    ProductName TEXT NULL,
    Packing TEXT NULL,
    HsnCode TEXT NULL,
    BatchNumber TEXT NOT NULL,
    ExpiryDate TEXT NOT NULL,
    Quantity INTEGER NOT NULL DEFAULT 0,
    FreeQuantity INTEGER NOT NULL DEFAULT 0,
    PTR NUMERIC NOT NULL DEFAULT 0,
    PTS NUMERIC NOT NULL DEFAULT 0,
    MRP NUMERIC NOT NULL DEFAULT 0,
    GstPercentage NUMERIC NOT NULL DEFAULT 0,
    TaxableAmount NUMERIC NOT NULL DEFAULT 0,
    GstAmount NUMERIC NOT NULL DEFAULT 0,
    LineTotal NUMERIC NOT NULL DEFAULT 0,
    FOREIGN KEY (PurchaseId) REFERENCES Purchases(Id) ON DELETE CASCADE,
    FOREIGN KEY (ProductId) REFERENCES Products(Id)
);

CREATE TABLE IF NOT EXISTS PurchaseReturns (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ReturnNumber TEXT NOT NULL UNIQUE,
    PurchaseId INTEGER NOT NULL,
    SupplierId INTEGER NOT NULL,
    ReturnDate TEXT NOT NULL,
    SubTotal NUMERIC NOT NULL DEFAULT 0,
    GstAmount NUMERIC NOT NULL DEFAULT 0,
    RoundOffAmount NUMERIC NOT NULL DEFAULT 0,
    NetAmount NUMERIC NOT NULL DEFAULT 0,
    Notes TEXT NULL,
    CreatedBy INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL,
    FOREIGN KEY (PurchaseId) REFERENCES Purchases(Id),
    FOREIGN KEY (SupplierId) REFERENCES Suppliers(Id)
);

CREATE TABLE IF NOT EXISTS PurchaseReturnItems (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    PurchaseReturnId INTEGER NOT NULL,
    PurchaseItemId INTEGER NOT NULL,
    ProductId INTEGER NOT NULL,
    BatchNumber TEXT NOT NULL,
    ReturnQuantity INTEGER NOT NULL DEFAULT 0,
    ReturnFreeQuantity INTEGER NOT NULL DEFAULT 0,
    PTR NUMERIC NOT NULL DEFAULT 0,
    GstPercentage NUMERIC NOT NULL DEFAULT 0,
    TaxableAmount NUMERIC NOT NULL DEFAULT 0,
    GstAmount NUMERIC NOT NULL DEFAULT 0,
    LineTotal NUMERIC NOT NULL DEFAULT 0,
    FOREIGN KEY (PurchaseReturnId) REFERENCES PurchaseReturns(Id) ON DELETE CASCADE,
    FOREIGN KEY (PurchaseItemId) REFERENCES PurchaseItems(Id),
    FOREIGN KEY (ProductId) REFERENCES Products(Id)
);

CREATE TABLE IF NOT EXISTS Invoices (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    InvoiceNumber TEXT NOT NULL UNIQUE,
    CustomerId INTEGER NOT NULL,
    InvoiceDate TEXT NOT NULL,
    PaymentMode TEXT NULL,
    SubTotal NUMERIC NOT NULL DEFAULT 0,
    DiscountAmount NUMERIC NOT NULL DEFAULT 0,
    SchemeAmount NUMERIC NOT NULL DEFAULT 0,
    GstAmount NUMERIC NOT NULL DEFAULT 0,
    RoundOffAmount NUMERIC NOT NULL DEFAULT 0,
    NetAmount NUMERIC NOT NULL DEFAULT 0,
    AmountPaid NUMERIC NOT NULL DEFAULT 0,
    BalanceAmount NUMERIC NOT NULL DEFAULT 0,
    Notes TEXT NULL,
    CreatedBy INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL,
    FOREIGN KEY (CustomerId) REFERENCES Customers(Id)
);

CREATE TABLE IF NOT EXISTS InvoiceItems (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    InvoiceId INTEGER NOT NULL,
    ProductId INTEGER NOT NULL,
    BatchNumber TEXT NOT NULL,
    ExpiryDate TEXT NOT NULL,
    Quantity INTEGER NOT NULL DEFAULT 0,
    FreeQuantity INTEGER NOT NULL DEFAULT 0,
    Rate NUMERIC NOT NULL DEFAULT 0,
    MRP NUMERIC NOT NULL DEFAULT 0,
    DiscountPercentage NUMERIC NOT NULL DEFAULT 0,
    DiscountAmount NUMERIC NOT NULL DEFAULT 0,
    SchemeDescription TEXT NULL,
    GstPercentage NUMERIC NOT NULL DEFAULT 0,
    TaxableAmount NUMERIC NOT NULL DEFAULT 0,
    GstAmount NUMERIC NOT NULL DEFAULT 0,
    LineTotal NUMERIC NOT NULL DEFAULT 0,
    FOREIGN KEY (InvoiceId) REFERENCES Invoices(Id) ON DELETE CASCADE,
    FOREIGN KEY (ProductId) REFERENCES Products(Id)
);

CREATE TABLE IF NOT EXISTS CustomerPayments (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ReceiptNumber TEXT NOT NULL UNIQUE,
    CustomerId INTEGER NOT NULL,
    PaymentDate TEXT NOT NULL,
    PaymentMode TEXT NULL,
    ReferenceNumber TEXT NULL,
    Amount NUMERIC NOT NULL DEFAULT 0,
    BalanceAfterPayment NUMERIC NOT NULL DEFAULT 0,
    Notes TEXT NULL,
    CreatedBy INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL,
    FOREIGN KEY (CustomerId) REFERENCES Customers(Id)
);

CREATE TABLE IF NOT EXISTS SupplierPayments (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    PaymentNumber TEXT NOT NULL UNIQUE,
    SupplierId INTEGER NOT NULL,
    PaymentDate TEXT NOT NULL,
    PaymentMode TEXT NULL,
    ReferenceNumber TEXT NULL,
    Amount NUMERIC NOT NULL DEFAULT 0,
    BalanceAfterPayment NUMERIC NOT NULL DEFAULT 0,
    Notes TEXT NULL,
    CreatedBy INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL,
    FOREIGN KEY (SupplierId) REFERENCES Suppliers(Id)
);

CREATE TABLE IF NOT EXISTS StockLedger (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ProductId INTEGER NOT NULL,
    BatchNumber TEXT NOT NULL,
    TransactionType TEXT NOT NULL,
    ReferenceType TEXT NOT NULL,
    ReferenceId INTEGER NOT NULL,
    QuantityIn INTEGER NOT NULL DEFAULT 0,
    QuantityOut INTEGER NOT NULL DEFAULT 0,
    BalanceQuantity INTEGER NOT NULL DEFAULT 0,
    UnitCost NUMERIC NOT NULL DEFAULT 0,
    Remarks TEXT NULL,
    TransactionDate TEXT NOT NULL,
    CreatedAt TEXT NOT NULL,
    FOREIGN KEY (ProductId) REFERENCES Products(Id)
);

CREATE TABLE IF NOT EXISTS StockAdjustments (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    AdjustmentNumber TEXT NOT NULL UNIQUE,
    AdjustmentDate TEXT NOT NULL,
    Notes TEXT NULL,
    CreatedBy INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS StockAdjustmentItems (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    StockAdjustmentId INTEGER NOT NULL,
    ProductId INTEGER NOT NULL,
    BatchNumber TEXT NOT NULL,
    AdjustmentMode TEXT NOT NULL,
    Quantity INTEGER NOT NULL DEFAULT 0,
    BalanceQuantity INTEGER NOT NULL DEFAULT 0,
    UnitCost NUMERIC NOT NULL DEFAULT 0,
    Remarks TEXT NULL,
    FOREIGN KEY (StockAdjustmentId) REFERENCES StockAdjustments(Id) ON DELETE CASCADE,
    FOREIGN KEY (ProductId) REFERENCES Products(Id)
);

CREATE TABLE IF NOT EXISTS Settings (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    SettingKey TEXT NOT NULL UNIQUE,
    SettingValue TEXT NOT NULL,
    Description TEXT NULL,
    UpdatedAt TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS IX_Products_ProductName ON Products(ProductName);
CREATE INDEX IF NOT EXISTS IX_Products_BatchNumber ON Products(BatchNumber);
CREATE INDEX IF NOT EXISTS IX_Products_ExpiryDate ON Products(ExpiryDate);
CREATE INDEX IF NOT EXISTS IX_Products_Barcode ON Products(Barcode);
CREATE INDEX IF NOT EXISTS IX_Customers_CustomerName ON Customers(CustomerName);
CREATE INDEX IF NOT EXISTS IX_Suppliers_SupplierName ON Suppliers(SupplierName);
CREATE INDEX IF NOT EXISTS IX_Invoices_InvoiceDate ON Invoices(InvoiceDate);
CREATE INDEX IF NOT EXISTS IX_Invoices_CustomerId ON Invoices(CustomerId);
CREATE INDEX IF NOT EXISTS IX_CustomerPayments_PaymentDate ON CustomerPayments(PaymentDate);
CREATE INDEX IF NOT EXISTS IX_CustomerPayments_CustomerId ON CustomerPayments(CustomerId);
CREATE INDEX IF NOT EXISTS IX_Purchases_PurchaseDate ON Purchases(PurchaseDate);
CREATE INDEX IF NOT EXISTS IX_SupplierPayments_PaymentDate ON SupplierPayments(PaymentDate);
CREATE INDEX IF NOT EXISTS IX_SupplierPayments_SupplierId ON SupplierPayments(SupplierId);
CREATE INDEX IF NOT EXISTS IX_PurchaseReturns_ReturnDate ON PurchaseReturns(ReturnDate);
CREATE INDEX IF NOT EXISTS IX_PurchaseReturnItems_PurchaseItemId ON PurchaseReturnItems(PurchaseItemId);
CREATE INDEX IF NOT EXISTS IX_StockLedger_ProductId ON StockLedger(ProductId);
CREATE INDEX IF NOT EXISTS IX_StockAdjustments_AdjustmentDate ON StockAdjustments(AdjustmentDate);
