/****** Object:  Database NorthwindDW    Script Date: 5/13/2024 2:17:59 AM ******/
/*
Kimball Group, The Microsoft Data Warehouse Toolkit
Generate a database from the datamodel worksheet, version: 4

You can use this Excel workbook as a data modeling tool during the logical design phase of your project.
As discussed in the book, it is in some ways preferable to a real data modeling tool during the inital design.
We expect you to move away from this spreadsheet and into a real modeling tool during the physical design phase.
The authors provide this macro so that the spreadsheet isn't a dead-end. You can 'import' into your
data modeling tool by generating a database using this script, then reverse-engineering that database into
your tool.

Uncomment the next lines if you want to drop and create the database
*/
/*
DROP DATABASE NorthwindDW
GO
CREATE DATABASE NorthwindDW
GO
ALTER DATABASE NorthwindDW
SET RECOVERY SIMPLE
GO
*/
CREATE DATABASE OlistDW
USE OlistDW


/* Drop table dbo.DimCustomers */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimCustomers') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimCustomers 
;

/* Create table dbo.DimCustomers */
CREATE TABLE dbo.DimCustomers (
   [customerKey]  int IDENTITY
,  [customer_id]  varchar(50)
,  [customer_city]  varchar(50)
,  [customer_state]  varchar(50)
, CONSTRAINT [PK_dbo.DimCustomers] PRIMARY KEY CLUSTERED 
( [customerKey])
) ON [PRIMARY]
;

/* Drop table dbo.DimProducts */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimProducts') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimProducts 
;

/* Create table dbo.DimProducts */
CREATE TABLE dbo.DimProducts (
   [productKey]  int IDENTITY 
,  [product_id]  varchar(50)
,  [product_category_name]  varchar(50) 
,  [product_weight_g]  decimal(10, 1)
,  [product_length_cm]  decimal(10, 1)
,  [product_height_cm]  decimal(10, 1)
,  [product_width_cm]  decimal(10, 1)
, CONSTRAINT [PK_dbo.DimProducts] PRIMARY KEY CLUSTERED 
( [productKey])
) ON [PRIMARY]
;

/* Drop table dbo.DimSellers */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimSellers') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimSellers 
;

/* Create table dbo.DimSellers */
CREATE TABLE dbo.DimSellers (
   [sellerKey]  int IDENTITY
,  [Seller_id]  varchar(50)
,  [seller_city]  varchar(50)
,  [seller_state]  varchar(50)
, CONSTRAINT [PK_dbo.DimSellers] PRIMARY KEY CLUSTERED 
( [sellerKey])
) ON [PRIMARY]
;

/* Drop table dbo.DimDate */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.DimDate') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.DimDate 
;

/* Create table dbo.DimDate */
CREATE TABLE dbo.DimDate (
   [DateKey]  int 
,  [Date]  smalldatetime 
,  [DayOfWeek]  tinyint
,  [DayName]  varchar(9)
,  [DayOfMonth]  tinyint
,  [DayOfYear]  smallint
,  [WeekOfYear]  tinyint
,  [MonthName]  varchar(9)
,  [MonthOfYear]  tinyint 
,  [Quarter]  tinyint
,  [Year]  smallint
,  [weekdayFlag] char(10)
, CONSTRAINT [PK_dbo.DimDate] PRIMARY KEY CLUSTERED 
( [DateKey] )
) ON [PRIMARY]
;

/* Drop table dbo.FactSales */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.FactSales') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.FactSales 
;

/* Create table dbo.FactSales */
CREATE TABLE dbo.FactSales (
   [productKey] int
,  [customerKey] int
,  [sellerKey] int
,  [price]  decimal(8, 3)
,  [freight_value]  decimal(8, 3)
,  [OrderID]  varchar(50)
,  [OrderItemID] int
,  [orderDeliveredCustomerDateKey] int
) ON [PRIMARY]
;

/* Drop table dbo.FactOrderFullFilment */
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'dbo.FactOrderFullFilment') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE dbo.FactOrderFullFilment 
;

/* Create table dbo.FactOrderFullFilment */
CREATE TABLE dbo.FactOrderFullFilment (
   [customerKey] int
,  [sellerKey] int
,  [delivery_time]  int
,  [delivery_delay]  int
,  [accept_time] int
,  [orderPurchaseTimestampKey]  int
,  [orderApprovedAtKey] int
,  [orderDiliveredCarrierDateKey]  int
,  [orderDeliveredCustomerDateKey] int
,  [orderDeliveredEstimateDateKey] int
) ON [PRIMARY]
;
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_ProductKey FOREIGN KEY
   (
   productKey
   ) REFERENCES DimProducts
   ( productKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_CustomerKey FOREIGN KEY
   (
   customerKey
   ) REFERENCES DimCustomers
   ( customerKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_SellerKey FOREIGN KEY
   (
   sellerKey
   ) REFERENCES DimSellers
   ( sellerKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactSales ADD CONSTRAINT
   FK_dbo_FactSales_OrderDeliveredCustomerDateKey FOREIGN KEY
   (
   orderDeliveredCustomerDateKey
   ) REFERENCES DimDate
   ( DateKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;

ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_SellerKey FOREIGN KEY
   (
   sellerKey
   ) REFERENCES DimSellers
   ( sellerKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_CustomerKey FOREIGN KEY
   (
   customerKey
   ) REFERENCES DimCustomers
   ( customerKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_OrderDeliveredCustomerDateKey FOREIGN KEY
   (
   orderDeliveredCustomerDateKey
   ) REFERENCES DimDate
   ( DateKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_orderPurchaseTimestampKey FOREIGN KEY
   (
   orderPurchaseTimestampKey
   ) REFERENCES DimDate
   ( DateKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_orderApprovedAtKey FOREIGN KEY
   (
   orderApprovedAtKey
   ) REFERENCES DimDate
   ( DateKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;
ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_orderDiliveredCarrierDateKey FOREIGN KEY
   (
   orderDiliveredCarrierDateKey
   ) REFERENCES DimDate
   ( DateKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;

;
ALTER TABLE dbo.FactOrderFullFilment ADD CONSTRAINT
   FK_dbo_FactOrderFullFilment_orderDeliveredEstimateDateKey FOREIGN KEY
   (
   orderDeliveredEstimateDateKey
   ) REFERENCES DimDate
   ( DateKey )
     ON UPDATE  NO ACTION
     ON DELETE  NO ACTION
;

