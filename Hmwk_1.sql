if OBJECT_ID('tblReceiver') is NOT NULL
drop table tblReceiver;

if OBJECT_ID('tblPurchaseOrderLine') is NOT NULL
DROP TABLE tblPurchaseOrderLine;

if OBJECT_ID('tblEmployee') is NOT NULL
DROP TABLE tblEmployee;

if OBJECT_ID('tblProduct') is NOT NULL
DROP TABLE tblProduct;

if OBJECT_ID('tblPurchaseOrder') is NOT NULL
DROP TABLE tblPurchaseOrder;

if OBJECT_ID('tblPurchaseHistory') is NOT NULL
DROP TABLE tblPurchaseHistory;

if OBJECT_ID('tblVendor') is NOT NULL
DROP TABLE tblVendor;

if object_id('tblProductType') is NOT NULL
DROP TABLE tblProductType;

if OBJECT_ID('tblCondition') is NOT NULL
DROP TABLE tblCondition;

CREATE TABLE tblProductType
(ProductTypeID		char(2) NOT NULL	PRIMARY KEY,
Description			varchar(30));

INSERT INTO tblProductType VALUES
('CC', 'Camping and Cooking'),
('CS', 'Camping and Resting'),
('HT', 'Hiking and Trails'),
('LA', 'Comfort Essentials'),
('MS', 'Miscellaneous'),
('PG', 'Travel Bags'),
('UT', 'Utility Materials');


CREATE TABLE tblCondition
(ConditionID		char(2) NOT NULL,
Description			varchar(30) NOT NULL,
CONSTRAINT pk_ConditionID
PRIMARY KEY (ConditionID));

INSERT INTO tblCondition VALUES
('FD', 'Freight damage'),
('IP', 'Improper packaging'),
('OK', 'Acceptable'),
('OT', 'Other'),
('UK', 'Unknown'),
('WD', 'Water damage');


CREATE TABLE tblVendor
(VendorID			char(5) NOT NULL	PRIMARY KEY,
Name				varchar(30) NOT NULL,
Address1			varchar(30) NOT NULL,
Address2			varchar(30),
City				varchar(20) NOT NULL,
State				char(2) NOT NULL,
Zip					varchar(12) NOT NULL,
Email				varchar(30),
Contact				varchar(30),
Phone				char(15) NOT NULL,
FirstBuyDate		Datetime);

INSERT INTO tblVendor VALUES
('00216', 'PolySort Manufacturing', '2550 23rd Avenue', NULL, 'Denver', 'CO', '80568', 'fxd@polysort.com', 'Francisco Delgado', '3035558123', '07-28-2017'),
('09567', 'Apex Mills', '3500 Industrial Parkway', 'Unit 7', 'SPARKS', 'Nv', '89431', 'tcc@gmail.com', NULL, '7755552894', '03-13-2014'),
('13135', 'Adventure Materials', 'P.O. Box 2700', NULL, 'SALT LAKE CITY', 'UT', '84560-2700', 'info@advmat.com', NULL, '8015554500', '08-04-2017'),
('17453', 'Albemarle Corporation', '2355-B Vista Drive', 'Suite 765-B', 'Sparks', 'NV', '89431', 'sandp@msn.com', 'maryanne Jones', '7755553451', '06-14-1998'),   
('18567', 'Alcan Plastic, Inc.', 'P.O. Box 4456', NULL, 'Eagle Falls', 'AK', '99565', 'alcanp@plastic.com', NULL, '9075555268', '05-15-2017'),
('20566', 'BestCo Food Equipment', '1515 Kendall Mill Road', NULL, 'Allentown', 'PA', '15003', 'buyer@best.com', NULL, '4845556789', '09-14-2018'), 
('22890', 'Kitchen Chemicals Corp.', '7750 Rock Blvd.', NULL, 'sparks', 'nv', '89431-5602', 'info@kchem.corp', 'melinda', '7755552566', '07-23-2014'),
('36257', 'Injectomatic Mold Corp.', '14557 Hawthorne Blvd.', 'Unit 14', 'los angeles', 'CA', '90036-9960', 'hayman@inject.com', 'Sam hayes', '2135554963', '06-15-2018'),
('45899', 'Celanette Design, LLC', '9865 Sepulveda Blvd.', 'Suite B', 'Los Angeles', 'ca', '90045-3660', 'buyer@celanette.com', 'Linda Burch', '3105555545', '09-23-2018'),
('87654', 'Recycle Plastics Company', '10 Riverview Highway', NULL, 'detroit', 'mi', '48050', 'vendor@rpc.com', NULL, '3135554266', '08-16-2015');


CREATE TABLE tblEmployee
(EmpID			char(6) NOT NULL,
EmpLastName		varchar(30) NOT NULL,
EmpFirstName	varchar(30) NOT NULL,
EmpEmail		varchar(30),
EmpPhone		char(15) NOT NULL,
EmpMgrID		char(6),
CONSTRAINT pk_EmpID
PRIMARY KEY (EmpID),
CONSTRAINT fk_EmpMgrID
FOREIGN KEY (EmpMgrID)
REFERENCES tblEmployee (EmpID));

INSERT INTO tblEmployee VALUES
('E10003', 'Shamsudeen', 'Mumtaz', 'MumSham@gmail.com', '7755423212', NULL),
('E10009', 'Hernandez', 'Nathan', NULL, '7755312562', 'E10003'),
('E10015', 'Fetters', 'Sam', 'SFETTERS@gmail.com', '7753453821', 'E10003'),
('E10018', 'Schnitkowski', 'Michael', 'MSCHNIT@gmail.com', '7776553582', 'E10003'),
('E10026', 'Nguyen', 'Lieu', 'LNguen@gmail.com', '7755313834', 'E10015'),
('E10042', 'Van Meter', 'Juliette', 'JVANMETER@gmail.com', '7758453551', 'E10026'),
('E10055', 'Huang', 'Hai Jun', 'HJHuang@gmail.com', '7759153891', NULL),
('E10056', 'Chen', 'John', NULL, '7759053821', 'E10015'),
('E10057', 'Pinot', 'Jean Claude', 'JPINOT@gmail.com', '7756749002', NULL),
('E10077', 'MacAndrews-Abernethy', 'Elizabeth Victoria', 'EMACANDREWSABERNET@gmail.com', '7755553894', 'E10026'),
('E10085', 'Wong', 'Thomas', 'TNW@gmail.com', '7757783511', 'E10026'),
('E10087', 'Agarwal', 'Aarav', 'AAwal@gmail.com', '7757821341', 'E10026'),
('E10088', 'O''Toole', 'Timothy', 'TOTOOL@gmail.com', '7758423529', 'E10026'),
('E10101', 'Dabiri', 'Dilshad', 'DilDabiri@gmail.com', '7758413838', 'E10026'),
('E10192', 'Garcia', 'Leonardo', NULL, '7756219005', NULL);

CREATE TABLE tblPurchaseOrder
(PoNumber		char(6) NOT NULL,
PODatePlaced	Datetime NOT NULL,
PODatedNeeded	Datetime,
Terms			varchar(15),
Conditions		varchar(15),
BuyerEmpID		char(6),
VendorID		char(5),
CONSTRAINT pk_PoNumber
PRIMARY KEY (PoNumber),
CONSTRAINT fk_BuyerEmpID
FOREIGN KEY (BuyerEmpID)
REFERENCES tblEmployee (EmpID),
CONSTRAINT fk_VendorID
FOREIGN KEY (VendorID)
REFERENCES tblVendor (VendorID));

INSERT INTO tblPurchaseOrder VALUES ('025974','08/15/2018 00:00:00','08/18/2018 00:00:00','Net 15','FOB-AIR','E10015','18567');
INSERT INTO tblPurchaseOrder VALUES ('045687','08/21/2018 00:00:00','09/12/2018 00:00:00','COD',NULL,'E10055','00216');
INSERT INTO tblPurchaseOrder VALUES ('056489','08/04/2018 00:00:00','08/15/2018 00:00:00',NULL,'FOB','E10055','36257');
INSERT INTO tblPurchaseOrder VALUES ('112233','09/21/2018 00:00:00','10/25/2018 00:00:00','Net10',NULL,'E10026','17453');
INSERT INTO tblPurchaseOrder VALUES ('234607','09/04/2018 00:00:00','09/28/2018 00:00:00','Net 30','FOB','E10055','17453');
INSERT INTO tblPurchaseOrder VALUES ('256887','09/19/2018 00:00:00','10/15/2018 00:00:00','Net 30','FOB',NULL,'87654');
INSERT INTO tblPurchaseOrder VALUES ('329987','10/10/2018 00:00:00','01/12/2019 00:00:00','Net 30','FOB','E10101','18567');
INSERT INTO tblPurchaseOrder VALUES ('365870','09/14/2018 00:00:00','03/14/2019 00:00:00','Net 30',NULL,'E10101','17453');
INSERT INTO tblPurchaseOrder VALUES ('453313','09/19/2018 00:00:00','10/05/2018 00:00:00','COD',NULL,'E10015','17453');
INSERT INTO tblPurchaseOrder VALUES ('543791','09/15/2018 00:00:00','01/15/2019 00:00:00','Net 30',NULL,'E10055','45899');
INSERT INTO tblPurchaseOrder VALUES ('600124','10/01/2018 00:00:00','12/20/2018 00:00:00','COD',NULL,'E10055','00216');
INSERT INTO tblPurchaseOrder VALUES ('661677','09/30/2018 00:00:00','11/15/2018 00:00:00','Net 15','FOB','E10101','00216');
INSERT INTO tblPurchaseOrder VALUES ('781900','10/03/2018 00:00:00','11/12/2018 00:00:00',NULL,'FOB','E10055','09567');
INSERT INTO tblPurchaseOrder VALUES ('902347','09/16/2018 00:00:00','10/18/2018 00:00:00',NULL,'Pickup','E10087','09567');

Create table tblProduct
(ProductID char(5) Not Null, 
Description varchar(30) NOT NULL, 
UOM char(10) Not Null CHECK (UOM IN ('each', 'feet', 'inches', 'meters', 'cm', 'sheet', 'case')), 
EOQ decimal(6,2), 
QOH decimal(8,2) Not Null, 
QOHDateUpdated datetime, 
ProductTypeID char(2) NOT NULL, 
CONSTRAINT pk_ProductID 
Primary key (ProductID),
Constraint fk_ProductTypeId
Foreign key (ProductTypeID)
References tblProductType (ProductTypeID));

INSERT INTO tblProduct Values 
('A7879',	'Canvas, Non-Woven',	'feet',	'2000.00',	'1200.50',	null,	'CS'),
('C2399',	'Thermoplastic',	'sheet',	'10.00',	'5.00',	'2018-09-15 09:00:00',	'CS'),
('C8733',	'Nylon Cover',	'meters',	'2000.00',	'2245.56',	'2019-09-24 04:50:00',	'CS'),
('C9100',	'Unbleached Muslin',	'meters',	'8000.00',	'5500.00',	'2018-09-18 03:42:00',	'CS'),
('G0983',	'Alpine Small Pot',	'each',	'100.00',	'65.00',	'2018-09-15 09:00:00',	'HT'),
('G1258',	'Alpine Pot/Kettle Handle',	'each',	'50.00',	'96.00',	'2018-09-15 09:00:00',	'UT'),
('G1366',	'Alpine Pot/Kettle Insert',	'each',	'48.00',	'55.00',	'2018-09-15 09:00:00',	'UT'),
('G5698',	'Alpine Skillet Handle Set',	'each',	'48.00',	'50.00',	'2018-09-15 09:00:00',	'UT'),
('J8006',	'Microfilter tubing',	'feet',	'450.00',	'550.75',	'2018-09-15 09:00:00',	'MS'),
('L8500',	'Hiking Lounge Seating - Blue',	'each',	'15.00',	'1.00',	'2018-09-15 09:00:00',	'LA'),
('M2356',	'Cot Mesh - Ultralite',	'meters',	'1000.00',	'1200.00',	'2018-09-18 03:42:00',	'MS'),
('M3577',	'Cot Mesh - Sturdy',	'feet',	'300.00',	'650.90',	'2018-09-18 03:42:00',	'MS'),
('O1957',	'Poly pro tubing, 1/2"',	'feet',	'300.00',	'95.00',	'2018-09-18 03:44:00',	'MS'),
('P5678',	'Stuff Sacks - Pillow Size',	'case',	'48.00',	'40.00',	'2018-09-18 03:44:00',	'PG'),
('P7844',	'Down Baffle Liner',	'meters',	'50.00',	'45.00', null,	'MS'),
('R5660',	'Water Filtration Pump',	'each',	'30.00',	'25.00',	'2018-09-18 03:42:00',	'CC'),
('T0460',	'Alpine Water Bottle',	'each',	'100.00',	'15.00',	'2018-09-18 03:42:00',	'CC');

CREATE TABLE tblPurchaseHistory
(HistoryID			INT IDENTITY(1,10),
ProductID			char(5) NOT NULL,
DatePurchased		Datetime NOT NULL,
Qty					decimal(8,2) NOT NULL,
Price				money NOT NULL,
QtyReceivedOnTime	decimal(8,2),
VendorID			char(5),
CONSTRAINT pk_HistoryID
PRIMARY KEY (HistoryID),
CONSTRAINT fk_ProductID
FOREIGN KEY (ProductID)
REFERENCES tblProduct (ProductID),
CONSTRAINT fk_VendorID2
FOREIGN KEY (VendorID)
REFERENCES tblVendor (VendorID));

INSERT INTO tblPurchaseHistory VALUES ('C2399','06/04/2018 00:00:00',200.00,1.55,200.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('C2399','02/12/2018 00:00:00',100.00,1.99,95.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('C2399','08/15/2017 00:00:00',150.00,1.75,.00,'36257');
INSERT INTO tblPurchaseHistory VALUES ('C2399','08/12/2018 00:00:00',350.00,1.89,345.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('C2399','04/14/2015 00:00:00',3000.00,1.45,3000.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('C2399','03/22/2017 00:00:00',100.00,1.95,100.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('C2399','01/25/2017 00:00:00',150.00,1.75,150.00,'36257');
INSERT INTO tblPurchaseHistory VALUES ('C2399','08/02/2018 00:00:00',50.00,2.15,50.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('C2399','05/16/2018 00:00:00',350.00,1.94,300.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('G0983','03/12/2018 00:00:00',75.00,2.15,75.00,'17453');
INSERT INTO tblPurchaseHistory VALUES ('G0983','09/22/2017 00:00:00',200.00,2.38,190.00,'17453');
INSERT INTO tblPurchaseHistory VALUES ('G0983','01/18/2018 00:00:00',250.00,1.99,10.00,'00216');
INSERT INTO tblPurchaseHistory VALUES ('G0983','02/10/2016 00:00:00',90.00,2.15,90.00,'20566');
INSERT INTO tblPurchaseHistory VALUES ('G0983','07/20/2016 00:00:00',300.00,2.1,300.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('G0983','05/10/2016 00:00:00',250.00,2.11,250.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('G1258','07/21/2018 00:00:00',25.00,4.29,24.00,'20566');
INSERT INTO tblPurchaseHistory VALUES ('G1258','06/20/2018 00:00:00',25.00,4.29,24.00,'20566');
INSERT INTO tblPurchaseHistory VALUES ('G1366','06/20/2018 00:00:00',25.00,4.81,25.00,NULL);
INSERT INTO tblPurchaseHistory VALUES ('G5698','06/20/2018 00:00:00',25.00,2.21,25.00,NULL);
INSERT INTO tblPurchaseHistory VALUES ('G5698','06/15/2017 00:00:00',200.00,2.16,200.00,NULL);
INSERT INTO tblPurchaseHistory VALUES ('L8500','04/21/2016 00:00:00',2.00,29.45,2.00,'13135');
INSERT INTO tblPurchaseHistory VALUES ('L8500','02/02/2017 00:00:00',5.00,38.4,5.00,'13135');
INSERT INTO tblPurchaseHistory VALUES ('L8500','12/12/2017 00:00:00',10.00,26.22,10.00,'13135');
INSERT INTO tblPurchaseHistory VALUES ('L8501','12/12/2017 00:00:00',15.00,29.94,10.00,'13135');
INSERT INTO tblPurchaseHistory VALUES ('M3577','01/10/2018 00:00:00',100.00,5.91,50.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('M3577','02/10/2016 00:00:00',1200.00,5.62,600.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('M3577','07/12/2018 00:00:00',150.00,6.35,75.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('M3577','08/15/2016 00:00:00',200.00,5.95,100.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('M3577','05/12/2015 00:00:00',3200.00,4.85,3100.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('M3577','12/02/2015 00:00:00',1200.00,5.62,1190.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('M3577','03/10/2016 00:00:00',1509.00,5.59,1500.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('M3577','11/06/2017 00:00:00',300.00,6.25,300.00,'87654');
INSERT INTO tblPurchaseHistory VALUES ('O1957','01/12/2018 00:00:00',1400.00,0.53,1400.00,'00216');
INSERT INTO tblPurchaseHistory VALUES ('O1957','09/18/2017 00:00:00',2500.00,0.42,2500.00,'20566');
INSERT INTO tblPurchaseHistory VALUES ('O1957','07/23/2018 00:00:00',200.00,0.43,200.00,'20566');
INSERT INTO tblPurchaseHistory VALUES ('O1957','07/19/2017 00:00:00',450.00,0.41,450.00,'00216');
INSERT INTO tblPurchaseHistory VALUES ('O1957','11/02/2016 00:00:00',2500.00,0.42,2490.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('O1957','10/03/2017 00:00:00',100.00,0.65,90.00,'18567');
INSERT INTO tblPurchaseHistory VALUES ('O1957','05/07/2018 00:00:00',450.00,0.44,400.00,'00216');
INSERT INTO tblPurchaseHistory VALUES ('P5678','06/12/2018 00:00:00',250.00,23.51,200.00,NULL);
INSERT INTO tblPurchaseHistory VALUES ('P5678','01/08/2018 00:00:00',100.00,24.66,100.00,'00216');
INSERT INTO tblPurchaseHistory VALUES ('P7844','07/12/2018 00:00:00',2570.00,0.67,2000.00,NULL);
INSERT INTO tblPurchaseHistory VALUES ('R5660','05/23/2018 00:00:00',80.00,2.12,80.00,'17453');
INSERT INTO tblPurchaseHistory VALUES ('R5660','02/27/2017 00:00:00',80.00,2.15,80.00,'17453');
INSERT INTO tblPurchaseHistory VALUES ('R5660','06/18/2018 00:00:00',4500.00,2.1,4500.00,NULL);
INSERT INTO tblPurchaseHistory VALUES ('T0460','03/29/2016 00:00:00',1200.00,1.12,1200.00,'45899');
INSERT INTO tblPurchaseHistory VALUES ('T0460','03/12/2017 00:00:00',1200.00,1.76,1200.00,'45899');
INSERT INTO tblPurchaseHistory VALUES ('T0460','02/19/2018 00:00:00',1500.00,1.98,1500.00,'45899');
INSERT INTO tblPurchaseHistory VALUES ('A7879','07/21/2015 00:00:00',10000.00,0.07,10000.00,'09567');
INSERT INTO tblPurchaseHistory VALUES ('A7879','03/15/2016 00:00:00',10000.00,0.08,9995.00,'09567');
INSERT INTO tblPurchaseHistory VALUES ('A7879','10/16/2017 00:00:00',8000.00,0.1,8000.00,'17453');
INSERT INTO tblPurchaseHistory VALUES ('A7879','03/15/2018 00:00:00',7500.00,0.12,7500.00,'17453');
INSERT INTO tblPurchaseHistory VALUES ('J8006','08/12/2018 00:00:00',450.00,1.1,450.00,'00216');
INSERT INTO tblPurchaseHistory VALUES ('J8006','08/12/2018 00:00:00',450.00,0.99,450.00,'17453');

CREATE TABLE tblPurchaseOrderLine
(PONumber			char(6) NOT NULL,
ProductID			char(5) NOT NULL,
QtyOrdered			decimal(6,2) NOT NULL CHECK (QtyOrdered > 0),
Price				money NOT NULL,
DateNeeded			datetime NOT NULL,
CONSTRAINT pk_PoProductDate
PRIMARY KEY (PONumber, ProductID, DateNeeded),
CONSTRAINT fk_PONumber
FOREIGN KEY (PONumber)
REFERENCES tblPurchaseOrder (PoNumber),
CONSTRAINT fk_ProductID2
FOREIGN KEY (ProductID)
REFERENCES tblProduct (ProductID));

INSERT INTO tblPurchaseOrderLine VALUES ('025974','M3577',600.00,4.63,'08/18/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('025974','O1957',300.00,0.46,'08/18/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('045687','C2399',40.00,2.99,'09/05/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('045687','L8500',15.00,29.64,'09/10/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('045687','O1957',450.00,0.29,'09/12/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('056489','M3577',600.00,5.29,'08/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('112233','M3577',600.60,5.89,'10/07/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('112233','P7844',500.25,1.5,'09/24/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('234607','C2399',75.00,1.75,'09/06/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('234607','G0983',200.00,1.5,'09/08/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('234607','G1366',182.00,4.89,'09/10/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('234607','G5698',182.00,1.22,'09/19/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('234607','R5660',100.00,1.99,'09/26/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('234607','T0460',200.00,2.75,'09/26/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('256887','C2399',20.00,2.5,'10/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('256887','P5678',48.00,22.5,'10/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('329987','O1957',300.00,0.41,'11/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('329987','T0460',50.00,3.1,'01/12/19 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('365870','G0983',120.00,1.85,'12/14/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('365870','R5660',75.00,1.59,'11/10/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('365870','T0460',120.00,2.39,'03/14/19 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('453313','J8006',125.50,1.02,'09/26/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('453313','J8006',275.50,1.02,'09/28/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('453313','J8006',100.00,1.07,'10/05/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G0983',100.00,1.89,'09/17/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G0983',20.00,2.19,'09/28/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G0983',20.00,2.25,'10/02/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G0983',25.00,2.25,'10/12/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G0983',45.00,2.15,'11/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G0983',25.00,2.15,'12/10/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G1366',48.00,1.89,'01/10/19 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G5698',12.00,1.95,'10/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','G5698',48.00,2.38,'01/10/19 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','T0460',120.00,1.89,'10/12/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('543791','T0460',50.00,2.49,'12/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('600124','G0983',100.00,1.96,'10/18/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('600124','G1366',150.00,4.85,'10/18/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('600124','G5698',100.00,1.87,'10/18/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('661677','L8500',8.00,26.45,'11/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('781900','C9100',800.00,3.45,'10/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('781900','M3577',750.00,5.55,'11/10/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('902347','G0983',100.00,2.25,'10/15/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('902347','T0460',100.00,2.25,'09/22/18 00:00:00');
INSERT INTO tblPurchaseOrderLine VALUES ('902347','T0460',450.00,1.39,'12/05/18 00:00:00');

CREATE TABLE tblReceiver
(ReceiverID				int IDENTITY(1,1),
DateReceived			datetime NOT NULL,
PONumber				char(6),
ProductID				char(5),
DateNeeded				datetime,
QtyReceived				decimal(6,2) CHECK (QtyReceived > 0),
ConditionID				char(2) NOT NULL,
ReceiveEmpID			char(6),
CONSTRAINT pk_ReceiverID
PRIMARY KEY (ReceiverID),
CONSTRAINT fk_PoProductDate
FOREIGN KEY (PONumber, ProductID, DateNeeded)
REFERENCES tblPurchaseOrderLine (PONumber, ProductID, DateNeeded),
CONSTRAINT fk_ConditionID
FOREIGN KEY (ConditionID)
REFERENCES tblCondition (ConditionID),
CONSTRAINT fk_ReceiveEmpID
FOREIGN KEY (ReceiveEmpID)
REFERENCES tblEmployee (EmpID));



INSERT INTO tblReceiver VALUES ('08/15/2018 00:00:00','025974','M3577','08/18/2018 00:00:00',400.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('08/16/2018 00:00:00','025974','M3577','08/18/2018 00:00:00',100.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('08/17/2018 00:00:00','025974','M3577','08/18/2018 00:00:00',100.00,'WD',NULL);
INSERT INTO tblReceiver VALUES ('08/15/2018 00:00:00','025974','O1957','08/18/2018 00:00:00',100.00,'OK','E10018');
INSERT INTO tblReceiver VALUES ('08/17/2018 00:00:00','025974','O1957','08/18/2018 00:00:00',210.00,'OK','E10018');
INSERT INTO tblReceiver VALUES ('09/05/2018 00:00:00','045687','C2399','09/05/2018 00:00:00',40.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('09/14/2018 00:00:00','045687','C2399','09/05/2018 00:00:00',11.00,'FD','E10042');
INSERT INTO tblReceiver VALUES ('09/05/2018 00:00:00','045687','O1957','09/12/2018 00:00:00',400.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('09/15/2018 00:00:00','045687','O1957','09/12/2018 00:00:00',40.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('09/07/2018 00:00:00','234607','C2399','09/06/2018 00:00:00',13.00,'FD','E10018');
INSERT INTO tblReceiver VALUES ('09/08/2018 00:00:00','234607','C2399','09/06/2018 00:00:00',22.00,'IP','E10018');
INSERT INTO tblReceiver VALUES ('09/12/2018 00:00:00','234607','C2399','09/06/2018 00:00:00',40.00,'WD','E10018');
INSERT INTO tblReceiver VALUES ('09/07/2018 00:00:00','234607','G0983','09/08/2018 00:00:00',200.00,'WD','E10042');
INSERT INTO tblReceiver VALUES ('09/07/2018 00:00:00','234607','G1366','09/10/2018 00:00:00',150.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('09/10/2018 00:00:00','234607','G1366','09/10/2018 00:00:00',25.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('09/14/2018 00:00:00','234607','G1366','09/10/2018 00:00:00',7.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('09/22/2018 00:00:00','234607','G5698','09/19/2018 00:00:00',182.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('09/22/2018 00:00:00','234607','R5660','09/26/2018 00:00:00',50.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('09/26/2018 00:00:00','234607','R5660','09/26/2018 00:00:00',50.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('09/22/2018 00:00:00','234607','T0460','09/26/2018 00:00:00',200.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('09/23/2018 00:00:00','256887','P5678','10/15/2018 00:00:00',40.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('09/25/2018 00:00:00','256887','P5678','10/15/2018 00:00:00',9.00,'OK',NULL);
INSERT INTO tblReceiver VALUES ('09/26/2018 00:00:00','256887','C2399','10/15/2018 00:00:00',20.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('09/25/2018 00:00:00','543791','G0983','09/17/2018 00:00:00',50.00,'FD','E10042');
INSERT INTO tblReceiver VALUES ('09/24/2018 00:00:00','543791','G0983','09/17/2018 00:00:00',50.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('09/29/2018 00:00:00','543791','G0983','09/28/2018 00:00:00',20.00,'FD','E10042');
INSERT INTO tblReceiver VALUES ('09/25/2018 00:00:00','543791','G0983','10/02/2018 00:00:00',20.00,'IP','E10042');
INSERT INTO tblReceiver VALUES ('10/09/2018 00:00:00','543791','G0983','10/12/2018 00:00:00',20.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('10/14/2018 00:00:00','543791','G0983','10/12/2018 00:00:00',5.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('10/19/2018 00:00:00','543791','G5698','10/15/2018 00:00:00',12.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('10/08/2018 00:00:00','543791','T0460','10/12/2018 00:00:00',100.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('10/09/2018 00:00:00','543791','T0460','10/12/2018 00:00:00',10.00,'WD','E10087');
INSERT INTO tblReceiver VALUES ('10/11/2018 00:00:00','543791','T0460','10/12/2018 00:00:00',10.00,'UK',NULL);
INSERT INTO tblReceiver VALUES ('10/15/2018 00:00:00','600124','G0983','10/18/2018 00:00:00',10.00,'WD','E10087');
INSERT INTO tblReceiver VALUES ('10/16/2018 00:00:00','600124','G0983','10/18/2018 00:00:00',100.00,'OK','E10087');
INSERT INTO tblReceiver VALUES ('10/28/2018 00:00:00','600124','G1366','10/18/2018 00:00:00',150.00,'UK','E10087');
INSERT INTO tblReceiver VALUES ('10/14/2018 00:00:00','600124','G5698','10/18/2018 00:00:00',10.00,'UK','E10042');
INSERT INTO tblReceiver VALUES ('10/16/2018 00:00:00','600124','G5698','10/18/2018 00:00:00',60.00,'OK','E10042');
INSERT INTO tblReceiver VALUES ('10/17/2018 00:00:00','600124','G5698','10/18/2018 00:00:00',30.00,'UK',NULL);
INSERT INTO tblReceiver VALUES ('10/14/2018 00:00:00','902347','G0983','10/15/2018 00:00:00',45.00,'UK','E10087');
INSERT INTO tblReceiver VALUES ('10/15/2018 00:00:00','902347','G0983','10/15/2018 00:00:00',45.00,'OK','E10101');
INSERT INTO tblReceiver VALUES ('10/05/2018 00:00:00','902347','G0983','10/15/2018 00:00:00',10.00,'OK','E10101');
INSERT INTO tblReceiver VALUES ('09/19/2018 00:00:00','902347','T0460','09/22/2018 00:00:00',95.00,'OK','E10101');
INSERT INTO tblReceiver VALUES ('09/24/2018 00:00:00','902347','T0460','09/22/2018 00:00:00',5.00,'OK','E10101');
INSERT INTO tblReceiver VALUES ('09/19/2018 00:00:00','902347','T0460','12/05/2018 00:00:00',100.00,'UK',NULL);
INSERT INTO tblReceiver VALUES ('09/23/2018 00:00:00','112233','P7844','09/24/2018 00:00:00',475.25,'OK','E10101');
INSERT INTO tblReceiver VALUES ('09/28/2018 00:00:00','112233','M3577','10/07/2018 00:00:00',300.45,'OK','E10101');
INSERT INTO tblReceiver VALUES ('10/25/2018 00:00:00','112233','M3577','10/07/2018 00:00:00',340.50,'WD','E10101');
INSERT INTO tblReceiver VALUES ('09/26/2018 00:00:00','453313','J8006','09/26/2018 00:00:00',125.50,'WD','E10087');
INSERT INTO tblReceiver VALUES ('09/28/2018 00:00:00','453313','J8006','09/28/2018 00:00:00',270.50,'OK','E10087');
INSERT INTO tblReceiver VALUES ('09/29/2018 00:00:00','453313','J8006','09/28/2018 00:00:00',5.00,'OK','E10087');

/*SELECT STATEMENTS HMWK 1
/* Question 1 */

SELECT		EmpLastName as EmployeeLastName,
			EmpFirstName as EmployeeFirstName,
			'(' + SUBSTRING(EmpPhone,1,3) + ') ' + SUBSTRING(EmpPhone,4,3) + '-' + SUBSTRING(EmpPhone,5,4) as EmployeePhoneNumber
FROM		tblEmployee
WHERE		EmpEmail is null
ORDER BY	EmpLastName;

/* Question 2 */
SELECT		VendorID,
			Name as VendorName,
			CONCAT(Address1 + ', ' + Upper(City), + ', '+ State) as VendorAddress,
			'(' + SUBSTRING(Phone,1,3) + ')' + SUBSTRING(Phone,4,3) + '-' + SUBSTRING(Phone,5,4) as VendorPhone,
			CONVERT(VARCHAR,FirstBuyDate,107) as FirstBuyDate 
FROM		tblVendor
WHERE		State = 'nv' or State = 'Ca'
ORDER BY	State;

/* Question 12 */

SELECT *
FROM tblReceiver;

Select		PONumber as 'Purchase Order Number',
			ProductID as 'Product Number',
			CONVERT(VARCHAR, DateNeeded, 107) as 'Date Needed',
			Count(PONumber) as 'Number of Receivers',
			Cast(Sum(QtyReceived) as INT) as 'Total Quantity Received'
FROM tblReceiver
Group BY PONumber, ProductID, DateNeeded
Order BY PONumber, ProductID, DateNeeded;			*/

/*
/*HMWK */
/*Question 1 -- GOTTA FORMAT*/

SElECT *
FROM tblPurchaseOrder;
SELECT *
FROM tblVendor;

SELECT			tblPurchaseOrder.PODatedNeeded,
				tblPurchaseOrder.PoNumber,
				tblPurchaseOrder.PODatePlaced,
				tblVendor.Name,
				tblVendor.Phone,
				tblPurchaseOrder.BuyerEmpID
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
WHERE			PODatedNeeded > '2018-09-30' and PODatedNeeded < '2018-12-01'
ORDER BY		PODatedNeeded DESC;

/* Question 2 *Gotta Format* */ 

SELECT *
FROM tblEmployee;

SELECT			tblPurchaseOrder.PODatedNeeded,
				tblPurchaseOrder.PoNumber,
				tblPurchaseOrder.PODatePlaced,
				tblVendor.Name,
				tblVendor.Phone,
				ISNULL(tblEmployee.EmpFirstName + ' ' + tblEmployee.EmpLastName, 'No Buyer') as 'Buyer Name'
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
WHERE			PODatedNeeded > '2018-09-30' and PODatedNeeded < '2018-12-01'
ORDER BY		PODatedNeeded DESC;

/* Question 3 -- Manager Name needed -- format needed */
/* Look into creating another table titled tblEmployee1 w/ EmpID, EmpMgrID, EmpFirstName, EmpLastName for self join*/

SELECT * FROM tblEmployee;

SELECT			tblPurchaseOrder.PODatedNeeded,
				tblPurchaseOrder.PoNumber,
				tblPurchaseOrder.PODatePlaced,
				tblVendor.Name,
				tblVendor.Phone,
				ISNULL(tblEmployee.EmpFirstName + ' '+ tblEmployee.EmpLastName, 'No Buyer') as 'Buyer Name',
				ISNULL(tblEmployee.EmpMgrID, 'No Manager') as 'Manager Name'
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
WHERE			PODatedNeeded > '2018-09-30' and PODatedNeeded < '2018-12-01'
ORDER BY		PODatedNeeded DESC;



/* Question 4  DONE */

SELECT *
FROM tblPurchaseOrder;
SELECT *
FROM tblReceiver;

SELECT			ProductID as 'Product ID',
				CONVERT(VARCHAR,DateReceived,107)as 'Date Product Received',
				PoNumber as 'Purchase Order Number',
				QtyReceived as 'Quantity Received',
				tblCondition.Description as 'Condition Received'
FROM			tblReceiver
INNER JOIN		tblCondition
ON				tblReceiver.ConditionID = tblCondition.ConditionID			
WHERE			tblCondition.ConditionID = 'FD' or tblCondition.ConditionID = 'WD'
ORDER BY		ProductID, DateReceived;



/* Question 5 */

SELECT *
FROm tblProduct;
SELECT *
FROM tblVendor;
SELECT *
FROM tblProductType;
SELECT *
FROM tblReceiver;
SELECT *
FROM tblCondition;

SELECT			tblreceiver.ProductID as 'Product ID',
				tblProduct.Description as 'Product Description',
				tblProductType.Description as 'Product Type Description',
				CONVERT(VARCHAR,tblreceiver.DateReceived,107)as 'Date Product Received',
				tblreceiver.PoNumber as 'Purchase Order Number',
				tblVendor.Name as 'Vendor Name',
				tblVendor.Phone as 'Vendor Phone Number',
				tblreceiver.QtyReceived as 'Quantity Received',
				tblCondition.Description as 'Condition Received'
FROM			tblReceiver
INNER JOIN		tblCondition
ON				tblReceiver.ConditionID = tblCondition.ConditionID
INNER JOIN		tblProduct
ON				tblReceiver.ProductID = tblProduct.ProductID			
WHERE			tblCondition.ConditionID = 'FD' or tblCondition.ConditionID = 'WD'
ORDER BY		tblreceiver.ProductID, tblreceiver.DateReceived;

/* Question 8 */

SELECT *
FROM tblPurchaseOrder;
select *
FROM tblPurchaseOrderLine;
SELECT *
fROM tblVendor;
select * from tblProduct;
Select * from tblEmployee;

SELECT			tblPurchaseOrder.PoNumber as 'Purchase Order Number',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced,107) as 'Date of Purchase Order',
				tblVendor.Name as 'Vendor Name',
				ISNULL(tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1) + '.', 'No Buyer') as 'Buyer',
				tblPurchaseOrderLine.ProductID as 'Product ID',
				tblproduct.Description as 'Product Description',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded, 107) as 'Product Date Needed',
				tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
				tblPurchaseOrderLine.Price as 'Unit Price',
				CONVERT(decimal(10,2),(tblPurchaseOrderLine.QtyOrdered * tblPurchaseOrderLine.Price)) as 'Extended Price'
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN		tblProduct
ON				tblPurchaseOrderLine.ProductID = tblProduct.ProductID;

/* Question 9 */

SELECT * FROM tblReceiver;

SELECT			tblPurchaseOrder.PoNumber as 'Purchase Order Number',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced,107) as 'Date of Purchase Order',
				tblVendor.Name as 'Vendor Name',
				ISNULL(tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1) + '.', 'No Buyer') as 'Buyer',
				tblPurchaseOrderLine.ProductID as 'Product ID',
				tblproduct.Description as 'Product Description',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded, 107) as 'Product Date Needed',
				tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
				ISNULL(CONVERT(VARCHAR,tblReceiver.DateReceived,107), 'Not Received') as 'Date Product Received',
				ISNULL(tblReceiver.QtyReceived, 0.00) as 'Quantity Received'
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN		tblProduct
ON				tblPurchaseOrderLine.ProductID = tblProduct.ProductID
LEFT JOIN		tblReceiver
ON				tblPurchaseOrderLine.ProductID = tblReceiver.ProductID
AND				tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
AND				tblPurchaseOrderLine.DateNeeded = tblReceiver.DateNeeded;

/* question 10 */

SELECT			tblPurchaseOrder.PoNumber as 'Purchase Order Number',
				tblVendor.Name as 'Vendor Name',
				tblPurchaseOrderLine.ProductID as 'Product ID',
				tblproduct.Description as 'Product Description',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded, 107) as 'Product Date Needed',
				tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
				ISNULL(SUM(tblReceiver.QtyReceived),0.00) as 'Total Quantity Received',
				(tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) as 'Quantity Remaining to be Received',
				CASE
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) = 0
					THEN 'Complete'
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) < 0
					THEN 'Over Shipment'
					WHEN ISNULL(SUM(tblReceiver.QtyReceived),0.00) = 0.00
					THEN 'Not Received'
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) > 0
					THEN 'Partial Shipment'
					END 'Receiving Status'
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN		tblProduct
ON				tblPurchaseOrderLine.ProductID = tblProduct.ProductID
LEFT JOIN		tblReceiver
ON				tblPurchaseOrderLine.ProductID = tblReceiver.ProductID
AND				tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
AND				tblPurchaseOrderLine.DateNeeded = tblReceiver.DateNeeded
GROUP BY		tblPurchaseOrder.PoNumber, tblVendor.Name, tblPurchaseOrderLine.ProductID,tblproduct.Description,tblPurchaseOrder.PODatedNeeded, tblPurchaseOrderLine.QtyOrdered  
ORDER BY		tblPurchaseOrder.PoNumber, tblPurchaseOrderLine.ProductID, tblPurchaseOrder.PODatedNeeded;

/* QUESTION 12*/


select				tblPurchaseOrderLine.PONumber 'Purchase Order Number',
                    tblPurchaseOrderLine.productID 'Product ID',
                    tblProduct.description 'Description',
                    tblPurchaseOrderLine.Price 'Price',
					(SELECT MAX(tblpurchaseorderline.price) FROM tblPurchaseOrderLine pol WHERE tblPurchaseOrderLine.Price = pol.Price) 'Price1'
from				tblPurchaseOrderLine
inner join			tblProduct
on					tblPurchaseOrderLine.ProductID = tblProduct.ProductID   
where				tblProduct.Description='Alpine Small Pot';

select				tblPurchaseOrderLine.PONumber 'Purchase Order Number',
                    tblPurchaseOrderLine.productID 'Product ID',
                    (SELECT tblProduct.description FROM tblProduct WHERE tblproduct.Description = 'Alpine Small Pot') 'Description',
                    (SELECT MAX(tblPurchaseOrderLine.Price) 
					 FROM tblPurchaseOrderLine) 'Price'
from				tblPurchaseOrderLine
INNER join			tblProduct
on					tblPurchaseOrderLine.ProductID = tblProduct.ProductID;
																				
*/

/* hmwk */

/* Question 1 (Too many rows popping up)*/

CREATE VIEW vQuantityReceived AS
SELECT			tblReceiver.PONumber,
				tblReceiver.ProductID,
				tblReceiver.DateNeeded,
				ISNULL(SUM(tblReceiver.QtyReceived),0.00) SumQtyReceived
FROM			tblReceiver
GROUP BY		tblReceiver.PONumber, tblReceiver.ProductID, tblReceiver.DateNeeded;

SELECT *
FROM vQuantityReceived;

select qtyreceived
from tblReceiver;
				

SELECT			vQuantityReceived.PoNumber as 'PO Number',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced, 107) as 'PO Date',
				tblVendor.Name as 'Vendor Name',
				ISNULL(tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1) + '.', 'No Buyer on File') as 'Employee Buyer',
				CASE
                       	WHEN (tblEmployee.EmpMgrID) IS NULL 
						THEN 'No Manager on File'
                        ELSE (Manager.EmpLastName + ', ' + SUBSTRING(Manager.EmpFirstName,1,1))
                       	END 'Manager of Buyer',
				vQuantityReceived.ProductID as 'Product ID',
				tblproduct.Description as 'Product Description',
				CONVERT(VARCHAR,vQuantityReceived.DateNeeded, 107) as 'Date Needed',
				tblPurchaseOrderLine.Price as 'Product Price',
				tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
				vQuantityReceived.SumQtyReceived as 'Quantity Received',
				(tblPurchaseOrderLine.QtyOrdered - vQuantityReceived.SumQtyReceived) as 'Quantity Remaining',
				CASE
					WHEN (tblPurchaseOrderLine.QtyOrdered - vQuantityReceived.SumQtyReceived) = 0
					THEN 'Complete'
					WHEN (tblPurchaseOrderLine.QtyOrdered - vQuantityReceived.SumQtyReceived) < 0
					THEN 'Over Shipment'
					WHEN vQuantityReceived.SumQtyReceived = 0.00
					THEN 'Not Received'
					WHEN (tblPurchaseOrderLine.QtyOrdered - vQuantityReceived.SumQtyReceived) > 0
					THEN 'Partial Shipment'
					END 'Receiving Status'
FROM			tblPurchaseOrderLine
INNER JOIN		tblPurchaseOrder
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
INNER JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
LEFT JOIN		tblEmployee Manager
ON				tblEmployee.EmpMgrID = Manager.EmpID
INNER JOIN		tblProduct
ON				tblPurchaseOrderLine.ProductID = tblProduct.ProductID
LEFT JOIN		tblReceiver
ON				tblPurchaseOrderLine.ProductID = tblReceiver.ProductID
AND				tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
AND				tblPurchaseOrderLine.DateNeeded = tblReceiver.DateNeeded
INNER JOIN		vQuantityReceived
ON				tblPurchaseOrderLine.ProductID = vQuantityReceived.ProductID
AND				tblPurchaseOrderLine.DateNeeded = vQuantityReceived.DateNeeded
AND				tblPurchaseOrderLine.PoNumber = vQuantityReceived.PONumber
ORDER BY		vQuantityReceived.PONumber, vQuantityReceived.ProductID, vQuantityReceived.DateNeeded;

SELECT *
FROM tblReceiver;

/* Question 2 */

