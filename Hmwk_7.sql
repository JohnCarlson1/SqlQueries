/*HMWK 7*/

/* 1 */
SELECT    	emplastname 'EmployeeLastName',
   	   		empfirstname 'EmployeeFirstName',
   			'(' + substring(empphone,1,3) + ') ' +
              substring(empphone,4,3) + '-' +
              substring(empphone,7,4) 'Employee Phone Number'
FROM      	tblemployee
WHERE     	empemail IS NULL
ORDER BY	EmpLastName;

/* 2 */

SELECT    	vendorid 'VendorID',
			name 'VendorName',
          	concat(address1, ', ', address2, ', ', city, ', ', state) 'VendorAddress',
          	'(' + substring(phone,1,3) + ') ' +
              substring(phone,4,3) + '-' +
              substring(phone,7,4) 'VendorPhone',
          	convert(varchar,firstbuydate,107) 'FirstBuyDate'
FROM      	tblvendor
WHERE     	state='NV' or state='CA'
ORDER BY  	state;

/* 3 */

SELECT    	convert(varchar,DateNeeded,107) 'Date Needed',
          	ProductID 'Product Number',
          	PONumber 'Purchase Order Number',
          	QtyOrdered 'Quantity Ordered',
          	Price 'Price',
          	cast((Price * QtyOrdered) as decimal (20,2)) 'Extended Price'
FROM      	tblPurchaseOrderLine
WHERE     	DateNeeded > dateadd(day, 68, GetDate())
ORDER BY  	DateNeeded;

/* 4 */

SELECT    	max(Price) 'Most Expensive Selling Price for Product G0983'
FROM      	tblPurchaseOrderLine
WHERE     	ProductID = 'G0983';

/* 5 */

SELECT 	cast(sum(QtyOrdered * Price) as decimal (20,2)) 'Total Order Price for September'
FROM   	tblPurchaseOrderLine
WHERE 		MONTH(DateNeeded) = 9 AND YEAR(DateNeeded) = 2018;

/* 6 */

SELECT    	ProductID 'Product Number',
description 'Product Description',
EOQ 'Economic Order Quantity',
QOH 'Quantity On Hand',
(QOH - EOQ) 'Difference',
CASE
          	WHEN ((EOQ - QOH) >= 50) THEN 'Order Now'
          	ELSE NULL
          	END AS OrderMessage
FROM      	tblProduct
WHERE     	(QOH - EOQ) < 0
ORDER BY  	ProductID;

/* 7 */

SELECT    	ProductID 'Product Number',
			description 'Product Description',
			EOQ 'Economic Order Quantity',
			QOH 'Quantity On Hand',
			(QOH - EOQ) 'Difference',
			CASE
          		WHEN QOH = 0 THEN 'Order Immediately'
          		WHEN (QOH - EOQ) BETWEEN -5 AND -1 THEN NULL
          		WHEN (QOH - EOQ) BETWEEN -10 AND -6 THEN 'Order next month'
          		WHEN (QOH - EOQ) BETWEEN -35 AND -11 THEN 'Order next week'
          		WHEN (QOH - EOQ) < -35 THEN 'Order This Week'
          		END AS OrderMessage
FROM      	tblProduct
WHERE     	(QOH - EOQ) < 0
ORDER BY  	EOQ desc;

/* 8 */

SELECT DISTINCT  	ProductTypeID 'Product Type ID',
					count(ALL ProductTypeID) 'Count of Products',
					sum(QOH) 'Total Quantity On Hand',
					cast(AVG(QOH) as decimal (20,2)) 'Average Quantity On Hand'
FROM             	tblProduct
GROUP BY         	ProductTypeID;

/* 9 */

SELECT DISTINCT  	ProductTypeID 'Product Type ID',
					count(ALL ProductTypeID) 'Count of Products',
					sum(QOH) 'Total Quantity On Hand',
					cast(AVG(QOH) as decimal (20,2)) 'Average Quantity On Hand'
FROM             	tblProduct
GROUP BY         	ProductTypeID
HAVING           	avg(QOH) > 50;

/* 10 */

SELECT DISTINCT  	ProductID 'Product Number',
					count(HistoryID) 'Number Of Times Purchased',
					max(DatePurchased) 'Last Date Purchased',
					min(price) 'Minimum Purchase Price',
					max(price) 'Maximum Purchase Price',
					(max(price) - min(price)) 'Difference Between Maximum And Minimum   Price'
FROM             	tblPurchaseHistory
GROUP BY         	ProductID;

/* 11 */

SELECT DISTINCT 	ProductID 'Product Number',
					count(HistoryID) 'Number Of Times Purchased',
					max(DatePurchased) 'Last Date Purchased',
					min(price) 'Minimum Purchase Price',
					max(price) 'Maximum Purchase Price',
					(max(price) - min(price)) 'Difference Between Maximum And Minimum Price'
FROM 				tblPurchaseHistory
GROUP BY 			ProductID
HAVING 				(max(Price) - min(Price)) = 0;

/* 12 */

SELECT		ponumber AS 'Purchase Order Number',
			productid AS 'Product Number',
			CONVERT(VARCHAR, dateneeded, 107) AS 'Date Needed',
			COUNT(PONumber) AS 'Number of Receivers',
			CAST(Sum(qtyreceived) AS INT) AS 'Total Quantity Received'
FROM 		tblReceiver
GROUP BY 	ponumber, productid, dateneeded
ORDER BY 	ponumber, productid, dateneeded;

/* 13 */

SELECT			CONVERT(varchar, PODatedNeeded, 107) 'Date Needed',
				CONVERT(varchar, Getdate(), 107) 'Todays Date',
				DateDiff(day,  PODatedNeeded, Getdate())   'Days Overdue',
				PONumber 'Purchase Order Number',
				CONVERT(varchar, PODatePlaced, 107) 'Date Placed',
				BuyerEmpID  'Buyer',
				VendorID
FROM			tblPurchaseOrder
WHERE			DATEDIFF(day, PODatedNeeded, GetDate()) > 7;

/* 14 */

select		*
from		tblproduct
where		ProductID NOT IN
					(select ProductID
					 from tblpurchaseorderline);

/* 15 */

select		PONumber 'Purchase Order Number',
			ProductID 'Product Number',
			Price 'Purchase Price'
from		tblPurchaseOrderLine
where		ProductID = 'G0983'
and			Price = (select max(Price)
					 from tblPurchaseOrderLine
					 where ProductID = 'G0983');


