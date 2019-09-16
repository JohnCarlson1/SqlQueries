/* HMWK 8 */

/* 1 */

SELECT           	CONVERT(varchar, tblPurchaseOrder.PODatedNeeded, 107) 'Date of Required Delivery',
                 	tblPurchaseOrder.PoNumber 'Purchase Order Number',
                     CONVERT(varchar, tblPurchaseOrder.PODatePlaced,107) 'Date of Purchase Order',
                 	tblVendor.Name 'Vendor Name',
                 	'(' + substring(tblVendor.Phone,1,3) + ') ' +
                                     	substring(tblVendor.Phone,4,3) + '-' +
                                                   	substring(tblVendor.Phone,7,4) 'Vendor Phone Number',
                    tblPurchaseOrder.BuyerEmpID 'Buyer ID'
FROM             	tblPurchaseOrder
INNER JOIN       	tblVendor
ON                  tblPurchaseOrder.VendorID = tblVendor.VendorID
WHERE            	PODatedNeeded > '2018-09-30' and PODatedNeeded < '2018-12-01'
ORDER BY         	PODatedNeeded DESC;

/* 2 */

SELECT           	convert(varchar,tblPurchaseOrder.PODatedNeeded,107) 'Date of Required Delivery',
                 	tblPurchaseOrder.PoNumber 'Purchase Order Number',
                     convert(varchar,tblPurchaseOrder.PODatePlaced,107) 'Date of Purchase Order',
                 	tblVendor.Name 'Vendor Name',
                 	'(' + substring(tblVendor.Phone,1,3) + ') ' +
                 	substring(tblVendor.Phone,4,3) + '-' +             	   	          	   	substring(tblVendor.Phone,7,4) 'Vendor Phone Number',
                     ISNULL(tblEmployee.EmpFirstName + ' ' + tblEmployee.EmpLastName, 'No Buyer') as 'Buyer Name'
FROM             	tblPurchaseOrder
INNER JOIN       	tblVendor
ON                   tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN        	tblEmployee
ON                   tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
WHERE            	PODatedNeeded > '2018-09-30' and PODatedNeeded < '2018-12-01'
ORDER BY         	PODatedNeeded DESC;

/* 3 */

SELECT           	convert(varchar,tblPurchaseOrder.PODatedNeeded,107) 'Date of Required Delivery',
                 	tblPurchaseOrder.PoNumber 'Purchase Order Number',
                 	convert(varchar,tblPurchaseOrder.PODatePlaced,107) 'Date of Purchase Order',
                 	tblVendor.Name 'Vendor Name',
                 	'(' + substring(tblVendor.Phone,1,3) + ') ' +
                       	substring(tblVendor.Phone,4,3) + '-' +
                       	substring(tblVendor.Phone,7,4) 'Vendor Phone Number',
                     ISNULL(tblEmployee.EmpFirstName + ' '+ tblEmployee.EmpLastName, 'No Buyer') as 'Buyer Name',
                 	case
                       	when (tblEmployee.EmpMgrID) IS NULL then 'No Manager'
                              	else (manager.EmpFirstName + ' ' + manager.EmpLastName)
                       	end 'Manager Name'
FROM             	tblPurchaseOrder
 
INNER JOIN       	tblVendor
ON                   tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT OUTER JOIN  	tblEmployee
ON                   tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
LEFT OUTER JOIN        	 tblEmployee manager
on                            	 tblEmployee.EmpMgrID = manager.EmpID
 
WHERE            	PODatedNeeded > '2018-09-30' and PODatedNeeded < '2018-12-01'
ORDER BY         	PODatedNeeded DESC;

/* 4 */

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

/* 5 */

SELECT  	tblReceiver.ProductID as 'Product ID',
 			CONVERT(VARCHAR,DateReceived,107)as 'Date Product Received',
   	 		tblProduct.Description 'Product Description',
   	 		tblProductType.Description 'Product Type Description',
        	tblReceiver.PONumber as 'Purchase Order Number',
   	 		tblVendor.Name 'Vendor Name',
			'(' + substring(tblVendor.Phone,1,3) + ') ' +
          		substring(tblVendor.Phone,4,3) + '-' +
			substring(tblVendor.Phone,7,4) 'Vendor Phone Number',
        	QtyReceived as 'Quantity Received',
        	tblCondition.Description as 'Condition Received'
 
FROM    	tblReceiver
INNER JOIN	tblCondition
ON        	tblReceiver.ConditionID = tblCondition.ConditionID
INNER JOIN	tblProduct
on        	tblReceiver.ProductID = tblProduct.ProductID
INNER JOIN    tblProductType
on        	tblProductType.ProductTypeID = tblProduct.ProductTypeID
INNER JOIN    tblPurchaseOrder
on        	tblReceiver.PONumber = tblPurchaseOrder.PoNumber
INNER JOIN    tblVendor
on        	tblPurchaseOrder.VendorID = tblVendor.VendorID 
WHERE     	tblCondition.ConditionID = 'FD' or tblCondition.ConditionID = 'WD'
ORDER BY  	tblReceiver.ProductID, DateReceived;

/* 6 */

SELECT distinct 	tblVendor.Name 'Vendor Name',
                 	'(' + substring(tblVendor.Phone,1,3) + ') ' +
                       	substring(tblVendor.Phone,4,3) + '-' +
                       	substring(tblVendor.Phone,7,4) 'Vendor Phone Number'
FROM    	     	tblReceiver 
INNER JOIN	   		tblCondition
ON        	   		tblReceiver.ConditionID = tblCondition.ConditionID
inner join	   		tblProduct
on               	tblReceiver.ProductID = tblProduct.ProductID
inner join       	tblProductType
on               	tblProductType.ProductTypeID = tblProduct.ProductTypeID
inner join       	tblPurchaseOrder
on               	tblReceiver.PONumber = tblPurchaseOrder.PoNumber
inner join       	tblVendor
on               	tblPurchaseOrder.VendorID = tblVendor.VendorID
WHERE        		tblCondition.ConditionID = 'FD' or tblCondition.ConditionID = 'WD' AND 	          	tblReceiver.ProductID = 'G0983'
ORDER BY    	 	tblVendor.Name desc;

/* 7 */

SELECT distinct     tblProduct.ProductID 'Product ID',
          			CASE
                 		WHEN (count(tblPurchaseOrderLine.ProductID) =0 ) then '0'
                 		ELSE count(tblPurchaseOrderLine.ProductID)
                 		END 'Number of Purchase Orders',
          			CASE
                 		WHEN (count(tblPurchaseOrderLine.ProductID) =0 ) then '0'
                 		ELSE cast(sum (tblPurchaseOrderLine.QtyOrdered) as int)
                 		END 'Total Quantity Ordered',
          			CASE
                 		WHEN (count(tblPurchaseOrderLine.ProductID) =0 ) then '0.00'
                 		ELSE cast(max(tblPurchaseOrderLine.Price) as money)
                 		END 'Maximum Price Paid',
          			CASE
                 		WHEN (count(tblPurchaseOrderLine.ProductID) =0 ) then '0.00'
                 		ELSE cast(min(tblPurchaseOrderLine.Price) as money)
                 		END 'Minimum Price Paid',
          			CASE
                 		WHEN (count(tblPurchaseOrderLine.ProductID) =0 ) then '0.00'
                 		ELSE cast(avg(tblPurchaseOrderLine.Price) as decimal(18,2))
                 		END 'Average Price Paid' 
from				tblProduct
 
left outer join		tblPurchaseOrderLine
on					tblPurchaseOrderLine.ProductID = tblProduct.ProductID
 
group by			tblProduct.ProductID;

/* 8 */

SELECT           	tblPurchaseOrder.PoNumber as 'Purchase Order Number',
                    CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced,107) as 'Date of Purchase Order',
                    tblVendor.Name as 'Vendor Name',
                    ISNULL(tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1) + '.', 'No Buyer') as 'Buyer',
                    tblPurchaseOrderLine.ProductID as 'Product ID',
                    tblproduct.Description as 'Product Description',
                    CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded, 107) as 'Product Date Needed',
                    tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
                    tblPurchaseOrderLine.Price as 'Unit Price',
                    CONVERT(decimal(10,2),(tblPurchaseOrderLine.QtyOrdered * tblPurchaseOrderLine.Price)) as 'Extended Price'
FROM             	tblPurchaseOrder
INNER JOIN       	tblVendor
ON                  tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN        	tblEmployee
ON                  tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN       	tblPurchaseOrderLine
ON                  tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN       	tblProduct
ON                  tblPurchaseOrderLine.ProductID = tblProduct.ProductID;

/* 9 */

SELECT           	tblPurchaseOrder.PoNumber as 'Purchase Order Number',
                 	CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced,107) as 'Date of Purchase Order',
                 	tblVendor.Name as 'Vendor Name',
                 	ISNULL(tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1) + '.', 'No Buyer') as 'Buyer',
                 	tblPurchaseOrderLine.ProductID as 'Product ID',
                 	tblproduct.Description as 'Product Description',
                 	CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded, 107) as 'Product Date Needed',
                 	tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
                 	ISNULL(CONVERT(VARCHAR,tblReceiver.DateReceived,107), 'Not Received') as 'Date Product Received',
                 	ISNULL(tblReceiver.QtyReceived, 0.00) as 'Quantity Received'
FROM             	tblPurchaseOrder
INNER JOIN       	tblVendor
ON               	tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN        	tblEmployee
ON               	tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN       	tblPurchaseOrderLine
ON               	tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN       	tblProduct
ON               	tblPurchaseOrderLine.ProductID = tblProduct.ProductID
LEFT JOIN        	tblReceiver
ON               	tblPurchaseOrderLine.ProductID = tblReceiver.ProductID
AND              	tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
AND              	tblPurchaseOrderLine.DateNeeded = tblReceiver.DateNeeded;

/* 10 */

SELECT           	 tblPurchaseOrder.PoNumber as 'Purchase Order Number',
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
FROM             	 tblPurchaseOrder
INNER JOIN       	 tblVendor
ON                   tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN        	 tblEmployee
ON                   tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN       	 tblPurchaseOrderLine
ON                   tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN       	 tblProduct
ON                   tblPurchaseOrderLine.ProductID = tblProduct.ProductID
LEFT JOIN        	 tblReceiver
ON                   tblPurchaseOrderLine.ProductID = tblReceiver.ProductID
AND                  tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
AND                  tblPurchaseOrderLine.DateNeeded = tblReceiver.DateNeeded
GROUP BY         	 tblPurchaseOrder.PoNumber, tblVendor.Name, tblPurchaseOrderLine.ProductID,tblproduct.Description,tblPurchaseOrder.PODatedNeeded, tblPurchaseOrderLine.QtyOrdered 
ORDER BY         	 tblPurchaseOrder.PoNumber, tblPurchaseOrderLine.ProductID, tblPurchaseOrder.PODatedNeeded;

/* 11 */

SELECT     	  		tblPurchaseOrder.PoNumber as 'Purchase Order Number',
            	 	PODatePlaced as 'Date Placed',
            	 	PODatedNeeded as 'Date Needed',
            	 	Terms as 'Terms',
                    tblVendor.VendorID 'Vendor Number',
            	 	Name as 'Vendor Name'
FROM             	tblPurchaseOrder
RIGHT OUTER JOIN 	TblVendor
ON                  tblPurchaseOrder.VendorID = tblVendor.VendorID
where            	tblPurchaseOrder.PoNumber NOT IN (select ponumber from tblReceiver)
ORDER BY         	tblPurchaseOrder.PoNumber Asc;

/* 12 */

select distinct  	PONumber,
   	          		tblProduct.ProductID,
   	          		tblProduct.Description,
   	          		Price
from             	tblPurchaseOrderLine
INNER JOIN       	tblProduct
on               	tblProduct.ProductID = tblPurchaseOrderLine.ProductID
WHERE				price = (SELECT max(tblPurchaseOrderLine.Price) 
							 from tblPurchaseOrderLine 
							 where tblProduct.ProductID = tblPurchaseOrderLine.ProductID)
AND              	tblProduct.Description = 'Alpine Small Pot';

/* 13 */

select distinct  	tblPurchaseOrderLine.PONumber,
          	   		tblVendor.Name,
   	          		tblProduct.ProductID,
   	          		tblProduct.Description,
   	          		Price
from             	tblPurchaseOrderLine
INNER JOIN       	tblProduct
on               	tblProduct.ProductID = tblPurchaseOrderLine.ProductID
INNER JOIN       	tblPurchaseOrder
on               	tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN       	tblVendor
on               	tblVendor.VendorID = tblPurchaseOrder.VendorID
WHERE				price = (SELECT max (tblPurchaseOrderLine.Price) 
							 from tblPurchaseOrderLine 
							 where tblProduct.ProductID = tblPurchaseOrderLine.ProductID)
AND              	tblProduct.Description = 'Alpine Small Pot';

/* 14 */

select distinct  	tblPurchaseOrderLine.PONumber,
          	   		tblVendor.Name,
          	   		tblEmployee.EmpLastName 'Purchase Employee',
   	          		tblProduct.ProductID,
   	          		tblProduct.Description,
   	          		tblProductType.Description 'Product Type Description',
   	          		Price
from             	tblPurchaseOrderLine
INNER JOIN       	tblProduct
on        	   		tblProduct.ProductID = tblPurchaseOrderLine.ProductID
INNER JOIN       	tblProductType
on        	   		tblProduct.ProductTypeID = tblProductType.ProductTypeID
INNER JOIN       	tblPurchaseOrder
on        			tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN       	tblEmployee
on        	   		tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
INNER JOIN       	tblVendor
on        			tblVendor.VendorID = tblPurchaseOrder.VendorID
WHERE				price = (SELECT max (tblPurchaseOrderLine.Price) 
							 from tblPurchaseOrderLine 
							 where tblProduct.ProductID = tblPurchaseOrderLine.ProductID)
AND              	tblProduct.Description = 'Alpine Small Pot';

/* 15 */

select distinct  	tblProduct.productID 'Product ID',
                 	description 'Product Description',
                 	EOQ 'Product Economic Order Quantity',
                 	ISNULL(cast(tblPurchaseHistory.DatePurchased as varchar(100)),
                       	'Not In Purchase History') 'Most Recent Purchase Date',
					ISNULL(cast(tblPurchaseHistory.Qty as varchar(10)), '--') 'Quantity Purchased',
					ISNULL(cast(tblPurchaseHistory.Price as varchar(10)), '--') 'Purchase Price'
from             	tblProduct
LEFT JOIN        	tblPurchaseHistory
on               	tblProduct.ProductID = tblPurchaseHistory.ProductID
WHERE            	tblPurchaseHistory.DatePurchased IS NULL
OR               	tblPurchaseHistory.Price = (SELECT max (tblPurchaseHistory.Price)
												 from tblPurchaseHistory   
												 where tblProduct.ProductID = tblPurchaseHistory.ProductID);



