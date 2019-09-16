/* Question 1 */
WITH cteQtyReceived AS
(SELECT			SUM(qtyReceived) qreceived,
				PONumber,
				ProductID,
				DateNeeded
FROM			tblReceiver
GROUP BY		PONumber,ProductID,DateNeeded)
				
SELECT			tblPurchaseOrder.PoNumber as 'Purchase Order Number',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced,107) as 'PO Date',
				tblVendor.Name as 'Vendor Name',
				ISNULL((tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1) + '.'), 'No Buyer on File') as 'Employee Buyer',
				case
                       	when (tblEmployee.EmpMgrID) IS NULL then 'No Manager on File'
                        else (manager.EmpLastName + ', ' + SUBSTRING(manager.empfirstname,1,1))
                       	end 'Manager of Buyer',
				tblPurchaseOrderLine.ProductID as 'Product ID',
				tblproduct.Description as 'Product Description',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded, 107) as 'Product Date Needed',
				tblPurchaseOrderLine.Price as 'Product Price',
				tblPurchaseOrderLine.QtyOrdered as 'Quantity Ordered',
				ISNULL(cteQtyReceived.qreceived,0.00) as 'Quantity Received',
				(tblPurchaseOrderLine.QtyOrdered - ISNULL(cteQtyReceived.qreceived,0.00)) as 'Quantity Remaining',
				CASE
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(cteQtyReceived.qreceived,0.00)) = 0
					THEN 'Complete'
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(cteQtyReceived.qreceived,0.00)) < 0
					THEN 'Over Shipment'
					WHEN ISNULL(cteQtyReceived.qreceived,0.00) = 0.00
					THEN 'Not Received'
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(cteQtyReceived.qreceived,0.00)) > 0
					THEN 'Partially Received'
					END 'Receiving Status'
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblPurchaseOrder.VendorID = tblVendor.VendorID
LEFT JOIN		tblEmployee
ON				tblPurchaseOrder.BuyerEmpID = tblEmployee.EmpID
LEFT JOIN		tblEmployee Manager
ON				tblEmployee.EmpMgrID = Manager.EmpID
INNER JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN		tblProduct
ON				tblPurchaseOrderLine.ProductID = tblProduct.ProductID
LEFT JOIN		cteQtyReceived
ON				cteQtyReceived.ProductID = tblPurchaseOrderLine.ProductID
AND				tblPurchaseOrderLine.PONumber = cteQtyReceived.PONumber
AND				tblPurchaseOrderLine.DateNeeded = cteQtyReceived.DateNeeded  
ORDER BY		tblPurchaseOrder.PoNumber, tblPurchaseOrderLine.ProductID, tblPurchaseOrder.PODatedNeeded;

/* Question 2 */

WITH cteReceiver AS
(SELECT			SUM(qtyreceived) qreceived,
				PONumber,
				ProductID,
				Dateneeded
FROM			tblReceiver
GROUP BY		PONumber,ProductID,DateNeeded),

cteSubtract AS 
(SELECT			(QtyOrdered - cteReceiver.qreceived) SUB,
				tblpurchaseorderline.PONumber,
				tblpurchaseorderline.ProductID,
				tblPurchaseOrderLine.DateNeeded
FROM			tblPurchaseOrderLine
INNER JOIN		cteReceiver
ON				tblPurchaseOrderLine.PONumber = cteReceiver.PONumber)

SELECT	distinct		tblpurchaseorder.PoNumber as 'PONumber',
						CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced, 107) as 'PODatePlaced',
						CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded,107) as 'PODateNeeded',
						tblVendor.Name
FROM			tblPurchaseOrder
LEFT JOIN		tblVendor
ON				tblVendor.VendorID = tblPurchaseOrder.VendorID
LEFT JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
left JOIN		cteSubtract
ON				cteSubtract.PONumber = tblPurchaseOrderLine.PONumber
WHERE			tblPurchaseOrder.PoNumber IN (SELECT PoNumber FROM cteSubtract)
AND				cteSubtract.SUB <= 0;


WITH cteReceiver AS
(SELECT			SUM(qtyreceived) qreceived,
				PONumber,
				ProductID,
				Dateneeded
FROM			tblReceiver
GROUP BY		PONumber,ProductID,DateNeeded)

SELECT	distinct		tblpurchaseorder.PoNumber as 'PONumber',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced, 107) as 'PODatePlaced',
				CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded,107) as 'PODateNeeded',
				tblVendor.Name
FROM			tblPurchaseOrder
INNER JOIN		tblVendor
ON				tblVendor.VendorID = tblPurchaseOrder.VendorID
INNER JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN		cteReceiver
ON				cteReceiver.PONumber = tblPurchaseOrder.PoNumber
WHERE			cteReceiver.PoNumber NOT IN (SELECT distinct PoNumber FROM tblPurchaseOrderLine
											WHERE tblPurchaseOrderLine.QtyOrdered - cteReceiver.qreceived = 0);

/* Question 3 */
WITH cteReceiver AS
(SELECT			SUM(qtyreceived) qreceived,
				PONumber,
				ProductID,
				Dateneeded
FROM			tblReceiver
GROUP BY		PONumber,ProductID,DateNeeded),

cteSubtract AS 
(SELECT			CASE
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(cteReceiver.qreceived,0.00)) = 0
					THEN 'Complete'
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(cteReceiver.qreceived,0.00)) < 0
					THEN 'Over Shipment'
					WHEN ISNULL(cteReceiver.qreceived,0.00) = 0.00
					THEN 'Not Received'
					WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(cteReceiver.qreceived,0.00)) > 0
					THEN 'Partially Received'
					END 'Receiving Status',
				tblpurchaseorderline.PONumber,
				tblpurchaseorderline.ProductID,
				tblPurchaseOrderLine.DateNeeded
FROM			tblPurchaseOrderLine
INNER JOIN		cteReceiver
ON				tblPurchaseOrderLine.PONumber = cteReceiver.PONumber),

cteOrderStatus AS
(SELECT			[Receiving Status]	
from			cteSubtract
WHERE			[Receiving Status] NOT IN ('Complete', 'Over Shipment'))

SELECT	distinct		tblpurchaseorder.PoNumber as 'PONumber',
						CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced, 107) as 'PODatePlaced',
						CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded,107) as 'PODateNeeded',
						tblVendor.Name
FROM			tblPurchaseOrder
LEFT JOIN		tblVendor
ON				tblVendor.VendorID = tblPurchaseOrder.VendorID
LEFT JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
left JOIN		cteSubtract
ON				cteSubtract.PONumber = tblPurchaseOrderLine.PONumber
WHERE			tblPurchaseOrder.PoNumber NOT IN (SELECT [Receiving Status] FROM cteOrderStatus);

/* Question 4 */

WITH cteReceiver AS
(SELECT              SUM(qtyreceived) qreceived,
                     tblreceiver.PONumber,
                     tblreceiver.ProductID,
                     tblreceiver.DateNeeded,
					 SUM(qtyOrdered) qOrdered
FROM             	tblReceiver
left join			tblPurchaseOrderLine
ON					tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
AND					tblPurchaseOrderLine.ProductID = tblReceiver.ProductID
AND					tblPurchaseOrderLine.DateNeeded = tblReceiver.DateNeeded
GROUP BY         	tblreceiver.PONumber,tblreceiver.ProductID,tblreceiver.DateNeeded)
 
SELECT distinct      tblpurchaseorder.PoNumber as 'PONumber',
                     CONVERT(VARCHAR,tblPurchaseOrder.PODatePlaced, 107) as 'PODatePlaced',
                     CONVERT(VARCHAR,tblPurchaseOrder.PODatedNeeded,107) as 'PODateNeeded',
                      tblVendor.Name,
					  tblPurchaseOrderLine.ProductID as 'Product ID',
					cteReceiver.qOrdered as 'Quantity Ordered'
FROM             	tblPurchaseOrder
INNER JOIN       	tblVendor
ON                   tblVendor.VendorID = tblPurchaseOrder.VendorID
INNER JOIN       	tblPurchaseOrderLine
ON                   tblPurchaseOrder.PoNumber = tblPurchaseOrderLine.PONumber
INNER JOIN       	cteReceiver
ON                   cteReceiver.PONumber = tblPurchaseOrder.PoNumber
WHERE            	cteReceiver.PoNumber NOT IN (SELECT distinct PoNumber FROM tblPurchaseOrderLine
          	                                 	WHERE tblPurchaseOrderLine.QtyOrdered - cteReceiver.qreceived = 0);

/* Question 5 */

Create View vproduct as 
Select Distinct			tblProduct.productID 'Product ID',
                        description 'Product Description',
                        EOQ 'Product Economic Order Quantity',
                        ISNULL(cast(tblPurchaseHistory.DatePurchased as varchar(100)),
                        'Not In Purchase History') 'Most Recent Purchase Date',
						ISNULL(cast(tblPurchaseHistory.Qty as varchar(10)), '--') 'Quantity Purchased',
						ISNULL(tblPurchaseHistory.Price, 0) 'Purchase Price'
from                    tblProduct
LEFT JOIN               tblPurchaseHistory
on                      tblProduct.ProductID = tblPurchaseHistory.ProductID
WHERE                   tblPurchaseHistory.DatePurchased IS NULL
OR                      tblPurchaseHistory.Price = (SELECT max (tblPurchaseHistory.Price)
from                    tblPurchaseHistory   where tblProduct.ProductID = tblPurchaseHistory.ProductID);

SELECT					tblpurchaseorderline.productid 'ProductID', 
						tblproduct.description 'Product Description',
						vproduct.[Purchase Price] 'Recent History Price',
						tblpurchaseorderline.price 'Current Price' , 
						tblpurchaseorder.ponumber ,
						tblVendor.name 'Vendor Name' 
from					tblPurchaseOrderLine 
Inner join				vproduct
On						tblPurchaseOrderline.productID = vproduct.[Product ID]
inner join				tblproduct
on						tblPurchaseorderline.ProductID = tblproduct.ProductID 
inner join				tblpurchaseorder
on						tblpurchaseorderline.ponumber = tblpurchaseorder.ponumber
inner join				tblVendor
on						tblpurchaseorder.vendorid = tblVendor.vendorid
Where					price > 1.2 * [Purchase Price];

/* Question 6 NOT DONE, HAVE TO GO TO CLASS*/

WITH cteQ AS 
(SELECT			CASE
                           WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) = 0
                           THEN 'Complete'
                           WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) < 0
                           THEN 'Over Shipment'
                           WHEN ISNULL(SUM(tblReceiver.QtyReceived),0.00) = 0.00
                           THEN 'Not Received'
                           WHEN (tblPurchaseOrderLine.QtyOrdered - ISNULL(SUM(tblReceiver.QtyReceived),0.00)) > 0
                           THEN 'Partial Shipment'
                           END ReceivingStatus,
				tblReceiver.PONumber,
				tblReceiver.ProductID,
				tblReceiver.DateNeeded
FROM			tblReceiver
INNER JOIN		tblPurchaseOrderLine
ON				tblPurchaseOrderLine.PONumber = tblReceiver.PONumber
GROUP BY		tblreceiver.PONumber,tblReceiver.ProductID,tblReceiver.DateNeeded, tblPurchaseOrderLine.QtyOrdered),

cteW AS
(SELECT			COUNT(ReceivingStatus) CountR,
				PONumber,
				ProductID,
				DateNeeded
FROM			cteQ
GROUP BY		PONumber,ProductID,DateNeeded)

SELECT			tblVendor.VendorID,
				tblVendor.Name,
				tblVendor.Email,
				MAX(cteW.CountR) as 'Count of Over-Shipped Purchase Order Lines'
FROM			tblVendor
INNER JOIN		tblPurchaseOrder
ON				tblVendor.VendorID = tblPurchaseOrder.VendorID
INNER JOIN		tblPurchaseOrderLine
ON				tblpurchaseorderline.PONumber = tblPurchaseOrder.PoNumber
INNER JOIN		cteQ
ON				tblPurchaseOrder.PoNumber = cteQ.PONumber
INNER JOIN		cteW
ON				tblPurchaseOrderLine.PONumber = cteW.PONumber
WHERE			tblVendor.Name = 'PolySort Manufacturing'
GROUP BY		tblVendor.VendorID, Name, Email;

/* question7 NOT DONE*/
select tblEmployee.EmpID 'EmployeeID',
   	 (tblEmployee.EmpLastName + ', ' + SUBSTRING(tblEmployee.EmpFirstName,1,1)) 'EmployeeName',
   	 tblEmployee.EmpEmail 'EmpEmail',
   	 case
   			 when (tblEmployee.EmpMgrID) IS NULL then 'No Manager'
   				 else (manager.EmpMgrID)
   			 end 'ManagerEmpID',
   	 case
   			 when (tblEmployee.EmpMgrID) IS NULL then 'No Manager'
   				 else (manager.EmpLastName + ', ' + SUBSTRING(manager.EmpFirstName,1,1))
   			 end 'ManagerName',
   	 case
   			 when (tblEmployee.EmpMgrID) IS NULL then 'No Manager'
   				 else (manager.EmpEmail)
   			 end 'Manager Email'

from tblEmployee

LEFT OUTER JOIN   	  tblEmployee manager
on   				  tblEmployee.EmpMgrID = manager.EmpID;
