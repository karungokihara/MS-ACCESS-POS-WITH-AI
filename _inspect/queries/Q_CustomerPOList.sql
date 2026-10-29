SELECT P.CustomerPOID, P.CustomerID, C.CustomerName, P.PONumber, P.PODate, P.PORevisionDate, P.POType, P.POStatus, P.GoodsConfirmed, P.GoodsConfirmedOn, P.PaymentTerms, P.ShipVia, P.ShipToName, P.ShipToContactName, P.ShipToPhone, P.ShipToEmail, P.AuthorizedBy, P.AuthorizedByEmail, Nz(L.LineCount,0) AS LineCount, Nz(L.TotalQty,0) AS TotalQty, Nz(L.LinesSubtotal,0) AS LinesSubtotal, P.SubtotalAmount, P.MiscChargeAmount, P.TaxAmount, P.TotalAmount, SO.SalesOrderID, SO.SalesOrderNo, INV.InvoiceID, INV.InvoiceNo, INV.InvoiceStatus, P.CreatedOn, P.UpdatedOn
FROM (((T_CustomerPOs AS P LEFT JOIN T_Customers AS C ON P.CustomerID = C.CustomerID) LEFT JOIN (SELECT
                CustomerPOID,
                Count(*) AS LineCount,
                Sum(Nz([QuantityOrdered],0)) AS TotalQty,
                Sum(Nz([LineSubtotal],0)) AS LinesSubtotal
            FROM T_CustomerPOLines
            GROUP BY CustomerPOID
        )  AS L ON P.CustomerPOID = L.CustomerPOID) LEFT JOIN T_SalesOrders AS SO ON P.CustomerPOID = SO.CustomerPOID) LEFT JOIN T_Invoices AS INV ON P.CustomerPOID = INV.CustomerPOID
ORDER BY P.PODate DESC , P.PONumber;

