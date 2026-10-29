SELECT INV.InvoiceID, INV.InvoiceNo, INV.InvoiceDate, INV.InvoiceType, INV.InvoiceStatus, INV.CustomerID, C.CustomerName, INV.CustomerPOID, P.PONumber, INV.SalesOrderID, SO.SalesOrderNo, INV.PaymentTerms, INV.BillToName, INV.SubtotalAmount, INV.TaxAmount, INV.GrandTotalAmount, Nz(L.LineCount,0) AS LineCount, Nz(L.TotalInvoiceQty,0) AS TotalInvoiceQty, Nz(L.LinesSubtotal,0) AS LinesSubtotal, INV.CreatedOn, INV.UpdatedOn
FROM (((T_Invoices AS INV LEFT JOIN T_Customers AS C ON INV.CustomerID = C.CustomerID) LEFT JOIN T_CustomerPOs AS P ON INV.CustomerPOID = P.CustomerPOID) LEFT JOIN T_SalesOrders AS SO ON INV.SalesOrderID = SO.SalesOrderID) LEFT JOIN (SELECT
                InvoiceID,
                Count(*) AS LineCount,
                Sum(Nz([Quantity],0)) AS TotalInvoiceQty,
                Sum(Nz([LineSubtotal],0)) AS LinesSubtotal
            FROM T_InvoiceLines
            GROUP BY InvoiceID
        )  AS L ON INV.InvoiceID = L.InvoiceID
ORDER BY INV.InvoiceDate DESC , INV.InvoiceNo;

