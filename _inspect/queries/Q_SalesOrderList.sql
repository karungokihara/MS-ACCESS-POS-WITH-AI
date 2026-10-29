SELECT SO.SalesOrderID, SO.SalesOrderNo, SO.CustomerID, C.CustomerName, SO.CustomerPOID, P.PONumber, SO.OrderDate, SO.RequiredDate, SO.ShipDate, SO.SalesOrderStatus, SO.PaymentTerms, SO.ShipVia, SO.ShipToName, Nz(DCount("*","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) AS LineCount, Nz(DSum("Nz([QuantityOrdered],0)","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) AS TotalOrderedQty, Nz(DSum("Nz([QuantityShipped],0)","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) AS TotalShippedQty, Nz(DSum("Nz([QuantityOrdered],0)-Nz([QuantityShipped],0)","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) AS RemainingQty, Nz(DSum("Nz([LineSubtotal],0)","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) AS OrderSubtotal, IIf(
        Nz(DSum("Nz([QuantityOrdered],0)-Nz([QuantityShipped],0)","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) <= 0,
        "Fully Shipped",
        IIf(
            Nz(DSum("Nz([QuantityShipped],0)","T_SalesOrderLines","SalesOrderID=" & [SO].[SalesOrderID]),0) > 0,
            "Partially Shipped",
            "Not Shipped"
        )
    ) AS FulfillmentStatus, DMax("InvoiceID","T_Invoices","SalesOrderID=" & [SO].[SalesOrderID]) AS InvoiceID, DLookup(
        "InvoiceNo",
        "T_Invoices",
        "InvoiceID=" & Nz(DMax("InvoiceID","T_Invoices","SalesOrderID=" & [SO].[SalesOrderID]),0)
    ) AS InvoiceNo, DLookup(
        "InvoiceStatus",
        "T_Invoices",
        "InvoiceID=" & Nz(DMax("InvoiceID","T_Invoices","SalesOrderID=" & [SO].[SalesOrderID]),0)
    ) AS InvoiceStatus, Nz(DSum("Nz([GrandTotalAmount],0)","T_Invoices","SalesOrderID=" & [SO].[SalesOrderID]),0) AS InvoiceValue, SO.CreatedOn, SO.UpdatedOn
FROM (T_SalesOrders AS SO LEFT JOIN T_Customers AS C ON SO.CustomerID = C.CustomerID) LEFT JOIN T_CustomerPOs AS P ON SO.CustomerPOID = P.CustomerPOID
ORDER BY SO.OrderDate DESC , SO.SalesOrderNo;

