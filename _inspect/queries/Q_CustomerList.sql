SELECT C.CustomerID, C.CustomerName, C.DefaultPaymentTerms, C.DefaultShipVia, Nz(P.TotalPOs,0) AS TotalPOs, Nz(P.OpenPOs,0) AS OpenPOs, Nz(SO.TotalSalesOrders,0) AS TotalSalesOrders, Nz(SO.OpenSalesOrders,0) AS OpenSalesOrders, Nz(INV.TotalInvoices,0) AS TotalInvoices, Nz(INV.OpenInvoices,0) AS OpenInvoices, Nz(INV.InvoiceValue,0) AS InvoiceValue
FROM ((T_Customers AS C LEFT JOIN (SELECT
                CustomerID,
                Count(*) AS TotalPOs,
                Sum(IIf(Nz([POStatus],'') Not In ('Closed','Cancelled'),1,0)) AS OpenPOs
            FROM T_CustomerPOs
            GROUP BY CustomerID
        )  AS P ON C.CustomerID = P.CustomerID) LEFT JOIN (SELECT
                CustomerID,
                Count(*) AS TotalSalesOrders,
                Sum(IIf(Nz([SalesOrderStatus],'') Not In ('Closed','Cancelled','Shipped'),1,0)) AS OpenSalesOrders
            FROM T_SalesOrders
            GROUP BY CustomerID
        )  AS SO ON C.CustomerID = SO.CustomerID) LEFT JOIN (SELECT
                CustomerID,
                Count(*) AS TotalInvoices,
                Sum(IIf(Nz([InvoiceStatus],'') Not In ('Paid','Cancelled'),1,0)) AS OpenInvoices,
                Sum(Nz([GrandTotalAmount],0)) AS InvoiceValue
            FROM T_Invoices
            GROUP BY CustomerID
        )  AS INV ON C.CustomerID = INV.CustomerID
ORDER BY C.CustomerName;

