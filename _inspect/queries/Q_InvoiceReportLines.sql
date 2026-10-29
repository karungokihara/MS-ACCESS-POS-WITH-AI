SELECT L.InvoiceLineID, L.InvoiceID, L.ItemID, L.ItemCode, L.LineDescription, L.Quantity, L.UnitID, U.UnitCode, L.UnitPrice, L.Taxable, L.LineSubtotal, Nz(L.ItemCode,"") AS ReportItem, Nz(L.LineDescription,"") AS ReportDescription, Nz(U.UnitCode,"") AS ReportUnit
FROM T_InvoiceLines AS L LEFT JOIN T_Units AS U ON L.UnitID = U.UnitID
ORDER BY L.InvoiceLineID;

