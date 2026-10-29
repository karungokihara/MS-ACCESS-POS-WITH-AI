SELECT L.SupplierPOLineID, L.SupplierPOID, P.PONumber, P.SupplierID, L.LineNumber, L.ItemID, I.InternalSKU, I.ItemName, I.ItemClass, I.RequiresSerial, L.SupplierPartNumber, L.ItemDescription, L.OrderedQuantity, Nz(R.QuantityReceivedToDate,0) AS QuantityReceivedToDate, L.OrderedQuantity
        - Nz(R.QuantityReceivedToDate,0) AS RemainingQuantity, L.UnitID, U.UnitCode, U.UnitName, L.UnitPrice, L.DiscountAmount, L.TaxPercent, L.RequiredDate, L.LineStatus, L.Notes
FROM (((T_SupplierPOLines AS L INNER JOIN T_SupplierPOs AS P ON L.SupplierPOID = P.SupplierPOID) INNER JOIN T_Items AS I ON L.ItemID = I.ItemID) INNER JOIN T_Units AS U ON L.UnitID = U.UnitID) LEFT JOIN Q_SupplierPOLineReceived AS R ON L.SupplierPOLineID = R.SupplierPOLineID
WHERE Nz(L.LineStatus,"Open") <> "Cancelled"
    AND
    (
        L.OrderedQuantity
        - Nz(R.QuantityReceivedToDate,0)
    ) > 0
ORDER BY L.SupplierPOID, L.LineNumber;

