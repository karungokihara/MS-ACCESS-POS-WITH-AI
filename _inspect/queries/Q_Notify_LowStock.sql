SELECT B.StockBalanceID, I.ItemID, I.InternalSKU, I.ItemName, C.CategoryName, W.WarehouseName, L.LocationCode, L.LocationName, L.BinNumber, L.ShelfNumber, Nz(B.QuantityOnHand,0) AS QuantityOnHand, Nz(B.QuantityReserved,0) AS QuantityReserved, Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) AS AvailableQty, Nz(I.ReorderLevel,0) AS ReorderPoint, IIf(
        Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= 0,
        'RESTOCK REQUIRED',
        'LOW STOCK'
    ) AS AlertStatus, IIf(
        Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= 0,
        'Restock required for ' & Nz(I.ItemName,''),
        'Low stock for ' & Nz(I.ItemName,'')
    ) AS AlertMessage, B.LastUpdated
FROM (((T_StockBalances AS B INNER JOIN T_Items AS I ON B.ItemID = I.ItemID) LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Locations AS L ON B.CurrentLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID
WHERE Nz(C.RequiresSerial,False) = False
    AND Nz(I.ReorderLevel,0) > 0
    AND Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= Nz(I.ReorderLevel,0)
ORDER BY IIf(Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= 0, 1, 2), I.ItemName;

