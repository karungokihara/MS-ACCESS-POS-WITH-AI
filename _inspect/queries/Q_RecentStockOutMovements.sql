SELECT TOP 100 SM.StockMovementID, SM.MovementDate, I.InternalSKU, I.ItemName, II.SerialNumber, SM.Quantity, W.WarehouseName, L.LocationName, SM.ReferenceDocument, SM.Reason, SM.Notes, U.FullName AS PostedBy
FROM (((((T_StockMovements AS SM INNER JOIN T_Items AS I ON SM.ItemID = I.ItemID) INNER JOIN T_StockMovementTypes AS MT ON SM.MovementTypeID = MT.MovementTypeID) LEFT JOIN T_ItemInstances AS II ON SM.ItemInstanceID = II.ItemInstanceID) LEFT JOIN T_Locations AS L ON SM.FromLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID) LEFT JOIN T_Users AS U ON SM.PerformedByUserID = U.UserID
WHERE MT.MovementTypeCode = "STOCK_OUT"
ORDER BY SM.MovementDate DESC , SM.StockMovementID DESC;

