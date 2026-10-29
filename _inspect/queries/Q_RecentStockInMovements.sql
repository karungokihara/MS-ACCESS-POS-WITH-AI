SELECT TOP 100 I.ItemID, SM.StockMovementID, SM.MovementDate, I.InternalSKU, I.ItemName, SM.Quantity, W.WarehouseName, L.LocationName, SM.ReferenceDocument, SM.Reason, SM.UnitCost
FROM T_StockMovementTypes AS MT INNER JOIN (T_Items AS I INNER JOIN ((T_StockMovements AS SM LEFT JOIN T_Locations AS L ON SM.ToLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID) ON I.ItemID = SM.ItemID) ON MT.MovementTypeID = SM.MovementTypeID
WHERE (((MT.MovementTypeCode)="STOCK_IN"))
ORDER BY SM.MovementDate DESC , SM.StockMovementID DESC;

