SELECT W.WarehouseID, W.WarehouseCode, W.WarehouseName, W.Address, W.ContactPerson, Count(L.LocationID) AS LocationCount
FROM T_Warehouses AS W LEFT JOIN T_Locations AS L ON W.WarehouseID = L.WarehouseID
GROUP BY W.WarehouseID, W.WarehouseCode, W.WarehouseName, W.Address, W.ContactPerson
ORDER BY W.WarehouseName, W.WarehouseCode;

