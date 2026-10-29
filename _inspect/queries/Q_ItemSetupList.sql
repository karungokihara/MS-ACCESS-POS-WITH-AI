SELECT I.ItemID, I.InternalSKU, I.ItemName, I.Description, I.CategoryID, C.CategoryName, IIf(Nz(C.RequiresSerial,False)=True,'Serialized','Bulk') AS StockType, I.BaseUnitID, U.UnitCode, I.ReorderLevel, I.DefaultLocationID, L.LocationCode AS DefaultLocationCode, L.LocationName AS DefaultLocationName, W.WarehouseName AS DefaultWarehouseName, I.CreatedOn, I.UpdatedOn
FROM (((T_Items AS I LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID) LEFT JOIN T_Locations AS L ON I.DefaultLocationID = L.LocationID) LEFT JOIN T_Warehouses AS W ON L.WarehouseID = W.WarehouseID
ORDER BY I.ItemName, I.InternalSKU;

