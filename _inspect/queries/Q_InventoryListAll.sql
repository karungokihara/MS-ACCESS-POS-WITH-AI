SELECT
    'Bulk' AS StockType,
    B.StockBalanceID AS SourceRecordID,
    I.ItemID AS ItemID,
    Null AS ItemInstanceID,
    I.CategoryID AS CategoryID,
    L.WarehouseID AS WarehouseID,
    B.CurrentLocationID AS LocationID,
    I.InternalSKU AS InternalSKU,
    I.ItemName AS ItemName,
    I.Description AS Description,
    C.CategoryName AS CategoryName,
    Null AS SerialNumber,
    Nz(B.QuantityOnHand,0) AS QuantityOnHand,
    Nz(B.QuantityReserved,0) AS QuantityReserved,
    Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) AS AvailableQuantity,
    IIf(
        Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= 0,
        'Out of Stock',
        IIf(
            Nz(I.ReorderLevel,0) > 0
            AND Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= Nz(I.ReorderLevel,0),
            'Low Stock',
            'In Stock'
        )
    ) AS StockStatus,
    IIf(
        Nz(I.ReorderLevel,0) > 0
        AND Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= 0,
        'Restock Required',
        IIf(
            Nz(I.ReorderLevel,0) > 0
            AND Nz(B.QuantityOnHand,0) - Nz(B.QuantityReserved,0) <= Nz(I.ReorderLevel,0),
            'Low Stock',
            'OK'
        )
    ) AS StockAlert,
    W.WarehouseName AS WarehouseName,
    L.LocationCode AS LocationCode,
    L.LocationName AS LocationName,
    L.BinNumber AS BinNumber,
    L.ShelfNumber AS ShelfNumber,
    U.UnitCode AS UnitCode,
    Nz(I.ReorderLevel,0) AS ReorderLevel,
    B.LastUpdated AS LastUpdated
FROM
    ((((T_StockBalances AS B
    LEFT JOIN T_Items AS I
        ON B.ItemID = I.ItemID)
    LEFT JOIN T_Categories AS C
        ON I.CategoryID = C.CategoryID)
    LEFT JOIN T_Units AS U
        ON I.BaseUnitID = U.UnitID)
    LEFT JOIN T_Locations AS L
        ON B.CurrentLocationID = L.LocationID)
    LEFT JOIN T_Warehouses AS W
        ON L.WarehouseID = W.WarehouseID
WHERE
    Nz(C.RequiresSerial,False) = False

UNION ALL SELECT
    'Serialized' AS StockType,
    S.ItemInstanceID AS SourceRecordID,
    I.ItemID AS ItemID,
    S.ItemInstanceID AS ItemInstanceID,
    I.CategoryID AS CategoryID,
    L.WarehouseID AS WarehouseID,
    S.CurrentLocationID AS LocationID,
    I.InternalSKU AS InternalSKU,
    I.ItemName AS ItemName,
    I.Description AS Description,
    C.CategoryName AS CategoryName,
    S.SerialNumber AS SerialNumber,
    1 AS QuantityOnHand,
    0 AS QuantityReserved,
    IIf(Nz(S.CurrentStatus,'')='In Stock',1,0) AS AvailableQuantity,
    Nz(S.CurrentStatus,'Unknown') AS StockStatus,
    IIf(
        Nz(S.CurrentStatus,'')='In Stock',
        'OK',
        IIf(
            Nz(S.CurrentStatus,'')='Issued',
            'Issued',
            IIf(
                Nz(S.CurrentStatus,'')='Out of Stock',
                'Not Available',
                Nz(S.CurrentStatus,'Unknown')
            )
        )
    ) AS StockAlert,
    W.WarehouseName AS WarehouseName,
    L.LocationCode AS LocationCode,
    L.LocationName AS LocationName,
    L.BinNumber AS BinNumber,
    L.ShelfNumber AS ShelfNumber,
    U.UnitCode AS UnitCode,
    Nz(I.ReorderLevel,0) AS ReorderLevel,
    S.AcquisitionDate AS LastUpdated
FROM
    ((((T_ItemInstances AS S
    LEFT JOIN T_Items AS I
        ON S.ItemID = I.ItemID)
    LEFT JOIN T_Categories AS C
        ON I.CategoryID = C.CategoryID)
    LEFT JOIN T_Units AS U
        ON I.BaseUnitID = U.UnitID)
    LEFT JOIN T_Locations AS L
        ON S.CurrentLocationID = L.LocationID)
    LEFT JOIN T_Warehouses AS W
        ON L.WarehouseID = W.WarehouseID
WHERE
    Nz(C.RequiresSerial,False) = True
ORDER BY ItemName, StockType, SerialNumber;

