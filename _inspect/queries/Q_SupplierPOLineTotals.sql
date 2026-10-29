SELECT L.SupplierPOID, Count(L.SupplierPOLineID) AS LineCount, Sum(
        Nz(L.OrderedQuantity,0) * Nz(L.UnitPrice,0)
    ) AS GrossAmount, Sum(
        Nz(L.DiscountAmount,0)
    ) AS DiscountTotal, Sum(
        (Nz(L.OrderedQuantity,0) * Nz(L.UnitPrice,0))
        - Nz(L.DiscountAmount,0)
    ) AS CalculatedSubtotal, Sum(
        (
            (Nz(L.OrderedQuantity,0) * Nz(L.UnitPrice,0))
            - Nz(L.DiscountAmount,0)
        )
        * (Nz(L.TaxPercent,0) / 100)
    ) AS CalculatedTax
FROM T_SupplierPOLines AS L
GROUP BY L.SupplierPOID;

