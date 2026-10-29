SELECT RI.SupplierPOLineID, Sum(Nz(RI.QuantityReceived,0)) AS QuantityReceivedToDate
FROM T_GoodsReceipts AS R INNER JOIN T_GoodsReceiptItems AS RI ON R.GoodsReceiptID = RI.GoodsReceiptID
WHERE R.ReceiptStatus = "Posted"
    AND RI.SupplierPOLineID Is Not Null
GROUP BY RI.SupplierPOLineID;

