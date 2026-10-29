SELECT T_Items.ItemID, T_Items.InternalSKU, T_Items.ItemName, T_Items.ItemClass, Nz(T_StockBalances.QuantityOnHand,0) AS CurrentStock
FROM T_Items LEFT JOIN T_StockBalances ON T_Items.ItemID = T_StockBalances.ItemID
WHERE Nz(T_Items.ItemClass,'') Not In ('Tool','Equipment')
ORDER BY T_Items.ItemName;

