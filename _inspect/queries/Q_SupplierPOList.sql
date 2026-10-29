SELECT P.SupplierPOID, P.PONumber, P.PODate, P.ExpectedDeliveryDate, P.SupplierID, S.SupplierCode, S.CompanyName, S.ContactPerson, S.Email, S.Phone, P.POType, P.POStatus, P.CurrencyID, P.PaymentTerms, P.ShipVia, P.SentOn, Nz(T.LineCount,0) AS LineCount, Nz(T.GrossAmount,0) AS GrossAmount, Nz(T.DiscountTotal,0) AS DiscountTotal, Nz(T.CalculatedSubtotal,0) AS CalculatedSubtotal, Nz(T.CalculatedTax,0) AS CalculatedTax, Nz(P.FreightAmount,0) AS FreightAmount, Nz(T.CalculatedSubtotal,0)
        + Nz(T.CalculatedTax,0)
        + Nz(P.FreightAmount,0) AS CalculatedTotal, P.TotalAmount, P.CreatedOn, P.UpdatedOn
FROM (T_SupplierPOs AS P INNER JOIN T_Suppliers AS S ON P.SupplierID = S.SupplierID) LEFT JOIN Q_SupplierPOLineTotals AS T ON P.SupplierPOID = T.SupplierPOID
ORDER BY P.PODate DESC , P.SupplierPOID DESC;

