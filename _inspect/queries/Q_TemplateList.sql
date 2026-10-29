SELECT T.AssemblyTemplateID, T.TemplateCode, T.TemplateName, T.FinishedItemID, I.InternalSKU AS FinishedItemSKU, I.ItemName AS FinishedItemName, T.CustomerID, T.CustomerPartNumber, T.OutputQuantity, T.RevisionNo, T.Status, T.EstimatedLaborHours, T.EstimatedMaterialCost, T.EstimatedTotalCost, DCount("*","T_AssemblyTemplateMaterials","AssemblyTemplateID=" & [T].[AssemblyTemplateID]) AS MaterialLineCount, DCount("*","T_AssemblyTemplateTools","AssemblyTemplateID=" & [T].[AssemblyTemplateID]) AS ToolLineCount, T.CreatedOn, T.UpdatedOn, T.Notes
FROM T_AssemblyTemplates AS T LEFT JOIN T_Items AS I ON T.FinishedItemID = I.ItemID
ORDER BY T.TemplateName, T.TemplateCode;

