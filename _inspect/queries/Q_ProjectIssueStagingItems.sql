SELECT S.StagingID, S.ProjectID, S.IssueType, S.ItemID, I.ItemName, U.UnitName, S.RequiredQty, S.AvailableQty, S.IssueQty, S.Selected, S.Notes
FROM (T_ProjectIssueStaging AS S INNER JOIN T_Items AS I ON S.ItemID = I.ItemID) LEFT JOIN T_Units AS U ON I.BaseUnitID = U.UnitID;

