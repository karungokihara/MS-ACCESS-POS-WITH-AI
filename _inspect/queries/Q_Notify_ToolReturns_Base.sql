SELECT M.StockMovementID AS IssueMovementID, M.MovementDate AS IssueDate, M.ExpectedReturnDate, DateDiff('d', Date(), DateValue(M.ExpectedReturnDate)) AS DaysUntilDue, IIf(
        DateValue(M.ExpectedReturnDate) < Date(),
        'OVERDUE',
        IIf(
            DateValue(M.ExpectedReturnDate) <= Date() + 2,
            'DUE SOON',
            'NOT DUE'
        )
    ) AS AlertStatus, I.ItemID, I.InternalSKU, I.ItemName, C.CategoryName, II.ItemInstanceID, II.SerialNumber, M.ProjectID, M.AssemblyJobID, M.CustomerID, M.ReferenceDocument, M.Reason AS IssueReason, M.MovementNotes, M.Notes, DMax(
        'MovementDate',
        'T_StockMovements',
        'ItemInstanceID=' & [M].[ItemInstanceID] &
        ' AND ToLocationID Is Not Null' &
        ' AND MovementDate >= #' & Format([M].[MovementDate],'yyyy\/mm\/dd hh\:nn\:ss') & '#' &
        ' AND (Reason=''Return to Stock'' OR Reason=''Returned to Stock'' OR Reason=''Tool Return'' OR Reason=''Project Tool Return'' OR Reason=''Return From Project'')'
    ) AS ReturnedOn, IIf(
        IsNull(
            DMax(
                'MovementDate',
                'T_StockMovements',
                'ItemInstanceID=' & [M].[ItemInstanceID] &
                ' AND ToLocationID Is Not Null' &
                ' AND MovementDate >= #' & Format([M].[MovementDate],'yyyy\/mm\/dd hh\:nn\:ss') & '#' &
                ' AND (Reason=''Return to Stock'' OR Reason=''Returned to Stock'' OR Reason=''Tool Return'' OR Reason=''Project Tool Return'' OR Reason=''Return From Project'')'
            )
        ),
        'NOT RETURNED',
        'RETURNED'
    ) AS ReturnStatus, IIf(
        DateValue(M.ExpectedReturnDate) < Date(),
        'Tool overdue: ' & Nz(I.ItemName,'') & ' / Serial: ' & Nz(II.SerialNumber,''),
        IIf(
            DateValue(M.ExpectedReturnDate) <= Date() + 2,
            'Tool due soon: ' & Nz(I.ItemName,'') & ' / Serial: ' & Nz(II.SerialNumber,''),
            'Tool not due yet: ' & Nz(I.ItemName,'') & ' / Serial: ' & Nz(II.SerialNumber,'')
        )
    ) AS AlertMessage
FROM ((T_StockMovements AS M LEFT JOIN T_Items AS I ON M.ItemID = I.ItemID) LEFT JOIN T_Categories AS C ON I.CategoryID = C.CategoryID) LEFT JOIN T_ItemInstances AS II ON M.ItemInstanceID = II.ItemInstanceID
WHERE M.ExpectedReturnDate Is Not Null
    AND M.ItemInstanceID Is Not Null
ORDER BY M.ExpectedReturnDate, I.ItemName;

