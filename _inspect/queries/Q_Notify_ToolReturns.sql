SELECT IssueMovementID, IssueDate, ExpectedReturnDate, DaysUntilDue, AlertStatus, ItemID, InternalSKU, ItemName, CategoryName, ItemInstanceID, SerialNumber, ProjectID, AssemblyJobID, CustomerID, ReferenceDocument, IssueReason, ReturnStatus, AlertMessage
FROM Q_Notify_ToolReturns_Base
WHERE ReturnStatus = 'NOT RETURNED'
    AND AlertStatus In ('DUE SOON','OVERDUE')
ORDER BY IIf(AlertStatus='OVERDUE',1,2), ExpectedReturnDate, ItemName;

