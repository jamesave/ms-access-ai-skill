SELECT
    tOrderItems.*
FROM
    tOrderItems
WHERE
    tOrderItems.OrderID=[Forms]![fOrderDetail]![ID]
ORDER BY
    tOrderItems.ID;
