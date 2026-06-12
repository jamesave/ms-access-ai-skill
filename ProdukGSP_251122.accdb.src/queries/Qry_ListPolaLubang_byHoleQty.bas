dbMemo "SQL" ="SELECT DISTINCT JumlahLubang, PolaLubang, LokasiImage\015\012FROM T_ListPolaLuba"
    "ng\015\012ORDER BY JumlahLubang;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="LokasiImage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PolaLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="JumlahLubang"
        dbLong "AggregateType" ="-1"
    End
End
