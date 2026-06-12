dbMemo "SQL" ="SELECT T_BlankoKeping.BlankoKepingID, T_BlankoKeping.DeskripsiBlankoKeping, T_Pr"
    "ofilKeping.KepingID, T_ProfilKeping.[DeskripsiProfilKeping], T_ProfilKeping.[Des"
    "kripsiProfilKeping]\015\012FROM T_BlankoKeping RIGHT JOIN T_ProfilKeping ON T_Bl"
    "ankoKeping.BlankoKepingID = T_ProfilKeping.BlankoKepingID;\015\012"
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
        dbText "Name" ="T_ProfilKeping.DeskripsiKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DeskripsiBlankoKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.BlankoKepingID"
        dbInteger "ColumnWidth" ="1190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
End
