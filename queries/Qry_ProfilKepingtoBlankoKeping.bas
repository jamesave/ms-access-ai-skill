dbMemo "SQL" ="SELECT T_ProfilKeping.BlankoKepingAutoID, T_ProfilKeping.[DeskripsiProfilKeping]"
    ", T_BlankoKeping.DeskripsiBlankoKeping, T_BlankoKeping.Panjang, T_BlankoKeping.L"
    "ebar, T_BlankoKeping.Tebal, T_BlankoKeping.DiameterDrum, T_BlankoKeping.Diameter"
    "BrakeShoe, T_BlankoKeping.CatatanBlankoKeping\015\012FROM T_BlankoKeping RIGHT J"
    "OIN T_ProfilKeping ON T_BlankoKeping.BlankoKepingAutoID = T_ProfilKeping.BlankoK"
    "epingAutoID;\015\012"
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
        dbText "Name" ="T_BlankoKeping.DiameterBrakeShoe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DiameterDrum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Tebal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Lebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Panjang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DeskripsiBlankoKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.CatatanBlankoKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.BlankoKepingAutoID"
        dbLong "AggregateType" ="-1"
    End
End
