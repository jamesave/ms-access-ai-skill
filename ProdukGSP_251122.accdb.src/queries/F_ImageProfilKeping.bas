dbMemo "SQL" ="SELECT T_ProfilKeping.KepingAutoID, T_ProfilKeping.KepingID, T_ProfilKeping.Desk"
    "ripsiProfilKeping, T_ProfilKeping.ImageCode, T_ProfilKeping.NoteGambar, GetPKIma"
    "gePath ([KepingAutoID]) AS ImagePath\015\012FROM T_ProfilKeping;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ImagePath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.DeskripsiProfilKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.NoteGambar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.ImageCode"
        dbLong "AggregateType" ="-1"
    End
End
