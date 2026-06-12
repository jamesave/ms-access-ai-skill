dbMemo "SQL" ="SELECT T_ProfilKeping.KepingAutoID, T_ProfilKeping.KepingID, T_ProfilKeping.Blan"
    "koKepingAutoID, T_ProfilKeping.DeskripsiProfilKeping, T_ProfilKeping.PolaLubangI"
    "D, T_ProfilKeping.RivetID, T_ProfilKeping.[Indikator-Lis], T_ProfilKeping.Coak, "
    "T_ProfilKeping.Koak, T_ProfilKeping.UkuranLubangID, T_ProfilKeping.Catatan, T_Pr"
    "ofilKeping.LubangFlag, T_ProfilKeping.TebalTipisFlag, T_BlankoKeping.BlankoKepin"
    "gID, T_BlankoKeping.DeskripsiBlankoKeping, T_BlankoKeping.Panjang, T_BlankoKepin"
    "g.Lebar, T_BlankoKeping.Tebal, T_BlankoKeping.Tebal2, T_BlankoKeping.Panjang2, T"
    "_BlankoKeping.DiameterDrum, T_UkuranRivet.KodeRivet, T_UkuranRivet.Ukuran, T_Uku"
    "ranPolaLubang.PolaLubang, T_UkuranPolaLubang.DetailPolaLubang, T_ListPolaLubang."
    "LokasiImage, T_ListPolaLubang.JumlahLubang, T_BlankoKeping.DiameterBrakeShoe, T_"
    "ProfilKeping.ImageCode, T_ProfilKeping.NoteGambar, GetPKImagePath ([T_ProfilKepi"
    "ng.KepingAutoID]) AS PKImageCode, GetHPImagePath (T_UkuranPolaLubang.PolaLubang)"
    " AS HPImageCode, T_UkuranLubang.UkuranLubang\015\012FROM T_UkuranRivet RIGHT JOI"
    "N ((T_ListPolaLubang RIGHT JOIN T_UkuranPolaLubang ON T_ListPolaLubang.PolaLuban"
    "g = T_UkuranPolaLubang.PolaLubang) RIGHT JOIN (T_UkuranLubang RIGHT JOIN (T_Blan"
    "koKeping RIGHT JOIN T_ProfilKeping ON T_BlankoKeping.BlankoKepingAutoID = T_Prof"
    "ilKeping.BlankoKepingAutoID) ON T_UkuranLubang.UkuranLubangID = T_ProfilKeping.U"
    "kuranLubangID) ON T_UkuranPolaLubang.PolaLubangID = T_ProfilKeping.PolaLubangID)"
    " ON T_UkuranRivet.RivetID = T_ProfilKeping.RivetID;\015\012"
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
        dbText "Name" ="T_BlankoKeping.DiameterDrum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Coak"
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
        dbText "Name" ="T_ProfilKeping.RivetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Panjang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.PolaLubangID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DeskripsiBlankoKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.BlankoKepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.BlankoKepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Catatan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Koak"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Tebal2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranRivet.KodeRivet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranPolaLubang.PolaLubang"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="T_ListPolaLubang.LokasiImage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ListPolaLubang.JumlahLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DiameterBrakeShoe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.[Indikator-Lis]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranRivet.[Ukuran]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranRivet.Ukuran"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.[DeskripsiProfilKeping]"
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
    Begin
        dbText "Name" ="T_BlankoKeping.Panjang2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.DeskripsiProfilKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.LubangFlag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.TebalTipisFlag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HPImageCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PKImageCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranPolaLubang.DetailPolaLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.UkuranLubangID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranLubang.UkuranLubang"
        dbLong "AggregateType" ="-1"
    End
End
