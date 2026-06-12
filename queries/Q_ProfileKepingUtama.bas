dbMemo "SQL" ="SELECT T_ProfilKeping.KepingAutoID, T_ProfilKeping.KepingID, T_ProfilKeping.Blan"
    "koKepingAutoID, T_ProfilKeping.[DeskripsiProfilKeping], T_ProfilKeping.PolaLuban"
    "gID, T_ProfilKeping.RivetID, T_ProfilKeping.[Indikator-Lis], T_ProfilKeping.Coak"
    ", T_ProfilKeping.Koak, T_ProfilKeping.Catatan, T_BlankoKeping.DeskripsiBlankoKep"
    "ing, T_BlankoKeping.Panjang, T_BlankoKeping.Lebar, T_BlankoKeping.Tebal_TipisFla"
    "g, T_BlankoKeping.Tebal, T_BlankoKeping.DiameterDrum, T_BlankoKeping.Tebal2, T_B"
    "lankoKeping.DiameterBrakeShoe, T_UkuranRivet.KodeRivet, T_UkuranRivet.[Ukuran], "
    "T_UkuranPolaLubang.PolaLubang, T_UkuranPolaLubang.PolaLubangID, T_ListPolaLubang"
    ".LokasiImage, T_ListPolaLubang.JumlahLubang\015\012FROM T_UkuranRivet RIGHT JOIN"
    " (T_ListPolaLubang RIGHT JOIN (T_BlankoKeping RIGHT JOIN (T_ProfilKeping LEFT JO"
    "IN T_UkuranPolaLubang ON T_ProfilKeping.PolaLubangID = T_UkuranPolaLubang.PolaLu"
    "bangID) ON T_BlankoKeping.BlankoKepingID = T_ProfilKeping.BlankoKepingID) ON T_L"
    "istPolaLubang.PolaLubang = T_UkuranPolaLubang.PolaLubang) ON T_UkuranRivet.Rivet"
    "ID = T_ProfilKeping.RivetID;\015\012"
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
        dbText "Name" ="T_BlankoKeping.Lebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Tebal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.PolaLubangID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Panjang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Coak"
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
        dbText "Name" ="T_ProfilKeping.RivetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Catatan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DeskripsiBlankoKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Tebal_TipisFlag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.Tebal2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BlankoKeping.DiameterBrakeShoe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranRivet.KodeRivet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranPolaLubang.PolaLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranPolaLubang.PolaLubangID"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="T_ProfilKeping.BlankoKepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.[DeskripsiProfilKeping]"
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
End
