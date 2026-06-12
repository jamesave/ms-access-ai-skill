Operation =1
Option =0
Where ="(((T_UkuranPolaLubang.PolaLubang)=\"12-A\"))"
Begin InputTables
    Name ="T_ProfilKeping"
    Name ="T_UkuranPolaLubang"
End
Begin OutputColumns
    Expression ="T_ProfilKeping.*"
    Expression ="T_UkuranPolaLubang.PolaLubang"
End
Begin Joins
    LeftTable ="T_UkuranPolaLubang"
    RightTable ="T_ProfilKeping"
    Expression ="T_UkuranPolaLubang.PolaLubangID = T_ProfilKeping.PolaLubangID"
    Flag =3
End
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
        dbText "Name" ="T_ProfilKeping.PolaLubangID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.KepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.PolaLubangID_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_UkuranPolaLubang.PolaLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.DeskripsiProfilKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.BlankoKepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.UkuranRivet_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.JenisRivet_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.LubangFlag_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.UkuranLubang_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.RivetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Indikator-Lis"
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
        dbText "Name" ="T_ProfilKeping.Catatan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Lebar_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Panjang_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Tebal_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.DiameterDrum_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.PolaLubangID_deprecated2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.JumlahLubang_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.UkuranRivetID_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.TebalTipisFlag_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Tebal/Tipis_deprecated"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =185
    Top =236
    Right =2384
    Bottom =1103
    Left =-1
    Top =-1
    Right =2164
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =441
        Top =0
        Name ="T_ProfilKeping"
        Name =""
    End
    Begin
        Left =387
        Top =43
        Right =603
        Bottom =259
        Top =0
        Name ="T_UkuranPolaLubang"
        Name =""
    End
End
