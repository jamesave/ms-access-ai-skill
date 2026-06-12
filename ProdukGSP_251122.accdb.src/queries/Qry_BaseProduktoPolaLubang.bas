Operation =1
Option =0
Begin InputTables
    Name ="T_ProfilKeping"
    Name ="T_PcsSetBOM"
End
Begin OutputColumns
    Expression ="T_PcsSetBOM.BaseProdukID"
    Expression ="T_PcsSetBOM.KepingID"
    Expression ="T_ProfilKeping.BlankoKepingAutoID"
    Expression ="T_ProfilKeping.DeskripsiProfilKeping"
    Expression ="T_ProfilKeping.RivetID"
    Expression ="T_ProfilKeping.[Indikator-Lis]"
    Expression ="T_ProfilKeping.Coak"
    Expression ="T_ProfilKeping.Koak"
    Expression ="T_ProfilKeping.Catatan"
End
Begin Joins
    LeftTable ="T_ProfilKeping"
    RightTable ="T_PcsSetBOM"
    Expression ="T_ProfilKeping.KepingID = T_PcsSetBOM.KepingID"
    Flag =1
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
        dbText "Name" ="T_ProfilKeping.BlankoKepingAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Koak"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.Coak"
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
        dbText "Name" ="T_ProfilKeping.[DeskripsiProfilKeping]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.[Indikator-Lis]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ProfilKeping.DeskripsiProfilKeping"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-16
    Top =0
    Right =1303
    Bottom =812
    Left =-1
    Top =-1
    Right =1284
    Bottom =429
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =268
        Top =13
        Right =568
        Bottom =506
        Top =0
        Name ="T_ProfilKeping"
        Name =""
    End
    Begin
        Left =21
        Top =14
        Right =237
        Bottom =230
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
End
