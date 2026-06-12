Operation =1
Option =0
Begin InputTables
    Name ="T_ProfilKeping"
    Name ="T_PcsSetBOM"
End
Begin OutputColumns
    Expression ="T_PcsSetBOM.BaseProdukID"
    Expression ="T_PcsSetBOM.KepingID"
    Expression ="T_ProfilKeping.BlankoKepingID"
    Expression ="T_ProfilKeping.[DeskripsiProfilKeping]"
End
Begin Joins
    LeftTable ="T_ProfilKeping"
    RightTable ="T_PcsSetBOM"
    Expression ="T_ProfilKeping.KepingID = T_PcsSetBOM.KepingID"
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
        dbText "Name" ="T_ProfilKeping.BlankoKepingID"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
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
End
Begin
    State =0
    Left =153
    Top =14
    Right =2547
    Bottom =1115
    Left =-1
    Top =-1
    Right =2359
    Bottom =680
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =438
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
    Begin
        Left =360
        Top =18
        Right =576
        Bottom =437
        Top =0
        Name ="T_ProfilKeping"
        Name =""
    End
End
