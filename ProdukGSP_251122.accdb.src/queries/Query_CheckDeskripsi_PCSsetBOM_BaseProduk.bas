Operation =1
Option =0
Begin InputTables
    Name ="T_BaseProduk"
    Name ="T_ProdukOEM"
    Name ="T_PcsSetBOM"
End
Begin OutputColumns
    Expression ="T_BaseProduk.SetAutoID"
    Expression ="T_PcsSetBOM.SetAutoID"
    Expression ="T_PcsSetBOM.SetID"
    Expression ="T_PcsSetBOM.Deskripsi_deprecated"
    Expression ="T_BaseProduk.Deskripsi"
End
Begin Joins
    LeftTable ="T_PcsSetBOM"
    RightTable ="T_BaseProduk"
    Expression ="T_PcsSetBOM.SetAutoID = T_BaseProduk.SetAutoID"
    Flag =1
    LeftTable ="T_PcsSetBOM"
    RightTable ="T_ProdukOEM"
    Expression ="T_PcsSetBOM.SetAutoID = T_ProdukOEM.SetID"
    Flag =1
End
Begin OrderBy
    Expression ="T_BaseProduk.SetAutoID"
    Flag =0
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
        dbText "Name" ="T_BaseProduk.SetAutoID"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.Deskripsi_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetAutoID"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.Deskripsi"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =430
    Top =66
    Right =1408
    Bottom =590
    Left =-1
    Top =-1
    Right =954
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =74
        Top =25
        Right =292
        Bottom =215
        Top =0
        Name ="T_BaseProduk"
        Name =""
    End
    Begin
        Left =759
        Top =114
        Right =903
        Bottom =258
        Top =0
        Name ="T_ProdukOEM"
        Name =""
    End
    Begin
        Left =357
        Top =26
        Right =564
        Bottom =242
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
End
