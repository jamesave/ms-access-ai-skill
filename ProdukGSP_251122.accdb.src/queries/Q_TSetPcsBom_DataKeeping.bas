Operation =1
Option =0
Begin InputTables
    Name ="T_DataKeping"
    Name ="T_PcsSetBOM"
End
Begin OutputColumns
    Expression ="T_PcsSetBOM.SetAutoID"
    Expression ="T_PcsSetBOM.SetID"
    Expression ="T_PcsSetBOM.Qty"
    Expression ="T_PcsSetBOM.KategoriProdukID"
    Expression ="T_PcsSetBOM.KepingID"
    Expression ="T_DataKeping.DeskripsiKeping"
    Expression ="T_DataKeping.Lebar_deprecated"
    Expression ="T_DataKeping.Panjang_deprecated"
    Expression ="T_DataKeping.Tebal_deprecated"
    Expression ="T_DataKeping.DiameterDrum_deprecated"
End
Begin Joins
    LeftTable ="T_DataKeping"
    RightTable ="T_PcsSetBOM"
    Expression ="T_DataKeping.KepingID = T_PcsSetBOM.KepingID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query1].[SetID]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="T_PcsSetBOM.SetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.DeskripsiKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.KategoriProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.DiameterDrum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.Lebar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.Panjang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.Tebal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.[Lebar_deprecated]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.[Panjang_deprecated]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.[Tebal_deprecated]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_DataKeping.[DiameterDrum_deprecated]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1048
    Bottom =642
    Left =-1
    Top =-1
    Right =998
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =407
        Top =-2
        Right =551
        Bottom =388
        Top =0
        Name ="T_DataKeping"
        Name =""
    End
    Begin
        Left =74
        Top =4
        Right =338
        Bottom =279
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
End
