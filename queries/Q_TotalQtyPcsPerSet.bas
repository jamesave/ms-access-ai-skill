Operation =1
Option =0
Begin InputTables
    Name ="T_ProfilKeping"
    Name ="T_PcsSetBOM"
End
Begin OutputColumns
    Expression ="T_PcsSetBOM.SetID"
    Expression ="T_PcsSetBOM.Deskripsi_deprecated"
    Alias ="SumOfQty"
    Expression ="Sum(T_PcsSetBOM.Qty)"
    Expression ="T_PcsSetBOM.KategoriProdukID_deprecated"
End
Begin Joins
    LeftTable ="T_ProfilKeping"
    RightTable ="T_PcsSetBOM"
    Expression ="[T_ProfilKeping].KepingID=T_PcsSetBOM.KepingID"
    Flag =1
End
Begin Groups
    Expression ="T_PcsSetBOM.SetID"
    GroupLevel =0
    Expression ="T_PcsSetBOM.Deskripsi_deprecated"
    GroupLevel =0
    Expression ="T_PcsSetBOM.KategoriProdukID_deprecated"
    GroupLevel =0
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
        dbText "Name" ="SumOfQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.[Deskripsi_deprecated]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.[KategoriProdukID_deprecated]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.Deskripsi_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.KategoriProdukID_deprecated"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-11
    Top =-45
    Right =1587
    Bottom =820
    Left =-1
    Top =-1
    Right =1523
    Bottom =550
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =831
        Top =11
        Right =1436
        Bottom =618
        Top =0
        Name ="T_DataKeping"
        Name =""
    End
    Begin
        Left =234
        Top =249
        Right =450
        Bottom =669
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
End
