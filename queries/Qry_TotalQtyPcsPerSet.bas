Operation =1
Option =0
Begin InputTables
    Name ="T_ListCustomer"
    Name ="T_PcsSetBOM"
    Name ="T_BaseProduk"
    Name ="T_Produk"
End
Begin OutputColumns
    Expression ="T_BaseProduk.SetAutoID"
    Expression ="T_ListCustomer.NamaCustomer"
    Expression ="T_BaseProduk.KategoriProdukID"
    Expression ="T_BaseProduk.Deskripsi"
    Expression ="T_BaseProduk.AxlePosID"
    Alias ="SumOfQty"
    Expression ="Sum(T_PcsSetBOM.Qty)"
End
Begin Joins
    LeftTable ="T_PcsSetBOM"
    RightTable ="T_BaseProduk"
    Expression ="T_PcsSetBOM.SetAutoID = T_BaseProduk.SetAutoID"
    Flag =1
    LeftTable ="T_PcsSetBOM"
    RightTable ="T_Produk"
    Expression ="T_PcsSetBOM.SetAutoID = T_Produk.SetAutoID"
    Flag =1
    LeftTable ="T_BaseProduk"
    RightTable ="T_Produk"
    Expression ="T_BaseProduk.SetAutoID = T_Produk.SetAutoID"
    Flag =1
    LeftTable ="T_ListCustomer"
    RightTable ="T_Produk"
    Expression ="T_ListCustomer.CustomerID = T_Produk.CustomerID"
    Flag =1
End
Begin OrderBy
    Expression ="T_ListCustomer.NamaCustomer"
    Flag =0
    Expression ="T_BaseProduk.KategoriProdukID"
    Flag =0
End
Begin Groups
    Expression ="T_BaseProduk.SetAutoID"
    GroupLevel =0
    Expression ="T_ListCustomer.NamaCustomer"
    GroupLevel =0
    Expression ="T_BaseProduk.KategoriProdukID"
    GroupLevel =0
    Expression ="T_BaseProduk.Deskripsi"
    GroupLevel =0
    Expression ="T_BaseProduk.AxlePosID"
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
        dbText "Name" ="T_BaseProduk.AxlePosID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.Deskripsi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.KategoriProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.SetAutoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[T_ListCustomer].NamaCustomer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ListCustomer.NamaCustomer"
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
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =361
        Top =20
        Right =577
        Bottom =236
        Top =0
        Name ="T_ListCustomer"
        Name =""
    End
    Begin
        Left =812
        Top =57
        Right =1144
        Bottom =545
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
    Begin
        Left =488
        Top =284
        Right =704
        Bottom =626
        Top =0
        Name ="T_BaseProduk"
        Name =""
    End
    Begin
        Left =96
        Top =18
        Right =312
        Bottom =429
        Top =0
        Name ="T_Produk"
        Name =""
    End
End
