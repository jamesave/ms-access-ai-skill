Operation =1
Option =2
Begin InputTables
    Name ="T_BaseProduk"
    Name ="T_KategoriProduk"
End
Begin OutputColumns
    Expression ="T_BaseProduk.KategoriProdukID"
    Expression ="T_KategoriProduk.KategoriProduk"
End
Begin Joins
    LeftTable ="T_KategoriProduk"
    RightTable ="T_BaseProduk"
    Expression ="T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID"
    Flag =2
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
        dbText "Name" ="T_KategoriProduk.KategoriProduk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.KategoriProdukID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-11
    Top =-45
    Right =1454
    Bottom =835
    Left =-1
    Top =-1
    Right =1949
    Bottom =732
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =465
        Top =0
        Name ="T_BaseProduk"
        Name =""
    End
    Begin
        Left =360
        Top =18
        Right =576
        Bottom =430
        Top =0
        Name ="T_KategoriProduk"
        Name =""
    End
End
