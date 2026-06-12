Operation =1
Option =0
Begin InputTables
    Name ="T_BaseProduk"
    Name ="T_ListAxlePos"
    Name ="T_KategoriProduk"
    Name ="T_ListStatusBaseProduk"
End
Begin OutputColumns
    Expression ="T_BaseProduk.BaseProdukID"
    Expression ="T_BaseProduk.BaseProdukName"
    Expression ="T_BaseProduk.Deskripsi"
    Expression ="T_BaseProduk.Catatan"
    Expression ="T_ListStatusBaseProduk.StatusProduk"
    Expression ="T_KategoriProduk.KategoriProduk"
    Expression ="T_ListAxlePos.AxlePos"
    Alias ="SortedBaseProdukName"
    Expression ="Val(Mid([BaseProdukName],2,4))"
    Expression ="T_KategoriProduk.KategoriProdukID"
    Alias ="SortKey"
    Expression ="ExtractMiddleID([BaseProdukName])"
End
Begin Joins
    LeftTable ="T_KategoriProduk"
    RightTable ="T_BaseProduk"
    Expression ="T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID"
    Flag =3
    LeftTable ="T_ListAxlePos"
    RightTable ="T_BaseProduk"
    Expression ="T_ListAxlePos.AxlePosID = T_BaseProduk.AxlePosID"
    Flag =3
    LeftTable ="T_ListStatusBaseProduk"
    RightTable ="T_BaseProduk"
    Expression ="T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID"
    Flag =3
End
Begin OrderBy
    Expression ="ExtractMiddleID([BaseProdukName])"
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
        dbText "Name" ="T_KategoriProduk.KategoriProduk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ListStatusBaseProduk.StatusProduk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.Catatan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.Deskripsi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.BaseProdukName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ListAxlePos.AxlePos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SortedBaseProdukName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_KategoriProduk.KategoriProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SortKey"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-11
    Top =-45
    Right =983
    Bottom =835
    Left =-1
    Top =-1
    Right =947
    Bottom =296
    Left =782
    Top =0
    ColumnsShown =539
    Begin
        Left =-317
        Top =24
        Right =42
        Bottom =398
        Top =0
        Name ="T_BaseProduk"
        Name =""
    End
    Begin
        Left =116
        Top =24
        Right =332
        Bottom =203
        Top =0
        Name ="T_ListAxlePos"
        Name =""
    End
    Begin
        Left =117
        Top =243
        Right =333
        Bottom =387
        Top =0
        Name ="T_KategoriProduk"
        Name =""
    End
    Begin
        Left =-619
        Top =93
        Right =-403
        Bottom =309
        Top =0
        Name ="T_ListStatusBaseProduk"
        Name =""
    End
End
