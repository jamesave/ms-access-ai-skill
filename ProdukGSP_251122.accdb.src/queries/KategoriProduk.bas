Operation =1
Option =0
Begin InputTables
    Name ="T_KategoriProduk"
End
Begin OutputColumns
    Expression ="T_KategoriProduk.KategoriProdukID"
    Expression ="T_KategoriProduk.KategoriProduk"
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
        dbText "Name" ="T_KategoriProduk.KategoriProdukID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =195
    Top =274
    Right =2354
    Bottom =1219
    Left =-1
    Top =-1
    Right =2124
    Bottom =524
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =234
        Top =0
        Name ="T_KategoriProduk"
        Name =""
    End
End
