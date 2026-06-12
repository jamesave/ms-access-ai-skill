Operation =1
Option =0
Begin InputTables
    Name ="T_BaseProduk"
End
Begin OutputColumns
    Expression ="T_BaseProduk.*"
    Alias ="SortKey"
    Expression ="ExtractMiddleID([BaseProdukName])"
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
        dbText "Name" ="SortKey"
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
        dbText "Name" ="T_BaseProduk.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.KategoriProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.AxlePosID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.Catatan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.StatusProdukID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-11
    Top =-45
    Right =660
    Bottom =835
    Left =-1
    Top =-1
    Right =624
    Bottom =380
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =234
        Top =0
        Name ="T_BaseProduk"
        Name =""
    End
End
