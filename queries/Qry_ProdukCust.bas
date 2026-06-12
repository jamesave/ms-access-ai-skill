Operation =1
Option =0
Begin InputTables
    Name ="T_Produk"
    Name ="T_ListCustomer"
End
Begin OutputColumns
    Expression ="T_Produk.ProdukID"
    Expression ="T_Produk.KodeBarang"
    Expression ="T_Produk.KodeBarangCust"
    Expression ="T_Produk.DeskripsiCust"
    Expression ="T_Produk.CustomerID"
    Expression ="T_Produk.SetAutoID"
    Expression ="T_Produk.BaseProdukID"
    Expression ="T_ListCustomer.NamaCustomer"
End
Begin Joins
    LeftTable ="T_ListCustomer"
    RightTable ="T_Produk"
    Expression ="T_ListCustomer.CustomerID = T_Produk.CustomerID"
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
        dbText "Name" ="T_ListCustomer.NamaCustomer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.ProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.KodeBarang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.CustomerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.DeskripsiCust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.KodeBarangCust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_Produk.SetAutoID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1897
    Bottom =1245
    Left =-1
    Top =-1
    Right =1862
    Bottom =758
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =361
        Top =0
        Name ="T_Produk"
        Name =""
    End
    Begin
        Left =360
        Top =18
        Right =576
        Bottom =234
        Top =0
        Name ="T_ListCustomer"
        Name =""
    End
End
