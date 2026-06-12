Operation =1
Option =2
Begin InputTables
    Name ="T_PolaLubang"
    Name ="T_ListPolaLubang"
End
Begin OutputColumns
    Expression ="T_ListPolaLubang.JumlahLubang"
    Expression ="T_ListPolaLubang.PolaLubang"
    Expression ="T_ListPolaLubang.LokasiImage"
End
Begin Joins
    LeftTable ="T_ListPolaLubang"
    RightTable ="T_PolaLubang"
    Expression ="T_ListPolaLubang.PolaLubang = T_PolaLubang.PolaLubang"
    Flag =1
End
Begin OrderBy
    Expression ="T_ListPolaLubang.JumlahLubang"
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
        dbText "Name" ="T_ListPolaLubang.LokasiImage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ListPolaLubang.JumlahLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_ListPolaLubang.PolaLubang"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =681
    Top =303
    Right =2263
    Bottom =1237
    Left =-1
    Top =-1
    Right =1547
    Bottom =461
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =62
        Top =44
        Right =326
        Bottom =355
        Top =0
        Name ="T_PolaLubang"
        Name =""
    End
    Begin
        Left =417
        Top =84
        Right =748
        Bottom =363
        Top =0
        Name ="T_ListPolaLubang"
        Name =""
    End
End
