Operation =1
Option =0
Begin InputTables
    Name ="T_ListPolaLubang"
    Name ="T_UkuranPolaLubang"
End
Begin OutputColumns
    Expression ="T_UkuranPolaLubang.*"
    Expression ="T_ListPolaLubang.LokasiImage"
    Alias ="ImageCode"
    Expression ="GetHPImagePath([T_ListPolaLubang.PolaLubang])"
End
Begin Joins
    LeftTable ="T_ListPolaLubang"
    RightTable ="T_UkuranPolaLubang"
    Expression ="T_ListPolaLubang.PolaLubang = T_UkuranPolaLubang.PolaLubang"
    Flag =3
End
Begin OrderBy
    Expression ="T_ListPolaLubang.JumlahLubang"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ImageCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =436
    Top =34
    Right =2768
    Bottom =1160
    Left =-1
    Top =-1
    Right =2297
    Bottom =653
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =234
        Top =0
        Name ="T_ListPolaLubang"
        Name =""
    End
    Begin
        Left =360
        Top =18
        Right =576
        Bottom =554
        Top =0
        Name ="T_UkuranPolaLubang"
        Name =""
    End
End
