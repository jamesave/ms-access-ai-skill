Operation =1
Option =0
Begin InputTables
    Name ="T_UkuranPolaLubang"
    Name ="T_ProfilKeping"
End
Begin OutputColumns
    Expression ="T_ProfilKeping.KepingID"
    Alias ="Expr1"
    Expression ="[T_ProfilKeping].BlankoKepingID"
    Expression ="T_UkuranPolaLubang.PolaLubangID"
    Expression ="T_UkuranPolaLubang.PolaLubang"
    Expression ="T_ProfilKeping.DeskripsiProfilKeping"
    Expression ="T_UkuranPolaLubang.Catatan"
End
Begin Joins
    LeftTable ="T_UkuranPolaLubang"
    RightTable ="T_ProfilKeping"
    Expression ="T_UkuranPolaLubang.PolaLubangID = T_ProfilKeping.PolaLubangID"
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
        dbText "Name" ="[T_UkuranPolaLubang].PolaLubangID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[T_ProfilKeping].KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[T_ProfilKeping].[DeskripsiProfilKeping]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[T_UkuranPolaLubang].PolaLubang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[T_UkuranPolaLubang].Catatan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =130
    Top =196
    Right =1650
    Bottom =1219
    Left =-1
    Top =-1
    Right =1485
    Bottom =328
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =234
        Top =0
        Name ="T_UkuranPolaLubang"
        Name =""
    End
    Begin
        Left =360
        Top =18
        Right =576
        Bottom =234
        Top =0
        Name ="T_ProfilKeping"
        Name =""
    End
End
