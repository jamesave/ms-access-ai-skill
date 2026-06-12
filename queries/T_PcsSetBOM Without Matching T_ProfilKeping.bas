Operation =1
Option =0
Where ="(((T_ProfilKeping.KepingID) Is Null))"
Begin InputTables
    Name ="T_PcsSetBOM"
    Name ="T_ProfilKeping"
End
Begin OutputColumns
    Expression ="T_PcsSetBOM.*"
End
Begin Joins
    LeftTable ="T_PcsSetBOM"
    RightTable ="T_ProfilKeping"
    Expression ="T_PcsSetBOM.[KepingID] = T_ProfilKeping.[KepingID]"
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
        dbText "Name" ="T_PcsSetBOM.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetBOMID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.KepingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetID_deprecated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.Qty"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =130
    Top =196
    Right =2389
    Bottom =1219
    Left =-1
    Top =-1
    Right =2224
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =234
        Top =0
        Name ="T_PcsSetBOM"
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
