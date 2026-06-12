Operation =1
Option =2
Begin InputTables
    Name ="T_PcsSetBOM"
End
Begin OutputColumns
    Expression ="T_PcsSetBOM.BaseProdukID"
    Alias ="pcs/set"
    Expression ="DSum(\"Qty\",\"T_PcsSetBOM\",\"BaseProdukID=\" & [T_PcsSetBOM]![BaseProdukID])"
    Alias ="QtyJenisKeping"
    Expression ="DCount(\"KepingID\",\"T_PcsSetBOM\",\"BaseProdukID=\" & [T_PcsSetBOM]![BaseProdu"
        "kID])"
    Alias ="CampurKepingFlagCalc"
    Expression ="IIf([QtyJenisKeping]=1,0,1)"
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
        dbText "Name" ="pcs/set"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.SetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CampurKepingFlagCalc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QtyJenisKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_PcsSetBOM.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =313
    Top =424
    Right =2243
    Bottom =1832
    Left =-1
    Top =-1
    Right =1895
    Bottom =198
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =430
        Bottom =343
        Top =0
        Name ="T_PcsSetBOM"
        Name =""
    End
End
