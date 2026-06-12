dbMemo "SQL" ="SELECT qcal_PcsPerSetCampurKeping.BaseProdukID, qcal_PcsPerSetCampurKeping.[pcs/"
    "set], qcal_PcsPerSetCampurKeping.QtyJenisKeping, qcal_PcsPerSetCampurKeping.Camp"
    "urKepingFlagCalc, T_BaseProduk.BaseProdukName\015\012FROM qcal_PcsPerSetCampurKe"
    "ping INNER JOIN T_BaseProduk ON qcal_PcsPerSetCampurKeping.BaseProdukID = T_Base"
    "Produk.BaseProdukID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qcal_PcsPerSetCampurKeping.BaseProdukID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qcal_PcsPerSetCampurKeping.CampurKepingFlagCalc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qcal_PcsPerSetCampurKeping.QtyJenisKeping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qcal_PcsPerSetCampurKeping.[pcs/set]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="T_BaseProduk.BaseProdukName"
        dbLong "AggregateType" ="-1"
    End
End
