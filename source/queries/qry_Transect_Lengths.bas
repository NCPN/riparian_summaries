dbMemo "SQL" ="SELECT Unit_Code, Plot_ID, Stream_Name, 1 AS Transect, T1_Length AS Transect_Len"
    "gth\015\012FROM tbl_Locations\015\012UNION\015\012SELECT Unit_Code, Plot_ID, Str"
    "eam_Name, 2 AS Transect, T2_Length AS Transect_Length\015\012FROM tbl_Locations\015"
    "\012UNION\015\012SELECT Unit_Code, Plot_ID, Stream_Name, 3 AS Transect, T3_Lengt"
    "h AS Transect_Length\015\012FROM tbl_Locations\015\012UNION\015\012SELECT Unit_C"
    "ode, Plot_ID, Stream_Name, 4 AS Transect, T4_Length AS Transect_Length\015\012FR"
    "OM tbl_Locations\015\012UNION\015\012SELECT Unit_Code, Plot_ID, Stream_Name, 5 A"
    "S Transect, T5_Length AS Transect_Length\015\012FROM tbl_Locations\015\012UNION\015"
    "\012SELECT Unit_Code, Plot_ID, Stream_Name, 6 AS Transect, T6_Length AS Transect"
    "_Length\015\012FROM tbl_Locations\015\012UNION SELECT Unit_Code, Plot_ID, Stream"
    "_Name, 7 AS Transect, T7_Length AS Transect_Length\015\012FROM tbl_Locations;\015"
    "\012"
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
        dbText "Name" ="Transect_Length"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_Name"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect"
        dbLong "AggregateType" ="-1"
    End
End
