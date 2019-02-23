dbMemo "SQL" ="SELECT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Lifeform_Type, COUNT(Lifefor"
    "m_Type) AS LiveCount\015\012FROM qry_NORM_Lifeform_Transect_Pt_Alive\015\012WHER"
    "E 1=1\015\012GROUP BY Unit_Code, Stream_Name, Visit_Year, Plot_ID, Lifeform_Type"
    "\015\012ORDER BY Unit_Code, Stream_Name, Visit_Year, Plot_ID;\015\012"
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
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lifeform_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiveCount"
        dbLong "AggregateType" ="-1"
    End
End
