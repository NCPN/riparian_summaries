dbMemo "SQL" ="SELECT DISTINCT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Lifeform_"
    "Type, Transect_Point_Lifeform, IIF(COUNT(Transect_Point_Lifeform) \015>1, 1, COU"
    "NT(Transect_Point_Lifeform)) AS NumALL\015\012FROM (SELECT  DISTINCT\015\012Unit"
    "_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \015\012IIF(Lifeform I"
    "N ('Graminoid', 'DwarfShrub'),\015\012SWITCH(\015\012Lifeform&Duration='Graminoi"
    "dPerennial','PGrass', \015\012Lifeform&Duration='GraminoidAnnual','AGrass', \015"
    "\012Lifeform&Duration='DwarfShrubPerennial','Shrub') \015\012, Lifeform)\015\012"
    "AS Lifeform_Type, \015\012Lifeform_Type&\"|\"&Transect&\"-\"&Point AS Transect_P"
    "oint_Lifeform,\015\012Lifeform, Duration\015\012FROM qry_NORM_Point_Cover_Specie"
    "s_Lifeform\015\012WHERE\015\012Lifeform IS NOT NULL\015\012ORDER BY \015\012Unit"
    "_Code, Stream_Name, Plot_ID, Transect, Lifeform\015\012)  AS t0\015\012WHERE 1=1"
    "\015\012GROUP BY Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Transect"
    "_Point_Lifeform, Lifeform_Type\015\012ORDER BY Plot_ID, Transect;\015\012"
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
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Point_Lifeform"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lifeform_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumALL"
        dbLong "AggregateType" ="-1"
    End
End
