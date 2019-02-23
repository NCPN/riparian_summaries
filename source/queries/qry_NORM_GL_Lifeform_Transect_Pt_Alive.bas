dbMemo "SQL" ="SELECT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Lifeform_Type, Tra"
    "nsect_Point_Lifeform, COUNT(Transect_Point_Lifeform) AS NumAlive\015\012FROM (SE"
    "LECT  DISTINCT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, IIF"
    "(Lifeform IN ('Graminoid', 'DwarfShrub'), SWITCH(Lifeform&Duration='GraminoidPer"
    "ennial','PGrass', Lifeform&Duration='GraminoidAnnual','AGrass', Lifeform&Duratio"
    "n='DwarfShrubPerennial','Shrub') , Lifeform) AS Lifeform_Type, Lifeform_Type&\"|"
    "\"&Transect&\"-\"&Point AS Transect_Point_Lifeform,Lifeform, Duration FROM qry_N"
    "ORM_Point_Cover_GL_Lifeform WHERE Is_Alive = 1 AND Lifeform IS NOT NULL ORDER BY"
    " Unit_Code, Stream_Name, Plot_ID, Transect, Lifeform)  AS t0\015\012WHERE 1=1\015"
    "\012GROUP BY Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Transect_Poi"
    "nt_Lifeform, Lifeform_Type\015\012ORDER BY Plot_ID, Transect;\015\012"
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
        dbText "Name" ="Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Point_Lifeform"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumAlive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lifeform_Type"
        dbLong "AggregateType" ="-1"
    End
End
