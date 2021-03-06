﻿dbMemo "SQL" ="SELECT \015\012Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \"T"
    "C\" AS Point_Strata, \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.D"
    "uration, \015\012Top AS Species, \015\012ABS(Alive) AS Is_Alive,\015\012Null AS "
    "Disturbance, Null AS Geomorphic_Surface\015\012FROM qry_Point_Cover_Species, tlu"
    "_NCPN_Plants \015\012WHERE \015\012Master_PLANT_Code = Top\015\012AND NOT IsNULL"
    "(tlu_NCPN_Plants.Lifeform)\015\012AND\015\012Len(Top) > 0;\015\012\015\012UNION "
    "ALL\015\012\015\012SELECT \015\012Unit_Code, Stream_Name, Visit_Year, Plot_ID, T"
    "ransect, Point, \"S\" AS Point_Strata, \015\012tlu_NCPN_Plants.Lifeform, \015\012"
    "tlu_NCPN_Plants.Duration, \015\012Surface AS Species, \015\012ABS(Surface_Alive)"
    " AS Is_Alive, \015\012Null AS Disturbance, Null AS Geomorphic_Surface\015\012FRO"
    "M qry_Point_Cover_Species, tlu_NCPN_Plants\015\012WHERE\015\012Master_PLANT_Code"
    " = Surface\015\012AND NOT IsNULL(tlu_NCPN_Plants.Lifeform)\015\012AND\015\012Len"
    "(Surface) > 0;\015\012\015\012UNION ALL\015\012\015\012SELECT Unit_Code, Stream_"
    "Name, Visit_Year, Plot_ID, Transect, Point, \"LC\" AS Point_Strata, \015\012tlu_"
    "NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.Duration,\015\012LCS1 AS Species, "
    "\015\012ABS([LCA1]) AS Is_Alive, \015\012Null AS Disturbance, Null AS Geomorphic"
    "_Surface\015\012FROM qry_Point_Cover_Species, tlu_NCPN_Plants\015\012WHERE\015\012"
    "Master_PLANT_Code = LCS1\015\012AND NOT IsNULL(tlu_NCPN_Plants.Lifeform)\015\012"
    "AND\015\012Len([LCS1])>0;\015\012\015\012UNION ALL\015\012\015\012SELECT Unit_Co"
    "de, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \"LC\" AS 'Point_Strata',"
    " \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.Duration, \015\012LCS"
    "2 AS Species, \015\012ABS([LCA2]) AS Is_Alive, \015\012Null AS Disturbance, Null"
    " AS Geomorphic_Surface\015\012FROM qry_Point_Cover_Species, tlu_NCPN_Plants\015\012"
    "WHERE \015\012Master_PLANT_Code = LCS2\015\012AND NOT IsNULL(tlu_NCPN_Plants.Lif"
    "eform)\015\012AND\015\012Len([LCS2])>0;\015\012\015\012UNION ALL\015\012\015\012"
    "SELECT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \"LC\" AS P"
    "oint_Strata, \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.Duration,"
    "\015\012LCS3 AS Species, \015\012ABS([LCA3]) AS Alive, \015\012Null AS Disturban"
    "ce, Null AS Geomorphic_Surface\015\012FROM qry_Point_Cover_Species, tlu_NCPN_Pla"
    "nts\015\012WHERE \015\012Master_PLANT_Code = LCS3\015\012AND NOT IsNULL(tlu_NCPN"
    "_Plants.Lifeform)\015\012AND\015\012Len([LCS3])>0;\015\012\015\012UNION ALL\015\012"
    "\015\012SELECT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \"L"
    "C\" AS Point_Strata, \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.D"
    "uration,\015\012LCS4 AS Species, \015\012ABS([LCA4]) AS Is_Alive, \015\012Null A"
    "S Disturbance, Null AS Geomorphic_Surface\015\012FROM qry_Point_Cover_Species, t"
    "lu_NCPN_Plants\015\012WHERE \015\012Master_PLANT_Code = LCS4\015\012AND NOT IsNU"
    "LL(tlu_NCPN_Plants.Lifeform)\015\012AND\015\012Len([LCS4])>0;\015\012\015\012UNI"
    "ON ALL\015\012\015\012SELECT Unit_Code, Stream_Name, Visit_Year, Plot_ID, Transe"
    "ct, Point, \"LC\" AS Point_Strata, \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu"
    "_NCPN_Plants.Duration,\015\012LCS5 AS Species, \015\012ABS([LCA5])AS Is_Alive, \015"
    "\012Null AS Disturbance, Null AS Geomorphic_Surface\015\012FROM qry_Point_Cover_"
    "Species, tlu_NCPN_Plants\015\012WHERE \015\012Master_PLANT_Code = LCS5\015\012AN"
    "D NOT IsNULL(tlu_NCPN_Plants.Lifeform)\015\012AND\015\012Len([LCS5])>0;\015\012\015"
    "\012UNION ALL\015\012\015\012SELECT Unit_Code, Stream_Name, Visit_Year, Plot_ID,"
    " Transect, Point, \"LC\" AS Point_Strata, \015\012tlu_NCPN_Plants.Lifeform, \015"
    "\012tlu_NCPN_Plants.Duration,  \015\012LCS6 AS Species, \015\012ABS([LCA6]) AS I"
    "s_Alive, \015\012Null AS Disturbance, Null AS Geomorphic_Surface\015\012FROM qry"
    "_Point_Cover_Species, tlu_NCPN_Plants\015\012WHERE \015\012Master_PLANT_Code = L"
    "CS6\015\012AND NOT IsNULL(tlu_NCPN_Plants.Lifeform)\015\012AND\015\012Len([LCS6]"
    ")>0;\015\012\015\012UNION ALL\015\012\015\012SELECT Unit_Code, Stream_Name, Visi"
    "t_Year, Plot_ID, Transect, Point, \"LC\" AS Point_Strata, \015\012tlu_NCPN_Plant"
    "s.Lifeform, \015\012tlu_NCPN_Plants.Duration, \015\012LCS7 AS Species, \015\012A"
    "BS([LCA7]) AS Is_Alive, \015\012Null AS Disturbance, Null AS Geomorphic_Surface\015"
    "\012FROM qry_Point_Cover_Species, tlu_NCPN_Plants\015\012WHERE \015\012Master_PL"
    "ANT_Code = LCS7\015\012AND NOT IsNULL(tlu_NCPN_Plants.Lifeform)\015\012AND\015\012"
    "Len([LCS7])>0;\015\012\015\012UNION ALL\015\012\015\012SELECT Unit_Code, Stream_"
    "Name, Visit_Year, Plot_ID, Transect, Point, \"LC\" AS Point_Strata, \015\012tlu_"
    "NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.Duration,  \015\012LCS8 AS Species"
    ", \015\012ABS([LCA8]) AS Is_Alive, \015\012Null AS Disturbance, Null AS Geomorph"
    "ic_Surface\015\012FROM qry_Point_Cover_Species, tlu_NCPN_Plants\015\012WHERE \015"
    "\012Master_PLANT_Code = LCS8\015\012AND NOT IsNULL(tlu_NCPN_Plants.Lifeform)\015"
    "\012AND\015\012Len([LCS8])>0;\015\012\015\012UNION ALL\015\012\015\012SELECT Uni"
    "t_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \"LC\" AS Point_Strat"
    "a, \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.Duration,\015\012LC"
    "S9 AS Species, \015\012ABS([LCA9]) AS Is_Alive, \015\012Null AS Disturbance, Nul"
    "l AS Geomorphic_Surface\015\012FROM qry_Point_Cover_Species, tlu_NCPN_Plants\015"
    "\012WHERE \015\012Master_PLANT_Code = LCS9\015\012AND NOT IsNULL(tlu_NCPN_Plants"
    ".Lifeform)\015\012AND\015\012Len([LCS9])>0;\015\012\015\012UNION ALL SELECT Unit"
    "_Code, Stream_Name, Visit_Year, Plot_ID, Transect, Point, \"LC\" AS Point_Strata"
    ", \015\012tlu_NCPN_Plants.Lifeform, \015\012tlu_NCPN_Plants.Duration,\015\012LCS"
    "10 AS Species, \015\012ABS([LCA10]) AS Is_Alive, \015\012Null AS Disturbance, Nu"
    "ll AS Geomorphic_Surface\015\012FROM qry_Point_Cover_Species, tlu_NCPN_Plants\015"
    "\012WHERE \015\012Master_PLANT_Code = LCS10\015\012AND NOT IsNULL(tlu_NCPN_Plant"
    "s.Lifeform)\015\012AND\015\012Len([LCS10])>0;\015\012"
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
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Point"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Point_Strata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Is_Alive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Disturbance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Geomorphic_Surface"
        dbLong "AggregateType" ="-1"
    End
End
