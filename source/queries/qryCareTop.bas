Operation =1
Option =0
Where ="(((tbl_Locations.Unit_Code)=\"care\") AND ((tbl_Locations.Plot_ID)=1) AND ((Year"
    "([Start_Date]))=2010) AND ((tbl_LP_Intercept.Alive)=Yes) AND ((tlu_NCPN_Plants.W"
    "etland_Code)=\"FACU\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Transect"
    Name ="tbl_LP_Intercept"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_LP_Transect.Transect"
    Expression ="tbl_LP_Intercept.Point"
    Expression ="tbl_LP_Intercept.Top"
    Expression ="tbl_LP_Intercept.Alive"
    Expression ="tlu_NCPN_Plants.LU_Code"
    Expression ="tlu_NCPN_Plants.Wetland_Code"
End
Begin Joins
    LeftTable ="tbl_LP_Intercept"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_LP_Intercept.Top = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_LP_Transect"
    RightTable ="tbl_LP_Intercept"
    Expression ="tbl_LP_Transect.Transect_ID = tbl_LP_Intercept.Transect_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Transect"
    Expression ="tbl_Events.Event_ID = tbl_LP_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_LP_Transect.Transect"
    Flag =0
    Expression ="tbl_LP_Intercept.Point"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =18
    Top =14
    Right =1400
    Bottom =516
    Left =-1
    Top =-1
    Right =1363
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =2
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =124
        Top =1
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =124
        Top =2
        Name ="tbl_LP_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =124
        Top =1
        Name ="tbl_LP_Intercept"
        Name =""
    End
    Begin
        Left =574
        Top =6
        Right =723
        Bottom =124
        Top =38
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
