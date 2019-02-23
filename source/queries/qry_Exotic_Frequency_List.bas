Operation =1
Option =0
Where ="(((tbl_LP_Exotic_Freq.Exotic_ID) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Exotic_Freq"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Stream_Name"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_LP_Belt_Transect.Transect"
    Expression ="tbl_LP_Exotic_Freq.Exotic_ID"
    Expression ="tbl_LP_Exotic_Freq.Species"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Expression ="tlu_NCPN_Plants.Duration"
    Expression ="tlu_NCPN_Plants.Nativity"
    Expression ="tbl_LP_Exotic_Freq.M0"
    Expression ="tbl_LP_Exotic_Freq.M5"
    Expression ="tbl_LP_Exotic_Freq.M10"
    Expression ="tbl_LP_Exotic_Freq.M15"
    Expression ="tbl_LP_Exotic_Freq.M20"
    Expression ="tbl_LP_Exotic_Freq.M25"
    Expression ="tbl_LP_Exotic_Freq.M30"
    Expression ="tbl_LP_Exotic_Freq.M35"
    Expression ="tbl_LP_Exotic_Freq.M40"
    Expression ="tbl_LP_Exotic_Freq.M45"
    Expression ="tbl_LP_Exotic_Freq.M50"
    Expression ="tbl_LP_Exotic_Freq.M55"
    Expression ="tbl_LP_Exotic_Freq.M60"
    Expression ="tbl_LP_Exotic_Freq.M65"
    Expression ="tbl_LP_Exotic_Freq.M70"
    Expression ="tbl_LP_Exotic_Freq.M75"
    Expression ="tbl_LP_Exotic_Freq.M80"
    Expression ="tbl_LP_Exotic_Freq.M85"
    Expression ="tbl_LP_Exotic_Freq.M90"
    Expression ="tbl_LP_Exotic_Freq.M95"
End
Begin Joins
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Exotic_Freq"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Exotic_Freq.Transect_ID"
    Flag =2
    LeftTable ="tbl_LP_Exotic_Freq"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_LP_Exotic_Freq.Species = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Belt_Transect"
    Expression ="tbl_Events.Event_ID = tbl_LP_Belt_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="Year([Start_Date])"
    Flag =0
    Expression ="tbl_LP_Belt_Transect.Transect"
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
    Right =1306
    Bottom =338
    Left =-1
    Top =-1
    Right =1269
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =3
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =124
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =451
        Bottom =124
        Top =0
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
    Begin
        Left =523
        Top =6
        Right =678
        Bottom =124
        Top =0
        Name ="tbl_LP_Exotic_Freq"
        Name =""
    End
    Begin
        Left =716
        Top =6
        Right =867
        Bottom =124
        Top =34
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
