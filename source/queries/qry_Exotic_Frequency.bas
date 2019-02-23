Operation =1
Option =0
Having ="(((tbl_LP_Exotic_Freq.Species) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Exotic_Freq"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Stream_Name"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_LP_Exotic_Freq.Species"
End
Begin Joins
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Exotic_Freq"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Exotic_Freq.Transect_ID"
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
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Plot_ID"
    GroupLevel =0
    Expression ="tbl_Locations.Stream_Name"
    GroupLevel =0
    Expression ="Year([Start_Date])"
    GroupLevel =0
    Expression ="tbl_LP_Exotic_Freq.Species"
    GroupLevel =0
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
    Right =1273
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
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
        Top =19
        Name ="tbl_LP_Exotic_Freq"
        Name =""
    End
End
