﻿Operation =1
Option =0
Where ="(((tbl_LP_Tree.Shrub_ID) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Tree"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_LP_Belt_Transect.Transect"
    Expression ="tbl_LP_Tree.Shrub_ID"
    Expression ="tbl_LP_Tree.Species"
    Expression ="tbl_LP_Tree.Alive"
    Expression ="tbl_LP_Tree.DBH"
End
Begin Joins
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Tree"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Tree.Transect_ID"
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
    Right =1400
    Bottom =338
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
        Left =330
        Top =3
        Right =485
        Bottom =121
        Top =0
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
    Begin
        Left =518
        Top =4
        Right =614
        Bottom =122
        Top =0
        Name ="tbl_LP_Tree"
        Name =""
    End
End
