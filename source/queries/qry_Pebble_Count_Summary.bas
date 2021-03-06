﻿Operation =1
Option =0
Where ="(((tbl_Pebble_Count.Pebble_ID) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Pebble_Transect"
    Name ="tbl_Pebble_Count"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_Pebble_Count.*"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Pebble_Transect"
    Expression ="tbl_Events.Event_ID = tbl_Pebble_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Pebble_Transect"
    RightTable ="tbl_Pebble_Count"
    Expression ="tbl_Pebble_Transect.Transect_ID = tbl_Pebble_Count.Transect_ID"
    Flag =2
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
        Left =306
        Top =6
        Right =402
        Bottom =124
        Top =0
        Name ="tbl_Pebble_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =124
        Top =0
        Name ="tbl_Pebble_Count"
        Name =""
    End
End
