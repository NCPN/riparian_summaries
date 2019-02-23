Operation =1
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
    Expression ="tbl_Pebble_Transect.Transect"
    Expression ="tbl_Pebble_Count.Pebble_ID"
    Expression ="tbl_Pebble_Count.P1"
    Expression ="tbl_Pebble_Count.P2"
    Expression ="tbl_Pebble_Count.P3"
    Expression ="tbl_Pebble_Count.P4"
    Expression ="tbl_Pebble_Count.P5"
    Expression ="tbl_Pebble_Count.P6"
    Expression ="tbl_Pebble_Count.P7"
    Expression ="tbl_Pebble_Count.P8"
    Expression ="tbl_Pebble_Count.P9"
    Expression ="tbl_Pebble_Count.P10"
    Expression ="tbl_Pebble_Count.P11"
    Expression ="tbl_Pebble_Count.P12"
    Expression ="tbl_Pebble_Count.P13"
    Expression ="tbl_Pebble_Count.P14"
    Expression ="tbl_Pebble_Count.P15"
    Expression ="tbl_Pebble_Count.P16"
    Expression ="tbl_Pebble_Count.P17"
    Expression ="tbl_Pebble_Count.P18"
    Expression ="tbl_Pebble_Count.P19"
    Expression ="tbl_Pebble_Count.P20"
    Expression ="tbl_Pebble_Count.P21"
    Expression ="tbl_Pebble_Count.P22"
    Expression ="tbl_Pebble_Count.P23"
    Expression ="tbl_Pebble_Count.P24"
    Expression ="tbl_Pebble_Count.P25"
    Expression ="tbl_Pebble_Count.P26"
    Expression ="tbl_Pebble_Count.P27"
    Expression ="tbl_Pebble_Count.P28"
    Expression ="tbl_Pebble_Count.P29"
    Expression ="tbl_Pebble_Count.P30"
    Expression ="tbl_Pebble_Count.P31"
    Expression ="tbl_Pebble_Count.P32"
    Expression ="tbl_Pebble_Count.P33"
    Expression ="tbl_Pebble_Count.P34"
    Expression ="tbl_Pebble_Count.P35"
    Expression ="tbl_Pebble_Count.P36"
    Expression ="tbl_Pebble_Count.P37"
    Expression ="tbl_Pebble_Count.P38"
    Expression ="tbl_Pebble_Count.P39"
    Expression ="tbl_Pebble_Count.P40"
    Expression ="tbl_Pebble_Count.P41"
    Expression ="tbl_Pebble_Count.P42"
    Expression ="tbl_Pebble_Count.P43"
    Expression ="tbl_Pebble_Count.P44"
    Expression ="tbl_Pebble_Count.P45"
    Expression ="tbl_Pebble_Count.P46"
    Expression ="tbl_Pebble_Count.P47"
    Expression ="tbl_Pebble_Count.P48"
    Expression ="tbl_Pebble_Count.P49"
    Expression ="tbl_Pebble_Count.P50"
    Expression ="tbl_Pebble_Count.P51"
    Expression ="tbl_Pebble_Count.P52"
    Expression ="tbl_Pebble_Count.P53"
    Expression ="tbl_Pebble_Count.P54"
    Expression ="tbl_Pebble_Count.P55"
    Expression ="tbl_Pebble_Count.P56"
    Expression ="tbl_Pebble_Count.P57"
    Expression ="tbl_Pebble_Count.P58"
    Expression ="tbl_Pebble_Count.P59"
    Expression ="tbl_Pebble_Count.P60"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Pebble_Transect"
    Expression ="tbl_Events.Event_ID = tbl_Pebble_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Pebble_Transect"
    RightTable ="tbl_Pebble_Count"
    Expression ="tbl_Pebble_Transect.Transect_ID = tbl_Pebble_Count.Transect_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="Year([Start_Date])"
    Flag =0
    Expression ="tbl_Pebble_Transect.Transect"
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
        Top =2
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
        Top =58
        Name ="tbl_Pebble_Count"
        Name =""
    End
End
