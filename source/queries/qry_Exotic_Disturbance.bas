Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Site_Impact"
    Name ="tbl_Dist_Exotic"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Stream_Name"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_Dist_Exotic.Species"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Site_Impact"
    Expression ="tbl_Events.Event_ID=tbl_Site_Impact.Event_ID"
    Flag =2
    LeftTable ="tbl_Site_Impact"
    RightTable ="tbl_Dist_Exotic"
    Expression ="tbl_Site_Impact.Impact_ID=tbl_Dist_Exotic.Impact_ID"
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
    Left =61
    Top =43
    Right =1306
    Bottom =367
    Left =-1
    Top =-1
    Right =1230
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
        Left =312
        Top =6
        Right =422
        Bottom =124
        Top =0
        Name ="tbl_Site_Impact"
        Name =""
    End
    Begin
        Left =467
        Top =5
        Right =576
        Bottom =123
        Top =0
        Name ="tbl_Dist_Exotic"
        Name =""
    End
End
