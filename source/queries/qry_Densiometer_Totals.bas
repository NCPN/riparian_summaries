Operation =1
Option =0
Where ="(((tbl_LP_Densiometer.Total1)>=0))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Densiometer"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Alias ="SumOfTotal1"
    Expression ="Sum(tbl_LP_Densiometer.Total1)"
    Alias ="SumOfTotal2"
    Expression ="Sum(tbl_LP_Densiometer.Total2)"
    Alias ="SumOfTotal3"
    Expression ="Sum(tbl_LP_Densiometer.Total3)"
    Alias ="SumOfTotal4"
    Expression ="Sum(tbl_LP_Densiometer.Total4)"
    Alias ="CountOfPoint"
    Expression ="Count(tbl_LP_Densiometer.Point)"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Belt_Transect"
    Expression ="tbl_Events.Event_ID=tbl_LP_Belt_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Densiometer"
    Expression ="tbl_LP_Belt_Transect.Transect_ID=tbl_LP_Densiometer.Transect_ID"
    Flag =2
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Stream_Name"
    GroupLevel =0
    Expression ="tbl_Locations.Plot_ID"
    GroupLevel =0
    Expression ="Year([Start_Date])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Visit_Year"
    End
    Begin
        dbText "Name" ="SumOfTotal1"
    End
    Begin
        dbText "Name" ="SumOfTotal2"
    End
    Begin
        dbText "Name" ="SumOfTotal3"
    End
    Begin
        dbText "Name" ="SumOfTotal4"
    End
    Begin
        dbText "Name" ="CountOfPoint"
    End
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
    ColumnsShown =543
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
        Left =305
        Top =9
        Right =401
        Bottom =112
        Top =0
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =711
        Bottom =124
        Top =0
        Name ="tbl_LP_Densiometer"
        Name =""
    End
End
