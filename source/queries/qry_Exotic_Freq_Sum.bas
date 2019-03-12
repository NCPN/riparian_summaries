Operation =1
Option =0
Where ="(((tbl_LP_Exotic_Freq.Species) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Belt_Transect"
    Name ="tbl_LP_Exotic_Freq"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_LP_Exotic_Freq.Species"
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
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Belt_Transect"
    Expression ="tbl_Events.Event_ID = tbl_LP_Belt_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_LP_Belt_Transect"
    RightTable ="tbl_LP_Exotic_Freq"
    Expression ="tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Exotic_Freq.Transect_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Stream_Name"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="Year([Start_Date])"
    Flag =0
    Expression ="tbl_LP_Exotic_Freq.Species"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Stream_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M70"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M75"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M80"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M85"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M90"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M95"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M35"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M40"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M45"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M50"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M55"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Exotic_Freq.M65"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_LP_Belt_Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =16
    Top =10
    Right =1047
    Bottom =334
    Left =-1
    Top =-1
    Right =999
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
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
        Left =337
        Top =5
        Right =459
        Bottom =123
        Top =0
        Name ="tbl_LP_Belt_Transect"
        Name =""
    End
    Begin
        Left =551
        Top =2
        Right =708
        Bottom =120
        Top =0
        Name ="tbl_LP_Exotic_Freq"
        Name =""
    End
End
