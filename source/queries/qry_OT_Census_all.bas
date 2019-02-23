Operation =1
Option =0
Where ="(((tbl_OT_Census.Census_ID) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_OT_Census"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Stream_Name"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_OT_Census.Census_ID"
    Expression ="tbl_OT_Census.Tag_No"
    Expression ="tbl_OT_Census.Species"
    Expression ="tbl_OT_Census.DBH"
    Expression ="tbl_OT_Census.Crown_Health"
    Expression ="tbl_OT_Census.Notes"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_OT_Census"
    Expression ="tbl_Events.Event_ID = tbl_OT_Census.Event_ID"
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
        dbText "Name" ="tbl_OT_Census.Notes"
        dbInteger "ColumnWidth" ="5355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Stream_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Census_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Tag_No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_OT_Census.Crown_Health"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1306
    Bottom =398
    Left =-1
    Top =-1
    Right =1256
    Bottom =146
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =37
        Top =2
        Right =133
        Bottom =120
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =205
        Top =4
        Right =301
        Bottom =122
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =376
        Top =2
        Right =501
        Bottom =120
        Top =0
        Name ="tbl_OT_Census"
        Name =""
    End
End
