Operation =1
Option =0
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
    Expression ="tbl_OT_Census.Species"
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
        Left =37
        Top =2
        Right =133
        Bottom =120
        Top =3
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =205
        Top =4
        Right =301
        Bottom =122
        Top =2
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =376
        Top =2
        Right =501
        Bottom =120
        Top =1
        Name ="tbl_OT_Census"
        Name =""
    End
End
