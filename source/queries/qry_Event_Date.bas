Operation =1
Option =0
Having ="(((Year([Start_Date])) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="Year([Start_Date])"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="Year([Start_Date])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Visit_Year"
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =985
    Bottom =338
    Left =-1
    Top =-1
    Right =952
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =353
        Top =10
        Right =449
        Bottom =113
        Top =2
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =140
        Top =3
        Right =236
        Bottom =107
        Top =1
        Name ="tbl_Locations"
        Name =""
    End
End
