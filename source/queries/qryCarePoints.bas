Operation =1
Option =0
Where ="(((tbl_Locations.Unit_Code)=\"care\") AND ((tbl_Locations.Plot_ID)=1) AND ((Year"
    "([Start_Date]))=2010) AND ((tbl_LP_Transect.Transect)=1))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Transect"
    Name ="tbl_LP_Intercept"
End
Begin OutputColumns
    Expression ="tbl_LP_Intercept.Point"
    Expression ="tbl_LP_Intercept.Top"
    Expression ="tbl_LP_Intercept.LCS1"
    Expression ="tbl_LP_Intercept.LCS2"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Transect"
    Expression ="tbl_Events.Event_ID=tbl_LP_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_LP_Transect"
    RightTable ="tbl_LP_Intercept"
    Expression ="tbl_LP_Transect.Transect_ID=tbl_LP_Intercept.Transect_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_LP_Intercept.Point"
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
    Right =1367
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =1
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
        Top =2
        Name ="tbl_LP_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =124
        Top =6
        Name ="tbl_LP_Intercept"
        Name =""
    End
End
