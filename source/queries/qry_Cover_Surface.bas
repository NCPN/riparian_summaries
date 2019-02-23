Operation =1
Option =0
Where ="(((tbl_LP_Intercept.Intercept_ID) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_LP_Transect"
    Name ="tbl_LP_Intercept"
    Name ="tlu_LP_Soil_Surface"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_LP_Transect.Transect"
    Expression ="tbl_LP_Intercept.Point"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Alias ="Surface"
    Expression ="tlu_LP_Soil_Surface.Surface_Code"
    Expression ="tbl_LP_Intercept.D1"
    Expression ="tbl_LP_Intercept.D2"
    Expression ="tbl_LP_Intercept.D3"
    Expression ="tbl_LP_Intercept.D4"
    Expression ="tbl_LP_Intercept.D5"
    Expression ="tbl_LP_Intercept.Geomorphic_Surface"
    Expression ="tbl_LP_Intercept.Intercept_ID"
End
Begin Joins
    LeftTable ="tbl_LP_Intercept"
    RightTable ="tlu_LP_Soil_Surface"
    Expression ="tbl_LP_Intercept.Surface=tlu_LP_Soil_Surface.Surface_Code"
    Flag =2
    LeftTable ="tbl_LP_Transect"
    RightTable ="tbl_LP_Intercept"
    Expression ="tbl_LP_Transect.Transect_ID=tbl_LP_Intercept.Transect_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_LP_Transect"
    Expression ="tbl_Events.Event_ID=tbl_LP_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Stream_Name"
    Flag =0
    Expression ="tbl_Locations.Plot_ID"
    Flag =0
    Expression ="tbl_LP_Transect.Transect"
    Flag =0
    Expression ="tbl_LP_Intercept.Point"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =35
    Top =186
    Right =1288
    Bottom =510
    Left =-1
    Top =-1
    Right =1238
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =142
        Bottom =109
        Top =7
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =109
        Top =1
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =412
        Bottom =109
        Top =0
        Name ="tbl_LP_Transect"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =644
        Bottom =109
        Top =0
        Name ="tbl_LP_Intercept"
        Name =""
    End
    Begin
        Left =682
        Top =6
        Right =864
        Bottom =109
        Top =0
        Name ="tlu_LP_Soil_Surface"
        Name =""
    End
End
