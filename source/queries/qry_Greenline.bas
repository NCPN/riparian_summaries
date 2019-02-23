Operation =1
Option =0
Where ="(((tbl_GL_Intercept.Point) Is Not Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_GL_Transect"
    Name ="tbl_GL_Intercept"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tbl_Locations.Plot_ID"
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
    Expression ="tbl_GL_Transect.Transect"
    Expression ="tbl_GL_Intercept.Point"
    Expression ="tbl_GL_Intercept.Top"
    Expression ="tbl_GL_Intercept.Alive"
    Expression ="tbl_GL_Intercept.LCS1"
    Expression ="tbl_GL_Intercept.LCA1"
    Expression ="tbl_GL_Intercept.LCS2"
    Expression ="tbl_GL_Intercept.LCA2"
    Expression ="tbl_GL_Intercept.LCS3"
    Expression ="tbl_GL_Intercept.LCA3"
    Expression ="tbl_GL_Intercept.LCS4"
    Expression ="tbl_GL_Intercept.LCA4"
    Expression ="tbl_GL_Intercept.LCS5"
    Expression ="tbl_GL_Intercept.LCA5"
    Expression ="tbl_GL_Intercept.LCS6"
    Expression ="tbl_GL_Intercept.LCA6"
    Expression ="tbl_GL_Intercept.LCS7"
    Expression ="tbl_GL_Intercept.LCA7"
    Expression ="tbl_GL_Intercept.LCS8"
    Expression ="tbl_GL_Intercept.LCA8"
    Expression ="tbl_GL_Intercept.LCS9"
    Expression ="tbl_GL_Intercept.LCA9"
    Expression ="tbl_GL_Intercept.LCS10"
    Expression ="tbl_GL_Intercept.LCA10"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_GL_Transect"
    Expression ="tbl_Events.Event_ID = tbl_GL_Transect.Event_ID"
    Flag =2
    LeftTable ="tbl_GL_Transect"
    RightTable ="tbl_GL_Intercept"
    Expression ="tbl_GL_Transect.Transect_ID = tbl_GL_Intercept.Transect_ID"
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
    Expression ="tbl_GL_Transect.Transect"
    Flag =0
    Expression ="tbl_GL_Intercept.Point"
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
    Right =1306
    Bottom =345
    Left =-1
    Top =-1
    Right =1269
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =46
        Top =4
        Right =142
        Bottom =122
        Top =2
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =200
        Top =4
        Right =314
        Bottom =122
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =370
        Top =4
        Right =491
        Bottom =122
        Top =0
        Name ="tbl_GL_Transect"
        Name =""
    End
    Begin
        Left =555
        Top =5
        Right =682
        Bottom =123
        Top =28
        Name ="tbl_GL_Intercept"
        Name =""
    End
End
