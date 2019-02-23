Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Site_Impact"
    Name ="tbl_Dist_Exotic"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Stream_Name"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Alias ="Reach_ID"
    Expression ="([Plot_ID])"
    Alias ="Visit_Year"
    Expression ="Year([start_date])"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Expression ="tlu_NCPN_Plants.Duration"
    Expression ="tbl_Dist_Exotic.Notes"
End
Begin Joins
    LeftTable ="tbl_Site_Impact"
    RightTable ="tbl_Dist_Exotic"
    Expression ="tbl_Site_Impact.Impact_ID = tbl_Dist_Exotic.Impact_ID"
    Flag =2
    LeftTable ="tbl_Dist_Exotic"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_Dist_Exotic.Species = tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Site_Impact"
    Expression ="tbl_Events.Event_ID = tbl_Site_Impact.Event_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Unit_Code"
    Flag =0
    Expression ="tbl_Locations.Stream_Name"
    Flag =0
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Flag =0
    Expression ="([Plot_ID])"
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
    Top =40
    Right =1247
    Bottom =392
    Left =-1
    Top =-1
    Right =1210
    Bottom =172
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
        Left =208
        Top =14
        Right =347
        Bottom =132
        Top =2
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =431
        Top =17
        Right =603
        Bottom =135
        Top =2
        Name ="tbl_Site_Impact"
        Name =""
    End
    Begin
        Left =668
        Top =13
        Right =764
        Bottom =131
        Top =1
        Name ="tbl_Dist_Exotic"
        Name =""
    End
    Begin
        Left =847
        Top =13
        Right =1021
        Bottom =131
        Top =2
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
