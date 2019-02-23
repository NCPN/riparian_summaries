Operation =1
Option =0
Where ="(((tbl_wrk_SR_Reach.Species_Code) Not Like \"unk*\") AND ((tlu_NCPN_Plants.Uniqu"
    "e_Species)=0))"
Begin InputTables
    Name ="tbl_wrk_SR_Reach"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Expression ="tbl_wrk_SR_Reach.Unit_Code"
    Expression ="tbl_wrk_SR_Reach.Plot_ID"
    Alias ="Species"
    Expression ="tbl_wrk_SR_Reach.Species_Code"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Expression ="tlu_NCPN_Plants.Duration"
    Expression ="tlu_NCPN_Plants.Unique_Species"
    Expression ="tlu_NCPN_Plants.Nativity"
    Expression ="tbl_wrk_SR_Reach.Stream_Name"
    Expression ="tlu_NCPN_Plants.Wetland_Code"
End
Begin Joins
    LeftTable ="tbl_wrk_SR_Reach"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_SR_Reach.Species_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_wrk_SR_Reach.Unit_Code"
    Flag =0
    Expression ="tbl_wrk_SR_Reach.Plot_ID"
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
    Left =88
    Top =271
    Right =1072
    Bottom =595
    Left =-1
    Top =-1
    Right =969
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =73
        Top =4
        Right =201
        Bottom =122
        Top =0
        Name ="tbl_wrk_SR_Reach"
        Name =""
    End
    Begin
        Left =239
        Top =6
        Right =423
        Bottom =109
        Top =39
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
