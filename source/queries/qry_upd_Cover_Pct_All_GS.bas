Operation =4
Option =0
Begin InputTables
    Name ="tbl_wrk_Cover_Pct_All_GS"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Name ="tbl_wrk_Cover_Pct_All_GS.SpeciesName"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Name ="tbl_wrk_Cover_Pct_All_GS.Lifeform"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Name ="tbl_wrk_Cover_Pct_All_GS.Duration"
    Expression ="tlu_NCPN_Plants.Duration"
    Name ="tbl_wrk_Cover_Pct_All_GS.Nativity"
    Expression ="tlu_NCPN_Plants.Nativity"
    Name ="tbl_wrk_Cover_Pct_All_GS.CommonName"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
End
Begin Joins
    LeftTable ="tbl_wrk_Cover_Pct_All_GS"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_Cover_Pct_All_GS.SpeciesCode=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =14
    Right =984
    Bottom =338
    Left =-1
    Top =-1
    Right =951
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =310
        Top =10
        Right =472
        Bottom =113
        Top =1
        Name ="tlu_NCPN_Plants"
        Name =""
    End
    Begin
        Left =40
        Top =10
        Right =219
        Bottom =113
        Top =2
        Name ="tbl_wrk_Cover_Pct_All_GS"
        Name =""
    End
End
