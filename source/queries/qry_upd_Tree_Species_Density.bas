Operation =4
Option =0
Begin InputTables
    Name ="tbl_wrk_Tree_Species_Density"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Name ="tbl_wrk_Tree_Species_Density.Species_Name"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Name ="tbl_wrk_Tree_Species_Density.Master_Common_Name"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
End
Begin Joins
    LeftTable ="tbl_wrk_Tree_Species_Density"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_Tree_Species_Density.Species = tlu_NCPN_Plants.Master_PLANT_Code"
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
    Right =1400
    Bottom =338
    Left =-1
    Top =-1
    Right =1367
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =426
        Top =6
        Right =613
        Bottom =124
        Top =0
        Name ="tbl_wrk_Tree_Species_Density"
        Name =""
    End
    Begin
        Left =670
        Top =4
        Right =835
        Bottom =122
        Top =2
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
