Operation =4
Option =0
Begin InputTables
    Name ="tbl_wrk_GL_Cover_Pct_All"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Name ="tbl_wrk_GL_Cover_Pct_All.SpeciesName"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Name ="tbl_wrk_GL_Cover_Pct_All.Lifeform"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Name ="tbl_wrk_GL_Cover_Pct_All.Duration"
    Expression ="tlu_NCPN_Plants.Duration"
    Name ="tbl_wrk_GL_Cover_Pct_All.Nativity"
    Expression ="tlu_NCPN_Plants.Nativity"
    Name ="tbl_wrk_GL_Cover_Pct_All.CommonName"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Name ="tbl_wrk_GL_Cover_Pct_All.Wetland_Code"
    Expression ="tlu_NCPN_Plants.Wetland_Code"
End
Begin Joins
    LeftTable ="tbl_wrk_GL_Cover_Pct_All"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_GL_Cover_Pct_All.SpeciesCode=tlu_NCPN_Plants.Master_PLANT_Code"
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
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =7
        Name ="tbl_wrk_GL_Cover_Pct_All"
        Name =""
    End
    Begin
        Left =602
        Top =4
        Right =764
        Bottom =107
        Top =1
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
