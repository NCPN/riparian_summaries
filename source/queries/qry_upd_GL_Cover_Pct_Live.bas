Operation =4
Option =0
Begin InputTables
    Name ="tbl_wrk_GL_Cover_Pct_Live"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Name ="tbl_wrk_GL_Cover_Pct_Live.SpeciesName"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Name ="tbl_wrk_GL_Cover_Pct_Live.Lifeform"
    Expression ="tlu_NCPN_Plants.Lifeform"
    Name ="tbl_wrk_GL_Cover_Pct_Live.Duration"
    Expression ="tlu_NCPN_Plants.Duration"
    Name ="tbl_wrk_GL_Cover_Pct_Live.Nativity"
    Expression ="tlu_NCPN_Plants.Nativity"
    Name ="tbl_wrk_GL_Cover_Pct_Live.CommonName"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
    Name ="tbl_wrk_GL_Cover_Pct_Live.Wetland_Code"
    Expression ="tlu_NCPN_Plants.Wetland_Code"
End
Begin Joins
    LeftTable ="tbl_wrk_GL_Cover_Pct_Live"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_GL_Cover_Pct_Live.SpeciesCode=tlu_NCPN_Plants.Master_PLANT_Code"
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
        Left =23
        Top =6
        Right =238
        Bottom =109
        Top =8
        Name ="tbl_wrk_GL_Cover_Pct_Live"
        Name =""
    End
    Begin
        Left =285
        Top =8
        Right =447
        Bottom =111
        Top =1
        Name ="tlu_NCPN_Plants"
        Name =""
    End
End
