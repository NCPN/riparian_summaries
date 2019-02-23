Operation =4
Option =0
Begin InputTables
    Name ="tbl_wrk_Exotic_Frequency_Summary"
    Name ="tlu_NCPN_Plants"
End
Begin OutputColumns
    Name ="tbl_wrk_Exotic_Frequency_Summary.SpeciesName"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Name ="tbl_wrk_Exotic_Frequency_Summary.CommonName"
    Expression ="tlu_NCPN_Plants.Master_Common_Name"
End
Begin Joins
    LeftTable ="tbl_wrk_Exotic_Frequency_Summary"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_Exotic_Frequency_Summary.SpeciesCode=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =14
    Right =1306
    Bottom =338
    Left =-1
    Top =-1
    Right =1273
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =768
        Top =2
        Right =892
        Bottom =120
        Top =0
        Name ="tlu_NCPN_Plants"
        Name =""
    End
    Begin
        Left =248
        Top =4
        Right =490
        Bottom =122
        Top =3
        Name ="tbl_wrk_Exotic_Frequency_Summary"
        Name =""
    End
End
