﻿Operation =4
Option =0
Begin InputTables
    Name ="tbl_LP_Intercept"
End
Begin OutputColumns
    Name ="tbl_LP_Intercept.Geomorphic_Surface"
    Expression ="IIf([point]<15,\"channel\",IIf([point]>30,\"floodplain\",\"terrace\"))"
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
        Left =38
        Top =6
        Right =225
        Bottom =124
        Top =29
        Name ="tbl_LP_Intercept"
        Name =""
    End
End
