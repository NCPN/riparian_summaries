Operation =1
Option =0
Begin InputTables
    Name ="qry_Densiometer_Totals"
End
Begin OutputColumns
    Expression ="qry_Densiometer_Totals.Unit_Code"
    Expression ="qry_Densiometer_Totals.Stream_Name"
    Expression ="qry_Densiometer_Totals.Plot_ID"
    Expression ="qry_Densiometer_Totals.Visit_Year"
    Alias ="Canopy_Closure"
    Expression ="(([SumOfTotal1]+[SumOfTotal2]+[SumOfTotal3]+[SumOfTotal4])/([Countofpoint]*68))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="qry_Densiometer_Totals.Unit_Code"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Densiometer_Totals.Plot_ID"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Densiometer_Totals.Stream_Name"
        dbInteger "ColumnWidth" ="2430"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Canopy_Closure"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Densiometer_Totals.Visit_Year"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Canopy_Closure2"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Canopy_ClosureO"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
    End
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
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =255
        Bottom =124
        Top =5
        Name ="qry_Densiometer_Totals"
        Name =""
    End
End
