Operation =1
Option =0
Begin InputTables
    Name ="tbl_wrk_Tree_Size_Class"
End
Begin OutputColumns
    Expression ="tbl_wrk_Tree_Size_Class.UnitCode"
    Expression ="tbl_wrk_Tree_Size_Class.Stream_Name"
    Expression ="tbl_wrk_Tree_Size_Class.Visit_Year"
    Expression ="tbl_wrk_Tree_Size_Class.PlotID"
    Alias ="SumOfTotal_Seedlings"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Total_Seedlings)"
    Alias ="SumOfLive_5"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_5)"
    Alias ="SumOfLive_10"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_10)"
    Alias ="SumOfLive_15"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_15)"
    Alias ="SumOfLive_20"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_20)"
    Alias ="SumOfLive_30"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_30)"
    Alias ="SumOfLive_40"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_40)"
    Alias ="SumOfLive_50"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_50)"
    Alias ="SumOfLive_Over_50"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.Live_Over_50)"
    Alias ="SumOfAll_5"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_5)"
    Alias ="SumOfAll_10"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_10)"
    Alias ="SumOfAll_15"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_15)"
    Alias ="SumOfAll_20"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_20)"
    Alias ="SumOfAll_30"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_30)"
    Alias ="SumOfAll_40"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_40)"
    Alias ="SumOfAll_50"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_50)"
    Alias ="SumOfAll_Over_50"
    Expression ="Sum(tbl_wrk_Tree_Size_Class.All_Over_50)"
    Alias ="Total_Live"
    Expression ="Sum(([Live_5]+[Live_10]+[Live_15]+[Live_20]+[Live_30]+[Live_40]+[Live_50]+[Live_"
        "Over_50]))"
    Alias ="Total_All"
    Expression ="Sum(([All_5]+[All_10]+[All_15]+[All_20]+[All_30]+[All_40]+[All_50]+[All_Over_50]"
        "))"
End
Begin Groups
    Expression ="tbl_wrk_Tree_Size_Class.UnitCode"
    GroupLevel =0
    Expression ="tbl_wrk_Tree_Size_Class.Stream_Name"
    GroupLevel =0
    Expression ="tbl_wrk_Tree_Size_Class.Visit_Year"
    GroupLevel =0
    Expression ="tbl_wrk_Tree_Size_Class.PlotID"
    GroupLevel =0
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
    Right =1400
    Bottom =338
    Left =-1
    Top =-1
    Right =1363
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =204
        Bottom =124
        Top =20
        Name ="tbl_wrk_Tree_Size_Class"
        Name =""
    End
End
