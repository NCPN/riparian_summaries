Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =9
    ItemSuffix =24
    Left =6450
    Top =1320
    Right =11955
    Bottom =8250
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9c87ff146f9de340
    End
    Caption ="Summary Queries"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =705
                    Top =360
                    Width =4335
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Summary of Raw Data"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1440
                    Top =1020
                    Width =2880
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label1"
                    Caption ="Table Data Queries"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =720
                    Top =1560
                    Width =1739
                    Height =300
                    Name ="ButtonPointIntercept"
                    Caption ="Point-Intercept"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =2040
                    Width =1739
                    Height =300
                    TabIndex =1
                    Name ="ButtonExotic"
                    Caption ="Exotic Frequency"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =720
                    Top =2520
                    Width =1739
                    Height =300
                    TabIndex =2
                    Name ="ButtonSeedlings"
                    Caption ="5-m Belt Tree Data"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2220
                    Top =6480
                    Width =1739
                    Height =299
                    TabIndex =3
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =720
                    Top =2040
                    Width =1739
                    Height =299
                    TabIndex =4
                    Name ="ButtonDensiometer"
                    Caption ="Spherical Densiometer"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =2520
                    Width =1740
                    Height =300
                    TabIndex =5
                    Name ="ButtonSoilStability"
                    Caption ="Pebble Count"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =3000
                    Width =1740
                    Height =300
                    TabIndex =6
                    Name ="ButtonOverstoryCensus"
                    Caption ="Overstory Census"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =3000
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =3300
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =1560
                    Width =1740
                    Height =300
                    TabIndex =7
                    Name ="ButtonFuels"
                    Caption ="Greenline PI"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =720
                    Top =3000
                    Width =1739
                    Height =299
                    TabIndex =8
                    Name ="Button_Plant_Walk"
                    Caption ="Exotic Plant Walk"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ButtonPointIntercept_Click()
On Error GoTo Err_ButtonPointIntercept_Click

    Dim stDocName As String

    stDocName = "qry_Point_Intercept"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonPointIntercept_Click:
    Exit Sub

Err_ButtonPointIntercept_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPointIntercept_Click
    
End Sub

Private Sub ButtonExotic_Click()
On Error GoTo Err_ButtonExotic_Click

    Dim stDocName As String

    stDocName = "qry_Exotic_Frequency_List"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonExotic_Click:
    Exit Sub

Err_ButtonExotic_Click:
    MsgBox Err.Description
    Resume Exit_ButtonExotic_Click
    
End Sub
Private Sub ButtonSeedlings_Click()
On Error GoTo Err_ButtonSeedlings_Click

    Dim stDocName As String

    stDocName = "qry_Belt_DBH_All"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonSeedlings_Click:
    Exit Sub

Err_ButtonSeedlings_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSeedlings_Click
    
End Sub



Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub ButtonDensiometer_Click()
On Error GoTo Err_ButtonDensiometer_Click

    Dim stDocName As String

    stDocName = "qry_Spherical_Densiometer"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonDensiometer_Click:
    Exit Sub

Err_ButtonDensiometer_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDensiometer_Click
    
End Sub
Private Sub ButtonSoilStability_Click()
On Error GoTo Err_ButtonSoilStability_Click

    Dim stDocName As String

    stDocName = "qry_Pebble_Count"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonSoilStability_Click:
    Exit Sub

Err_ButtonSoilStability_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSoilStability_Click
    
End Sub




Private Sub ButtonOverstoryCensus_Click()
On Error GoTo Err_ButtonOverstoryCensus_Click

    Dim stDocName As String

    stDocName = "qry_OT_Census_List"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonOverstoryCensus_Click:
    Exit Sub

Err_ButtonOverstoryCensus_Click:
    MsgBox Err.Description
    Resume Exit_ButtonOverstoryCensus_Click
    
End Sub
Private Sub ButtonFuels_Click()
On Error GoTo Err_ButtonFuels_Click

    Dim stDocName As String

    stDocName = "qry_Greenline"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonFuels_Click:
    Exit Sub

Err_ButtonFuels_Click:
    MsgBox Err.Description
    Resume Exit_ButtonFuels_Click
    
End Sub



Private Sub Button_Plant_Walk_Click()
On Error GoTo Err_Button_Plant_Walk_Click

    Dim stDocName As String

    stDocName = "qry_Dist_Exotic"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_Button_Plant_Walk_Click:
    Exit Sub

Err_Button_Plant_Walk_Click:
    MsgBox Err.Description
    Resume Exit_Button_Plant_Walk_Click
    
End Sub
