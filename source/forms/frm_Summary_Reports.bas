Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =9
    ItemSuffix =56
    Left =675
    Top =1020
    Right =15075
    Bottom =9645
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =8640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =525
                    Left =7320
                    Top =720
                    Width =900
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Locations.Unit_Code FROM tbl_Locations ORDER BY tbl_Location"
                        "s.Unit_Code; "
                    ColumnWidths ="525"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5940
                            Top =720
                            Width =1260
                            Height =245
                            FontWeight =700
                            Name ="Select a park if desired_Label"
                            Caption ="Select a park"
                            EventProcPrefix ="Select_a_park_if_desired_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =1740
                    Width =2219
                    Height =300
                    TabIndex =1
                    Name ="Button_Line_Summary"
                    Caption ="All Species by Reach"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =7980
                    Width =1739
                    Height =300
                    TabIndex =2
                    Name ="Button_Close"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =975
                    Left =7320
                    Top =1140
                    Width =900
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Date"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Event_Date.Visit_Year FROM qry_Event_Date; "
                    ColumnWidths ="975"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5940
                            Top =1140
                            Width =1260
                            Height =245
                            FontWeight =700
                            Name ="Select a date if desired_Label"
                            Caption ="Select a year"
                            EventProcPrefix ="Select_a_date_if_desired_Label"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5040
                    Top =180
                    Width =4320
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Reach Summaries"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =1740
                    Width =2219
                    Height =300
                    TabIndex =4
                    Name ="ButtonPointHit"
                    Caption ="Richness by Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =2220
                    Width =2219
                    Height =300
                    TabIndex =5
                    Name ="RichnessbyWetland"
                    Caption ="Richness by Wetland Status"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =2220
                    Width =2219
                    Height =300
                    TabIndex =6
                    Name ="AllSpecies"
                    Caption ="All Plant Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =2700
                    Width =2219
                    Height =300
                    TabIndex =7
                    Name ="ButtonPercentCoverSpecies"
                    Caption ="Species Cover % Live"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =2700
                    Width =2159
                    Height =300
                    TabIndex =8
                    Name ="Button_Cover_Pct_All"
                    Caption ="Species Cover % All"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =3180
                    Width =2219
                    Height =300
                    TabIndex =9
                    Name ="ButtonCoverLifeform"
                    Caption ="% Cover by Lifeform"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =3180
                    Width =2219
                    Height =300
                    TabIndex =10
                    Name ="ButtonCoverNativity"
                    Caption ="% Cover by Nativity"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =3660
                    Width =2219
                    Height =300
                    TabIndex =11
                    Name ="ButtonCoverWetland"
                    Caption ="% Cover by Wetland Status"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =3660
                    Width =2219
                    Height =300
                    TabIndex =12
                    Name ="ButtonCoverSurface"
                    Caption ="% Cover Surface"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8280
                    Top =1740
                    Width =4320
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label31"
                    Caption ="Cover by Geomorphic Surface"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =2340
                    Width =2219
                    Height =300
                    TabIndex =13
                    Name ="Command32"
                    Caption ="Species Cover % All"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7860
                    Top =2820
                    Width =2219
                    Height =300
                    TabIndex =14
                    Name ="ButtonLifeformGS"
                    Caption ="% Cover by Lifeform"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =2820
                    Width =2219
                    Height =300
                    TabIndex =15
                    Name ="Command34"
                    Caption ="% Cover by Nativity"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7860
                    Top =3300
                    Width =2219
                    Height =300
                    TabIndex =16
                    Name ="ButtonCoverWetlandGS"
                    Caption ="% Cover by Wetland Status"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =3300
                    Width =2219
                    Height =300
                    TabIndex =17
                    Name ="ButtonCoverSurfaceGS"
                    Caption ="% Cover Surface"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Left =7200
                    Top =1740
                    Width =0
                    Height =5460
                    Name ="Line37"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7860
                    Top =4920
                    Width =2219
                    Height =300
                    TabIndex =18
                    Name ="Command38"
                    Caption ="Species Cover % Live"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8625
                    Top =4440
                    Width =3675
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label39"
                    Caption ="Greenline Cover"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =4920
                    Width =2219
                    Height =300
                    TabIndex =19
                    Name ="ButtonGLCoverAll"
                    Caption ="Species Cover % All"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7860
                    Top =5400
                    Width =2219
                    Height =300
                    TabIndex =20
                    Name ="ButtonCoverLifeformGL"
                    Caption ="% Cover by Lifeform"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =5400
                    Width =2219
                    Height =300
                    TabIndex =21
                    Name ="ButtonCoverGLNativity"
                    Caption ="% Cover by Nativity"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7860
                    Top =5880
                    Width =2219
                    Height =300
                    TabIndex =22
                    Name ="ButtonCoverWetlandGL"
                    Caption ="% Cover by Wetland Status"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =4140
                    Width =2219
                    Height =300
                    TabIndex =23
                    Name ="ButtonClosureOverstory"
                    Caption ="% Closure by Overstory"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =4620
                    Width =2219
                    Height =300
                    TabIndex =24
                    Name ="ButtonExoticFrequency"
                    Caption ="Exotic Frequency Summary"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =4620
                    Width =2220
                    Height =300
                    TabIndex =25
                    Name ="ButtonTreeSize"
                    Caption ="Trees by Size Class"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =5100
                    Width =2220
                    Height =300
                    TabIndex =26
                    Name ="ButtonTreeDensity"
                    Caption ="Tree Density"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =5580
                    Width =2220
                    Height =300
                    TabIndex =27
                    Name ="ButtonBasalArea"
                    Caption ="Basal Area"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7860
                    Top =2340
                    Width =2219
                    Height =300
                    TabIndex =28
                    Name ="ButtonCoverPctLiveGS"
                    Caption ="Species Cover % Live"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =4140
                    Width =2220
                    Height =299
                    TabIndex =29
                    Name ="ButtonTotalCover"
                    Caption ="Total % Cover"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7860
                    Top =3780
                    Width =2220
                    Height =299
                    TabIndex =30
                    Name ="ButtonTotalCoverGS"
                    Caption ="Total % Cover"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =5580
                    Width =2220
                    Height =300
                    TabIndex =31
                    Name ="ButtonPebbleCount"
                    Caption ="Pebble Count Summary"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10800
                    Top =5880
                    Width =2220
                    Height =299
                    TabIndex =32
                    Name ="ButtonGLTotalCover"
                    Caption ="Total % Cover"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =5100
                    Width =2220
                    Height =300
                    TabIndex =33
                    Name ="ButtonTreeSpeciesDensity"
                    Caption ="Tree Species  Density"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =6060
                    Width =2220
                    Height =330
                    TabIndex =34
                    Name ="ButtonTreeCensus"
                    Caption ="Tree Census Summary"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =6060
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =6390
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


Private Sub AllSpecies_Click()
On Error GoTo Err_AllSpecies_Click

  Dim strSQL As String
  Dim stDocName As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim SpeciesIn As DAO.Recordset
  Dim points As DAO.Recordset
  Dim CanopyIndex As Integer
  Dim RecordCount As Long
  Dim SpeciesColumn As String
  Dim Unit As String
  Dim plot As Integer
  Dim Stream As String
  Dim Species As String
  Dim Plant_Name As Variant
  Dim LCIndex As String
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Reach_Species"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_SR_Intercept where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If

  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0

  ' Start with cross-section point intercept
   Set WorkOutput = db.OpenRecordset("tbl_Reach_Species")
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF  ' Load all 10 lower canopy species into work table.
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       Species = points!Top
       GoSub Write_Detail
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And points!Surface <> "" And points!Surface <> " " And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       Species = points!Surface
       GoSub Write_Detail
     End If  ' End if for null soil surface check

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipEntry
       Else
         Species = points(SpeciesColumn)
         GoSub Write_Detail
       End If  ' End if for null lower canopy check
SkipEntry:
       LCIndex = LCIndex + 1
     Loop
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for greenline point intercept
  strSQL = "SELECT * FROM qry_GL_Intercept where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do Greenline point intercept
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF  ' Load all 10 lower canopy species into work table.
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       Species = points!Top
       GoSub Write_Detail
     End If  ' End if for null top check

     '  Soil Surface next
     ' If Not IsNull(Points!Surface) And Points!Surface <> "" And Points!Surface <> " " And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & Points!Surface & "'")) Then
       ' Species = Points!Surface
       ' GoSub Write_Detail
     ' End If  ' End if for null soil surface check

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipGLEntry
       Else
         Species = points(SpeciesColumn)
         GoSub Write_Detail
       End If  ' End if for null lower canopy check
SkipGLEntry:
       LCIndex = LCIndex + 1
     Loop
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for exotic frequency
  strSQL = "SELECT * FROM qry_Exotic_Frequency where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do Exotic Frequency
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     Species = points!Species
     GoSub Write_Detail
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for belt trees
  strSQL = "SELECT * FROM qry_Belt_Tree where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do belt trees
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     Species = points!Species
     GoSub Write_Detail
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for OT Census
  strSQL = "SELECT * FROM qry_OT_Census where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do OT Census
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     Species = points!Species
     GoSub Write_Detail
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for Exotic Disturbance (exotic plant walk)
  strSQL = "SELECT * FROM qry_Exotic_Disturbance where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do OT Census
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     Species = points!Species
     GoSub Write_Detail
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
   WorkOutput.Close
   Set WorkOutput = Nothing

Exit_AllSpecies_Click:
    DoCmd.Hourglass False
    MsgBox RecordCount & " records written.  Results are in tbl_Reach_Species."
    Exit Sub

Err_AllSpecies_Click:
    MsgBox Err.Description
    Resume Exit_AllSpecies_Click
    
Write_Detail:
  If IsNull(DLookup("Unit_Code", "tbl_Reach_Species", "[Unit_Code]='" & Unit & "' AND Species_Code = '" & Species & "' AND Plot_ID = " & plot)) Then
    WorkOutput.AddNew
    WorkOutput!Unit_Code = Unit
    WorkOutput!Stream_Name = Stream
    WorkOutput!Plot_ID = plot
    WorkOutput!Visit_Year = Me!Visit_Date
    WorkOutput!Species_Code = Species
    Plant_Name = DLookup("Utah_Species", "tlu_NCPN_Plants", "[Master_Plant_Code]= '" & Species & "'")
    If Not IsNull(Plant_Name) Then
      WorkOutput!Species_Name = Plant_Name
    End If
    WorkOutput.Update
    RecordCount = RecordCount + 1
  End If
Return
End Sub

Private Sub Button_Cover_Pct_All_Click()
On Error GoTo Err_CoverSpeciesAll_Click

  Dim strSQL As String
  Dim lifeForm As Variant
  Dim SpeciesColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim ArrayEnd As Integer
  Dim PointArray(12) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 is species code
  Dim PlotArray(300, 1) As Variant ' Array for species in a plot
  ' Species hits per plot array
  ' Column 1 is species code
  ' Column 2 is alive count
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_All"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_Species where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpeciesAll_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 299           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary
   StreamSave = points!Stream_Name
   Point_Count = 0

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Is it a new plot
       PointIndex = 0  ' yes - add in last point
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextAIndex  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextAIndex:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       ' *** End of plot processing ***
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_All")
       ArrayIndex = 0
       Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for last plot
         If PlotArray(ArrayIndex, 1) > 0 Then
           WorkOutput.AddNew
           WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
           WorkOutput!Visit_Year = Me!Visit_Date
           WorkOutput!Stream_Name = StreamSave
           WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
           WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
           WorkOutput!PercentCoverAll = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           WorkOutput.Update  ' Write previous output record
           RecordCount = RecordCount + 1
         End If
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 299    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
       Point_Count = 0
     End If
     If PointSave <> points!point Then  ' Is it a new point
     '  *** End of point processing ***
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Top = PointArray(PointIndex) Then  ' is the species the same
             Exit Do   ' Already have the species for this point
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Top  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Surface = PointArray(PointIndex) Then  ' is the species the same
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Surface  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipSpecies
       Else
         PointIndex = 0
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex) = points(SpeciesColumn)  ' set species
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipSpecies:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   ' End of file - add in last point
   PointIndex = 0
     Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
       ArrayIndex = 0
       Do Until ArrayIndex > 299
         If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
           If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           Else
             GoTo LastPlotArray  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
           Exit Do
         End If  ' end if for array slot open test
LastPlotArray:
         ArrayIndex = ArrayIndex + 1
       Loop
       PointIndex = PointIndex + 1
     Loop
     Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_All")  ' Write last output record
     ArrayIndex = 0
     Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "
       If PlotArray(ArrayIndex, 1) > 0 Then
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!PercentCoverAll = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
         WorkOutput.Update  ' Write previous output record
         RecordCount = RecordCount + 1
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_upd_Cover_Pct_All"   ' Update species names.
   DoCmd.SetWarnings True
Exit_CoverSpeciesAll_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_All."
    Exit Sub

Err_CoverSpeciesAll_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpeciesAll_Click
End Sub

Private Sub Button_Line_Summary_Click()
On Error GoTo Err_Button_Line_Summary_Click

  Dim strSQL As String
  Dim stDocName As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim SpeciesIn As DAO.Recordset
  Dim points As DAO.Recordset
  Dim CanopyIndex As Integer
  Dim RecordCount As Long
  Dim SpeciesColumn As String
  Dim Unit As String
  Dim plot As Integer
  Dim Stream As String
  Dim Species As String
  Dim Plant_Name As Variant
  Dim LCIndex As String
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_SR_Reach"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_SR_Intercept where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If

  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0

  ' Start with cross-section point intercept
   Set WorkOutput = db.OpenRecordset("tbl_wrk_SR_Reach")
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF  ' Load all lower canopy species into work table.
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       Species = points!Top
       GoSub Write_Detail
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And points!Surface <> "" And points!Surface <> " " And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       Species = points!Surface
       GoSub Write_Detail
     End If  ' End if for null soil surface check

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipEntry
       Else
         Species = points(SpeciesColumn)
         GoSub Write_Detail
       End If  ' End if for null lower canopy check
SkipEntry:
       LCIndex = LCIndex + 1
     Loop
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for greenline point intercept
  strSQL = "SELECT * FROM qry_GL_Intercept where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do Greenline point intercept
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF  ' Load all lower canopy species into work table.
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       Species = points!Top
       GoSub Write_Detail
     End If  ' End if for null top check

     '  Soil Surface next
     ' If Not IsNull(Points!Surface) And Points!Surface <> "" And Points!Surface <> " " And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & Points!Surface & "'")) Then
       ' Species = Points!Surface
       ' GoSub Write_Detail
     ' End If  ' End if for null soil surface check

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipGLEntry
       Else
         Species = points(SpeciesColumn)
         GoSub Write_Detail
       End If  ' End if for null lower canopy check
SkipGLEntry:
       LCIndex = LCIndex + 1
     Loop
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for exotic frequency
  strSQL = "SELECT * FROM qry_Exotic_Frequency where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do Exotic Frequency
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     Species = points!Species
     GoSub Write_Detail
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   
'  Build SQL statement for belt trees
  strSQL = "SELECT * FROM qry_Belt_Tree where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' Now do belt trees
   Set points = db.OpenRecordset(strSQL)
   Do Until points.EOF
     Unit = points!Unit_Code
     plot = points!Plot_ID
     Stream = points!Stream_Name
     Species = points!Species
     GoSub Write_Detail
     points.MoveNext
   Loop
   points.Close
   Set points = Nothing
   WorkOutput.Close
   Set WorkOutput = Nothing

Exit_Button_Line_Summary_Click:
    DoCmd.Hourglass False
    MsgBox RecordCount & " records written.  Results are in tbl_wrk_SR_Reach."
    Exit Sub

Err_Button_Line_Summary_Click:
    MsgBox Err.Description
    Resume Exit_Button_Line_Summary_Click
    
Write_Detail:
  If IsNull(DLookup("Unit_Code", "tbl_Wrk_SR_Reach", "[Unit_Code]='" & Unit & "' AND Species_Code = '" & Species & "' AND Plot_ID = " & plot)) Then
    WorkOutput.AddNew
    WorkOutput!Unit_Code = Unit
    WorkOutput!Stream_Name = Stream
    WorkOutput!Plot_ID = plot
    WorkOutput!Visit_Year = Me!Visit_Date
    WorkOutput!Species_Code = Species
    Plant_Name = DLookup("Utah_Species", "tlu_NCPN_Plants", "[Master_Plant_Code]= '" & Species & "'")
    If Not IsNull(Plant_Name) Then
      WorkOutput!Species_Name = Plant_Name
    End If
    WorkOutput.Update
    RecordCount = RecordCount + 1
  End If
Return
    
End Sub
Private Sub Button_Close_Click()
On Error GoTo Err_Button_Close_Click


    DoCmd.Close

Exit_Button_Close_Click:
    Exit Sub

Err_Button_Close_Click:
    MsgBox Err.Description
    Resume Exit_Button_Close_Click
    
End Sub

Private Sub ButtonBasalArea_Click()
On Error GoTo Err_BasalArea_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Trees As DAO.Recordset
  Dim DBHLive As Double
  Dim DBHAll As Double
  Dim PlotSave As Variant
  Dim SpeciesSave As String
  Dim NameSave As String
  Dim SizeColumn As String
  Dim ReachData As DAO.Recordset
  Dim TreeIndex As Integer
  Dim TransectSum As Double
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Basal_Area"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Tree_DBH where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' strSQL = strSQL & " AND Plot_ID = " & 1 & " AND Transect = " & 1
  strSQL = strSQL & " ORDER BY Unit_Code, Visit_Year, Plot_ID, Species"
  
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get tree size info
   Set Trees = db.OpenRecordset(strSQL)
   If Trees.EOF Then
     MsgBox "No valid tree records found."
     Trees.Close
     Set Trees = Nothing
     GoTo Exit_BasalArea_Click
   End If
   Trees.MoveFirst
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Basal_Area")  ' Open output table
   PlotSave = Trees!Unit_Code & Trees!Plot_ID     ' Save necessary fields
   SpeciesSave = Trees!Species
   NameSave = Trees!Stream_Name
   DBHLive = 0
   DBHAll = 0
   Do Until Trees.EOF
     If (PlotSave <> Trees!Unit_Code & Trees!Plot_ID) Or (SpeciesSave <> Trees!Species) Then
       strSQL = "SELECT * FROM tbl_Locations WHERE Unit_Code = '" & Left(PlotSave, 4) & "' AND Plot_ID = " & Right(PlotSave, Len(PlotSave) - 4)
       Set ReachData = db.OpenRecordset(strSQL)
       If ReachData.EOF Then
         MsgBox "Reach data not found.  Unit " & Trees!unitCode & " Reach " & Trees!plotID & "."
         GoTo Exit_BasalArea_Click
       End If
       TreeIndex = 1   ' Initialize index
       TransectSum = 0
       Do Until TreeIndex > 7  ' Go through transect length fields and sum them
         SizeColumn = "T" & TreeIndex & "_Length" ' Get the transect length field
         If Not IsNull(ReachData(SizeColumn)) Then
           TransectSum = TransectSum + ReachData(SizeColumn)  ' accumulate total transect lengths
         End If
         TreeIndex = TreeIndex + 1
       Loop
       TransectSum = TransectSum * 0.0005 ' Sum of transects times 5 plus a bunch of decimals for divisor
       ReachData.Close
       Set ReachData = Nothing
       WorkOutput.AddNew  ' New species - write an output record
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = NameSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Species = SpeciesSave
       WorkOutput!Basal_Area_L = DBHLive / TransectSum
       WorkOutput!Basal_Area_A = DBHAll / TransectSum
       WorkOutput.Update  ' Write it
       PlotSave = Trees!Unit_Code & Trees!Plot_ID     ' Save necessary fields
       SpeciesSave = Trees!Species
       NameSave = Trees!Stream_Name
       DBHLive = 0
       DBHAll = 0
     End If
     DBHAll = DBHAll + (((Trees!DBH / 200) * (Trees!DBH / 200)) * 3.142)
     If Trees!alive Then
       DBHLive = DBHLive + (((Trees!DBH / 200) * (Trees!DBH / 200)) * 3.142)
     End If
     Trees.MoveNext
   Loop
       ' Write last output record
       strSQL = "SELECT * FROM tbl_Locations WHERE Unit_Code = '" & Left(PlotSave, 4) & "' AND Plot_ID = " & Right(PlotSave, Len(PlotSave) - 4)
       Set ReachData = db.OpenRecordset(strSQL)
       If ReachData.EOF Then
         MsgBox "Reach data not found.  Unit " & Trees!unitCode & " Reach " & Trees!plotID & "."
         GoTo Exit_BasalArea_Click
       End If
       TreeIndex = 1   ' Initialize index
       TransectSum = 0
       Do Until TreeIndex > 7  ' Go through transect length fields and sum them
         SizeColumn = "T" & TreeIndex & "_Length" ' Get the transect length field
         If Not IsNull(ReachData(SizeColumn)) Then
           TransectSum = TransectSum + ReachData(SizeColumn)  ' accumulate total transect lengths
         End If
         TreeIndex = TreeIndex + 1
       Loop
       TransectSum = TransectSum * 0.0005 ' Sum of transects times 5 plus a bunch of decimals for divisor
       ReachData.Close
       Set ReachData = Nothing
       WorkOutput.AddNew  ' New species - write an output record
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = NameSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Species = SpeciesSave
       WorkOutput!Basal_Area_L = (DBHLive) / TransectSum
       WorkOutput!Basal_Area_A = (DBHAll) / TransectSum
       WorkOutput.Update  ' Write it
       Trees.Close
       Set Trees = Nothing
Exit_BasalArea_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Basal_Area."
    Exit Sub

Err_BasalArea_Click:
    MsgBox Err.Description
    Resume Exit_BasalArea_Click
End Sub

Private Sub ButtonClosureOverstory_Click()
On Error GoTo Err_ClosureOverstory_Click

  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim db As DAO.Database

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Closure_Overstory"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Canopy_Closure where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid densiometer records found."
     points.Close
     Set points = Nothing
     GoTo Exit_ClosureOverstory_Click
   End If

   Set WorkOutput = db.OpenRecordset("tbl_wrk_Canopy_Closure")
   Do Until points.EOF
       WorkOutput.AddNew
       WorkOutput!unitCode = points!Unit_Code  ' Set unit code
       WorkOutput!Stream_Name = points!Stream_Name
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = points!Plot_ID  ' Set plot ID
       WorkOutput!Canopy_Closure = points!Canopy_Closure * 100
       WorkOutput.Update  ' Write plot record
       points.MoveNext
   Loop
       WorkOutput.Close
       Set WorkOutput = Nothing

Exit_ClosureOverstory_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Canopy_Closure."
    Exit Sub

Err_ClosureOverstory_Click:
    MsgBox Err.Description
    Resume Exit_ClosureOverstory_Click
End Sub

Private Sub ButtonCoverGLNativity_Click()
On Error GoTo Err_CoverNativity_Click
  Dim strSQL As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Nativity As String
  Dim PointSave As Double
  Dim LCIndex As Integer
  Dim PlotTotalA As Integer
  Dim PlotTotalL As Integer
  Dim PointNativeL As Byte
  Dim PointNativeA As Byte
  Dim PointNonNativeL As Byte
  Dim PointNonNativeA As Byte
  Dim PlotTotalNativeL As Integer
  Dim PlotTotalNativeA As Integer
  Dim PlotTotalNonNativeL As Integer
  Dim PlotTotalNonNativeA As Integer

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_GL_Cover_Pct_Nativity"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GL where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverNativity_Click
   End If
   PlotTotalL = 0
   PlotTotalA = 0
   PointNativeL = 0
   PointNativeA = 0
   PointNonNativeL = 0
   PointNonNativeA = 0
   PlotTotalNativeL = 0
   PlotTotalNativeA = 0
   PlotTotalNonNativeL = 0
   PlotTotalNonNativeA = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - process last point totals from previous plot first
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Nativity")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' WorkOutput!TotalL = (PlotTotalL / 560) * 100
       ' WorkOutput!TotalA = (PlotTotalA / 560) * 100
       WorkOutput!NativeL = (PlotTotalNativeL / 560) * 100
       WorkOutput!NativeA = (PlotTotalNativeA / 560) * 100
       WorkOutput!NonNativeL = (PlotTotalNonNativeL / 560) * 100
       WorkOutput!NonNativeA = (PlotTotalNonNativeA / 560) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       PlotTotalL = 0
       PlotTotalA = 0
       PointNativeL = 0
       PointNativeA = 0
       PointNonNativeL = 0
       PointNonNativeA = 0
       PlotTotalNativeL = 0
       PlotTotalNativeA = 0
       PlotTotalNonNativeL = 0
       PlotTotalNonNativeA = 0
     End If
     If PointSave <> points!point Then  ' End of point - add counts to plot array
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       PointLive = 0
       PointAll = 0
       PointSave = points!point  '  Save new point
       PointNativeL = 0
       PointNativeA = 0
       PointNonNativeL = 0
       PointNonNativeA = 0
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Nativity])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       Nativity = SpeciesLU!Nativity
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If Nativity = "Native" Then
         PointNativeA = 1
         If points!alive Then
           PointNativeL = 1
         End If
       Else
         PointNonNativeA = 1
         If points!alive Then
           PointNonNativeL = 1
         End If
       End If
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface removed 4/24/2013 RD.

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Nativity])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
           Nativity = SpeciesLU!Nativity
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Nativity = "Native" Then
           PointNativeA = 1
           If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
             PointNativeL = 1
           End If
         Else
           PointNonNativeA = 1
           If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
             PointNonNativeL = 1
           End If
         End If
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
   Loop
   '  Process last point totals
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Nativity")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' WorkOutput!TotalL = (PlotTotalL / 560) * 100
       ' WorkOutput!TotalA = (PlotTotalA / 560) * 100
       WorkOutput!NativeL = (PlotTotalNativeL / 560) * 100
       WorkOutput!NativeA = (PlotTotalNativeA / 560) * 100
       WorkOutput!NonNativeL = (PlotTotalNonNativeL / 560) * 100
       WorkOutput!NonNativeA = (PlotTotalNonNativeA / 560) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverNativity_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_GL_Cover_Pct_Nativity."
    Exit Sub

Err_CoverNativity_Click:
    MsgBox Err.Description
    Resume Exit_CoverNativity_Click
End Sub

' ---------------------------------
' SUB:          ButtonCoverLifeform_Click
' Description:  Calculate lifeform percentage cover
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 9, 2014 - for NCPN Riparian summaries tool
' Revisions:
'       7/9/2014        BLC     initial version
' ---------------------------------
Private Sub ButtonCoverLifeform_Click()
On Error GoTo Err_Handler:
Dim visitYear As Integer
Dim parkCode As String
    
    'handle NULLs
    If IsNull(Me!Visit_Date) Then
        visitYear = 0
    Else
        visitYear = CInt(Me!Visit_Date)
    End If
    If IsNull(Me!Park_Code) Then
        parkCode = ""
    Else
        parkCode = Me!Park_Code
    End If
    
    getLifeformCounts parkCode, visitYear, "", "riparian"

Exit_Sub:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonCoverLifeform_Click[Form_frm_Summary_Reports])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ButtonCoverLifeformGL_Click
' Description:  Calculate lifeform percentage cover for greenline point intercept data
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 9, 2014 - for NCPN Riparian summaries tool
' Revisions:
'       8/6/2014        BLC     initial version
' ---------------------------------
Private Sub ButtonCoverLifeformGL_Click()
On Error GoTo Err_Handler:
Dim visitYear As Integer
Dim parkCode As String
    
    'handle NULLs
    If IsNull(Me!Visit_Date) Then
        visitYear = 0
    Else
        visitYear = CInt(Me!Visit_Date)
    End If
    If IsNull(Me!Park_Code) Then
        parkCode = ""
    Else
        parkCode = Me!Park_Code
    End If
    
    getLifeformCounts parkCode, visitYear, "GL", "riparian"

Exit_Sub:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonCoverLifeformGL_Click[Form_frm_Summary_Reports])"
    End Select
    Resume Exit_Sub
End Sub

Private Sub ButtonCoverNativity_Click()
On Error GoTo Err_CoverNativity_Click
  Dim strSQL As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Nativity As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim LCIndex As Integer
  Dim PlotTotalA As Integer
  Dim PlotTotalL As Integer
  Dim PointNativeL As Byte
  Dim PointNativeA As Byte
  Dim PointNonNativeL As Byte
  Dim PointNonNativeA As Byte
  Dim PlotTotalNativeL As Integer
  Dim PlotTotalNativeA As Integer
  Dim PlotTotalNonNativeL As Integer
  Dim PlotTotalNonNativeA As Integer

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Nativity"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_Species where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverNativity_Click
   End If
   PlotTotalL = 0
   PlotTotalA = 0
   PointNativeL = 0
   PointNativeA = 0
   PointNonNativeL = 0
   PointNonNativeA = 0
   PlotTotalNativeL = 0
   PlotTotalNativeA = 0
   PlotTotalNonNativeL = 0
   PlotTotalNonNativeA = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   Point_Count = 0
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - process last point totals from previous plot first
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Nativity")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       '  WorkOutput!TotalL = (PlotTotalL / Point_Count) * 100  dropped 8/23/2011 rd
       '  WorkOutput!TotalA = (PlotTotalA / Point_Count) * 100
       WorkOutput!NativeL = (PlotTotalNativeL / Point_Count) * 100
       WorkOutput!NativeA = (PlotTotalNativeA / Point_Count) * 100
       WorkOutput!NonNativeL = (PlotTotalNonNativeL / Point_Count) * 100
       WorkOutput!NonNativeA = (PlotTotalNonNativeA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Point_Count = 0
       PlotTotalL = 0
       PlotTotalA = 0
       PointNativeL = 0
       PointNativeA = 0
       PointNonNativeL = 0
       PointNonNativeA = 0
       PlotTotalNativeL = 0
       PlotTotalNativeA = 0
       PlotTotalNonNativeL = 0
       PlotTotalNonNativeA = 0
     End If
     If PointSave <> points!point Then  ' End of point - add counts to plot array
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       PointLive = 0
       PointAll = 0
       PointSave = points!point  '  Save new point
       PointNativeL = 0
       PointNativeA = 0
       PointNonNativeL = 0
       PointNonNativeA = 0
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Nativity])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       Nativity = SpeciesLU!Nativity
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If Nativity = "Native" Then
         PointNativeA = 1
         If points!alive Then
           PointNativeL = 1
         End If
       Else
         PointNonNativeA = 1
         If points!alive Then
           PointNonNativeL = 1
         End If
       End If
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Surface & "' AND NOT IsNull([Nativity])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       Nativity = SpeciesLU!Nativity
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If Nativity = "Native" Then
         PointNativeA = 1
         If points!Surface_Alive Then
           PointNativeL = 1
         End If
       Else
         PointNonNativeA = 1
         If points!Surface_Alive Then
           PointNonNativeL = 1
         End If
       End If
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Nativity])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
           Nativity = SpeciesLU!Nativity
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Nativity = "Native" Then
           PointNativeA = 1
           If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
             PointNativeL = 1
           End If
         Else
           PointNonNativeA = 1
           If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
             PointNonNativeL = 1
           End If
         End If
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   '  Process last point totals
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Nativity")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       '  WorkOutput!TotalL = (PlotTotalL / Point_Count) * 100
       '  WorkOutput!TotalA = (PlotTotalA / Point_Count) * 100
       WorkOutput!NativeL = (PlotTotalNativeL / Point_Count) * 100
       WorkOutput!NativeA = (PlotTotalNativeA / Point_Count) * 100
       WorkOutput!NonNativeL = (PlotTotalNonNativeL / Point_Count) * 100
       WorkOutput!NonNativeA = (PlotTotalNonNativeA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverNativity_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Nativity."
    Exit Sub

Err_CoverNativity_Click:
    MsgBox Err.Description
    Resume Exit_CoverNativity_Click
End Sub

Private Sub ButtonCoverPctLiveGS_Click()
On Error GoTo Err_CoverSpeciesLiveGS_Click

  Dim strSQL As String
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim ACount As Integer
  Dim Point_Count As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim ArrayEnd As Integer
  Dim PointArray(12) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 is species code
  Dim PlotArray(300, 1) As Variant ' Array for species in a plot
  ' Species hits per plot array
  ' Column 1 is species code
  ' Column 2 is alive count
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Live_GS"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GS where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpeciesLiveGS_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 299           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary
   StreamSave = points!Stream_Name
   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
     GeomorphIn = points!Geomorphic_Surface
   Else
     Geomorph = "None"
     GeomorphIn = "None"
   End If
   Point_Count = 0

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Or GeomorphIn <> Geomorph Then   ' Is it a new plot
       PointIndex = 0  ' yes - add in last point
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextAIndex  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextAIndex:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       ' *** End of plot processing ***
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Live_GS")
       ArrayIndex = 0
       Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for last plot
         If PlotArray(ArrayIndex, 1) > 0 Then
           WorkOutput.AddNew
           WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
           WorkOutput!Stream_Name = StreamSave
           WorkOutput!Geomorph = Geomorph
           WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
           WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
           WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           WorkOutput.Update  ' Write previous output record
           RecordCount = RecordCount + 1
         End If
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Geomorph = GeomorphIn
       ArrayIndex = 0
       Do Until ArrayIndex > 299    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
       Point_Count = 0
     End If
     If PointSave <> points!point Then  ' Is it a new point
     '  *** End of point processing ***
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) And points!alive Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Top = PointArray(PointIndex) Then  ' is the species the same
             Exit Do   ' Already have the species for this point
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Top  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And points!Surface_Alive And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Surface = PointArray(PointIndex) Then  ' is the species the same
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Surface  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipSpecies
       Else
         PointIndex = 0
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex) = points(SpeciesColumn)  ' set species
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipSpecies:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
   Loop
   ' End of file - add in last point
   PointIndex = 0
     Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
       ArrayIndex = 0
       Do Until ArrayIndex > 299
         If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
           If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           Else
             GoTo LastPlotArray  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
           Exit Do
         End If  ' end if for array slot open test
LastPlotArray:
         ArrayIndex = ArrayIndex + 1
       Loop
       PointIndex = PointIndex + 1
     Loop
     Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Live_GS")  ' Write last output record
     ArrayIndex = 0
     Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "
       If PlotArray(ArrayIndex, 1) > 0 Then
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!Geomorph = Geomorph
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
         WorkOutput.Update  ' Write previous output record
         RecordCount = RecordCount + 1
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_upd_Cover_Pct_Live_GS"   ' Update species names.
   DoCmd.SetWarnings True
Exit_CoverSpeciesLiveGS_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Live_GS."
    Exit Sub

Err_CoverSpeciesLiveGS_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpeciesLiveGS_Click
End Sub

Private Sub ButtonCoverSurface_Click()
On Error GoTo Err_CoverSurface_Click

  Dim strSQL As String
  Dim Surface As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim dblDivisor As Double
  Dim SpeciesColumn As String
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim PointArray(10) As Variant ' Array for surface features and disturbances at a point
  ' Feature hits per point array
  ' Column 1 feature
  Dim PlotArray(21, 1) As Variant ' Array for Surface features in a plot
  ' Species hits per plot array
  ' Column 1 is feature
  ' Column x, 0 is count

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Surface"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Cover_Surface where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSurface_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 9            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 20           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   Point_Count = 0
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - process last point in previous plot first
       PointIndex = 0
       Do Until PointIndex > 9 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 20
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex) Then  ' is this the correct feature
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextArrayEntry  ' Different feature - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set feature
             PlotArray(ArrayIndex, 1) = 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for feature processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Surface")
       ArrayIndex = 0
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' MsgBox PlotTotalL & " " & PlotTotalA
       Do Until ArrayIndex > 20 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "BL"
             WorkOutput!BL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BR"
             WorkOutput!BR = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BSC"
             WorkOutput!BSC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "CB"
             WorkOutput!CB = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "F"
             WorkOutput!f = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GC"
             WorkOutput!GC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GF"
             WorkOutput!GF = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "L"
             WorkOutput!L = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "NR"
             WorkOutput!NR = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SL"
             WorkOutput!SL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SW"
             WorkOutput!SW = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WA"
             WorkOutput!WA = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WD"
             WorkOutput!WD = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Ant"
             WorkOutput!Ant = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Auto"
             WorkOutput!Auto = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Bike"
             WorkOutput!Bike = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Camp"
             WorkOutput!Camp = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Flood"
             WorkOutput!Flood = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Graze"
             WorkOutput!Graze = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Other"
             WorkOutput!Other = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Trail"
             WorkOutput!Trail = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case Else
             MsgBox "Unknown Surface code " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Point_Count = 0
   PointIndex = 0
   Do Until PointIndex > 9            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 20           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
       PointLive = 0
       PointAll = 0
       PlotTotalA = 0
       PlotTotalL = 0
     End If
     If PointSave <> points!point Then  ' End of point - add lifeforms to plot array
       PointIndex = 0
       Do Until PointIndex > 9 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 20
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex) Then  ' is this the correct feature
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different feature - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set feature
             PlotArray(ArrayIndex, 1) = 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for feature processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 9            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Soil Surface first
     If Not IsNull(points!Surface) And Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       Do Until PointIndex > 9
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Surface = PointArray(PointIndex) Then  ' is the species the same
             GoTo NextSurface  ' Different species - go to next array entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Surface  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 5  ' Go through disturbance fields
       SpeciesColumn = "D" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Dist_Code FROM tlu_LP_Disturbance WHERE Dist_Code = '" & points(SpeciesColumn) & "'"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' Skip unknown disturbance code.
         End If
         Do Until PointIndex > 9
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next array entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex) = points(SpeciesColumn)  ' set feature
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   '  Process last point
       PointIndex = 0
       Do Until PointIndex > 9 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 20
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex) Then  ' is this the correct feature
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo LastArrayEntry  ' Different feature - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set feature
             PlotArray(ArrayIndex, 1) = 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
LastArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for feature processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Surface")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ArrayIndex = 0
       Do Until ArrayIndex > 19 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "BL"
             WorkOutput!BL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BR"
             WorkOutput!BR = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BSC"
             WorkOutput!BSC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "CB"
             WorkOutput!CB = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "F"
             WorkOutput!f = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GC"
             WorkOutput!GC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GF"
             WorkOutput!GF = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "L"
             WorkOutput!L = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "NR"
             WorkOutput!NR = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SL"
             WorkOutput!SL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SW"
             WorkOutput!SW = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WA"
             WorkOutput!WA = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WD"
             WorkOutput!WD = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Ant"
             WorkOutput!Ant = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Auto"
             WorkOutput!Auto = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Bike"
             WorkOutput!Bike = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Camp"
             WorkOutput!Camp = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Flood"
             WorkOutput!Flood = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Graze"
             WorkOutput!Graze = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Other"
             WorkOutput!Other = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Trail"
             WorkOutput!Trail = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case Else
             MsgBox "Unknown Surface code " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverSurface_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Surface."
    Exit Sub

Err_CoverSurface_Click:
    MsgBox Err.Description
    Resume Exit_CoverSurface_Click
End Sub

Private Sub ButtonCoverSurfaceGS_Click()
On Error GoTo Err_CoverSurface_Click

  Dim strSQL As String
  Dim Surface As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim dblDivisor As Double
  Dim SpeciesColumn As String
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim PointArray(5) As Variant ' Array for surface features at a point
  ' Feature hits per point array
  ' Column 1 feature
  Dim PlotArray(19, 1) As Variant ' Array for Surface features in a plot
  ' Species hits per plot array
  ' Column 1 is feature
  ' Column x, 0 is count

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Surface_GS"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Cover_Surface where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSurface_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 5            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 19           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
     GeomorphIn = points!Geomorphic_Surface
   Else
     Geomorph = "None"
     GeomorphIn = "None"
   End If
   Point_Count = 0
   Do Until points.EOF
     If (PlotSave <> points!Unit_Code & points!Plot_ID) Or (Geomorph <> GeomorphIn) Then  ' Check for new plot code
       ' New plot - process last point in previous plot first
       PointIndex = 0
       Do Until PointIndex > 5 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 19
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex) Then  ' is this the correct feature
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextArrayEntry  ' Different feature - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set feature
             PlotArray(ArrayIndex, 1) = 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for feature processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Surface_GS")
       ArrayIndex = 0
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' MsgBox PlotTotalL & " " & PlotTotalA
       Do Until ArrayIndex > 19 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "BL"
             WorkOutput!BL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BR"
             WorkOutput!BR = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BSC"
             WorkOutput!BSC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "CB"
             WorkOutput!CB = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "F"
             WorkOutput!f = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GC"
             WorkOutput!GC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GF"
             WorkOutput!GF = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "L"
             WorkOutput!L = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SL"
             WorkOutput!SL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SW"
             WorkOutput!SW = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WA"
             WorkOutput!f = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WD"
             WorkOutput!GC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Ant"
             WorkOutput!Ant = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Auto"
             WorkOutput!Auto = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Bike"
             WorkOutput!Bike = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Camp"
             WorkOutput!Camp = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Flood"
             WorkOutput!Flood = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Graze"
             WorkOutput!Graze = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Other"
             WorkOutput!Other = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Trail"
             WorkOutput!Trail = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case Else
             MsgBox "Unknown Surface code " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Geomorph = GeomorphIn
       Point_Count = 0
   PointIndex = 0
   Do Until PointIndex > 5            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 19           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
       PointLive = 0
       PointAll = 0
       PlotTotalA = 0
       PlotTotalL = 0
     End If
     If PointSave <> points!point Then  ' End of point - add lifeforms to plot array
       PointIndex = 0
       Do Until PointIndex > 5 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 19
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex) Then  ' is this the correct feature
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different feature - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set feature
             PlotArray(ArrayIndex, 1) = 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for feature processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 5            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Soil Surface first
     If Not IsNull(points!Surface) And points!Surface <> "NR" And Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       Do Until PointIndex > 5
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Surface = PointArray(PointIndex) Then  ' is the species the same
             GoTo NextSurface  ' Different species - go to next array entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Surface  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 5  ' Go through disturbance fields
       SpeciesColumn = "D" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Dist_Code FROM tlu_LP_Disturbance WHERE Dist_Code = '" & points(SpeciesColumn) & "'"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' Skip unknown disturbance code.
         End If
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next array entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex) = points(SpeciesColumn)  ' set feature
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
   Loop
   '  Process last point
       PointIndex = 0
       Do Until PointIndex > 5 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 19
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex) Then  ' is this the correct feature
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo LastArrayEntry  ' Different feature - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set feature
             PlotArray(ArrayIndex, 1) = 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
LastArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for feature processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Surface_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ArrayIndex = 0
       Do Until ArrayIndex > 19 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "BL"
             WorkOutput!BL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BR"
             WorkOutput!BR = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "BSC"
             WorkOutput!BSC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "CB"
             WorkOutput!CB = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "F"
             WorkOutput!f = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GC"
             WorkOutput!GC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "GF"
             WorkOutput!GF = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "L"
             WorkOutput!L = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SL"
             WorkOutput!SL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "SW"
             WorkOutput!SW = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WA"
             WorkOutput!f = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "WD"
             WorkOutput!GC = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Ant"
             WorkOutput!Ant = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Auto"
             WorkOutput!Auto = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Bike"
             WorkOutput!Bike = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Camp"
             WorkOutput!Camp = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Flood"
             WorkOutput!Flood = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Graze"
             WorkOutput!Graze = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Other"
             WorkOutput!Other = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case "Trail"
             WorkOutput!Trail = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           Case Else
             MsgBox "Unknown Surface code " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverSurface_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Surface_GS."
    Exit Sub

Err_CoverSurface_Click:
    MsgBox Err.Description
    Resume Exit_CoverSurface_Click
End Sub

Private Sub ButtonCoverWetland_Click()
On Error GoTo Err_CoverWetland_Click

  Dim strSQL As String
  Dim Wetland As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim dblDivisor As Double
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim PointLive As Byte
  Dim PointAll As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer
  Dim PointArray(12, 3) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 lifeform
  ' Column x,0 is alive flag
  ' Column x, 1 is dead flag
  Dim PlotArray(7, 3) As Variant ' Array for wetland status in a plot
  ' Species hits per plot array
  ' Column 1 is lifeform
  ' Column x, 0 is alive count
  ' Column x, 1 is total count

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Wetland"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_Species where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  
  
  ' strSQL = strSQL & " AND Plot_ID = 1 AND Transect = 2"
  
  
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverWetland_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex, 0) = " "
     PointArray(PointIndex, 1) = 0
     PointArray(PointIndex, 2) = 0
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 6           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     PlotArray(ArrayIndex, 2) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   PointLive = 0
   PointAll = 0
   PlotTotalA = 0
   PlotTotalL = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   Point_Count = 0
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - process last point in previous plot first
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 7
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextArrayEntry  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive

             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for lifeform processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Wetland")
       ArrayIndex = 0
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' MsgBox PlotTotalL & " " & PlotTotalA
       ' WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       ' WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       Do Until ArrayIndex > 6 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "FAC"
             WorkOutput!FACL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             ' MsgBox PlotArray(ArrayIndex, 1) & " FACL"
             WorkOutput!FACA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACU"
             WorkOutput!FACUL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             ' MsgBox PlotArray(ArrayIndex, 1) & " FACUL"
             WorkOutput!FACUA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
             ' MsgBox PlotArray(ArrayIndex, 2) & " FACUA"
           Case "FACW"
             WorkOutput!FACWL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             ' MsgBox PlotArray(ArrayIndex, 1) & " FACWL"
             WorkOutput!FACWA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
             ' MsgBox PlotArray(ArrayIndex, 2) & " FACWA"
           Case "OBL"
             WorkOutput!OBLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!OBLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "UPL"
             WorkOutput!UPLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!UPLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "CULT"
             WorkOutput!CULTL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!CULTA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case Else
             MsgBox "Unknown wetland status " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Point_Count = 0
       ArrayIndex = 0
       Do Until ArrayIndex > 6    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         PlotArray(ArrayIndex, 2) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
       PlotTotalA = 0
       PlotTotalL = 0
     End If
     If PointSave <> points!point Then  ' End of point - add lifeforms to plot array
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 6
         '  If PointArray(PointIndex, 0) = "FACU" Then
         '    MsgBox PointSave
         '  End If
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
SkipPointSpecies:
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Wetland_Code])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!Wetland_Code = "" Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       Wetland = SpeciesLU!Wetland_Code
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!alive Then
               PointArray(PointIndex, 1) = 1  ' Set alive flag
             Else
               PointArray(PointIndex, 2) = 1  ' set dead flag
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = Wetland  ' set species
           If points!alive Then
             PointArray(PointIndex, 1) = 1  ' count it as alive
           Else
             PointArray(PointIndex, 2) = 1  ' set dead flag
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Surface & "' AND NOT IsNull([Wetland_Code])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!Wetland_Code = "" Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       Wetland = SpeciesLU!Wetland_Code
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!Surface_Alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!Surface_Alive Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = Wetland  ' set species
           If points!Surface_Alive Then
             PointArray(PointIndex, 1) = 1  ' flag it as alive
           Else
             PointArray(PointIndex, 2) = 1  ' flag it as dead
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Wetland_Code])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         If SpeciesLU!Wetland_Code = "" Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         Wetland = SpeciesLU!Wetland_Code
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
         Do Until PointIndex > 11
           If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
             If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
               If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
                 PointArray(PointIndex, 1) = 1  ' flag it as alive
               Else
                 PointArray(PointIndex, 2) = 1  ' flag it as dead
               End If ' end if for alive test
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex, 0) = Wetland  ' set species
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   '  Process last point
   PointIndex = 0
   Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
     ArrayIndex = 0
     Do Until ArrayIndex > 7
       If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
         If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
           If PointArray(PointIndex, 1) = 1 Then
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
           End If
           PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
           Exit Do
         Else
           GoTo LastPlotArray  ' Different species - go to next entry
         End If  ' End if for species compare
       Else
         PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
         If PointArray(PointIndex, 1) = 1 Then
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
         End If
         PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
         Exit Do
       End If  ' end if for array slot open test
LastPlotArray:
       ArrayIndex = ArrayIndex + 1
     Loop
     PointIndex = PointIndex + 1
   Loop

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Wetland")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ArrayIndex = 0
       ' WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       ' WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       Do Until ArrayIndex > 6 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "FAC"
             WorkOutput!FACL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACU"
            '  MsgBox PlotArray(ArrayIndex, 1) & " FACUL"
            '  MsgBox PlotArray(ArrayIndex, 2) & " FACUA"
             WorkOutput!FACUL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACUA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACW"
             WorkOutput!FACWL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACWA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "OBL"
             WorkOutput!OBLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!OBLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "UPL"
             WorkOutput!UPLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!UPLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "CULT"
             WorkOutput!CULTL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!CULTA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case Else
             MsgBox "Unknown wetland status " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverWetland_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Wetland."
    Exit Sub

Err_CoverWetland_Click:
    MsgBox Err.Description
    Resume Exit_CoverWetland_Click
End Sub

Private Sub ButtonCoverWetlandGL_Click()
On Error GoTo Err_CoverWetland_Click

  Dim strSQL As String
  Dim Wetland As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim dblDivisor As Double
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim PointLive As Byte
  Dim PointAll As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer
  Dim PointArray(12, 3) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 lifeform
  ' Column x,0 is alive flag
  ' Column x, 1 is dead flag
  Dim PlotArray(7, 3) As Variant ' Array for wetland status in a plot
  ' Species hits per plot array
  ' Column 1 is lifeform
  ' Column x, 0 is alive count
  ' Column x, 1 is total count

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_GL_Cover_Pct_Wetland"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GL where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverWetland_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex, 0) = " "
     PointArray(PointIndex, 1) = 0
     PointArray(PointIndex, 2) = 0
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 6           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     PlotArray(ArrayIndex, 2) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   PointLive = 0
   PointAll = 0
   PlotTotalA = 0
   PlotTotalL = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - process last point in previous plot first
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 7
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextArrayEntry  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive

             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for lifeform processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Wetland")
       ArrayIndex = 0
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' MsgBox PlotTotalL & " " & PlotTotalA
       ' WorkOutput!Total_Live = (PlotTotalL / 560) * 100
       ' WorkOutput!Total_Cover = (PlotTotalA / 560) * 100
       Do Until ArrayIndex > 6 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "FAC"
             WorkOutput!FACL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!FACA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "FACU"
             WorkOutput!FACUL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!FACUA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "FACW"
             WorkOutput!FACWL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!FACWA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "OBL"
             WorkOutput!OBLL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!OBLA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "UPL"
             WorkOutput!UPLL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!UPLA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "CULT"
             WorkOutput!CULTL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!CULTA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case Else
             MsgBox "Unknown wetland status " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 6    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         PlotArray(ArrayIndex, 2) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
       PlotTotalA = 0
       PlotTotalL = 0
     End If
     If PointSave <> points!point Then  ' End of point - add lifeforms to plot array
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 5
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
SkipPointSpecies:
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Wetland_Code])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!Wetland_Code = "" Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       Wetland = SpeciesLU!Wetland_Code
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!alive Then
               PointArray(PointIndex, 1) = 1  ' Set alive flag
             Else
               PointArray(PointIndex, 2) = 1  ' set dead flag
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = Wetland  ' set species
           If points!alive Then
             PointArray(PointIndex, 1) = 1  ' count it as alive
           Else
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface removed 4/18/13 RD.

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Wetland_Code])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         If SpeciesLU!Wetland_Code = "" Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         Wetland = SpeciesLU!Wetland_Code
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
         Do Until PointIndex > 11
           If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
             If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
               If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
                 PointArray(PointIndex, 1) = 1  ' flag it as alive
               Else
                 PointArray(PointIndex, 2) = 1  ' flag it as dead
               End If ' end if for alive test
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex, 0) = Wetland  ' set species
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     points.MoveNext
   Loop
   '  Process last point
   PointIndex = 0
   Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
     ArrayIndex = 0
     Do Until ArrayIndex > 7
       If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
         If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
           If PointArray(PointIndex, 1) = 1 Then
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
           End If
           PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
           Exit Do
         Else
           GoTo LastPlotArray  ' Different species - go to next entry
         End If  ' End if for species compare
       Else
         PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
         If PointArray(PointIndex, 1) = 1 Then
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
         End If
         PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
         Exit Do
       End If  ' end if for array slot open test
LastPlotArray:
       ArrayIndex = ArrayIndex + 1
     Loop
     PointIndex = PointIndex + 1
   Loop

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Wetland")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ArrayIndex = 0
       ' WorkOutput!Total_Live = (PlotTotalL / 560) * 100
       ' WorkOutput!Total_Cover = (PlotTotalA / 560) * 100
       Do Until ArrayIndex > 6 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "FAC"
             WorkOutput!FACL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!FACA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "FACU"
             WorkOutput!FACUL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!FACUA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "FACW"
             WorkOutput!FACWL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!FACWA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "OBL"
             WorkOutput!OBLL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!OBLA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "UPL"
             WorkOutput!UPLL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!UPLA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case "CULT"
             WorkOutput!CULTL = (PlotArray(ArrayIndex, 1) / 560) * 100
             WorkOutput!CULTA = (PlotArray(ArrayIndex, 2) / 560) * 100
           Case Else
             MsgBox "Unknown wetland status " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverWetland_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_GL_Cover_Pct_Wetland."
    Exit Sub

Err_CoverWetland_Click:
    MsgBox Err.Description
    Resume Exit_CoverWetland_Click
End Sub

Private Sub ButtonCoverWetlandGS_Click()
On Error GoTo Err_CoverWetland_Click

  Dim strSQL As String
  Dim Wetland As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim dblDivisor As Double
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim PointLive As Byte
  Dim PointAll As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer
  Dim PointArray(12, 3) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 lifeform
  ' Column x,0 is alive flag
  ' Column x, 1 is dead flag
  Dim PlotArray(7, 3) As Variant ' Array for wetland status in a plot
  ' Species hits per plot array
  ' Column 1 is lifeform
  ' Column x, 0 is alive count
  ' Column x, 1 is total count

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Wetland_GS"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GS where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverWetland_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex, 0) = " "
     PointArray(PointIndex, 1) = 0
     PointArray(PointIndex, 2) = 0
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 6           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     PlotArray(ArrayIndex, 2) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   PointLive = 0
   PointAll = 0
   PlotTotalA = 0
   PlotTotalL = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
     GeomorphIn = points!Geomorphic_Surface
   Else
     Geomorph = "None"
     GeomorphIn = "None"
   End If
   Point_Count = 0
   Do Until points.EOF
     If (PlotSave <> points!Unit_Code & points!Plot_ID) Or (Geomorph <> GeomorphIn) Then  ' Check for new plot code
       ' New plot - process last point in previous plot first
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 7
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextArrayEntry  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive

             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for lifeform processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Wetland_GS")
       ArrayIndex = 0
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ' MsgBox PlotTotalL & " " & PlotTotalA
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       Do Until ArrayIndex > 6 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "FAC"
             WorkOutput!FACL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACU"
             WorkOutput!FACUL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACUA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACW"
             WorkOutput!FACWL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACWA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "OBL"
             WorkOutput!OBLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!OBLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "UPL"
             WorkOutput!UPLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!UPLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "CULT"
             WorkOutput!CULTL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!CULTA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case Else
             MsgBox "Unknown wetland status " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       Geomorph = GeomorphIn
       StreamSave = points!Stream_Name
       Point_Count = 0
       ArrayIndex = 0
       Do Until ArrayIndex > 6    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         PlotArray(ArrayIndex, 2) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
       PlotTotalA = 0
       PlotTotalL = 0
     End If
     If PointSave <> points!point Then  ' End of point - add lifeforms to plot array
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 5
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
SkipPointSpecies:
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Wetland_Code])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!Wetland_Code = "" Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       Wetland = SpeciesLU!Wetland_Code
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!alive Then
               PointArray(PointIndex, 1) = 1  ' Set alive flag
             Else
               PointArray(PointIndex, 2) = 1  ' set dead flag
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = Wetland  ' set species
           If points!alive Then
             PointArray(PointIndex, 1) = 1  ' count it as alive
           Else
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Surface & "' AND NOT IsNull([Wetland_Code])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!Wetland_Code = "" Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       Wetland = SpeciesLU!Wetland_Code
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!Surface_Alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!Surface_Alive Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = Wetland  ' set species
           If points!Surface_Alive Then
             PointArray(PointIndex, 1) = 1  ' flag it as alive
           Else
             PointArray(PointIndex, 2) = 1  ' flag it as dead
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Wetland_Code FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Wetland_Code])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         If SpeciesLU!Wetland_Code = "" Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         Wetland = SpeciesLU!Wetland_Code
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
         Do Until PointIndex > 11
           If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
             If Wetland = PointArray(PointIndex, 0) Then  ' is the species the same
               If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
                 PointArray(PointIndex, 1) = 1  ' flag it as alive
               Else
                 PointArray(PointIndex, 2) = 1  ' flag it as dead
               End If ' end if for alive test
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex, 0) = Wetland  ' set species
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     points.MoveNext
     Point_Count = Point_Count + 1
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
   Loop
   '  Process last point
   PointIndex = 0
   Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
     ArrayIndex = 0
     Do Until ArrayIndex > 7
       If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
         If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
           If PointArray(PointIndex, 1) = 1 Then
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
           End If
           PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
           Exit Do
         Else
           GoTo LastPlotArray  ' Different species - go to next entry
         End If  ' End if for species compare
       Else
         PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
         If PointArray(PointIndex, 1) = 1 Then
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
         End If
         PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
         Exit Do
       End If  ' end if for array slot open test
LastPlotArray:
       ArrayIndex = ArrayIndex + 1
     Loop
     PointIndex = PointIndex + 1
   Loop

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Wetland_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ArrayIndex = 0
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       Do Until ArrayIndex > 6 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "FAC"
             WorkOutput!FACL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACU"
             WorkOutput!FACUL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACUA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "FACW"
             WorkOutput!FACWL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FACWA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "OBL"
             WorkOutput!OBLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!OBLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "UPL"
             WorkOutput!UPLL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!UPLA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "CULT"
             WorkOutput!CULTL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!CULTA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case Else
             MsgBox "Unknown wetland status " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverWetland_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Wetland_GS."
    Exit Sub

Err_CoverWetland_Click:
    MsgBox Err.Description
    Resume Exit_CoverWetland_Click
End Sub

Private Sub ButtonExoticFrequency_Click()
On Error GoTo Err_ExoticFrequency_Click

  Dim strSQL As String
  Dim Surface As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim LatinSave As String
  Dim CommonSave As String
  Dim PointSave As Double
  Dim Quad_Count As Integer
  Dim SpeciesColumn As String
  Dim LCIndex As Integer
  Dim ArrayIndex As Integer
  Dim PlotArray(25, 1) As Variant ' Array for species in a reach
  ' Species hits per plot array
  ' Column 1 is species
  ' Column x, 0 is count of quads

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Exotic_Frequency"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Exotic_Freq_Sum where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Species"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid Exotic records found."
     points.Close
     Set points = Nothing
     GoTo Exit_ExoticFrequency_Click
   End If
   ArrayIndex = 0
   Do Until ArrayIndex > 24           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   Quad_Count = 0
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' Now write reach record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Exotic_Frequency_Summary")
       ArrayIndex = 0
       Do Until ArrayIndex > 24 Or PlotArray(ArrayIndex, 0) = " "  ' Write species totals for plot
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!ExoticFrequency = (PlotArray(ArrayIndex, 1) / Quad_Count) * 100
         ArrayIndex = ArrayIndex + 1
         WorkOutput.Update  ' Write plot record
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Quad_Count = 0
       ArrayIndex = 0
       Do Until ArrayIndex > 24           ' Initialize plot array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
     End If
     Quad_Count = Quad_Count + 20
     ArrayIndex = 0
     Do Until ArrayIndex > 24
       If PlotArray(ArrayIndex, 0) = " " Then
         PlotArray(ArrayIndex, 0) = points!Species
       End If
       If PlotArray(ArrayIndex, 0) = points!Species Then
         ' Accumulate quad counts for this species
         LCIndex = 0   ' Initialize index
         Do Until LCIndex > 95  ' Go through disturbance fields
           SpeciesColumn = "M" & LCIndex ' Get the quad field
           If points(SpeciesColumn) Then
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1    ' Count quadrat hit
           End If
           LCIndex = LCIndex + 5
         Loop   ' Loop for quadrats in a record
         Exit Do
       End If
       ArrayIndex = ArrayIndex + 1
     Loop  ' Loop for plot array entry
     points.MoveNext
   Loop   ' Loop for Points input file

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Exotic_Frequency_Summary")
       ArrayIndex = 0
       Do Until ArrayIndex > 24 Or PlotArray(ArrayIndex, 0) = " "  ' Write species totals for plot
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!ExoticFrequency = (PlotArray(ArrayIndex, 1) / Quad_Count) * 100
         ArrayIndex = ArrayIndex + 1
         WorkOutput.Update  ' Write plot record
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       DoCmd.SetWarnings False
       DoCmd.OpenQuery "qry_upd_Exotic_Frequency_Summary"   ' Update species names.
       DoCmd.SetWarnings True
Exit_ExoticFrequency_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Exotic_Frequency_Summary."
    Exit Sub

Err_ExoticFrequency_Click:
    MsgBox Err.Description
    Resume Exit_ExoticFrequency_Click
End Sub

Private Sub ButtonGLCoverAll_Click()
On Error GoTo Err_CoverSpeciesAll_Click

  Dim strSQL As String
  Dim lifeForm As Variant
  Dim SpeciesColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim ArrayEnd As Integer
  Dim PointArray(12) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 is species code
  Dim PlotArray(300, 1) As Variant ' Array for species in a plot
  ' Species hits per plot array
  ' Column 1 is species code
  ' Column 2 is alive count
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_GL_Cover_Pct_All"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GL where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpeciesAll_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 299           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary
   StreamSave = points!Stream_Name

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Is it a new plot
       PointIndex = 0  ' yes - add in last point
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextAIndex  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextAIndex:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       ' *** End of plot processing ***
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_All")
       ArrayIndex = 0
       Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for last plot
         If PlotArray(ArrayIndex, 1) > 0 Then
           WorkOutput.AddNew
           WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
           WorkOutput!Stream_Name = StreamSave
           WorkOutput!Visit_Year = Me!Visit_Date
           WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
           WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
           WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / 560) * 100
           WorkOutput.Update  ' Write previous output record
           RecordCount = RecordCount + 1
         End If
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 299    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If
     If PointSave <> points!point Then  ' Is it a new point
     '  *** End of point processing ***
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Top = PointArray(PointIndex) Then  ' is the species the same
             Exit Do   ' Already have the species for this point
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Top  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check

     '  Soil Surface removed 4/18/2013 RD.
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipSpecies
       Else
         PointIndex = 0
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex) = points(SpeciesColumn)  ' set species
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipSpecies:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
   Loop
   ' End of file - add in last point
   PointIndex = 0
     Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
       ArrayIndex = 0
       Do Until ArrayIndex > 299
         If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
           If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           Else
             GoTo LastPlotArray  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
           Exit Do
         End If  ' end if for array slot open test
LastPlotArray:
         ArrayIndex = ArrayIndex + 1
       Loop
       PointIndex = PointIndex + 1
     Loop
     Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_All")  ' Write last output record
     ArrayIndex = 0
     Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "
       If PlotArray(ArrayIndex, 1) > 0 Then
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / 560) * 100
         WorkOutput.Update  ' Write previous output record
         RecordCount = RecordCount + 1
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_upd_GL_Cover_Pct_All"   ' Update species names.
   DoCmd.SetWarnings True
Exit_CoverSpeciesAll_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_GL_Cover_Pct_All."
    Exit Sub

Err_CoverSpeciesAll_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpeciesAll_Click
End Sub

Private Sub ButtonGLTotalCover_Click()
On Error GoTo Err_GLTotalCover_Click

  Dim strSQL As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Point_Count As Integer
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointAll As Byte
  Dim PointLive As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_GL_Cover_Pct_Totals"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GL where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
'   strSQL = strSQL & " AND Plot_ID = 51"
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_GLTotalCover_Click
   End If
   PlotTotalA = 0
   PlotTotalL = 0
   PointAll = 0
   PointLive = 0
   Point_Count = 0
   points.MoveFirst
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Totals")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       PlotTotalA = 0
       PlotTotalL = 0
       Point_Count = 0
     End If
     
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
     End If  ' End if for null top check

     '  Soil Surface removed 4/18/13 RD.
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipEntry
       Else
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
       End If  ' End if for null lower canopy check
SkipEntry:
       LCIndex = LCIndex + 1
     Loop
   ' Now update plot totals
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     PointLive = 0  ' Clear hit indicators
     PointAll = 0   ' For the next point
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Totals")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_GLTotalCover_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_GL_Cover_Pct_Totals."
    Exit Sub

Err_GLTotalCover_Click:
    MsgBox Err.Description
    Resume Exit_GLTotalCover_Click
End Sub

Private Sub ButtonPebbleCount_Click()
On Error GoTo Err_PebbleCount_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Pebbles As DAO.Recordset
  Dim PebbleDA As DAO.Recordset
  Dim PlotSave As Variant
  Dim NameSave As String
  Dim FieldIndex As Integer
  Dim ArrayIndex As Integer
  Dim intDivisor As Integer
  Dim intEntry As Integer
  Dim PlotArray(6) As Integer
  ' Array for pebble count accumulators
  ' Index 0 is total Pebbles < .2mm
  ' Index 1 is total Pebbles .2-6.3mm
  ' Index 2 is total Pebbles 6.4-25.5mm
  ' Index 3 is total Pebbles 25.6-409.6mm
  ' Index 4 is total Pebbles >409.6mm
  ' Index 5 is total not recorded

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Pebble_Summary"
  DoCmd.OpenQuery "qry_Clear_Pebbles"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Pebble_Count_Summary where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
'  strSQL = strSQL & " and Plot_ID = 1"
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID"
  DoCmd.Hourglass True
  Set db = CurrentDb
  ' Load work table
   Set Pebbles = db.OpenRecordset(strSQL)
   If Pebbles.EOF Then
     MsgBox "No valid pebble records found."
     Pebbles.Close
     Set Pebbles = Nothing
     GoTo Exit_PebbleCount_Click
   End If
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Pebbles")
   Do Until Pebbles.EOF
     FieldIndex = 6   ' Set Field Index to seventh field
     Do Until FieldIndex > 65  ' Write an output record for each of the count fields
       If Not IsNull(Pebbles.Fields(FieldIndex)) Then
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Pebbles!Unit_Code  ' Set unit code
         WorkOutput!Stream_Name = Pebbles!Stream_Name
         WorkOutput!Plot_ID = Pebbles!Plot_ID  ' Set plot ID
         WorkOutput!Pebble_Size = Pebbles.Fields(FieldIndex)
         WorkOutput.Update  ' Write plot record
       End If
       FieldIndex = FieldIndex + 1
     Loop
     Pebbles.MoveNext
   Loop
   WorkOutput.Close
   Set WorkOutput = Nothing
   Pebbles.Close
   Set Pebbles = Nothing

   ArrayIndex = 0
   Do Until ArrayIndex > 5    ' clear array
     PlotArray(ArrayIndex) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   strSQL = "SELECT * FROM tbl_wrk_Pebbles ORDER BY Unit_Code, Plot_ID, Pebble_Size"
   Set Pebbles = db.OpenRecordset(strSQL)
   Pebbles.MoveFirst
   PlotSave = Pebbles!Unit_Code & Pebbles!Plot_ID     ' Save necessary fields
   NameSave = Pebbles!Stream_Name
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Pebble_Summary")
   Do Until Pebbles.EOF
     If (PlotSave <> Pebbles!Unit_Code & Pebbles!Plot_ID) Then  ' Test for new plot
         intDivisor = 420 - PlotArray(5) ' Subtract out non-recorded entries
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = NameSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Plot_ID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!Pct_2 = (PlotArray(0) / intDivisor) * 100
         WorkOutput!Pct_6 = (PlotArray(1) / intDivisor) * 100
         WorkOutput!Pct_25 = (PlotArray(2) / intDivisor) * 100
         WorkOutput!Pct_409 = (PlotArray(3) / intDivisor) * 100
         WorkOutput!Pct_Over_409 = (PlotArray(4) / intDivisor) * 100
         WorkOutput!Count_NR = PlotArray(5)
         strSQL = "SELECT * FROM tbl_wrk_Pebbles WHERE Pebble_Size <> -999 AND Unit_Code = '" & Left(PlotSave, 4) & "' AND Plot_ID = " & Right(PlotSave, Len(PlotSave) - 4)
         strSQL = strSQL & " ORDER BY Pebble_Size"
         Set PebbleDA = db.OpenRecordset(strSQL)
         PebbleDA.MoveFirst
         If PlotArray(5) = 0 Then
           PebbleDA.Move 209
         Else
           intEntry = Int(intDivisor * 0.5)
           PebbleDA.Move intEntry
         End If
         If Not PebbleDA.EOF Then
           WorkOutput!D50 = PebbleDA!Pebble_Size
         End If
         PebbleDA.MoveFirst
         If PlotArray(5) = 0 Then
           PebbleDA.Move 352
         Else
           intEntry = Int(intDivisor * 0.84)
           PebbleDA.Move intEntry
         End If
         If Not PebbleDA.EOF Then
           WorkOutput!D84 = PebbleDA!Pebble_Size
         End If
         PebbleDA.Close
         Set PebbleDA = Nothing
         WorkOutput.Update  ' Write plot record
       PlotSave = Pebbles!Unit_Code & Pebbles!Plot_ID
       NameSave = Pebbles!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 5    ' clear array
         PlotArray(ArrayIndex) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
     End If ' End if for new species test
     Select Case Pebbles!Pebble_Size
       Case Is < 0
         PlotArray(5) = PlotArray(5) + 1  ' Count non-recorded entries
       Case 0 To 0.1
         PlotArray(0) = PlotArray(0) + 1  ' Count pebble size classes
       Case 0.2 To 6.3
         PlotArray(1) = PlotArray(1) + 1
       Case 6.4 To 25.5
         PlotArray(2) = PlotArray(2) + 1
       Case 25.6 To 409.6
         PlotArray(3) = PlotArray(3) + 1
       Case Else
         PlotArray(4) = PlotArray(4) + 1
     End Select
     Pebbles.MoveNext
   Loop
   ' Write last output record
         intDivisor = 420 - PlotArray(5)
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = NameSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Plot_ID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!Pct_2 = (PlotArray(0) / intDivisor) * 100
         WorkOutput!Pct_6 = (PlotArray(1) / intDivisor) * 100
         WorkOutput!Pct_25 = (PlotArray(2) / intDivisor) * 100
         WorkOutput!Pct_409 = (PlotArray(3) / intDivisor) * 100
         WorkOutput!Pct_Over_409 = (PlotArray(4) / intDivisor) * 100
         WorkOutput!Count_NR = PlotArray(5)
         strSQL = "SELECT * FROM tbl_wrk_Pebbles WHERE Pebble_Size <> -999 AND Unit_Code = '" & Left(PlotSave, 4) & "' AND Plot_ID = " & Right(PlotSave, Len(PlotSave) - 4)
         strSQL = strSQL & " ORDER BY Pebble_Size"
         Set PebbleDA = db.OpenRecordset(strSQL)
         PebbleDA.MoveFirst
         PebbleDA.Move 209
         If Not PebbleDA.EOF Then
           WorkOutput!D50 = PebbleDA!Pebble_Size
         End If
         PebbleDA.MoveFirst
         PebbleDA.Move 352
         If Not PebbleDA.EOF Then
           WorkOutput!D84 = PebbleDA!Pebble_Size
         End If
         PebbleDA.Close
         Set PebbleDA = Nothing
         WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   Pebbles.Close
   Set Pebbles = Nothing
Exit_PebbleCount_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Pebble_Summary."
    Exit Sub

Err_PebbleCount_Click:
    MsgBox Err.Description
    Resume Exit_PebbleCount_Click

End Sub

Private Sub ButtonPercentCoverSpecies_Click()
On Error GoTo Err_CoverSpecies_Click

  Dim strSQL As String
  Dim lifeForm As Variant
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim ACount As Integer
  Dim Point_Count As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim ArrayEnd As Integer
  Dim PointArray(12) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 is species code
  Dim PlotArray(300, 1) As Variant ' Array for species in a plot
  ' Species hits per plot array
  ' Column 1 is species code
  ' Column 2 is alive count
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Live"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_Species where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpecies_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 299           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary
   StreamSave = points!Stream_Name
   Point_Count = 0

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Is it a new plot
       PointIndex = 0  ' yes - add in last point
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextAIndex  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextAIndex:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       ' *** End of plot processing ***
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Live")
       ArrayIndex = 0
       Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for last plot
         If PlotArray(ArrayIndex, 1) > 0 Then
           WorkOutput.AddNew
           WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
           WorkOutput!Visit_Year = Me!Visit_Date
           WorkOutput!Stream_Name = StreamSave
           WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
           WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
           WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           WorkOutput.Update  ' Write previous output record
           RecordCount = RecordCount + 1
         End If
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 299    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
       Point_Count = 0
     End If
     If PointSave <> points!point Then  ' Is it a new point
     '  *** End of point processing ***
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) And points!alive Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Top = PointArray(PointIndex) Then  ' is the species the same
             Exit Do   ' Already have the species for this point
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Top  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And points!Surface_Alive And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Surface = PointArray(PointIndex) Then  ' is the species the same
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Surface  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipSpecies
       Else
         PointIndex = 0
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex) = points(SpeciesColumn)  ' set species
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipSpecies:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   ' End of file - add in last point
   PointIndex = 0
     Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
       ArrayIndex = 0
       Do Until ArrayIndex > 299
         If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
           If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           Else
             GoTo LastPlotArray  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
           Exit Do
         End If  ' end if for array slot open test
LastPlotArray:
         ArrayIndex = ArrayIndex + 1
       Loop
       PointIndex = PointIndex + 1
     Loop
     Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Live")  ' Write last output record
     ArrayIndex = 0
     Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "
       If PlotArray(ArrayIndex, 1) > 0 Then
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
         WorkOutput.Update  ' Write previous output record
         RecordCount = RecordCount + 1
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
   DoCmd.SetWarnings False
'   DoCmd.OpenQuery "qry_del_Cover_Pct_Live_SS"   ' Delete any soil surface codes that may have been picked up in lower canopy.
   DoCmd.OpenQuery "qry_upd_Cover_Pct_Live"   ' Update species names.
   DoCmd.SetWarnings True
Exit_CoverSpecies_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Live."
    Exit Sub

Err_CoverSpecies_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpecies_Click
End Sub

Private Sub ButtonPointHit_Click()
On Error GoTo Err_ButtonPointHit_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim Plot_Save As Integer
  Dim Unit_Save As String
  Dim Stream_Name As String
  Dim reply As Integer
  Dim Species_Total As Integer
  Dim Species_TotalE As Integer
  Dim Tree_Total As Integer
  Dim Tree_TotalE As Integer
  Dim Shrub_Total As Integer
  Dim Shrub_TotalE As Integer
  Dim Perennial_Total As Integer
  Dim Perennial_TotalE As Integer
  Dim Annual_Total As Integer
  Dim Annual_TotalE As Integer
  Dim Forb_Total As Integer
  Dim Forb_TotalE As Integer
  Dim Fern_Total As Integer
  Dim Fern_TotalE As Integer
  Dim Vine_Total As Integer
  Dim Vine_TotalE As Integer
  
  reply = MsgBox("You must run All Species by Reach first.", vbOKCancel, "Richness by Species")
  If reply = vbCancel Then
    Exit Sub
  End If
  
  DoCmd.Hourglass True
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_SR_Richness"
  DoCmd.SetWarnings True
  
  Set db = CurrentDb
  RecordCount = 0
  ' Get species info
   Set WorkOutput = db.OpenRecordset("tbl_wrk_SR_Richness")
   Set points = db.OpenRecordset("qry_SR_Species_Lifeform")
   If points.EOF Then
     MsgBox "Species by plot table is empty."
     GoTo Exit_ButtonPointHit_Click:
   End If
   points.MoveFirst
   Unit_Save = points!Unit_Code
   Stream_Name = points!Stream_Name
   Plot_Save = points!Plot_ID
   Do Until points.EOF  ' Load lifeform breakdown into work table.
     If points!Unit_Code <> Unit_Save Or points!Plot_ID <> Plot_Save Then
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Unit_Save
         WorkOutput!Plot_ID = Plot_Save
         WorkOutput!Stream_Name = Stream_Name
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!TotalA = Species_Total
         WorkOutput!TreeA = Tree_Total
         WorkOutput!ShrubA = Shrub_Total
         WorkOutput!PerennialGrassA = Perennial_Total
         WorkOutput!AnnualGrassA = Annual_Total
         WorkOutput!Forb_HerbA = Forb_Total
         WorkOutput!FernA = Fern_Total
         WorkOutput!VineA = Vine_Total
         WorkOutput!TotalE = Species_TotalE
         WorkOutput!TreeE = Tree_TotalE
         WorkOutput!ShrubE = Shrub_TotalE
         WorkOutput!PerennialGrassE = Perennial_TotalE
         WorkOutput!AnnualGrassE = Annual_TotalE
         WorkOutput!Forb_HerbE = Forb_TotalE
         WorkOutput!FernE = Fern_TotalE
         WorkOutput!VineE = Vine_TotalE
         WorkOutput.Update
         RecordCount = RecordCount + 1
         Unit_Save = points!Unit_Code
         Stream_Name = points!Stream_Name
         Plot_Save = points!Plot_ID
         Species_Total = 0
         Tree_Total = 0
         Shrub_Total = 0
         Perennial_Total = 0
         Annual_Total = 0
         Forb_Total = 0
         Fern_Total = 0
         Vine_Total = 0
         Species_TotalE = 0
         Tree_TotalE = 0
         Shrub_TotalE = 0
         Perennial_TotalE = 0
         Annual_TotalE = 0
         Forb_TotalE = 0
         Fern_TotalE = 0
         Vine_TotalE = 0
      End If
      Species_Total = Species_Total + 1
      If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
        Species_TotalE = Species_TotalE + 1
      End If
      Select Case points!lifeForm
        Case "DwarfShrub", "Shrub"
          Shrub_Total = Shrub_Total + 1
          If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
            Shrub_TotalE = Shrub_TotalE + 1
          End If
        Case "forb"
          Forb_Total = Forb_Total + 1
          If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
            Forb_TotalE = Forb_TotalE + 1
          End If
        Case "Tree"
          Tree_Total = Tree_Total + 1
          If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
            Tree_TotalE = Tree_TotalE + 1
          End If
        Case "Graminoid"
          If points!Duration = "Perennial" Then
            Perennial_Total = Perennial_Total + 1
            If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
              Perennial_TotalE = Perennial_TotalE + 1
            End If
          Else
            Annual_Total = Annual_Total + 1
            If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
              Annual_TotalE = Annual_TotalE + 1
            End If
          End If
        Case "Fern"
          Fern_Total = Fern_Total + 1
          If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
            Fern_TotalE = Fern_TotalE + 1
          End If
        Case "Vine"
          Vine_Total = Vine_Total + 1
          If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
            Vine_TotalE = Vine_TotalE + 1
          End If
        Case Else
          MsgBox "Unrecognized lifeform " & points!lifeForm
      End Select
     points.MoveNext
   Loop
   ' Write last record
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Unit_Save
         WorkOutput!Plot_ID = Plot_Save
         WorkOutput!Stream_Name = Stream_Name
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!TotalA = Species_Total
         WorkOutput!TreeA = Tree_Total
         WorkOutput!ShrubA = Shrub_Total
         WorkOutput!PerennialGrassA = Perennial_Total
         WorkOutput!AnnualGrassA = Annual_Total
         WorkOutput!Forb_HerbA = Forb_Total
         WorkOutput!FernA = Fern_Total
         WorkOutput!VineA = Vine_Total
         WorkOutput!TotalE = Species_TotalE
         WorkOutput!TreeE = Tree_TotalE
         WorkOutput!ShrubE = Shrub_TotalE
         WorkOutput!PerennialGrassE = Perennial_TotalE
         WorkOutput!AnnualGrassE = Annual_TotalE
         WorkOutput!Forb_HerbE = Forb_TotalE
         WorkOutput!FernE = Fern_TotalE
         WorkOutput!VineE = Vine_TotalE
         WorkOutput.Update
         RecordCount = RecordCount + 1
   points.Close
   Set points = Nothing
   WorkOutput.Close
   Set WorkOutput = Nothing


Exit_ButtonPointHit_Click:
    DoCmd.Hourglass False
    MsgBox RecordCount & " records written.  Results are in tbl_wrk_SR_Richness."
    Exit Sub

Err_ButtonPointHit_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPointHit_Click
  
End Sub

Private Sub ButtonTotalCoverGS_Click()
On Error GoTo Err_TotalCoverGS_Click

  Dim strSQL As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim Point_Count As Integer
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointAll As Byte
  Dim PointLive As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Totals_GS"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GS where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
'   strSQL = strSQL & " AND Plot_ID = 51"
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_TotalCoverGS_Click
   End If
   PlotTotalA = 0
   PlotTotalL = 0
   PointAll = 0
   PointLive = 0
   Point_Count = 0
   points.MoveFirst
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name

   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
   Else
     Geomorph = "None"
   End If
   GeomorphIn = Geomorph
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Or Geomorph <> GeomorphIn Then  ' Check for new plot code
       ' New plot - write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Totals_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Total_Live = PlotTotalL / Point_Count
       WorkOutput!Total_Cover = PlotTotalA / Point_Count
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       PlotTotalA = 0
       PlotTotalL = 0
       Point_Count = 0
       Geomorph = GeomorphIn
     End If
     
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And points!Surface <> "" And points!Surface <> " " And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       If points!Surface_Alive Then
         PointLive = 1
       End If
       PointAll = 1
     End If  ' End if for null soil surface check

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipEntry
       Else
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
       End If  ' End if for null lower canopy check
SkipEntry:
       LCIndex = LCIndex + 1
     Loop
   ' Now update plot totals
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     PointLive = 0  ' Clear hit indicators
     PointAll = 0   ' For the next point
     points.MoveNext
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
     Point_Count = Point_Count + 1
   Loop
   ' Output last plot record
    '   MsgBox "A=" & [PlotTotalL] & "  T=" & [PlotTotalA]
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Totals_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Total_Live = PlotTotalL / Point_Count
       WorkOutput!Total_Cover = PlotTotalA / Point_Count
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_TotalCoverGS_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Totals_GS."
    Exit Sub

Err_TotalCoverGS_Click:
    MsgBox Err.Description
    Resume Exit_TotalCoverGS_Click
End Sub

Private Sub ButtonTreeCensus_Click()
On Error GoTo Err_ButtonTreeCensus_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Trees As DAO.Recordset
  Dim PlotSave As Variant
  Dim YearSave As Integer
  Dim SpeciesSave As String
  Dim NameSave As String
  Dim ArrayIndex As Integer
  Dim PlotArray(16) As Integer
  ' Array for tree count accumulators
  ' Index 0 is total Trees 25-30cm
  ' Index 1 is total Trees 30.1-40cm
  ' Index 2 is total Trees 40.1-50cm
  ' Index 3 is total Trees 50.1-60cm
  ' Index 4 is total Trees 60.1-70cm
  ' Index 5 is total Trees 70.1-80cm
  ' Index 6 is total Trees 80.1-90cm
  ' Index 7 is total Trees 90.1-100cm
  ' Index 8 is total Trees >100cm
  ' Index 9 is total Trees with crown health 1
  ' Index 10 is total Trees with crown health 2
  ' Index 11 is total Trees with crown health 3
  ' Index 12 is total Trees with crown health 4
  ' Index 13 is total Trees with crown health 5
  ' Index 14 is total Trees with no crown health recorded
  ' Index 15 is total Trees for species

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_OT_Trees"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Tree_Census_Summary WHERE 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
'  strSQL = strSQL & " and Plot_ID = 1"
  strSQL = strSQL & " ORDER BY Unit_Code, Visit_Year, Plot_ID, Species"
  DoCmd.Hourglass True
  Set db = CurrentDb
  ' Load work table
   Set Trees = db.OpenRecordset(strSQL)
   If Trees.EOF Then
     MsgBox "No valid tree records found."
     Trees.Close
     Set Trees = Nothing
     GoTo Exit_ButtonTreeCensus_Click
   End If
   Set WorkOutput = db.OpenRecordset("tbl_wrk_OT_Summary")
   ArrayIndex = 0
   Do Until ArrayIndex > 15    ' clear array
     PlotArray(ArrayIndex) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   PlotSave = Trees!Unit_Code & Trees!Plot_ID      ' Save necessary fields
   YearSave = Trees!Visit_Year
   SpeciesSave = Trees!Utah_Species
   NameSave = Trees!Stream_Name
   Do Until Trees.EOF
     If (PlotSave <> Trees!Unit_Code & Trees!Plot_ID) Or SpeciesSave <> Trees!Utah_Species Or YearSave <> Trees!Visit_Year Then  ' Test for new plot
       WorkOutput.AddNew
       WorkOutput!Unit_Code = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = NameSave
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Reach_ID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Utah_Species = SpeciesSave
       WorkOutput!Tree_Total = PlotArray(15)
       WorkOutput!DBH25_30 = PlotArray(0)
       WorkOutput!DBH30_40 = PlotArray(1)
       WorkOutput!DBH40_50 = PlotArray(2)
       WorkOutput!DBH50_60 = PlotArray(3)
       WorkOutput!DBH60_70 = PlotArray(4)
       WorkOutput!DBH70_80 = PlotArray(5)
       WorkOutput!DBH80_90 = PlotArray(6)
       WorkOutput!DBH90_100 = PlotArray(7)
       WorkOutput!Over100 = PlotArray(8)
       WorkOutput!CH1 = (PlotArray(9) / PlotArray(15)) * 100
       WorkOutput!CH2 = (PlotArray(10) / PlotArray(15)) * 100
       WorkOutput!CH3 = (PlotArray(11) / PlotArray(15)) * 100
       WorkOutput!CH4 = (PlotArray(12) / PlotArray(15)) * 100
       WorkOutput!CH5 = (PlotArray(13) / PlotArray(15)) * 100
       WorkOutput!NCH = (PlotArray(14) / PlotArray(15)) * 100
       WorkOutput.Update  ' Write plot record
       PlotSave = Trees!Unit_Code & Trees!Plot_ID
       YearSave = Trees!Visit_Year
       SpeciesSave = Trees!Utah_Species
       NameSave = Trees!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 15    ' clear array
         PlotArray(ArrayIndex) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
     End If ' End if for new species test
     PlotArray(15) = PlotArray(15) + 1
     Select Case Trees!DBH
       Case 25 To 30
         PlotArray(0) = PlotArray(0) + 1  ' Count
       Case 30.1 To 40
         PlotArray(1) = PlotArray(1) + 1  ' Trees
       Case 40.1 To 50
         PlotArray(2) = PlotArray(2) + 1  ' by
       Case 50.1 To 60
         PlotArray(3) = PlotArray(3) + 1  ' Size
       Case 60.1 To 70
         PlotArray(4) = PlotArray(4) + 1  ' Class
       Case 70.1 To 80
         PlotArray(5) = PlotArray(5) + 1
       Case 80.1 To 90
         PlotArray(6) = PlotArray(6) + 1
       Case 90.1 To 100
         PlotArray(7) = PlotArray(7) + 1
       Case Else
         PlotArray(8) = PlotArray(8) + 1
     End Select
     Select Case Trees!Crown_Health
       Case 1
         PlotArray(9) = PlotArray(9) + 1  ' Count
       Case 2
         PlotArray(10) = PlotArray(10) + 1  ' Trees
       Case 3
         PlotArray(11) = PlotArray(11) + 1  ' by
       Case 4
         PlotArray(12) = PlotArray(12) + 1  ' Crown
       Case 5
         PlotArray(13) = PlotArray(13) + 1  ' Health
       Case Else
         PlotArray(14) = PlotArray(14) + 1
     End Select
     Trees.MoveNext
   Loop
   ' Write last output record
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = NameSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Reach_ID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!Utah_Species = SpeciesSave
         WorkOutput!Tree_Total = PlotArray(15)
         WorkOutput!DBH25_30 = PlotArray(0)
         WorkOutput!DBH30_40 = PlotArray(1)
         WorkOutput!DBH40_50 = PlotArray(2)
         WorkOutput!DBH50_60 = PlotArray(3)
         WorkOutput!DBH60_70 = PlotArray(4)
         WorkOutput!DBH70_80 = PlotArray(5)
         WorkOutput!DBH80_90 = PlotArray(6)
         WorkOutput!DBH90_100 = PlotArray(7)
         WorkOutput!Over100 = PlotArray(8)
         WorkOutput!CH1 = (PlotArray(9) / PlotArray(15)) * 100
         WorkOutput!CH2 = (PlotArray(10) / PlotArray(15)) * 100
         WorkOutput!CH3 = (PlotArray(11) / PlotArray(15)) * 100
         WorkOutput!CH4 = (PlotArray(12) / PlotArray(15)) * 100
         WorkOutput!CH5 = (PlotArray(13) / PlotArray(15)) * 100
         WorkOutput!NCH = (PlotArray(14) / PlotArray(15)) * 100
         WorkOutput.Update  ' Write plot record
         WorkOutput.Close
         Set WorkOutput = Nothing
         Trees.Close
         Set Trees = Nothing
Exit_ButtonTreeCensus_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_OT_Summary."
    Exit Sub

Err_ButtonTreeCensus_Click:
    MsgBox Err.Description
    Resume Exit_ButtonTreeCensus_Click
End Sub

Private Sub ButtonTreeDensity_Click()
On Error GoTo Err_TreeDensity_Click

  Dim strSQL As String
  Dim SizeColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Trees As DAO.Recordset
  Dim ReachData As DAO.Recordset
  Dim reply As Integer
  Dim TreeIndex As Integer
  Dim TransectSum As Single
  Dim DBHLive As Double
  Dim DBHAll As Double
  Dim LiveCount As Integer
  Dim AllCount As Integer
  Dim YearSave As Integer
  
  reply = MsgBox("You must run Trees by Size Class first.", vbOKCancel, "Tree Density")
  If reply = vbCancel Then
    Exit Sub
  End If
  If IsNull(Me!Visit_Date) Then
    MsgBox "Visit year required", vbOKOnly, "Tree Density"
    Exit Sub
  End If
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Tree_Density"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Sum_Tree_Density ORDER BY UnitCode, PlotID"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get tree size info
   Set Trees = db.OpenRecordset(strSQL)
   If Trees.EOF Then
     MsgBox "No valid tree records found."
     Trees.Close
     Set Trees = Nothing
     GoTo Exit_TreeDensity_Click
   End If

   Trees.MoveFirst
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Tree_Density")  ' Open output table
   Do Until Trees.EOF
     strSQL = "SELECT * FROM tbl_Locations WHERE Unit_Code = '" & Trees!unitCode & "' AND Plot_ID = " & Trees!plotID
     Set ReachData = db.OpenRecordset(strSQL)
     If ReachData.EOF Then
       MsgBox "Reach data not found.  Unit " & Trees!unitCode & " Reach " & Trees!plotID & "."
       GoTo Exit_TreeDensity_Click
     End If
     TreeIndex = 1   ' Initialize index
     TransectSum = 0
     Do Until TreeIndex > 7  ' Go through transect length fields and sum them
       SizeColumn = "T" & TreeIndex & "_Length" ' Get the transect length field
       If Not IsNull(ReachData(SizeColumn)) Then
         TransectSum = TransectSum + ReachData(SizeColumn)  ' accumulate total transect lengths
       End If
       TreeIndex = TreeIndex + 1
     Loop
     TransectSum = TransectSum * 5  ' Sum of transects times 5 - dont know why
     ReachData.Close
     Set ReachData = Nothing
     If TransectSum > 0 Then
       WorkOutput.AddNew  ' Now write an output record
       WorkOutput!unitCode = Trees!unitCode  ' Set unit code
       WorkOutput!Stream_Name = Trees!Stream_Name
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Trees!plotID  ' Set plot ID
       WorkOutput!Seedlings_ha = (Trees!SumOfTotal_Seedlings * 10000) / TransectSum
       WorkOutput!Live_Total_ha = (Trees!Total_Live * 10000) / TransectSum
       WorkOutput!Live_5_ha = (Trees!SumOfLive_5 * 10000) / TransectSum
       WorkOutput!Live_10_ha = (Trees!SumOfLive_10 * 10000) / TransectSum
       WorkOutput!Live_15_ha = (Trees!SumOfLive_15 * 10000) / TransectSum
       WorkOutput!Live_20_ha = (Trees!SumOfLive_20 * 10000) / TransectSum
       WorkOutput!Live_30_ha = (Trees!SumOfLive_30 * 10000) / TransectSum
       WorkOutput!Live_40_ha = (Trees!SumOfLive_40 * 10000) / TransectSum
       WorkOutput!Live_50_ha = (Trees!SumOfLive_50 * 10000) / TransectSum
       WorkOutput!Live_Over_50_ha = (Trees!SumOfLive_Over_50 * 10000) / TransectSum
       WorkOutput!All_Total_ha = (Trees!Total_All * 10000) / TransectSum
       WorkOutput!All_5_ha = (Trees!SumOfAll_5 * 10000) / TransectSum
       WorkOutput!All_10_ha = (Trees!SumOfAll_10 * 10000) / TransectSum
       WorkOutput!All_15_ha = (Trees!SumOfAll_15 * 10000) / TransectSum
       WorkOutput!All_20_ha = (Trees!SumOfAll_20 * 10000) / TransectSum
       WorkOutput!All_30_ha = (Trees!SumOfAll_30 * 10000) / TransectSum
       WorkOutput!All_40_ha = (Trees!SumOfAll_40 * 10000) / TransectSum
       WorkOutput!All_50_ha = (Trees!SumOfAll_50 * 10000) / TransectSum
       WorkOutput!All_Over_50_ha = (Trees!SumOfAll_Over_50 * 10000) / TransectSum
       WorkOutput.Update  ' Write previous output record
     End If
     Trees.MoveNext
   Loop
   WorkOutput.Close
   Set WorkOutput = Nothing
   Trees.Close
   Set Trees = Nothing
   ' Disabling DBH fields for now, but do not trust them, so code stays
   ' Now update DBH fields
   '  Set Trees = db.OpenRecordset("tbl_wrk_Tree_Density")
   '  Do Until Trees.EOF
   '    strSQL = "SELECT * FROM qry_Tree_DBH WHERE Unit_Code = '" & Trees!UnitCode & "' AND Species = '" & Trees!Species & "' AND Plot_ID = " & Trees!PlotID & " AND Visit_Year = " & Me!Visit_Date
   '    Set ReachData = db.OpenRecordset(strSQL)
   '    If Not ReachData.EOF Then  ' Check for no DBH data
   '      LiveCount = 0
   '      AllCount = 0
   '      DBHLive = 0
   '      DBHAll = 0
   '      Do Until ReachData.EOF
   '        AllCount = AllCount + 1
   '        DBHAll = DBHAll + ReachData!dbh
   '        If ReachData!alive Then
   '          LiveCount = LiveCount + 1
   '          DBHLive = DBHLive + ReachData!dbh
   '        End If
   '        ReachData.MoveNext
   '      Loop
   '      Trees.Edit
   '      If LiveCount > 0 Then
   '        Trees!Mean_DBH_L = DBHLive / LiveCount
   '      End If
   '      If AllCount > 0 Then
   '        Trees!Mean_DBH_A = DBHAll / AllCount
   '      End If
   '      Trees.Update
   '    End If  ' End if for no dbh data check
   '    ReachData.Close
   '    Set ReachData = Nothing
   '    Trees.MoveNext
   '  Loop
   '  Trees.Close
   '  Set Trees = Nothing
Exit_TreeDensity_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Tree_Density."
    Exit Sub

Err_TreeDensity_Click:
    MsgBox Err.Description
    Resume Exit_TreeDensity_Click
End Sub

Private Sub ButtonTreeSpeciesDensity_Click()
On Error GoTo Err_Tree_Species_Density_Click

  Dim strSQL As String
  Dim SizeColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Trees As DAO.Recordset
  Dim ReachData As DAO.Recordset
  Dim reply As Integer
  Dim TreeIndex As Integer
  Dim TransectSum As Single
  Dim DBHLive As Double
  Dim DBHAll As Double
  Dim LiveCount As Integer
  Dim AllCount As Integer
  Dim YearSave As Integer
  
  reply = MsgBox("You must run Trees by Size Class first.", vbOKCancel, "Tree Density")
  If reply = vbCancel Then
    Exit Sub
  End If
  If IsNull(Me!Visit_Date) Then
    MsgBox "Visit year required", vbOKOnly, "Tree Density"
    Exit Sub
  End If
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Tree_Species_Density"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Sum_Tree_Size ORDER BY UnitCode, PlotID"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get tree size info
   Set Trees = db.OpenRecordset(strSQL)
   If Trees.EOF Then
     MsgBox "No valid tree records found."
     Trees.Close
     Set Trees = Nothing
     GoTo Exit_Tree_Species_Density_Click
   End If

   Trees.MoveFirst
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Tree_Species_Density")  ' Open output table
   Do Until Trees.EOF
     strSQL = "SELECT * FROM tbl_Locations WHERE Unit_Code = '" & Trees!unitCode & "' AND Plot_ID = " & Trees!plotID
     Set ReachData = db.OpenRecordset(strSQL)
     If ReachData.EOF Then
       MsgBox "Reach data not found.  Unit " & Trees!unitCode & " Reach " & Trees!plotID & "."
       GoTo Exit_Tree_Species_Density_Click
     End If
     TreeIndex = 1   ' Initialize index
     TransectSum = 0
     Do Until TreeIndex > 7  ' Go through transect length fields and sum them
       SizeColumn = "T" & TreeIndex & "_Length" ' Get the transect length field
       If Not IsNull(ReachData(SizeColumn)) Then
         TransectSum = TransectSum + ReachData(SizeColumn)  ' accumulate total transect lengths
       End If
       TreeIndex = TreeIndex + 1
     Loop
     TransectSum = TransectSum * 5  ' Sum of transects times 5 - dont know why
     ReachData.Close
     Set ReachData = Nothing
     If TransectSum > 0 Then
       WorkOutput.AddNew  ' Now write an output record
       WorkOutput!unitCode = Trees!unitCode  ' Set unit code
       WorkOutput!Stream_Name = Trees!Stream_Name
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!plotID = Trees!plotID  ' Set plot ID
       WorkOutput!Species = Trees!Species
       WorkOutput!Seedlings_ha = (Trees!SumOfTotal_Seedlings * 10000) / TransectSum
       WorkOutput!Live_Total_ha = (Trees!Total_Live * 10000) / TransectSum
       WorkOutput!Live_5_ha = (Trees!SumOfLive_5 * 10000) / TransectSum
       WorkOutput!Live_10_ha = (Trees!SumOfLive_10 * 10000) / TransectSum
       WorkOutput!Live_15_ha = (Trees!SumOfLive_15 * 10000) / TransectSum
       WorkOutput!Live_20_ha = (Trees!SumOfLive_20 * 10000) / TransectSum
       WorkOutput!Live_30_ha = (Trees!SumOfLive_30 * 10000) / TransectSum
       WorkOutput!Live_40_ha = (Trees!SumOfLive_40 * 10000) / TransectSum
       WorkOutput!Live_50_ha = (Trees!SumOfLive_50 * 10000) / TransectSum
       WorkOutput!Live_Over_50_ha = (Trees!SumOfLive_Over_50 * 10000) / TransectSum
       WorkOutput!All_Total_ha = (Trees!Total_All * 10000) / TransectSum
       WorkOutput!All_5_ha = (Trees!SumOfAll_5 * 10000) / TransectSum
       WorkOutput!All_10_ha = (Trees!SumOfAll_10 * 10000) / TransectSum
       WorkOutput!All_15_ha = (Trees!SumOfAll_15 * 10000) / TransectSum
       WorkOutput!All_20_ha = (Trees!SumOfAll_20 * 10000) / TransectSum
       WorkOutput!All_30_ha = (Trees!SumOfAll_30 * 10000) / TransectSum
       WorkOutput!All_40_ha = (Trees!SumOfAll_40 * 10000) / TransectSum
       WorkOutput!All_50_ha = (Trees!SumOfAll_50 * 10000) / TransectSum
       WorkOutput!All_Over_50_ha = (Trees!SumOfAll_Over_50 * 10000) / TransectSum
       WorkOutput!Total_Seedlings_Pct = (Trees!SumOfTotal_Seedlings / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_5_Pct = (Trees!SumOfLive_5 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_10_Pct = (Trees!SumOfLive_10 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_15_Pct = (Trees!SumOfLive_15 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_20_Pct = (Trees!SumOfLive_20 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_30_Pct = (Trees!SumOfLive_30 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_40_Pct = (Trees!SumOfLive_40 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_50_Pct = (Trees!SumOfLive_50 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!Live_Over_50_Pct = (Trees!SumOfLive_Over_50 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_5_Pct = (Trees!SumOfAll_5 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_10_Pct = (Trees!SumOfAll_10 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_15_Pct = (Trees!SumOfAll_15 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_20_Pct = (Trees!SumOfAll_20 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_30_Pct = (Trees!SumOfAll_30 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_40_Pct = (Trees!SumOfAll_40 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_50_Pct = (Trees!SumOfAll_50 / Trees!SumOfTotal_Trees) * 100
       WorkOutput!All_Over_50_Pct = (Trees!SumOfAll_Over_50 / Trees!SumOfTotal_Trees) * 100
       WorkOutput.Update  ' Write previous output record
     End If
     Trees.MoveNext
   Loop
   WorkOutput.Close
   Set WorkOutput = Nothing
   Trees.Close
   Set Trees = Nothing
   ' Now update DBH fields
   Set Trees = db.OpenRecordset("tbl_wrk_Tree_Species_Density")
   Do Until Trees.EOF
     strSQL = "SELECT * FROM qry_Tree_DBH WHERE Unit_Code = '" & Trees!unitCode & "' AND Species = '" & Trees!Species & "' AND Plot_ID = " & Trees!plotID & " AND Visit_Year = " & Me!Visit_Date
     Set ReachData = db.OpenRecordset(strSQL)
     If Not ReachData.EOF Then  ' Check for no DBH data
       LiveCount = 0
       AllCount = 0
       DBHLive = 0
       DBHAll = 0
       Do Until ReachData.EOF
         AllCount = AllCount + 1
         DBHAll = DBHAll + ReachData!DBH
         If ReachData!alive Then
           LiveCount = LiveCount + 1
           DBHLive = DBHLive + ReachData!DBH
         End If
         ReachData.MoveNext
       Loop
       Trees.Edit
       If LiveCount > 0 Then
         Trees!Mean_DBH_L = DBHLive / LiveCount
       End If
       If AllCount > 0 Then
         Trees!Mean_DBH_A = DBHAll / AllCount
       End If
       Trees.Update
     End If  ' End if for no dbh data check
     ReachData.Close
     Set ReachData = Nothing
     Trees.MoveNext
   Loop
   Trees.Close
   Set Trees = Nothing
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_upd_Tree_Species_Density"   ' Update species names.
   DoCmd.SetWarnings True
Exit_Tree_Species_Density_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Tree_Species_Density."
    Exit Sub

Err_Tree_Species_Density_Click:
    MsgBox Err.Description
    Resume Exit_Tree_Species_Density_Click
End Sub

Private Sub Command32_Click()
On Error GoTo Err_CoverSpeciesAll_Click

  Dim strSQL As String
  Dim lifeForm As Variant
  Dim SpeciesColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim PlotSave As Variant
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim StreamSave As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim ArrayEnd As Integer
  Dim PointArray(12) As Variant ' Array for species at a GS
  ' Species hits per point array
  ' Column 1 is species code
  Dim PlotArray(300, 1) As Variant ' Array for species in a plot
  ' Species hits per plot array
  ' Column 1 is species code
  ' Column 2 is alive count
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_All_GS"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GS where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpeciesAll_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 299           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary
   StreamSave = points!Stream_Name                  ' Fields
   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
     GeomorphIn = points!Geomorphic_Surface
   Else
     Geomorph = "None"
     GeomorphIn = "None"
   End If
   Point_Count = 0

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Or GeomorphIn <> Geomorph Then  ' Is it a new GS
       PointIndex = 0  ' yes - add in last point
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextAIndex  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextAIndex:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       ' *** End of plot processing ***
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_All_GS")
       ArrayIndex = 0
       Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for last plot
         If PlotArray(ArrayIndex, 1) > 0 Then
           WorkOutput.AddNew
           WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
           WorkOutput!Geomorph = Geomorph
           WorkOutput!Stream_Name = StreamSave
           WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
           WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
           WorkOutput!PercentCoverAll = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
           WorkOutput.Update  ' Write previous output record
           RecordCount = RecordCount + 1
         End If
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Geomorph = GeomorphIn
       ArrayIndex = 0
       Do Until ArrayIndex > 299    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
       Point_Count = 0
     End If
     If PointSave <> points!point Then  ' Is it a new point
     '  *** End of point processing ***
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Top = PointArray(PointIndex) Then  ' is the species the same
             Exit Do   ' Already have the species for this point
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Top  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Surface = PointArray(PointIndex) Then  ' is the species the same
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Surface  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipSpecies
       Else
         PointIndex = 0
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex) = points(SpeciesColumn)  ' set species
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipSpecies:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
   Loop
   ' End of file - add in last point
   PointIndex = 0
     Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
       ArrayIndex = 0
       Do Until ArrayIndex > 299
         If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
           If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           Else
             GoTo LastPlotArray  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
           Exit Do
         End If  ' end if for array slot open test
LastPlotArray:
         ArrayIndex = ArrayIndex + 1
       Loop
       PointIndex = PointIndex + 1
     Loop
     Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_All_GS")  ' Write last output record
     ArrayIndex = 0
     Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "
       If PlotArray(ArrayIndex, 1) > 0 Then
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Geomorph = Geomorph
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!PercentCoverAll = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
         WorkOutput.Update  ' Write previous output record
         RecordCount = RecordCount + 1
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
   DoCmd.SetWarnings False
   DoCmd.OpenQuery "qry_upd_Cover_Pct_All_GS"   ' Update species names.
   DoCmd.SetWarnings True
Exit_CoverSpeciesAll_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_All_GS."
    Exit Sub

Err_CoverSpeciesAll_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpeciesAll_Click
End Sub

Private Sub Command34_Click()
On Error GoTo Err_CoverNativity_Click
  Dim strSQL As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Nativity As String
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim LCIndex As Integer
  Dim PlotTotalA As Integer
  Dim PlotTotalL As Integer
  Dim PointNativeL As Byte
  Dim PointNativeA As Byte
  Dim PointNonNativeL As Byte
  Dim PointNonNativeA As Byte
  Dim PlotTotalNativeL As Integer
  Dim PlotTotalNativeA As Integer
  Dim PlotTotalNonNativeL As Integer
  Dim PlotTotalNonNativeA As Integer

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Nativity_GS"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GS where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverNativity_Click
   End If
   PlotTotalL = 0
   PlotTotalA = 0
   PointNativeL = 0
   PointNativeA = 0
   PointNonNativeL = 0
   PointNonNativeA = 0
   PlotTotalNativeL = 0
   PlotTotalNativeA = 0
   PlotTotalNonNativeL = 0
   PlotTotalNonNativeA = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
     GeomorphIn = points!Geomorphic_Surface
   Else
     Geomorph = "None"
     GeomorphIn = "None"
   End If
   Point_Count = 0
   Do Until points.EOF
     If (PlotSave <> points!Unit_Code & points!Plot_ID) Or (Geomorph <> GeomorphIn) Then  ' Check for new geomorphic surface
       ' New plot - process last point totals from previous plot first
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Nativity_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!TotalL = (PlotTotalL / Point_Count) * 100
       WorkOutput!TotalA = (PlotTotalA / Point_Count) * 100
       WorkOutput!NativeL = (PlotTotalNativeL / Point_Count) * 100
       WorkOutput!NativeA = (PlotTotalNativeA / Point_Count) * 100
       WorkOutput!NonNativeL = (PlotTotalNonNativeL / Point_Count) * 100
       WorkOutput!NonNativeA = (PlotTotalNonNativeA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Point_Count = 0
       PlotTotalL = 0
       PlotTotalA = 0
       PointNativeL = 0
       PointNativeA = 0
       PointNonNativeL = 0
       PointNonNativeA = 0
       PlotTotalNativeL = 0
       PlotTotalNativeA = 0
       PlotTotalNonNativeL = 0
       PlotTotalNonNativeA = 0
     End If
     If PointSave <> points!point Then  ' End of point - add counts to plot array
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       PointLive = 0
       PointAll = 0
       PointSave = points!point  '  Save new point
       Geomorph = GeomorphIn     '  Save new geomorphic surface
       PointNativeL = 0
       PointNativeA = 0
       PointNonNativeL = 0
       PointNonNativeA = 0
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Nativity])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       Nativity = SpeciesLU!Nativity
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If Nativity = "Native" Then
         PointNativeA = 1
         If points!alive Then
           PointNativeL = 1
         End If
       Else
         PointNonNativeA = 1
         If points!alive Then
           PointNonNativeL = 1
         End If
       End If
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Surface & "' AND NOT IsNull([Nativity])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       Nativity = SpeciesLU!Nativity
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If Nativity = "Native" Then
         PointNativeA = 1
         If points!Surface_Alive Then
           PointNativeL = 1
         End If
       Else
         PointNonNativeA = 1
         If points!Surface_Alive Then
           PointNonNativeL = 1
         End If
       End If
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Nativity FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Nativity])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
           Nativity = SpeciesLU!Nativity
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Nativity = "Native" Then
           PointNativeA = 1
           If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
             PointNativeL = 1
           End If
         Else
           PointNonNativeA = 1
           If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
             PointNonNativeL = 1
           End If
         End If
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
     Point_Count = Point_Count + 1
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
   Loop
   '  Process last point totals
       If PointNativeL + PointNonNativeL > 0 Then  ' Accumulate
         PlotTotalL = PlotTotalL + 1               ' Live
       End If                                      ' And
       If PointNativeA + PointNonNativeA > 0 Then  ' Dead
         PlotTotalA = PlotTotalA + 1               ' Plot
       End If                                      ' Totals
       If PointNonNativeL = 1 Then
         PlotTotalNonNativeL = PlotTotalNonNativeL + 1
       End If
       If PointNonNativeA = 1 Then
         PlotTotalNonNativeA = PlotTotalNonNativeA + 1
       End If
       If PointNativeL = 1 Then
         PlotTotalNativeL = PlotTotalNativeL + 1
       End If
       If PointNativeA = 1 Then
         PlotTotalNativeA = PlotTotalNativeA + 1
       End If
       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Nativity_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!Geomorph = Geomorph
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!TotalL = (PlotTotalL / Point_Count) * 100
       WorkOutput!TotalA = (PlotTotalA / Point_Count) * 100
       WorkOutput!NativeL = (PlotTotalNativeL / Point_Count) * 100
       WorkOutput!NativeA = (PlotTotalNativeA / Point_Count) * 100
       WorkOutput!NonNativeL = (PlotTotalNonNativeL / Point_Count) * 100
       WorkOutput!NonNativeA = (PlotTotalNonNativeA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverNativity_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Nativity_GS."
    Exit Sub

Err_CoverNativity_Click:
    MsgBox Err.Description
    Resume Exit_CoverNativity_Click
End Sub

Private Sub Command38_Click()
On Error GoTo Err_CoverSpecies_Click

  Dim strSQL As String
  Dim lifeForm As Variant
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim RecordCount As Long
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim PointSave As Double
  Dim ACount As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim ArrayEnd As Integer
  Dim PointArray(12) As Variant ' Array for species at a point
  ' Species hits per point array
  ' Column 1 is species code
  Dim PlotArray(300, 1) As Variant ' Array for species in a plot
  ' Species hits per plot array
  ' Column 1 is species code
  ' Column 2 is alive count
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_GL_Cover_Pct_Live"
  DoCmd.SetWarnings True
  
'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GL where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpecies_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex) = " "
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 299           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary
   StreamSave = points!Stream_Name

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Is it a new plot
       PointIndex = 0  ' yes - add in last point
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextAIndex  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextAIndex:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       ' *** End of plot processing ***
       Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Live")
       ArrayIndex = 0
       Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for last plot
         If PlotArray(ArrayIndex, 1) > 0 Then
           WorkOutput.AddNew
           WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
           WorkOutput!Stream_Name = StreamSave
           WorkOutput!Visit_Year = Me!Visit_Date
           WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
           WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
           WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / 560) * 100
           WorkOutput.Update  ' Write previous output record
           RecordCount = RecordCount + 1
         End If
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 299    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If
     If PointSave <> points!point Then  ' Is it a new point
     '  *** End of point processing ***
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 299
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex) = " "
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) And points!alive Then
       PointIndex = 0
       Do Until PointIndex > 11
         If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
           If points!Top = PointArray(PointIndex) Then  ' is the species the same
             Exit Do   ' Already have the species for this point
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex) = points!Top  ' set species
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check

     '  Soil Surface removed 4/18/2013 RD.
     
     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipSpecies
       Else
         PointIndex = 0
         Do Until PointIndex > 11
           If PointArray(PointIndex) <> " " Then ' if spot is used - check it out
             If points(SpeciesColumn) = PointArray(PointIndex) Then  ' is the species the same
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex) = points(SpeciesColumn)  ' set species
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipSpecies:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     points.MoveNext
   Loop
   ' End of file - add in last point
   PointIndex = 0
     Do Until PointIndex > 11 Or PointArray(PointIndex) = " "
       ArrayIndex = 0
       Do Until ArrayIndex > 299
         If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
           If PointArray(PointIndex) = PlotArray(ArrayIndex, 0) Then  ' is the species the same
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
             Exit Do
           Else
             GoTo LastPlotArray  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PlotArray(ArrayIndex, 0) = PointArray(PointIndex)  ' set species
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it
           Exit Do
         End If  ' end if for array slot open test
LastPlotArray:
         ArrayIndex = ArrayIndex + 1
       Loop
       PointIndex = PointIndex + 1
     Loop
     Set WorkOutput = db.OpenRecordset("tbl_wrk_GL_Cover_Pct_Live")  ' Write last output record
     ArrayIndex = 0
     Do Until ArrayIndex > 299 Or PlotArray(ArrayIndex, 0) = " "
       If PlotArray(ArrayIndex, 1) > 0 Then
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = StreamSave
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!SpeciesCode = PlotArray(ArrayIndex, 0)
         WorkOutput!PercentCoverLive = (PlotArray(ArrayIndex, 1) / 560) * 100
         WorkOutput.Update  ' Write previous output record
         RecordCount = RecordCount + 1
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     WorkOutput.Close
     Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
   DoCmd.SetWarnings False
'   DoCmd.OpenQuery "qry_del_Cover_Pct_Live_SS"   ' Delete any soil surface codes that may have been picked up in lower canopy.
   DoCmd.OpenQuery "qry_upd_GL_Cover_Pct_Live"   ' Update species names.
   DoCmd.SetWarnings True
Exit_CoverSpecies_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_GL_Cover_Pct_Live."
    Exit Sub

Err_CoverSpecies_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpecies_Click
End Sub

Private Sub Park_Code_AfterUpdate()
  If Not IsNull(Me!Park_Code) Then
    Me!Visit_Date.RowSource = "SELECT Visit_Year FROM qry_Event_Date WHERE Unit_Code = '" & Me!Park_Code & "'"
  End If
End Sub

Private Sub RichnessbyWetland_Click()
On Error GoTo Err_RichnessbyWetland_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SQL As String
  Dim RecordCount As Long
  Dim Plot_Save As Integer
  Dim Unit_Save As String
  Dim Stream_Name As String
  Dim reply As Integer
  Dim FAC As Integer
  Dim FACU As Integer
  Dim FACW As Integer
  Dim OBL As Integer
  Dim UPL As Integer
  Dim CULT As Integer
  
  reply = MsgBox("You must run All Species by Reach first.", vbOKCancel, "Richness by Species")
  If reply = vbCancel Then
    Exit Sub
  End If
  
  DoCmd.Hourglass True
  
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Wetland_Richness"
  DoCmd.SetWarnings True
  
  Set db = CurrentDb
  RecordCount = 0
  ' Get species info
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Richness_Wetland")
   SQL = "SELECT * FROM qry_SR_Species_Lifeform where Wetland_Code is not null"
   Set points = db.OpenRecordset(SQL)
   If points.EOF Then
     MsgBox "Species by plot table is empty."
     GoTo Exit_RichnessbyWetland_Click:
   End If
   points.MoveFirst
   Unit_Save = points!Unit_Code
   Stream_Name = points!Stream_Name
   Plot_Save = points!Plot_ID
   Do Until points.EOF  ' Load lifeform breakdown into work table.
     If points!Unit_Code <> Unit_Save Or points!Plot_ID <> Plot_Save Then
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Unit_Save
         WorkOutput!Plot_ID = Plot_Save
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Stream_Name = Stream_Name
         WorkOutput!FAC = FAC
         WorkOutput!FACU = FACU
         WorkOutput!FACW = FACW
         WorkOutput!OBL = OBL
         WorkOutput!UPL = UPL
         WorkOutput!CULT = CULT
         WorkOutput.Update
         RecordCount = RecordCount + 1
         Unit_Save = points!Unit_Code
         Stream_Name = points!Stream_Name
         Plot_Save = points!Plot_ID
         FAC = 0
         FACU = 0
         FACW = 0
         OBL = 0
         UPL = 0
         CULT = 0
      End If
      Species_Total = Species_Total + 1
      If Not IsNull(points!Nativity) And points!Nativity = "NonNative" Then
        Species_TotalE = Species_TotalE + 1
      End If
      Select Case points!Wetland_Code
        Case "FAC"
          FAC = FAC + 1
        Case "FACU"
          FACU = FACU + 1
        Case "FACW"
          FACW = FACW + 1
        Case "OBL"
          OBL = OBL + 1
        Case "UPL"
          UPL = UPL + 1
        Case "CULT"
          CULT = CULT + 1
        Case Else
          MsgBox "Unrecognized wetland code " & points!lifeForm
      End Select
     points.MoveNext
   Loop
   ' Write last record
         WorkOutput.AddNew
         WorkOutput!Unit_Code = Unit_Save
         WorkOutput!Plot_ID = Plot_Save
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Stream_Name = Stream_Name
         WorkOutput!FAC = FAC
         WorkOutput!FACU = FACU
         WorkOutput!FACW = FACW
         WorkOutput!OBL = OBL
         WorkOutput!UPL = UPL
         WorkOutput!CULT = CULT
         WorkOutput.Update
         RecordCount = RecordCount + 1
   points.Close
   Set points = Nothing
   WorkOutput.Close
   Set WorkOutput = Nothing


Exit_RichnessbyWetland_Click:
    DoCmd.Hourglass False
    MsgBox RecordCount & " records written.  Results are in tbl_wrk_Richness_Wetland."
    Exit Sub

Err_RichnessbyWetland_Click:
    MsgBox Err.Description
    Resume Exit_RichnessbyWetland_Click
  
End Sub
Private Sub ButtonLifeformGS_Click()
On Error GoTo Err_CoverSpecies_Click

  Dim strSQL As String
  Dim lifeForm As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim SpeciesLU As DAO.Recordset
  Dim PlotSave As Variant
  Dim Geomorph As String
  Dim GeomorphIn As String
  Dim StreamSave As String
  Dim PointSave As Double
  Dim Point_Count As Integer
  Dim dblDivisor As Double
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointIndex As Integer
  Dim ArrayIndex As Integer
  Dim PointLive As Byte
  Dim PointAll As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer
  Dim PointArray(12, 3) As Variant ' Array for lifeform at a geomorphic surface
  ' Species hits per point array
  ' Column 1 lifeform
  ' Column x,0 is alive flag
  ' Column x, 1 is dead flag
  Dim PlotArray(8, 3) As Variant ' Array for lifeforms in a GC
  ' Species hits per plot array
  ' Column 1 is lifeform
  ' Column x, 0 is alive count
  ' Column x, 1 is total count

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Lifeform_GS"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_GS where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Geomorphic_Surface"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_CoverSpecies_Click
   End If
   PointIndex = 0
   Do Until PointIndex > 11            ' Initialize point array
     PointArray(PointIndex, 0) = " "
     PointArray(PointIndex, 1) = 0
     PointArray(PointIndex, 2) = 0
     PointIndex = PointIndex + 1
   Loop
   ArrayIndex = 0
   Do Until ArrayIndex > 7           ' Initialize plot array
     PlotArray(ArrayIndex, 0) = " "
     PlotArray(ArrayIndex, 1) = 0
     PlotArray(ArrayIndex, 2) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   PointLive = 0
   PointAll = 0
   PlotTotalA = 0
   PlotTotalL = 0
   points.MoveFirst
   PointSave = points!point                         ' Save
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name
   If Not IsNull(points!Geomorphic_Surface) Then
     Geomorph = points!Geomorphic_Surface
     GeomorphIn = points!Geomorphic_Surface
   Else
     Geomorph = "None"
     GeomorphIn = "None"
   End If
   Point_Count = 0
   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Or Geomorph <> GeomorphIn Then  ' Check for new plot code
       ' New plot - process last point in previous plot first
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 7
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextArrayEntry  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive

             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextArrayEntry:
           ArrayIndex = ArrayIndex + 1
         Loop  ' Loop for lifeform processing
         PointIndex = PointIndex + 1
       Loop  ' Loop for point processing

       ' Now write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Lifeform_GS")
       ArrayIndex = 0
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Geomorph = Geomorph
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!ShrubL = 0
       WorkOutput!ShrubA = 0
       ' MsgBox PlotTotalL & " " & PlotTotalA
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       Do Until ArrayIndex > 7 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "Tree"
             WorkOutput!TreeL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!TreeA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Shrub", "DwarfShrub"
             WorkOutput!ShrubL = WorkOutput!ShrubL + ((PlotArray(ArrayIndex, 1) / Point_Count) * 100)
             WorkOutput!ShrubA = WorkOutput!ShrubA + ((PlotArray(ArrayIndex, 2) / Point_Count) * 100)
           Case "Forb"
             WorkOutput!ForbL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!ForbA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Annual"
             WorkOutput!AGrassL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!AGrassA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Perennial"
             WorkOutput!PGrassL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!PGrassA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Fern"
             WorkOutput!FernL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FernA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Vine"
             WorkOutput!VineL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!VineA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case Else
             MsgBox "Unknown lifeform " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       Geomorph = GeomorphIn
       Point_Count = 0
       ArrayIndex = 0
       Do Until ArrayIndex > 5    ' Clear array
         PlotArray(ArrayIndex, 0) = " "
         PlotArray(ArrayIndex, 1) = 0
         PlotArray(ArrayIndex, 2) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
       PlotTotalA = 0
       PlotTotalL = 0
     End If
     If PointSave <> points!point Then  ' End of point - add lifeforms to plot array
       PointIndex = 0
       Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
         ArrayIndex = 0
         Do Until ArrayIndex > 5
           If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
             If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
               If PointArray(PointIndex, 1) = 1 Then
                 PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
               End If
               PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
               Exit Do
             Else
               GoTo NextPlotArray  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
             If PointArray(PointIndex, 1) = 1 Then
               PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
             End If
             PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
             Exit Do
           End If  ' end if for array slot open test
NextPlotArray:
           ArrayIndex = ArrayIndex + 1
         Loop
         PointIndex = PointIndex + 1
       Loop
       PointLive = 0
       PointAll = 0
SkipPointSpecies:
       PointSave = points!point  '  Save new point
       PointIndex = 0
       Do Until PointIndex > 11            ' Initialize point array
         PointArray(PointIndex, 0) = " "
         PointArray(PointIndex, 1) = 0
         PointArray(PointIndex, 2) = 0
         PointIndex = PointIndex + 1
       Loop
     End If  ' End if for new point test
     
     '  Top cover first
     If Not IsNull(points!Top) Then
       PointIndex = 0
       strSQL = "SELECT Lifeform, Duration FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Top & "' AND NOT IsNull([Lifeform])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipTop  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!lifeForm = "Graminoid" Then
         If IsNull(SpeciesLU!Duration) Then
           GoTo SkipTop ' Skip null duration
         Else
           If SpeciesLU!Duration = "Perennial" Then
             lifeForm = "Perennial"
           Else
             lifeForm = "Annual"
           End If
         End If
       Else
         lifeForm = SpeciesLU!lifeForm
       End If
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If lifeForm = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!alive Then
               PointArray(PointIndex, 1) = 1  ' Set alive flag
             Else
               PointArray(PointIndex, 2) = 1  ' set dead flag
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextTop  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = lifeForm  ' set species
           If points!alive Then
             PointArray(PointIndex, 1) = 1  ' count it as alive
           Else
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextTop:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null top check
SkipTop:

     '  Soil Surface next
     If Not IsNull(points!Surface) And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       PointIndex = 0
       strSQL = "SELECT Lifeform, Duration FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points!Surface & "' AND NOT IsNull([Lifeform])"
       Set SpeciesLU = db.OpenRecordset(strSQL)
       If SpeciesLU.EOF Then
         GoTo SkipSurface  ' It was probably a null lifeform, skip it.
       End If
       If SpeciesLU!lifeForm = "Graminoid" Then
         If IsNull(SpeciesLU!Duration) Then
           GoTo SkipSurface ' Skip null duration
         Else
           If SpeciesLU!Duration = "Perennial" Then
             lifeForm = "Perennial"
           Else
             lifeForm = "Annual"
           End If
         End If
       Else
         lifeForm = SpeciesLU!lifeForm
       End If
       SpeciesLU.Close
       Set SpeciesLU = Nothing
       If points!Surface_Alive Then
         PointLive = 1
       End If
       PointAll = 1
       Do Until PointIndex > 11
         If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
           If lifeForm = PointArray(PointIndex, 0) Then  ' is the species the same
             If points!Surface_Alive Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           Else
             GoTo NextSurface  ' Different species - go to next entry
           End If  ' End if for species compare
         Else
           PointArray(PointIndex, 0) = lifeForm  ' set species
           If points!Surface_Alive Then
             PointArray(PointIndex, 1) = 1  ' flag it as alive
           Else
             PointArray(PointIndex, 2) = 1  ' flag it as dead
           End If ' end if for alive test
           Exit Do
         End If  ' end if for array slot open test
NextSurface:
         PointIndex = PointIndex + 1  ' next array entry
       Loop
     End If  ' End if for null soil surface check
SkipSurface:

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Then
         Exit Do  ' If we hit a null, there will be no more
       Else
         PointIndex = 0
         strSQL = "SELECT Lifeform, Duration FROM tlu_NCPN_Plants WHERE Master_PLANT_Code = '" & points(SpeciesColumn) & "' AND NOT IsNull([Lifeform])"
         Set SpeciesLU = db.OpenRecordset(strSQL)
         If SpeciesLU.EOF Then
           GoTo SkipLC  ' It was probably a null lifeform, skip it.
         End If
         If SpeciesLU!lifeForm = "Graminoid" Then
           If IsNull(SpeciesLU!Duration) Then
             GoTo SkipLC ' Skip null duration
           Else
             If SpeciesLU!Duration = "Perennial" Then
               lifeForm = "Perennial"
             Else
               lifeForm = "Annual"
             End If
           End If
         Else
           lifeForm = SpeciesLU!lifeForm
         End If
         SpeciesLU.Close
         Set SpeciesLU = Nothing
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
         Do Until PointIndex > 11
           If PointArray(PointIndex, 0) <> " " Then ' if spot is used - check it out
             If lifeForm = PointArray(PointIndex, 0) Then  ' is the species the same
               If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
                 PointArray(PointIndex, 1) = 1  ' flag it as alive
               Else
                 PointArray(PointIndex, 2) = 1  ' flag it as dead
               End If ' end if for alive test
               Exit Do
             Else
               GoTo NextLC  ' Different species - go to next entry
             End If  ' End if for species compare
           Else
             PointArray(PointIndex, 0) = lifeForm  ' set species
             If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
               PointArray(PointIndex, 1) = 1  ' flag it as alive
             Else
               PointArray(PointIndex, 2) = 1  ' flag it as dead
             End If ' end if for alive test
             Exit Do
           End If  ' end if for array slot open test
NextLC:
           PointIndex = PointIndex + 1  ' next array entry
         Loop
       End If  ' End if for null lower canopy check
SkipLC:
       LCIndex = LCIndex + 1
     Loop
NextPoint:
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     points.MoveNext
     Point_Count = Point_Count + 1
     If Not points.EOF Then
       If IsNull(points!Geomorphic_Surface) Then
         GeomorphIn = "None"
       Else
         GeomorphIn = points!Geomorphic_Surface
       End If
     End If
   Loop
   '  Process last point
   PointIndex = 0
   Do Until PointIndex > 11 Or PointArray(PointIndex, 0) = " "
     ArrayIndex = 0
     Do Until ArrayIndex > 7
       If PlotArray(ArrayIndex, 0) <> " " Then ' if spot is used - check it out
         If PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0) Then  ' is this the correct lifeform
           If PointArray(PointIndex, 1) = 1 Then
             PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
           End If
           PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
           Exit Do
         Else
           GoTo LastPlotArray  ' Different species - go to next entry
         End If  ' End if for species compare
       Else
         PlotArray(ArrayIndex, 0) = PointArray(PointIndex, 0)  ' set lifeform
         If PointArray(PointIndex, 1) = 1 Then
           PlotArray(ArrayIndex, 1) = PlotArray(ArrayIndex, 1) + 1  ' count it as alive
         End If
         PlotArray(ArrayIndex, 2) = PlotArray(ArrayIndex, 2) + 1  ' count it for total
         Exit Do
       End If  ' end if for array slot open test
LastPlotArray:
       ArrayIndex = ArrayIndex + 1
     Loop
     PointIndex = PointIndex + 1
   Loop

   ' Output last plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Lifeform_GS")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Geomorph = Geomorph
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       ArrayIndex = 0
       WorkOutput!ShrubL = 0
       WorkOutput!ShrubA = 0
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       Do Until ArrayIndex > 7 Or PlotArray(ArrayIndex, 0) = " "  ' Write totals for plot
         Select Case PlotArray(ArrayIndex, 0)
           Case "Tree"
             WorkOutput!TreeL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!TreeA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Shrub", "DwarfShrub"
             WorkOutput!ShrubL = WorkOutput!ShrubL + ((PlotArray(ArrayIndex, 1) / Point_Count) * 100)
             WorkOutput!ShrubA = WorkOutput!ShrubA + ((PlotArray(ArrayIndex, 2) / Point_Count) * 100)
           Case "Forb"
             WorkOutput!ForbL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!ForbA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Annual"
             WorkOutput!AGrassL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!AGrassA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Perennial"
             WorkOutput!PGrassL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!PGrassA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Fern"
             WorkOutput!FernL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!FernA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case "Vine"
             WorkOutput!VineL = (PlotArray(ArrayIndex, 1) / Point_Count) * 100
             WorkOutput!VineA = (PlotArray(ArrayIndex, 2) / Point_Count) * 100
           Case Else
             MsgBox "Unknown lifeform " & PlotArray(ArrayIndex, 0)
         End Select
         ArrayIndex = ArrayIndex + 1
       Loop
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_CoverSpecies_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Lifeform_GS."
    Exit Sub

Err_CoverSpecies_Click:
    MsgBox Err.Description
    Resume Exit_CoverSpecies_Click
    
End Sub
Private Sub ButtonTreeSize_Click()
On Error GoTo Err_TreeSize_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Trees As DAO.Recordset
  Dim PlotSave As Variant
  Dim TransSave As String
  Dim SpeciesSave As String
  Dim NameSave As String
  Dim ArrayIndex As Integer
  Dim PlotArray(18) As Integer  ' Array for a bunch of tree counts
  ' Array for transect accumulators
  ' Index 0 is total trees
  ' Index 1 is total seedlings
  ' Index 2 is live trees 2.5-5cm
  ' Index 3 is live trees 5.1-10cm
  ' Index 4 is live trees 10.1-15cm
  ' Index 5 is live trees 15.1-20cm
  ' Index 6 is live trees 20.1-30cm
  ' Index 7 is live trees 30.1-40cm
  ' Index 8 is live trees 40.1-50cm
  ' Index 9 is live trees >50cm
  ' Index 10 is total trees 2.5-5cm
  ' Index 11 is total trees 5.1-10cm
  ' Index 12 is total trees 10.1-15cm
  ' Index 13 is total trees 15.1-20cm
  ' Index 14 is total trees 20.1-30cm
  ' Index 15 is total trees 30.1-40cm
  ' Index 16 is total trees 40.1-50cm
  ' Index 17 is total trees >50cm
  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Tree_Size_Class"
  DoCmd.SetWarnings True

  If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Date) Then
    MsgBox "Park code and visit year are required.", vbOKOnly, "Trees by Size Class"
    Exit Sub
  End If
'  Build SQL statement
  strSQL = "SELECT * FROM qry_All_Trees where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = " & Me!Visit_Date
  End If
  ' strSQL = strSQL & " AND Plot_ID = " & 1 & " AND Transect = " & 1
  strSQL = strSQL & " ORDER BY Unit_Code, Visit_Year, Plot_ID, Transect, Tree_Species"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set Trees = db.OpenRecordset(strSQL)
   If Trees.EOF Then
     MsgBox "No valid tree count records found."
     Trees.Close
     Set Trees = Nothing
     GoTo Exit_TreeSize_Click
   End If
   ArrayIndex = 0
   Do Until ArrayIndex > 17    ' clear array
     PlotArray(ArrayIndex) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   Trees.MoveFirst
   PlotSave = Trees!Unit_Code & Trees!Plot_ID     ' Save necessary fields
   TransSave = Trees!Transect
   SpeciesSave = Trees!Tree_Species
   NameSave = Trees!Stream_Name
   Set WorkOutput = db.OpenRecordset("tbl_wrk_Tree_Size_Class")
   Do Until Trees.EOF
     If (PlotSave <> Trees!Unit_Code & Trees!Plot_ID) Or (TransSave <> Trees!Transect) Or (SpeciesSave <> Trees!Tree_Species) Then  ' New species
       If PlotArray(0) > 0 Then  ' If total trees is zero, our work here is done
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = NameSave
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Transect = TransSave
         WorkOutput!Species = SpeciesSave
         WorkOutput!Total_Trees = PlotArray(0)
         WorkOutput!Total_Seedlings = PlotArray(1)
         WorkOutput!Live_5 = PlotArray(2)
         WorkOutput!Live_10 = PlotArray(3)
         WorkOutput!Live_15 = PlotArray(4)
         WorkOutput!Live_20 = PlotArray(5)
         WorkOutput!Live_30 = PlotArray(6)
         WorkOutput!Live_40 = PlotArray(7)
         WorkOutput!Live_50 = PlotArray(8)
         WorkOutput!Live_Over_50 = PlotArray(9)
         WorkOutput!All_5 = PlotArray(10)
         WorkOutput!All_10 = PlotArray(11)
         WorkOutput!All_15 = PlotArray(12)
         WorkOutput!All_20 = PlotArray(13)
         WorkOutput!All_30 = PlotArray(14)
         WorkOutput!All_40 = PlotArray(15)
         WorkOutput!All_50 = PlotArray(16)
         WorkOutput!All_Over_50 = PlotArray(17)
         '  Discontinued by DW as of 8/24/2011
         '  WorkOutput!Total_Seedlings_Pct = (PlotArray(1) / PlotArray(0)) * 100
         '  WorkOutput!Live_5_Pct = (PlotArray(2) / PlotArray(0)) * 100
         '  WorkOutput!Live_10_Pct = (PlotArray(3) / PlotArray(0)) * 100
         '  WorkOutput!Live_15_Pct = (PlotArray(4) / PlotArray(0)) * 100
         '  WorkOutput!Live_20_Pct = (PlotArray(5) / PlotArray(0)) * 100
         '  WorkOutput!Live_30_Pct = (PlotArray(6) / PlotArray(0)) * 100
         '  WorkOutput!Live_40_Pct = (PlotArray(7) / PlotArray(0)) * 100
         '  WorkOutput!Live_50_Pct = (PlotArray(8) / PlotArray(0)) * 100
         '  WorkOutput!Live_Over_50_Pct = (PlotArray(9) / PlotArray(0)) * 100
         '  WorkOutput!All_5_Pct = (PlotArray(10) / PlotArray(0)) * 100
         '  WorkOutput!All_10_Pct = (PlotArray(11) / PlotArray(0)) * 100
         '  WorkOutput!All_15_Pct = (PlotArray(12) / PlotArray(0)) * 100
         '  WorkOutput!All_20_Pct = (PlotArray(13) / PlotArray(0)) * 100
         '  WorkOutput!All_30_Pct = (PlotArray(14) / PlotArray(0)) * 100
         '  WorkOutput!All_40_Pct = (PlotArray(15) / PlotArray(0)) * 100
         '  WorkOutput!All_50_Pct = (PlotArray(16) / PlotArray(0)) * 100
         '  WorkOutput!All_Over_50_Pct = (PlotArray(17) / PlotArray(0)) * 100
         WorkOutput.Update  ' Write plot record
       End If
       PlotSave = Trees!Unit_Code & Trees!Plot_ID
       TransSave = Trees!Transect
       SpeciesSave = Trees!Tree_Species
       NameSave = Trees!Stream_Name
       ArrayIndex = 0
       Do Until ArrayIndex > 17    ' clear array
         PlotArray(ArrayIndex) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
     End If ' End if for new species test
     PlotArray(0) = PlotArray(0) + Trees!Tree_Count  ' Count total trees
     Select Case Trees!Tree_Size
       Case 1
         PlotArray(1) = PlotArray(1) + Trees!Tree_Count  ' Count tree size classes
       Case 3
         PlotArray(10) = PlotArray(10) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(2) = PlotArray(2) + Trees!Tree_Count  ' Count it as live
         End If
       Case 7
         PlotArray(11) = PlotArray(11) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(3) = PlotArray(3) + Trees!Tree_Count  ' Count it as live
         End If
       Case 13
         PlotArray(12) = PlotArray(12) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(4) = PlotArray(4) + Trees!Tree_Count  ' Count it as live
         End If
       Case 17
         PlotArray(13) = PlotArray(13) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(5) = PlotArray(5) + Trees!Tree_Count  ' Count it as live
         End If
       Case 25
         PlotArray(14) = PlotArray(14) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(6) = PlotArray(6) + Trees!Tree_Count  ' Count it as live
         End If
       Case 35
         PlotArray(15) = PlotArray(15) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(7) = PlotArray(7) + Trees!Tree_Count  ' Count it as live
         End If
       Case 45
         PlotArray(16) = PlotArray(16) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(8) = PlotArray(8) + Trees!Tree_Count  ' Count it as live
         End If
       Case 55
         PlotArray(17) = PlotArray(17) + Trees!Tree_Count
         If Trees!alive Then
           PlotArray(9) = PlotArray(9) + Trees!Tree_Count  ' Count it as live
         End If
       Case Else
         MsgBox "Undefined tree size " & Trees!Tree_Size & "."
         GoTo Exit_TreeSize_Click
     End Select
     Trees.MoveNext
   Loop
   ' Write last output record
       If PlotArray(0) > 0 Then  ' If total trees is zero, our work here is done
         WorkOutput.AddNew
         WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
         WorkOutput!Stream_Name = NameSave
         WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
         WorkOutput!Visit_Year = Me!Visit_Date
         WorkOutput!Transect = TransSave
         WorkOutput!Species = SpeciesSave
         WorkOutput!Total_Trees = PlotArray(0)
         WorkOutput!Total_Seedlings = PlotArray(1)
         WorkOutput!Live_5 = PlotArray(2)
         WorkOutput!Live_10 = PlotArray(3)
         WorkOutput!Live_15 = PlotArray(4)
         WorkOutput!Live_20 = PlotArray(5)
         WorkOutput!Live_30 = PlotArray(6)
         WorkOutput!Live_40 = PlotArray(7)
         WorkOutput!Live_50 = PlotArray(8)
         WorkOutput!Live_Over_50 = PlotArray(9)
         WorkOutput!All_5 = PlotArray(10)
         WorkOutput!All_10 = PlotArray(11)
         WorkOutput!All_15 = PlotArray(12)
         WorkOutput!All_20 = PlotArray(13)
         WorkOutput!All_30 = PlotArray(14)
         WorkOutput!All_40 = PlotArray(15)
         WorkOutput!All_50 = PlotArray(16)
         WorkOutput!All_Over_50 = PlotArray(17)
         '  WorkOutput!Total_Seedlings_Pct = (PlotArray(1) / PlotArray(0)) * 100
         '  WorkOutput!Live_5_Pct = (PlotArray(2) / PlotArray(0)) * 100
         '  WorkOutput!Live_10_Pct = (PlotArray(3) / PlotArray(0)) * 100
         '  WorkOutput!Live_15_Pct = (PlotArray(4) / PlotArray(0)) * 100
         '  WorkOutput!Live_20_Pct = (PlotArray(5) / PlotArray(0)) * 100
         '  WorkOutput!Live_30_Pct = (PlotArray(6) / PlotArray(0)) * 100
         '  WorkOutput!Live_40_Pct = (PlotArray(7) / PlotArray(0)) * 100
         '  WorkOutput!Live_50_Pct = (PlotArray(8) / PlotArray(0)) * 100
         '  WorkOutput!Live_Over_50_Pct = (PlotArray(9) / PlotArray(0)) * 100
         '  WorkOutput!All_5_Pct = (PlotArray(10) / PlotArray(0)) * 100
         '  WorkOutput!All_10_Pct = (PlotArray(11) / PlotArray(0)) * 100
         '  WorkOutput!All_15_Pct = (PlotArray(12) / PlotArray(0)) * 100
         '  WorkOutput!All_20_Pct = (PlotArray(13) / PlotArray(0)) * 100
         '  WorkOutput!All_30_Pct = (PlotArray(14) / PlotArray(0)) * 100
         '  WorkOutput!All_40_Pct = (PlotArray(15) / PlotArray(0)) * 100
         '  WorkOutput!All_50_Pct = (PlotArray(16) / PlotArray(0)) * 100
         '  WorkOutput!All_Over_50_Pct = (PlotArray(17) / PlotArray(0)) * 100
         WorkOutput.Update  ' Write plot record
       End If
       WorkOutput.Close
       Set WorkOutput = Nothing
   Trees.Close
   Set Trees = Nothing
Exit_TreeSize_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Tree_Size_Class."
    Exit Sub

Err_TreeSize_Click:
    MsgBox Err.Description
    Resume Exit_TreeSize_Click
    
End Sub
Private Sub ButtonTotalCover_Click()
On Error GoTo Err_TotalCover_Click

  Dim strSQL As String
  Dim SpeciesColumn As String
  Dim AliveColumn As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim points As DAO.Recordset
  Dim PlotSave As Variant
  Dim StreamSave As String
  Dim Point_Count As Integer
  Dim ACount As Integer
  Dim DCount As Integer
  Dim LCIndex As Integer
  Dim PointAll As Byte
  Dim PointLive As Byte
  Dim PlotTotalL As Integer
  Dim PlotTotalA As Integer

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_Cover_Pct_Totals"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_Point_Cover_Species where 1 = 1 "
  If Not IsNull(Me!Park_Code) Then
    strSQL = strSQL & " AND Unit_Code = '" & Me!Park_Code & "'"
  End If
  If Not IsNull(Me!Visit_Date) Then
    strSQL = strSQL & " AND Visit_Year = '" & Me!Visit_Date & "'"
  End If
'   strSQL = strSQL & " AND Plot_ID = 51"
  strSQL = strSQL & " ORDER BY Unit_Code, Plot_ID, Transect, Point"
  DoCmd.Hourglass True
  Set db = CurrentDb
  RecordCount = 0
  ' Get basic point info
   Set points = db.OpenRecordset(strSQL)
   If points.EOF Then
     MsgBox "No valid intercept records found."
     points.Close
     Set points = Nothing
     GoTo Exit_TotalCover_Click
   End If
   PlotTotalA = 0
   PlotTotalL = 0
   PointAll = 0
   PointLive = 0
   Point_Count = 0
   points.MoveFirst
   PlotSave = points!Unit_Code & points!Plot_ID     ' Necessary fields
   StreamSave = points!Stream_Name

   Do Until points.EOF
     If PlotSave <> points!Unit_Code & points!Plot_ID Then  ' Check for new plot code
       ' New plot - write previous plot record
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Totals")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
       PlotSave = points!Unit_Code & points!Plot_ID
       StreamSave = points!Stream_Name
       PlotTotalA = 0
       PlotTotalL = 0
       Point_Count = 0
     End If
     
     '  Top cover first
     If Not IsNull(points!Top) And points!Top <> "" And points!Top <> " " Then
       If points!alive Then
         PointLive = 1
       End If
       PointAll = 1
     End If  ' End if for null top check

     '  Soil Surface next
     If Not IsNull(points!Surface) And points!Surface <> "" And points!Surface <> " " And IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points!Surface & "'")) Then
       If points!Surface_Alive Then
         PointLive = 1
       End If
       PointAll = 1
     End If  ' End if for null soil surface check

     ' Now do lower canopy
     LCIndex = 1   ' Initialize index
     Do Until LCIndex > 10  ' Go through LC fields that I was forced to put in tbl_LC_Intercept
       SpeciesColumn = "LCS" & LCIndex ' Get the species field
       AliveColumn = "LCA" & LCIndex   ' Get the alive/dead field
       If IsNull(points(SpeciesColumn)) Or points(SpeciesColumn) = " " Then
         Exit Do  ' If we hit a null or spaces, there will be no more
       ElseIf Not IsNull(DLookup("Surface_Description", "tlu_LP_Soil_Surface", "[Surface_Code]='" & points(SpeciesColumn) & "'")) Then
         GoTo SkipEntry
       Else
         If Not IsNull(points(AliveColumn)) And points(AliveColumn) = -1 Then
           PointLive = 1
         End If
         PointAll = 1
       End If  ' End if for null lower canopy check
SkipEntry:
       LCIndex = LCIndex + 1
     Loop
   ' Now update plot totals
     If PointAll = 1 Then
       PlotTotalA = PlotTotalA + 1  ' accumulate total all
     End If
     If PointLive = 1 Then
       PlotTotalL = PlotTotalL + 1  ' accumulate total live
     End If
     PointLive = 0  ' Clear hit indicators
     PointAll = 0   ' For the next point
     points.MoveNext
     Point_Count = Point_Count + 1
   Loop
   ' Output last plot record
    '   MsgBox "A=" & [PlotTotalL] & "  T=" & [PlotTotalA]
       Set WorkOutput = db.OpenRecordset("tbl_wrk_Cover_Pct_Totals")
       WorkOutput.AddNew
       WorkOutput!unitCode = Left(PlotSave, 4)  ' Set unit code
       WorkOutput!Visit_Year = Me!Visit_Date
       WorkOutput!Stream_Name = StreamSave
       WorkOutput!plotID = Right(PlotSave, Len(PlotSave) - 4)  ' Set plot ID
       WorkOutput!Total_Live = (PlotTotalL / Point_Count) * 100
       WorkOutput!Total_Cover = (PlotTotalA / Point_Count) * 100
       WorkOutput.Update  ' Write plot record
       WorkOutput.Close
       Set WorkOutput = Nothing
   points.Close
   Set points = Nothing
Exit_TotalCover_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_Cover_Pct_Totals."
    Exit Sub

Err_TotalCover_Click:
    MsgBox Err.Description
    Resume Exit_TotalCover_Click
    
End Sub
