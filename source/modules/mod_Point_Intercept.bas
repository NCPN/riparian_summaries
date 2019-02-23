Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Point_Intercept
' VERSION:      1.03
' Description:  Point-intercept functions & subroutines
'
' Source/date:  Bonnie Campbell, 8/8/2014
' Revisions:    BLC, 8/8/2014   - 1.00 - initial version
'               BLC, 8/11/2014  - 1.01 - renamed module to mod_Point_Intercept from mod_Upland to accommodate multiple
'                                        project types (riparian, upland, ...)
'               BLC, 8/12/2014  - 1.02 - replaced mod_Riparian with mod_Point_Intercept (contains same functions, but
'                                        changed to accommodate upland project type)
'               BLC, 8/19/2014  - 1.03 - removed unused variables (k, numRecords) from getLifeformCounts,
'                                        added versioning, added writeCountsToTable fern skipping logic
' =================================

' ---------------------------------
' SUB:          getLifeformCounts
' Description:  Determine count of alive/total lifeform based on species & write it
'               to tbl_wrk_Cover_Pct_Lifeform or similar table for point intercept data (standard/greenline(GL))
' Parameters:   parkCode
'
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Upland tools
' Revisions:    BLC, 7/24/2014 - initial version
'               BLC, 7/30/2014 - removed aryLive2 & aryALL2 which were extra arrays since mergeArrays
'                                changes original arrays (aryLive & aryALL) which are passed ByRef by default
'                                fixed aryPctCover upper bound (UBound + 1 to accommodate zero-based array)
'               BLC, 8/6/2014  - added SOP type parameter to accommodate GL intercept calculations
'               BLC, 8/11/2014 - added projectType parameter to accommodate upland vs riparian calculations
'                                so that only one code source is necessary for point intercept calculations
'               BLC, 8/12/2014 - added iShiftIndex parameter to accommodate upland vs riparian arrays
'                                (index values differ since riparian doesn't include primary ecological site (PrimEcoSite))
' ---------------------------------
Public Sub getLifeformCounts(parkCode As String, visitYear As Integer, sopType As String, projectType As String)
On Error GoTo Err_Handler:
'===================================================
' aryTotalPoints array contains:
'  upland  riparian*
'   0       0       - Unit_Code (park name - ZION, ARCH, etc.)
'   1       1       - Stream_Name (riparian), Master_Stratification (upland)
'   2       2       - Visit_Year
'   3       -       - Primary_Eco_Site (upland) - unused (NULL) in riparian
'   4       3       - Plot_ID (reach)
'   5       4       - Total # points in the plot
'
' aryLive & aryALL arrays contain:
'  upland  riparian*
'   0       0       - Unit_Code (park name - ZION, ARCH, etc.)
'   1       1       - Stream_Name (riparian), Master_Stratification (upland)
'   2       2       - Visit_Year
'   3       -       - Primary_Eco_Site (upland) - unused (NULL) in riparian
'   4       3       - Plot_ID (reach)
'   5       4       - Lifeform_Type (Tree, Shrub, PGrass, AGrass, Forb, Fern, Vine)
'   6       5       - Count (# hits tallied per lifeform either Live or ALL (live & dead) )
'
' aryPercentCover array contains:
'   0 - Unit_Code (park name - ZION, ARCH, etc.)
'   1 - Stream_Name (riparian), Master_Stratification (upland)
'   2 - Visit_Year
'   3 - Primary_Eco_Site (upland) - unused (NULL) in riparian
'   4 - Plot_ID (reach)
'   5 - TotalPoints (# points per reach)
'   6 - Lifeform_Type (Tree, Shrub, PGrass, AGrass, Forb, Fern, Vine - A & L for each)
'   7 - Count (# hits tallied per lifeform either Live or ALL (live & dead) )
'   8 - Percent Cover ( (#hits/total pts) * 100 )
'
' * iIndexShift handles the difference in index between project types for these values
'
' These arrays are helpful for quick troubleshooting of values if needed.
'===================================================
    Dim rstTotalPoints As DAO.Recordset, rstLive As DAO.Recordset, rstALL As DAO.Recordset
    Dim aryTotalPoints As Variant, aryLive As Variant, aryALL As Variant, aryMerged As Variant
    Dim samplesite As Integer, j As Integer, iIndexShift As Integer
    Dim sop As String, tbl As String, liveQuery As String, allQuery As String, _
     ptsQuery As String, siteType As String
    
    'set values based on project type
    Select Case projectType
        Case "riparian"
            siteType = "Reach"
            iIndexShift = -1
        Case "upland"
            siteType = "Plot"
            iIndexShift = 0
    End Select
    
    'set queries & tables based on SOP being run
    Select Case sopType
        Case "GL"   'greenline
            sop = sopType & "_"
            ptsQuery = Replace("qry_Point_Cover_Species_Points_Per_" & siteType, "Species", "GL")
        Case Else   'standard point intercept
            sop = ""
            ptsQuery = "qry_Point_Cover_Species_Points_Per_" & siteType
    End Select
    
    tbl = "tbl_wrk_" & sop & "Cover_Pct_Lifeform"
    liveQuery = "qry_NORM_" & sop & "Lifeform_Transect_Pt_Alive_Counts"
    allQuery = "qry_NORM_" & sop & "Lifeform_Transect_Pt_ALL_Counts"
    
    'get total points
    Set rstTotalPoints = runLifeformQuery(ptsQuery, projectType, parkCode, visitYear)
      
    'check if rstTotalPoints is a valid recordset otherwise, exit (no points found for the reach)
    If rstTotalPoints.EOF And rstTotalPoints.BOF Then
        Exit Sub
    End If
       
    'get hit tallies
    Set rstLive = runLifeformQuery(liveQuery, projectType, parkCode, visitYear)
    Set rstALL = runLifeformQuery(allQuery, projectType, parkCode, visitYear)
    
    'convert to arrays
    aryTotalPoints = rstTotalPoints.GetRows(rstTotalPoints.RecordCount)
    aryLive = rstLive.GetRows(rstLive.RecordCount)
    aryALL = rstALL.GetRows(rstALL.RecordCount)
    
    'prep & merge the arrays
    'aryLive & aryALL are passed ByRef by default, so they are changed when renameArrayLifeforms runs
    renameArrayLifeforms aryLive, 5 + iIndexShift, "L" 'add "L" to lifeform types
    renameArrayLifeforms aryALL, 5 + iIndexShift, "A"  'add "A" to lifeform types
    
    'create an array that includes BOTH Live & All for the same values of 0,1,2,3,4,5(,6)
    aryMerged = mergeArrays(aryLive, aryALL, 2, 6 + iIndexShift)
    
    '-------------------------
    ' prepare an array that contains total points, live & ALL lifeform percentages for each reach
    '-------------------------
    ' statusbar feedback for user
    SysCmd acSysCmdSetStatus, "Calculating percent cover... "
    
    'set dimensions
    Dim aryPercentCover() As Variant
    'set array dimensions & preserve values -> add one for zero-based array
    ReDim Preserve aryPercentCover(8, (UBound(aryLive, 2) + UBound(aryALL, 2) + 1)) As Variant

    ' -----------------------
    '  lifeform type counts
    ' -----------------------
    'add the reach lifeform values from the matching merged array
    For j = 0 To UBound(aryMerged, 2)
        'add the common parameters to the combined array (using same positional values as LIVE & ALL arrays)
         aryPercentCover(0, j) = aryMerged(0, j) ' Unit_Code
         aryPercentCover(1, j) = aryMerged(1, j) ' Stream_Name (riparian), Master_Stratification (upland)
         aryPercentCover(2, j) = aryMerged(2, j) ' Visit_Year

         Select Case projectType
            Case "riparian"
                aryPercentCover(3, j) = Null ' Primary_Eco_Site (upland) - unused in riparian
            Case "upland"
                aryPercentCover(3, j) = aryMerged(3, j) ' Primary_Eco_Site (upland) - unused in riparian
         End Select
         
         aryPercentCover(4, j) = aryMerged(4 + iIndexShift, j) ' Plot_ID (reach)
        
        ' --------------------------------------------
        '  samplesite - reach (riparian)/plot(upland)
        ' --------------------------------------------
        'iterate through the plot - if Unit_Code, Stream_Name/Master_Stratification, Visit_Year, & Plot_ID (reach/plot) match
        'then fetch the total # points for that reach/plot
        For samplesite = 0 To UBound(aryTotalPoints, 2)
          
            If aryMerged(0, j) = aryTotalPoints(0, samplesite) And _
                aryMerged(1, j) = aryTotalPoints(1, samplesite) And _
                aryMerged(2, j) = aryTotalPoints(2, samplesite) And _
                aryMerged(4 + iIndexShift, j) = aryTotalPoints(4 + iIndexShift, samplesite) Then
            
                aryPercentCover(5, j) = aryTotalPoints(5 + iIndexShift, samplesite) ' TotalPoints
                
                Exit For
                
            End If
        
        Next samplesite
        
         'add the lifeform, hit count & percentage cover
         aryPercentCover(6, j) = aryMerged(5 + iIndexShift, j) 'lifeform type
         aryPercentCover(7, j) = aryMerged(6 + iIndexShift, j) 'hit count
         aryPercentCover(8, j) = calculatePercentCover(CInt(aryMerged(6 + iIndexShift, j)), CInt(aryPercentCover(5, j))) 'lifeform type % cover
    
    Next j
             
    'prepare pct counts table (tbl_wrk_Cover_Pct_Lifeform)
    writeCountsToTable aryPercentCover, sopType, projectType
    
    'cleanup
    Set rstTotalPoints = Nothing
    Set rstLive = Nothing
    Set rstALL = Nothing
    'release memory allocated to arrays
    DeallocateArrays Array(aryLive, aryALL, aryMerged, aryTotalPoints, aryPercentCover)

    ' clear statusbar
    SysCmd acSysCmdSetStatus, " "

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getSpeciesLifeformCounts[mod_Point_Intercept])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:     writeCountsToTable
' Description:  Enters the lifeform counts & percent cover into working table
'
' Note:         This denormalizes the arrays back to fit the working table's columnar structure.
'               As a result it iterates to write the first values in the percent cover array (aryPctCover)
'               When it encounters a new 'currentRecord'* that is different from the 'lastRecord'
'               it checks to see if that record already exists in the database. If so, it is edited
'               to add the additional lifeform type percentages. If no existing record is found,
'               a new one is added and the percentage values are entered into it.
'
'               * Records are identified by the unique primary key combination of
'                 UnitCode | StreamName/MasterStratification | VisitYear | PlotID
'                 which is a concatenation of the 4 of the first 5 values of the percent cover array.
'
'               When iterating through the aryPctCover, array values (rows) are skipped when there
'               are no points (aryPctCover(5,j) is empty).
'               This would occur when pulling values for all parks, streams/masterstratifications, years, & plots (reaches/plots)
'               when hits weren't recorded or points weren't identified for the reach/plot.
'               However, the latter would also produce rstTotalPoints (in getLifeformCounts) with
'               no records which is already trapped for prior to reaching this point.
'
'               The rstCover recordset is a table type (dbOpenTable), therefore the .Seek method is used to find
'               unique record combinations that already exist when writing to the working table.
'               This requires identification of the index and unique values to find within the table
'
' Parameters:   array - array of lifeform count & percent cover records to write
' Assumptions:  Lifeform count/percent cover array contains no NULL/Empty lifeform values
' Returns:      -
' Throws:       -
' References:   -
' Note:
' Source/date:  Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Adapted:      Bonnie Campbell, August, 2014 for NCPN Upland tools
' Revisions:    BLC, 7/24/2014 - initial version
'               BLC, 8/6/2014  - add sopType parameter to accommodate GL and other SOP writes to various tables
'               BLC, 8/11/2014 - added projectType parameter to accommodate upland vs riparian calculations
'                                so that only one code source is necessary for point intercept calculations
'               BLC, 8/19/2014 - added ferns to vine skipping logic since these are also skipped for uplands
' ---------------------------------
Public Sub writeCountsToTable(aryPctCover As Variant, sopType As String, projectType As String)
On Error GoTo Err_Handler:
    Dim currentRecord As String, lastRecord As String
    Dim aryFields As Variant, aryLiveFields As Variant, aryALLFields As Variant, aryMergedFields As Variant
    Dim rstCover As DAO.Recordset
    Dim vBookmark As Variant
    Dim blnNew As Boolean, blnSkipVinesAndFerns As Boolean
    Dim i As Integer, j As Integer
    Dim Response As String, sop As String, tbl As String, qryClear As String
    
    'default
    blnSkipVinesAndFerns = False

    'set queries & tables based on SOP being run
    Select Case sopType
        Case "GL"   'greenline
            sop = "_" & sopType
        Case Else   'standard point intercept
            sop = ""
    End Select

    tbl = "tbl_wrk" & sop & "_Cover_Pct_Lifeform"
    'If Len(sop) > 0 Then sop = "_" & Replace(sop, "_", "") 'flip underscore
    qryClear = "qry_Clear" & sop & "_Cover_Pct_Lifeform"

    If UBound(aryPctCover, 2) < 1 Then
        MsgBox "Sorry, no valid intercept records were found" & vbCrLf & vbCrLf & _
                "for the park and year(s) selected.", vbOKOnly, _
                "No Records for " & StrConv(projectType, vbProperCase) & " " & Replace(sop, "_", " ") & "Lifeform Percent Cover"
        GoTo Exit_Sub
    End If
    
    ' statusbar feedback for user
    SysCmd acSysCmdSetStatus, "Writing values to " & tbl & "... "
    
    'fields to populate
    aryFields = Array("Tree", "Shrub", "PGrass", "AGrass", "Forb", "Fern", "Vine")
    'prep & merge Live & ALL lifeform type field arrays
    aryLiveFields = aryFields 'since using aryFields twice need to separate so both suffixes aren't added for aryALLFields
    aryLiveFields = addArrayValueSuffix(aryLiveFields, "L") 'add "L" to lifeform types
    aryALLFields = addArrayValueSuffix(aryFields, "A")  'add "A" to lifeform types
    'create an array that includes BOTH Live & All fields
    aryMergedFields = mergeArrays(aryLiveFields, aryALLFields)
    
    'clear & open the table
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    DoCmd.OpenQuery qryClear
    Set rstCover = dbCurrent.OpenRecordset(tbl, dbOpenTable)

    With rstCover

        '------------------------------------
        ' lifeform type counts & percentages
        '------------------------------------
        For j = 0 To UBound(aryPctCover, 2)

            'default
            blnNew = False
   
            'prepare unique combined key
            Select Case projectType
                Case "riparian"
                    'combine Unit_Code|Stream_Name|Visit_Year|Plot_ID to form a unique record identifier
                    currentRecord = aryPctCover(0, j) & "|" & aryPctCover(1, j) & "|" & aryPctCover(2, j) & "|" & aryPctCover(4, j)
                Case "upland"
                    'combine Unit_Code|Master_Stratification|Visit_Year|PrimEcoSite|Plot_ID to form a unique record identifier
                    currentRecord = aryPctCover(0, j) & "|" & aryPctCover(1, j) & "|" & aryPctCover(2, j) & "|" & aryPctCover(3, j) & "|" & aryPctCover(4, j)
                    'skip writing ferns/vines to table (fields don't exist in upland working table)
                    blnSkipVinesAndFerns = True
            End Select
'------
'Debug.Print j & " - " & currentRecord & " - " & aryPctCover(6, j)
'------
            'each unit, stream/stratification, year & plot will have its own record in rstCover
            If currentRecord <> lastRecord Or j = 0 Then
                'add the record data (unless it's the first record, then add a new record)
                If j <> 0 Then
                    .Update
                    .Bookmark = .LastModified
                    vBookmark = .LastModified
                End If
           
                If .RecordCount <> 0 Then
                    'store the current position (last record)
                    .MoveLast
                    vBookmark = .Bookmark
                    
                    If .Bookmarkable Then
                        .MoveFirst
                        'ensure the same record doesn't already exist
                        .Index = "PrimaryKey"
                        
                        'set the primary key values based on project type
                        Select Case projectType
                            Case "riparian" 'Unit_Code, Stream_Name, Visit_Year, Plot_ID
                                .Seek "=", aryPctCover(0, j), aryPctCover(1, j), _
                                           aryPctCover(2, j), aryPctCover(4, j)
                            Case "upland"   'Unit_Code, Master_Stratification, Plot_ID, Visit_Year, Primary_Eco_Site
                                .Seek "=", aryPctCover(0, j), aryPctCover(1, j), _
                                           aryPctCover(4, j), aryPctCover(2, j), _
                                           aryPctCover(3, j)
                        End Select
                                                                               
                        'add record if no match is found and the array value has points (not empty)
                        If .NoMatch And Not IsEmpty(aryPctCover(5, j)) Then
                            .AddNew
                            blnNew = True
                            '.Bookmark = vBookmark
                        Else
                            'record exists - add values to it
                            vBookmark = .Bookmark
                            .Edit
                        End If
                    End If
                Else
                    'add record if the array value has points (not empty)
                    If Not IsEmpty(aryPctCover(5, j)) Then
                        'add new record
                        .AddNew
                        blnNew = True
                    End If
                End If

            End If

            '------------------------------------
            ' add reach/plot identifiers (unitCode, streamName/MasterStratification, VisitYear, plotID)
            '------------------------------------
            If blnNew Then
                     'add reach identifying values
                    ![unitCode] = aryPctCover(0, j)
                    
                    'handle various project types
                    Select Case projectType
                        Case "riparian"
                            ![Stream_Name] = aryPctCover(1, j)
                        Case "upland"
                            ![Master_Stratification] = aryPctCover(1, j)
                            ![PrimEcoSite] = aryPctCover(3, j)
                    End Select

                    ![Visit_Year] = aryPctCover(2, j)
                    ![plotID] = aryPctCover(4, j)
            End If
            
            '------------------------------------
            ' get counts & percentages for all fields
            '------------------------------------
            'handle NULL lifeform types
            If Not IsNull(aryPctCover(6, j)) Then
                For i = LBound(aryMergedFields) To UBound(aryMergedFields)
                    
                    'skip vines (uplands)
                    If blnSkipVinesAndFerns And _
                        (Left(aryPctCover(6, j), 4) = "Vine" Or Left(aryPctCover(6, j), 4) = "Fern") Then Exit For
                    
                    'add values for all lifeform type fields
                    If aryMergedFields(i) = aryPctCover(6, j) Then
                        
                        'add the percent cover
                        rstCover(aryMergedFields(i)) = aryPctCover(8, j)
                        
                        'go to next array value
                        Exit For
    
                    End If
                Next i
             End If
             
            'prepare for the next iteration
            lastRecord = currentRecord
            
        Next j
        
        'update last record
        .Update

        ' clear statusbar
        SysCmd acSysCmdSetStatus, "Calculations complete!"

        'present user with the choice of viewing the records or not
         Response = MsgBox("Finished calculating!" & vbCrLf & vbCrLf & _
                           "Do you want to view your results in the " & _
                           vbCrLf & vbCrLf & tbl & " table?" & _
                           vbCrLf & vbCrLf & "If not, they'll be there until you run this calculation again.", _
                           vbYesNo, StrConv(projectType, vbProperCase) & " " & Replace(sop, "_", "") & " Lifeform Percent Cover Complete!")
         
         If Response = vbYes Then    ' User chose Yes.
            'open the table "tbl_wrk_Cover_Pct_Lifeform"
            DoCmd.OpenTable tbl, acViewNormal, acEdit
         Else    ' User chose No.
            'do nothing
         End If

    End With

    'cleanup
    Set rstCover = Nothing
    'release memory allocated to arrays
    DeallocateArrays Array(aryPctCover, aryFields, aryLiveFields, aryALLFields, aryMergedFields)

Exit_Sub:
    DoCmd.Hourglass False
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3022
        Dim strErr As String
        strErr = vbCrLf & j & "-" & currentRecord _
                & vbCrLf & aryPctCover(1, j) & " " _
                & vbCrLf & aryPctCover(2, j) & " " _
                & vbCrLf & aryPctCover(3, j) & " " _
                & vbCrLf & aryPctCover(4, j) & " " _
                & vbCrLf & aryPctCover(5, j) & " " _
                & vbCrLf & aryPctCover(6, j) & " " _
                & vbCrLf & aryPctCover(7, j) & " " _
                & vbCrLf & aryPctCover(8, j) & " "
        Debug.Print strErr
        MsgBox "Sorry, but the following record cannot be added:" _
                & strErr, vbCritical, "Oops! Duplicate Record"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - writeCountsToTable[mod_Point_Intercept])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     runLifeformQueries
' Description:  Run the lifeform count or points queries for a given reach/transect
'               for specific park & visit year (or ALL) & return the recordset
'
' Note:         When adding queries to the list, ensure all queries have a "WHERE 1=1" clause
'               in their SQL statement otherwise the generated WHERE clause limiting to
'               specific parks, years, plots, streams or transects will not be run
'               since it replaces the existing WHERE 1=1 clause
'
' Parameters:   qry - query name (string)
'               optional parameters (if not included returns records for all parks, years, & plots
'               parkCode - 4 character park code (e.g. ARCH)
'               visitYear - year field data was collected
'               streamName - stream where field data was collected
'               plotID - plot (reach) of interest (integer)
' Returns:      recordset from query (if query produces one) or records affected by query (if no recordset produced)
' Throws:       -
' References:   -
' Note:
' Source/date:  Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Adapted:      -
' Revisions:    BLC, 7/22/2014 - initial version
' ---------------------------------
Public Function runLifeformQuery(qry As String, projectType As String, _
                                  Optional parkCode As String, _
                                  Optional visitYear As Integer, _
                                  Optional streamName As String, _
                                  Optional plotID As Integer, _
                                  Optional stratification As String, _
                                  Optional transectID As Integer) As DAO.Recordset
On Error GoTo Err_Handler:
'-------------------
' Included queries:
'
' POINT-INTERCEPT (RIPARIAN & UPLAND):
'   qry_NORM_Lifeform_Transect_Pt_Alive_Counts
'   qry_NORM_Lifeform_Transect_Pt_ALL_Counts
'
' RIPARIAN:
'   qry_Point_Cover_Species_Points_Per_Reach
'
' RIPARIAN (GREENLINE):
'   qry_NORM_GL_Lifeform_Transect_Pt_Alive_Counts
'   qry_NORM_GL_Lifeform_Transect_Pt_ALL_Counts
'   qry_Point_Cover_GL_Points_Per_Reach

' UPLAND:
'   qry_Point_Cover_Species_Points_Per_Plot
'-------------------

    Dim strSQL As String, strWhere As String
  
    'default
    strWhere = ""
    
    ' filter by park, year, & plot if desired
    If Not IsNull(parkCode) And Len(Trim(parkCode)) > 0 Then
        strWhere = " Unit_Code = '" & parkCode & "'"
        'park is required for transect
        If Not IsNull(transectID) And transectID > 0 Then
            strWhere = strWhere & " AND Transect = '" & transectID & "'"
            'park & transect is required for plot
            If Not IsNull(plotID) And plotID > 0 Then
                strWhere = strWhere & " AND plot_ID = '" & plotID & "'"
            End If
        End If
    End If
    
    'prepare parameters & generate SQL WHERE statement
    Dim params() As Variant
    ReDim params(1, 2)
    'handle visit year (if present)
    If visitYear > 0 Then
        params(0, 0) = visitYear
        params(0, 1) = "integer"
        params(0, 2) = "Visit_Year"
    End If
    
    'add project type specific fields (if present)
    Select Case projectType
        Case "upland"
            'ReDim Preserve params(2, UBound(params, 2) + 1) -- CAN ONLY REDIM OUTER LARGER BOUND?
            If Len(stratification) > 0 Then
                params(1, 0) = stratification
                params(1, 1) = "string"
                params(1, 2) = "Master_Stratification"
            End If
        Case "riparian"
            'ReDim Preserve params(2, UBound(params, 2) + 1) -- CAN ONLY REDIM OUTER LARGER BOUND?
            If Len(streamName) > 0 Then
                params(1, 0) = streamName
                params(1, 1) = "string"
                params(1, 2) = "Stream_Name"
            End If
    End Select
    
    strWhere = getWhereSQL(strWhere, params)
   
    'choose the desired query to run
    Select Case qry
      Case "qry_NORM_Lifeform_Transect_Pt_Alive_Counts", _
           "qry_NORM_Lifeform_Transect_Pt_ALL_Counts", _
           "qry_Point_Cover_Species_Points_Per_Reach", _
           "qry_NORM_GL_Lifeform_Transect_Pt_Alive_Counts", _
           "qry_NORM_GL_Lifeform_Transect_Pt_ALL_Counts", _
           "qry_Point_Cover_GL_Points_Per_Reach", _
           "qry_Point_Cover_Species_Points_Per_Plot"
          strSQL = getSQL(qry)
      Case Else
        'no existing SQL query identified
        GoTo Exit_Function
    End Select
    
    'filter the query if desired
    If Len(strWhere) > 0 Then
      'remove the standard SQL WHERE clause 1=1 that brings back all records
      '& replace with filtered version
      strSQL = Replace(strSQL, "1=1", strWhere)
    End If
'Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.Hourglass True
    SysCmd acSysCmdSetStatus, "Running " & qry & " query... "

    ' Get basic point info
     Set runLifeformQuery = dbCurrent.OpenRecordset(strSQL)
     If runLifeformQuery.EOF Then
        MsgBox "Sorry, no valid intercept records were found" & vbCrLf & vbCrLf & _
          "for the park(s)/year(s) selected when running the query.", vbOKOnly, _
          "No Records for " & projectType & " Lifeform Percent Cover"
      GoTo Exit_Function
     End If
    
    DoCmd.SetWarnings True
    SysCmd acSysCmdSetStatus, " "
    
    'cleanup
       
Exit_Function:
    DoCmd.Hourglass False
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - runLifeformQueries[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     calculatePercentCover
' Description:  Determine percent cover based on field hits
'                   % Cover = ( # hits / total # pts ) * 100
' Parameters:
'               numHits - # of hits for sample location (e.g. plot or transect)
'               totalPoints - total # of points for same sample location (plot/transect)
' Returns:      decimal percentage value
' Throws:       -
' References:   -
' Note:
' Source/date:  Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Adapted:      -
' Revisions:    BLC, 7/24/2014 - initial version
' ---------------------------------
Public Function calculatePercentCover(numHits As Integer, totalPoints As Integer) As Double

On Error GoTo Err_Handler:

    If totalPoints > 0 Then
        'calcuate percent cover
        calculatePercentCover = (numHits / totalPoints) * 100
    Else
        'Skip
        'MsgBox "Total number of points is not > 0.", vbOKOnly, "Oops! Number of Points Is Zero"
    End If
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - calculatePercentCover[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     renameArrayLifeforms
' Description:  Add an "L" or "A" suffix to lifeform type values
' Parameters:   array - lifeform array (live or ALL)
'               suffix - string to add to the lifeform type
' Returns:      array with lifeforms altered to (ForbA, ForbL, etc.)
' Throws:       -
' References:   -
' Note:
' Source/date:  Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Adapted:      -
' Revisions:    BLC, 7/24/2014 - initial version
'               BLC, 8/11/2014 - changed lifeform array index to reflect aryLive & aryALL lifeform array position change
'                                due to handling multiple project types (upland & riparian) for point-intercept data
'               BLC, 8/12/2014 - added array index parameter (idx) to handle difference in lifeform type index
'                                position between project types (riparian/upland)
' ---------------------------------
Public Function renameArrayLifeforms(ary As Variant, idx As Integer, suffix As String) As Variant
On Error GoTo Err_Handler:
'--------------
' aryLive & aryALL arrays contain:
'  upland  riparian
'   0       0       - Unit_Code (park name - ZION, ARCH, etc.)
'   1       1       - Stream_Name (riparian), Master_Stratification (upland)
'   2       2       - Visit_Year
'   3       -       - Primary_Eco_Site (upland) - unused (NULL) in riparian
'   4       3       - Plot_ID (reach)
'   5       4       - Lifeform_Type (Tree, Shrub, PGrass, AGrass, Forb, Fern, Vine)
'   6       5       - Count (# hits tallied per lifeform either Live or ALL (live & dead) )
'--------------
Dim i As Integer
    
    For i = 0 To UBound(ary, 2)
        If Not IsNull(ary(idx, i)) Then
            'rename lifeform
            ary(idx, i) = ary(idx, i) & suffix
        End If
    Next i
    
    renameArrayLifeforms = ary
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - renameArrayLifeforms[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function