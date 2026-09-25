Option Explicit

' Define a Public Constant for the Log Sheet Name from the prototype
Public Const LOG_SHEET_NAME As String = "Log"

Sub genfile_ByReviewer_V11_WithLog()

    ' --- [Index 1]: Variable Declarations ---
    ' Declare variables for timing the execution
    Dim StartTime As Double
    Dim endTime As Double
    Dim runTime As Double

    ' Declare worksheet objects for easier reference
    Dim wsMasterMapping As Worksheet        ' Worksheet object for "MasterMapping"
    Dim wsTemp As Worksheet                 ' Temporary worksheet object for cleaning up default sheets
    Dim wsNewWorkbookSheet As Worksheet     ' Worksheet object in the newly created workbook
    Dim wsOriginalTemplate As Worksheet     ' Worksheet object for the original template sheet
    Dim wsLog As Worksheet                  ' Worksheet object for logging

    ' Declare variables to store data in arrays for faster processing
    Dim arrMasterMapping As Variant
    Dim arrMasterData As Variant            ' Array to store data from Master (M) sheets
    Dim arrCurrentSheetFinalData As Variant ' Array to hold final data for the current sheet before writing
    Dim arrTemplateFormulas As Variant      ' Array to hold formulas from the original template sheet

    ' Declare Dictionary object for faster lookups of Master (M) data
    Dim dictAllMasterData As Object
    Set dictAllMasterData = CreateObject("Scripting.Dictionary")

    ' Declare variables for storing sheet mappings and loop counters
    Dim dictSheetTemplates As Object
    Set dictSheetTemplates = CreateObject("Scripting.Dictionary")

    ' Dictionary to group Profit Centers by Reviewer
    Dim dictReviewerFiles As Object
    Set dictReviewerFiles = CreateObject("Scripting.Dictionary")

    Dim i As Long, j As Long, k As Long
    Dim dataRowCounter As Long
    Dim col As Long

    ' Declare variables for file operations
    Dim newFilePath As String
    Dim folderPath As String
    Dim fileName As String
    Dim newWorkbook As Workbook
    Dim baseFileName As String

    ' Variables for progress and time tracking on the UserForm
    Dim frmProgress As New frmProgress
    Dim totalFilesToProcess As Long
    Dim filesProcessed As Long
    Dim currentTime As Double
    Dim minutes As Long
    Dim seconds As Long
    ' --- End [Index 1] ---

    ' --- [Index 2]: Initialize FileSystemObject and Application Settings ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    StartTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' --- LOGGING: Start of Macro ---
    Call WriteLog("INFO", "Macro started: genfile_ByReviewer_V10_WithLog")

    ' --- Prepare Log Sheet ---
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If Not wsLog Is Nothing Then
        ' Check if there's data beyond headers before clearing
        If wsLog.Cells(Rows.Count, "A").End(xlUp).Row > 1 Then
            wsLog.Range("A2:E" & wsLog.Cells(Rows.Count, "A").End(xlUp).Row).ClearContents
            Call WriteLog("INFO", "Cleared existing data in Log sheet.")
        End If
    End If

    
            Dim wsActionPage As Worksheet
            Dim sheetName As String

        ' ???????? wsActionPage ??????? "Action Page"
            Set wsActionPage = ThisWorkbook.Sheets("Action_Page")

        ' ????????????? Cell A6 ?????? Action Page
        sheetName = wsActionPage.Range("A6").Value

        ' ????? wsMasterMapping ???????????????????????????? sheetName
        Set wsMasterMapping = ThisWorkbook.Sheets(sheetName)
        'Set wsMasterMapping = ThisWorkbook.Sheets("MasterMapping")
    
    ' --- End [Index 2] ---

    ' --- [Index 3]: Load MasterMapping Data ---
    Dim lastRowMasterMapping As Long
    lastRowMasterMapping = wsMasterMapping.Cells(Rows.Count, "A").End(xlUp).Row
    If lastRowMasterMapping < 5 Then ' Modified to check from row 5
        MsgBox "MasterMapping sheet is empty or has only headers.", vbExclamation
        Call WriteLog("WARNING", "MasterMapping sheet is empty or has only headers.", "lastRowMasterMapping", lastRowMasterMapping)
        GoTo CleanUp
    End If
    arrMasterMapping = wsMasterMapping.Range("A5:AR" & lastRowMasterMapping).Value
    Call WriteLog("INFO", "MasterMapping data loaded into array.", "Rows Loaded", UBound(arrMasterMapping, 1))
    ' --- End [Index 3] ---

    ' --- [Index 4]: Identify Sheet Templates from MasterMapping ---
    Dim lastColSheets As Long
    Dim masterSourceSheetName As String
    Dim templateSheetName As String
    Dim createdSheetName As String

    On Error Resume Next
    lastColSheets = wsMasterMapping.Cells(2, Columns.Count).End(xlToLeft).Column
    If lastColSheets < 4 Then lastColSheets = 3 ' Ensure at least column D (index 4 in 1-based) is considered for sheet mappings
    On Error GoTo 0

    For j = 4 To lastColSheets
        masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, j).Value))
        templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, j).Value))

        If Len(masterSourceSheetName) > 0 And Len(templateSheetName) > 0 Then
            createdSheetName = masterSourceSheetName
            If InStr(1, createdSheetName, " (M)", vbTextCompare) = 0 Then
                createdSheetName = createdSheetName & " (M)"
            End If

            If Len(createdSheetName) > 0 Then
                If Not dictSheetTemplates.Exists(createdSheetName) Then
                    dictSheetTemplates.Add createdSheetName, templateSheetName
                End If
            End If
        End If
    Next j

    If dictSheetTemplates.Count = 0 Then
        MsgBox "No valid sheet mappings found in MasterMapping (Row 2 & 3, Columns D-AR).", vbExclamation
        Call WriteLog("WARNING", "No valid sheet mappings found in MasterMapping.", "dictSheetTemplates.Count", dictSheetTemplates.Count)
        GoTo CleanUp
    End If
    Call WriteLog("INFO", "Identified sheet templates and mappings.", "Total Mappings", dictSheetTemplates.Count)
    ' --- End [Index 4] ---

    ' --- [Index 5]: Pre-populate Dictionary with Master (M) Data ---
    Dim wsMasterSource As Worksheet
    Dim lastRowMasterSource As Long
    Dim lastColMasterSource As Long
    Dim masterKey As String
    Dim masterRowData_1D As Variant
    Dim masterRowsForProfitCenter As Collection
    Dim dictSingleMasterGrouped As Object

    For Each wsMasterSource In ThisWorkbook.Sheets
        If InStr(1, wsMasterSource.Name, " (M)", vbTextCompare) > 0 Then
            lastRowMasterSource = wsMasterSource.Cells(Rows.Count, "B").End(xlUp).Row
            If lastRowMasterSource >= 2 Then
                lastColMasterSource = wsMasterSource.Cells(1, wsMasterSource.Columns.Count).End(xlToLeft).Column
                If lastColMasterSource < 1 Then lastColMasterSource = 1
                arrMasterData = wsMasterSource.Range(wsMasterSource.Cells(2, "A"), wsMasterSource.Cells(lastRowMasterSource, lastColMasterSource)).Value
                Set dictSingleMasterGrouped = CreateObject("Scripting.Dictionary")
                For k = LBound(arrMasterData, 1) To UBound(arrMasterData, 1)
                    masterKey = Trim(CStr(arrMasterData(k, 2))) ' Assuming Profit Center ID is in Column B of Master sheets
                    ReDim masterRowData_1D(1 To UBound(arrMasterData, 2))
                    For col = LBound(arrMasterData, 2) To UBound(arrMasterData, 2)
                        masterRowData_1D(col) = arrMasterData(k, col)
                    Next col
                    If Not dictSingleMasterGrouped.Exists(masterKey) Then
                        Set masterRowsForProfitCenter = New Collection
                        dictSingleMasterGrouped.Add masterKey, masterRowsForProfitCenter
                    Else
                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(masterKey)
                    End If
                    masterRowsForProfitCenter.Add masterRowData_1D
                Next k
                dictAllMasterData.Add wsMasterSource.Name, dictSingleMasterGrouped
            End If
        End If
    Next wsMasterSource

    If Not IsEmpty(arrMasterData) Then Erase arrMasterData

    If dictAllMasterData.Count = 0 Then
        MsgBox "No Master sheets with names containing ' (M)' and data found in this workbook.", vbExclamation
        Call WriteLog("ERROR", "No Master (M) sheets found or loaded with data.", "dictAllMasterData.Count", dictAllMasterData.Count)
        GoTo CleanUp
    End If
    Call WriteLog("INFO", "Successfully pre-populated dictionary with all Master (M) data.", "Total Master Sheets", dictAllMasterData.Count)
    ' --- End [Index 5] ---

    ' --- [Index 6]: Populate Reviewer Files Dictionary ---
    Dim masterMappingRowData_1D As Variant
    Dim collProfitCentersForReviewer As Collection

    For i = 1 To UBound(arrMasterMapping, 1)
        ' The column for Reviewer Name is AE (31st column in 1-based array).
        ' The column for TRUE/FALSE checkbox is C (3rd column in 1-based array).
        ' Check if column AE exists (which is 31st column)
        If UBound(arrMasterMapping, 2) >= 31 Then
            Dim reviewerName As String
            reviewerName = Trim(CStr(arrMasterMapping(i, 31))) ' Reviewer Name from Column AE
            If arrMasterMapping(i, 3) = True And Len(reviewerName) > 0 Then ' Checkbox in Column C
                If Not dictReviewerFiles.Exists(reviewerName) Then
                    Set collProfitCentersForReviewer = New Collection
                    dictReviewerFiles.Add reviewerName, collProfitCentersForReviewer
                Else
                    Set collProfitCentersForReviewer = dictReviewerFiles.item(reviewerName)
                End If
                ReDim masterMappingRowData_1D(1 To UBound(arrMasterMapping, 2))
                For col = LBound(arrMasterMapping, 2) To UBound(arrMasterMapping, 2)
                    masterMappingRowData_1D(col) = arrMasterMapping(i, col)
                Next col
                collProfitCentersForReviewer.Add masterMappingRowData_1D
            End If
        End If
    Next i

    If dictReviewerFiles.Count = 0 Then
        MsgBox "No files found to generate by Reviewer. Please ensure 'TRUE' is marked in Column C and Reviewer names are provided in Column AE of MasterMapping.", vbExclamation
        Call WriteLog("WARNING", "No files found to generate by Reviewer.")
        GoTo CleanUp
    End If
    Call WriteLog("INFO", "Grouped data by Reviewer.", "Total Reviewers", dictReviewerFiles.Count)
    ' --- End [Index 6] ---

    ' --- [Index 7]: Initialize Progress Form ---
    totalFilesToProcess = dictReviewerFiles.Count
    With frmProgress
        .lblProgress.Caption = "Processing: 0 / " & totalFilesToProcess
        .lblTime.Caption = "Time Elapsed: 00:00"
        .Show vbModeless
    End With
    filesProcessed = 0
    Call WriteLog("INFO", "Progress UserForm initialized and displayed.")
    ' --- End [Index 7] ---

    ' --- [Index 8]: Main Loop for Generating Files by Reviewer ---
    Dim currentReviewerName_str As Variant
    For Each currentReviewerName_str In dictReviewerFiles.Keys

        filesProcessed = filesProcessed + 1
        Call WriteLog("INFO", "Starting file generation for Reviewer.", "Reviewer Name", currentReviewerName_str)

        frmProgress.lblProgress.Caption = "Processing: " & filesProcessed & " / " & totalFilesToProcess & " (" & currentReviewerName_str & ")"
        currentTime = Timer - StartTime
        minutes = Int(currentTime / 60)
        seconds = Int(currentTime Mod 60)
        frmProgress.lblTime.Caption = "Time Elapsed: " & Format(minutes, "00") & ":" & Format(seconds, "00")
        DoEvents

        Set collProfitCentersForReviewer = dictReviewerFiles.item(currentReviewerName_str)
        Set newWorkbook = Application.Workbooks.Add(xlWBATWorksheet)
        
        ' --- [Index 8.1.A]: Copy (R) Sheets as Values and Formats ---
        ' This new block handles the copying of sheets marked with (R) in row 4 of MasterMapping.
        Dim rSheetName As String
        Dim wsSourceR As Worksheet
        Dim wsCopiedR As Worksheet
        
        For j = 4 To lastColSheets
            rSheetName = Trim(CStr(wsMasterMapping.Cells(4, j).Value))
            
            ' Check if the sheet name contains "(R)" and is not empty.
            If InStr(1, rSheetName, "(R)", vbTextCompare) > 0 And Len(rSheetName) > 0 Then
                
                ' Check if the source sheet actually exists in the current workbook.
                Set wsSourceR = Nothing
                On Error Resume Next
                Set wsSourceR = ThisWorkbook.Sheets(rSheetName)
                On Error GoTo 0
                
                If Not wsSourceR Is Nothing Then
                    ' Copy the entire sheet to the new workbook.
                    wsSourceR.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                    Set wsCopiedR = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                    
                    ' Convert all formulas in the newly copied sheet to their values, keeping formats.
                    If wsCopiedR.UsedRange.Cells.Count > 0 Then
                        wsCopiedR.UsedRange.Value = wsCopiedR.UsedRange.Value
                    End If
                    
                    Call WriteLog("INFO", "Copied (R) sheet and converted to values.", "Sheet Name", rSheetName)
                    Set wsCopiedR = Nothing
                Else
                    Call WriteLog("WARNING", "(R) Sheet specified in MasterMapping not found in workbook.", "Sheet Name", rSheetName)
                End If
            End If
        Next j
        ' --- End [Index 8.1.A] ---
        
        ' --- [Index 8.1]: Loop through Selected Sheets for Current Reviewer ---
        Dim selectedSheetCol As Long
        Dim tempCollectionForSheetData As Collection
        Dim currentMasterDataCols As Long
        
        For selectedSheetCol = 4 To lastColSheets ' Loop through columns D onwards for sheet mappings
            If UBound(arrMasterMapping, 2) >= selectedSheetCol Then
                Dim isSheetEnabledForReviewer As Boolean: isSheetEnabledForReviewer = False
                Dim pcMappingRow As Variant
                For Each pcMappingRow In collProfitCentersForReviewer
                    If pcMappingRow(selectedSheetCol) = True Then ' Check if this sheet is enabled for any PC of this reviewer
                        isSheetEnabledForReviewer = True
                        Exit For
                    End If
                Next pcMappingRow

                If isSheetEnabledForReviewer Then
                    Set tempCollectionForSheetData = New Collection
                    currentMasterDataCols = 0
                    masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, selectedSheetCol).Value))
                    templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, selectedSheetCol).Value))
                    createdSheetName = masterSourceSheetName
                    If InStr(1, createdSheetName, " (M)", vbTextCompare) = 0 Then
                        createdSheetName = createdSheetName & " (M)"
                    End If

                    If dictSheetTemplates.Exists(createdSheetName) Then
                        ' Modified WriteLog call: Combine arguments into the message string
                        Call WriteLog("INFO", "Processing sheet '" & createdSheetName & "' for Reviewer '" & currentReviewerName_str & "'")
                        Set wsOriginalTemplate = Nothing
                        On Error Resume Next
                        Set wsOriginalTemplate = ThisWorkbook.Sheets(templateSheetName)
                        On Error GoTo 0

                        If wsOriginalTemplate Is Nothing Then
                            Call WriteLog("WARNING", "Template sheet not found, skipping. Template: '" & templateSheetName & "', Reviewer: '" & currentReviewerName_str & "'")
                            GoTo NextSheetLoopInner
                        End If

                        ' --- [Index 8.1.1]: Copy Template Sheet and Rename ---
                        On Error Resume Next
                        ThisWorkbook.Sheets(templateSheetName).Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                        Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                        wsNewWorkbookSheet.Name = masterSourceSheetName ' Rename to original master sheet name (without "(M)")
                        On Error GoTo 0
                        If wsNewWorkbookSheet Is Nothing Then
                             Call WriteLog("ERROR", "Failed to copy template sheet. Template: '" & templateSheetName & "', Reviewer: '" & currentReviewerName_str & "'")
                            GoTo NextSheetLoopInner
                        End If
                        Call WriteLog("INFO", "Template sheet copied and renamed. New Sheet: '" & wsNewWorkbookSheet.Name & "', Template: '" & templateSheetName & "'")
                        ' --- End [Index 8.1.1] ---

                        ' --- [Index 8.1.2]: Retrieve Formulas from Original Template ---
                        Dim maxColsInTemplateFormulas As Long
                        maxColsInTemplateFormulas = wsOriginalTemplate.Cells(2, wsOriginalTemplate.Columns.Count).End(xlToLeft).Column
                        If maxColsInTemplateFormulas = 0 Then maxColsInTemplateFormulas = 1
                        arrTemplateFormulas = wsOriginalTemplate.Range("A2").Resize(1, maxColsInTemplateFormulas).FormulaR1C1
                        Call WriteLog("DEBUG", "Formulas retrieved from template. Template: '" & templateSheetName & "', Columns: " & maxColsInTemplateFormulas)
                        ' --- End [Index 8.1.2] ---

                        ' --- [Index 8.1.3]: Data Processing for the Newly Created Sheet ---
                        If dictAllMasterData.Exists(createdSheetName) Then ' Use createdSheetName (e.g., "Sheet1 (M)") for lookup
                            Set dictSingleMasterGrouped = dictAllMasterData.item(createdSheetName)
                            Dim firstMasterDataRowSet As Boolean: firstMasterDataRowSet = False
                            Dim currentProfitCenterID As String, currentProfitCenterName As String

                            For Each pcMappingRow In collProfitCentersForReviewer
                                If pcMappingRow(selectedSheetCol) = True Then ' Check if this PC is enabled for the current sheet
                                    currentProfitCenterID = Trim(CStr(pcMappingRow(1))) ' Profit Center ID from Column A
                                    currentProfitCenterName = Trim(CStr(pcMappingRow(2))) ' Profit Center Name from Column B
                                    If dictSingleMasterGrouped.Exists(currentProfitCenterID) Then
                                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(currentProfitCenterID)
                                        If masterRowsForProfitCenter.Count > 0 Then
                                            If Not firstMasterDataRowSet Then
                                                If IsArray(masterRowsForProfitCenter.item(1)) Then
                                                    currentMasterDataCols = UBound(masterRowsForProfitCenter.item(1))
                                                    firstMasterDataRowSet = True
                                                End If
                                            End If
                                            If IsArray(masterRowsForProfitCenter.item(1)) Then
                                                If UBound(masterRowsForProfitCenter.item(1)) > currentMasterDataCols Then
                                                    currentMasterDataCols = UBound(masterRowsForProfitCenter.item(1))
                                                End If
                                            End If
                                            For Each masterRowData_1D In masterRowsForProfitCenter
                                                Dim rowToAdd As Variant
                                                ReDim rowToAdd(1 To UBound(masterRowData_1D))
                                                For k = LBound(masterRowData_1D) To UBound(masterRowData_1D)
                                                    If IsArray(arrTemplateFormulas) Then
                                                        If k <= UBound(arrTemplateFormulas, 2) Then
                                                            If Left(CStr(arrTemplateFormulas(1, k)), 1) = "=" Then
                                                                rowToAdd(k) = arrTemplateFormulas(1, k)
                                                            Else
                                                                rowToAdd(k) = masterRowData_1D(k)
                                                            End If
                                                        Else
                                                            rowToAdd(k) = masterRowData_1D(k)
                                                        End If
                                                    Else
                                                        rowToAdd(k) = masterRowData_1D(k)
                                                    End If
                                                Next k
                                                ' Column AF (32nd in MasterMapping array) to Column A of new sheet (if exists)
                                                If UBound(pcMappingRow) >= 32 Then
                                                    If UBound(rowToAdd) >= 1 Then rowToAdd(1) = pcMappingRow(32) ' Column AF
                                                End If
                                                ' Profit Center ID to Column B of new sheet
                                                If UBound(rowToAdd) >= 2 Then rowToAdd(2) = currentProfitCenterID
                                                ' Profit Center Name to Column C of new sheet
                                                If UBound(rowToAdd) >= 3 Then rowToAdd(3) = currentProfitCenterName
                                                
                                                tempCollectionForSheetData.Add rowToAdd
                                               
                                            Next masterRowData_1D
                                        Else
                                            Call WriteLog("WARNING", "No master data found for Profit Center. PC: '" & currentProfitCenterID & " (" & currentProfitCenterName & ")', Sheet: '" & wsNewWorkbookSheet.Name & "'")
                                        End If
                                    Else
                                        Call WriteLog("WARNING", "Profit Center ID not found in Master (M) data. PC ID: '" & currentProfitCenterID & "', Master Sheet: '" & createdSheetName & "'")
                                    End If
                                End If
                            Next pcMappingRow

                            If tempCollectionForSheetData.Count > 0 And currentMasterDataCols > 0 Then
                                ReDim arrCurrentSheetFinalData(1 To tempCollectionForSheetData.Count, 1 To currentMasterDataCols)
                                dataRowCounter = 0
                                For Each rowToAdd In tempCollectionForSheetData
                                    dataRowCounter = dataRowCounter + 1
                                    For k = 1 To Application.Min(UBound(rowToAdd), currentMasterDataCols)
                                        arrCurrentSheetFinalData(dataRowCounter, k) = rowToAdd(k)
                                    Next k
                                Next rowToAdd
                                wsNewWorkbookSheet.Range("A2").Resize(tempCollectionForSheetData.Count, currentMasterDataCols).FormulaR1C1 = arrCurrentSheetFinalData
                                Call WriteLog("INFO", "Data written to new workbook sheet. Sheet: '" & wsNewWorkbookSheet.Name & "', Rows: " & tempCollectionForSheetData.Count)
                                Erase arrCurrentSheetFinalData
                            Else
                                Call WriteLog("INFO", "No data to write to sheet after filtering. Sheet: '" & wsNewWorkbookSheet.Name & "', Reviewer: '" & currentReviewerName_str & "'")
                            End If
                        Else
                            Call WriteLog("WARNING", "Master (M) sheet data not found in dictAllMasterData. Master Sheet: '" & createdSheetName & "'")
                        End If
                        Set wsNewWorkbookSheet = Nothing
                        Set wsOriginalTemplate = Nothing
                    End If
                End If
            End If
NextSheetLoopInner:
        Next selectedSheetCol
        ' --- End [Index 8.1] ---

        


        ' --- [Index 8.2]: Delete Default Sheets in New Workbook ---
        On Error Resume Next
        For Each wsTemp In newWorkbook.Sheets
            If Left(wsTemp.Name, 5) = "Sheet" And IsNumeric(Mid(wsTemp.Name, 6)) Then
                wsTemp.Delete
            End If
        Next wsTemp
        On Error GoTo 0
        Call WriteLog("INFO", "Deleted default sheets in new workbook. Workbook: '" & newWorkbook.Name & "'")
        ' --- End [Index 8.2] ---
        
        ' --- [Index 8.3]: Clear Unused Rows in Each Sheet of the New Workbook ---
        Dim ws As Worksheet, lastRowInB As Long
        For Each ws In newWorkbook.Sheets
            lastRowInB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
            If lastRowInB < ws.Rows.Count Then
                ws.Rows(lastRowInB + 1 & ":" & ws.Rows.Count).Clear
                Call WriteLog("DEBUG", "Cleared unused rows in sheet. Sheet: '" & ws.Name & "', Last Row with Data: " & lastRowInB)
            End If
        Next ws
        ' --- End [Index 8.3] ---
        
        ' [8.4] Lock all Sheets with a Password (User's Request)
        '--------------------------------------------------------------------------------------------------------------------
        Dim ws2 As Worksheet
        For Each ws2 In newWorkbook.Worksheets
            
            ws2.Unprotect Password:="MTIEPBCS"
        
            ws2.Protect Password:="MTIEPBCS", _
               AllowInsertingRows:=False, _
               AllowDeletingRows:=False, _
               AllowInsertingColumns:=False, _
               AllowDeletingColumns:=False, _
               AllowFiltering:=True, _
               AllowFormattingColumns:=True, _
               AllowSorting:=True, _
               Contents:=True
        Next ws2
        Call WriteLog("INFO", "All sheets in the new workbook have been locked with a password.")

        ' --- [Index 8.5]: File Saving Operations ---
        ' Column AD (30th column in 1-based MasterMapping array) contains the new file path.
        If collProfitCentersForReviewer.Count > 0 Then
            newFilePath = CStr(collProfitCentersForReviewer.item(1)(30)) ' Column AD
        Else
            newFilePath = ""
        End If

        If Len(newFilePath) > 0 Then
            baseFileName = currentReviewerName_str
            baseFileName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(baseFileName, "/", "_"), "\", "_"), ":", "_"), "*", "_"), "?", "_"), Chr(34), "_"), "<", "_"), ">", "_"), "|", "_")
            folderPath = newFilePath & Application.PathSeparator & baseFileName
            If Not fso.FolderExists(folderPath) Then
                On Error Resume Next
                fso.CreateFolder folderPath
                If Err.Number <> 0 Then
                    MsgBox "Error creating folder: " & Err.Description & ". Folder: " & folderPath, vbCritical
                    Call WriteLog("ERROR", "Failed to create folder. Error: '" & Err.Description & "', Path: '" & folderPath & "'")
                    Err.Clear
                    GoTo SkipFileSave
                End If
                On Error GoTo 0
                Call WriteLog("INFO", "Created new output folder. Folder Path: '" & folderPath & "'")
            End If
            fileName = baseFileName & ".xlsx"
            Dim copyNum As Long: copyNum = 1
            Do While fso.fileExists(folderPath & Application.PathSeparator & fileName)
                copyNum = copyNum + 1
                fileName = baseFileName & " (" & copyNum & ").xlsx"
            Loop
            On Error Resume Next
            newWorkbook.SaveAs fileName:=folderPath & Application.PathSeparator & fileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            If Err.Number <> 0 Then
                MsgBox "Error saving file: " & Err.Description & ". File: " & folderPath & Application.PathSeparator & fileName, vbCritical
                Call WriteLog("ERROR", "Failed to save workbook. Error: '" & Err.Description & "', File: '" & fileName & "'")
                Err.Clear
            Else
                Call WriteLog("INFO", "Workbook saved successfully. Full Path: '" & folderPath & Application.PathSeparator & fileName & "'")
            End If
            On Error GoTo 0
SkipFileSave:
            newWorkbook.Close SaveChanges:=False
            Set newWorkbook = Nothing
        Else
            Call WriteLog("WARNING", "No file path provided, workbook closed without saving. Reviewer Name: '" & currentReviewerName_str & "'")
            newWorkbook.Close SaveChanges:=False
            Set newWorkbook = Nothing
        End If
    Next currentReviewerName_str
    ' --- End [Index 8] ---

    ' --- [Index 9]: CleanUp and Restore Excel Settings ---
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Call WriteLog("INFO", "Application settings restored.")

    Set wsMasterMapping = Nothing
    Set wsTemp = Nothing
    Set wsNewWorkbookSheet = Nothing
    Set wsOriginalTemplate = Nothing
    Set dictAllMasterData = Nothing
    Set dictSheetTemplates = Nothing
    Set dictReviewerFiles = Nothing
    Set dictSingleMasterGrouped = Nothing
    Set masterRowsForProfitCenter = Nothing
    Set collProfitCentersForReviewer = Nothing
    Set fso = Nothing
    Set wsLog = Nothing
    Call WriteLog("INFO", "All object variables cleared.")

    If Not IsEmpty(arrMasterMapping) Then Erase arrMasterMapping
    ' Check if arrMasterData is initialized and not empty before erasing
    If Not IsEmpty(arrMasterData) And IsArray(arrMasterData) Then Erase arrMasterData
    ' Check if arrCurrentSheetFinalData is initialized and not empty before erasing
    If Not IsEmpty(arrCurrentSheetFinalData) And IsArray(arrCurrentSheetFinalData) Then Erase arrCurrentSheetFinalData
    ' Check if arrTemplateFormulas is initialized and not empty before erasing
    If IsArray(arrTemplateFormulas) Then Erase arrTemplateFormulas
    Call WriteLog("INFO", "All array variables cleared.")

    endTime = Timer
    runTime = endTime - StartTime
    ' --- End [Index 9] ---

    ' --- [Index 10]: Update UserForm and Display Final Message ---
    If Not frmProgress Is Nothing Then
        With frmProgress
            .lblProgress.Caption = "Processing Complete!"
            minutes = Int(runTime / 60)
            seconds = Int(runTime Mod 60)
            .lblTime.Caption = "Total Time: " & Format(minutes, "00") & ":" & Format(seconds, "00")
            If .Visible Then Application.Wait Now + TimeValue("00:00:03")
            Unload frmProgress
        End With
        Set frmProgress = Nothing
        Call WriteLog("INFO", "Progress UserForm closed.")
    End If

    minutes = Int(runTime / 60)
    seconds = Int(runTime Mod 60)
    MsgBox "Completed Run Time: " & Format(minutes, "00") & " minutes and " & Format(seconds, "00") & " seconds"
    Call WriteLog("INFO", "Macro finished.", "Total Run Time", Format(minutes, "00") & ":" & Format(seconds, "00"))

End Sub

' --- Helper Function for Logging (from Prototype) ---
Sub WriteLog(logType As String, message As String, Optional varName As String = "", Optional varValue As Variant)
    ' Purpose: Writes a log entry to the specified log sheet.
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim logDetail As String
    Dim currentWorkbook As Workbook
    Set currentWorkbook = ThisWorkbook

    On Error GoTo ErrorHandler

    ' Check if the log sheet exists, create it if not.
    On Error Resume Next
    Set wsLog = currentWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = currentWorkbook.Sheets.Add(After:=currentWorkbook.Sheets(currentWorkbook.Sheets.Count))
        wsLog.Name = LOG_SHEET_NAME
        With wsLog
            .Cells(1, 1).Value = "Timestamp"
            .Cells(1, 2).Value = "Type"
            .Cells(1, 3).Value = "Message"
            .Cells(1, 4).Value = "Variable Name"
            .Cells(1, 5).Value = "Variable Value"
            .Rows(1).Font.Bold = True
            .Columns("A:E").AutoFit
        End With
    End If

    ' Find the next available row in the log sheet.
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    ' Construct the variable detail string if varName is provided.
    If varName <> "" Then
        If IsObject(varValue) And Not IsEmpty(varValue) Then
            On Error Resume Next
            logDetail = varValue.Name
            If Err.Number <> 0 Then
                logDetail = "Object"
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf IsError(varValue) Then
            logDetail = "Error: " & CStr(varValue)
        Else
            logDetail = CStr(varValue)
        End If
    Else
        logDetail = ""
    End If

    ' Write the log entry to the log sheet.
    With wsLog
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = logType
        .Cells(nextRow, 3).Value = message
        .Cells(nextRow, 4).Value = varName
        .Cells(nextRow, 5).Value = logDetail
        .Columns("A:E").AutoFit
    End With

    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteLog function: " & Err.Description & " (Log Type: " & logType & ", Message: " & message & ")"
    On Error GoTo 0
End Sub


