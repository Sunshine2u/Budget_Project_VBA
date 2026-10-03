Option Explicit

' กำหนดค่าคงที่สาธารณะสำหรับชื่อชีท Log จากต้นแบบ
Public Const LOG_SHEET_NAME As String = "Log"
Public Const SHEET_PASSWORD As String = "MTIEPBCS"

Sub genfile_ByReviewer_V11_WithLog()

    On Error GoTo ErrorHandler

    ' STREAMING_CHUNK:Declaring variables and initializing timing variables...
    ' --- [ดัชนี 1]: การประกาศตัวแปร ---
    ' ประกาศตัวแปรสำหรับการจับเวลาการทำงาน
    Dim StartTime As Double
    Dim endTime As Double
    Dim runTime As Double
    Dim savedScreenUpdating As Boolean
    Dim savedCalculation As XlCalculation
    Dim savedEnableEvents As Boolean
    Dim savedDisplayAlerts As Boolean
    Dim runtimeErrorNumber As Long
    Dim runtimeErrorDescription As String
    Dim runtimeErrorSource As String
    Dim runFailed As Boolean
    Dim progressFormShown As Boolean
    Dim applicationStateCaptured As Boolean

    ' ประกาศวัตถุ Worksheet เพื่อให้อ้างอิงได้ง่ายขึ้น
    Dim wsMasterMapping As Worksheet        ' วัตถุ Worksheet สำหรับ "MasterMapping"
    Dim wsTemp As Worksheet                 ' วัตถุ Worksheet ชั่วคราวสำหรับลบชีทเริ่มต้น
    Dim wsNewWorkbookSheet As Worksheet     ' วัตถุ Worksheet ในเวิร์กบุ๊กที่สร้างขึ้นใหม่
    Dim wsOriginalTemplate As Worksheet     ' วัตถุ Worksheet สำหรับชีทเทมเพลตดั้งเดิม
    Dim wsStandaloneSource As Worksheet     ' วัตถุ Worksheet ต้นทางสำหรับชีท (R)/(T) แบบ standalone
    Dim wsFormulaCell As Range
    Dim wsTargetFormulaCell As Range
    Dim wsStandaloneFormulaCells As Range
    Dim wsOutputTable As ListObject
    Dim wsOutputFilterRange As Range
    Dim wsOutputSheet As Worksheet
    Dim wsLog As Worksheet                  ' วัตถุ Worksheet สำหรับบันทึก Log
    Dim wsCalculation As Worksheet
    Dim wsTemplateShape As Shape
    Dim wsCopiedShape As Shape
    Dim dictTemplateShapeText As Object
    Dim formulaErrorsCount As Long          ' ตัวนับสำหรับข้อผิดพลาดของสูตร
    Dim standaloneSheetName As String
    Dim standaloneFormulaErrors As Long
    Dim firstStandaloneFormulaError As String
    Dim formulaCellText As String
    Dim formulaCellAddress As String
    Dim standaloneFormulaCellCount As Long
    Dim copiedTableCount As Long
    Dim tableIndex As Long
    Dim progressCheckCounter As Long

    ' ประกาศตัวแปรสำหรับเก็บข้อมูลในอาร์เรย์เพื่อการประมวลผลที่เร็วขึ้น
    Dim arrMasterMapping As Variant
    Dim arrMasterData As Variant            ' อาร์เรย์สำหรับเก็บข้อมูลจากชีท Master (M)
    Dim arrCurrentSheetFinalData As Variant ' อาร์เรย์สำหรับเก็บข้อมูลสุดท้ายของชีทปัจจุบันก่อนทำการเขียนลงชีท
    Dim arrTemplateFormulas As Variant      ' อาร์เรย์สำหรับเก็บสูตรจากชีทเทมเพลตดั้งเดิม

    ' STREAMING_CHUNK:Initializing dictionary objects...
    ' ประกาศวัตถุ Dictionary สำหรับการค้นหาข้อมูล Master (M) ที่รวดเร็ว
    Dim dictAllMasterData As Object

    ' ประกาศตัวแปรสำหรับเก็บการจับคู่ชีทและตัวนับรอบการวนลูป
    Dim dictSheetTemplates As Object

    ' Dictionary สำหรับจัดกลุ่ม Profit Center ตาม Reviewer
    Dim dictReviewerFiles As Object

    Dim i As Long, j As Long, k As Long
    Dim dataRowCounter As Long
    Dim col As Long

    ' ประกาศตัวแปรสำหรับการดำเนินการเกี่ยวกับไฟล์
    Dim newFilePath As String
    Dim folderPath As String
    Dim fileName As String
    Dim newWorkbook As Workbook
    Dim baseFileName As String

    ' ตัวแปรสำหรับติดตามความคืบหน้าและเวลาบน UserForm
    Dim frmProgress As New frmProgress
    Dim totalFilesToProcess As Long
    Dim filesProcessed As Long
    Dim currentTime As Double
    Dim minutes As Long
    Dim seconds As Long
    Dim currentReviewerName_str As Variant
    ' --- สิ้นสุด [ดัชนี 1] ---

    savedScreenUpdating = Application.ScreenUpdating
    savedCalculation = Application.Calculation
    savedEnableEvents = Application.EnableEvents
    savedDisplayAlerts = Application.DisplayAlerts
    applicationStateCaptured = True

    Set dictAllMasterData = CreateObject("Scripting.Dictionary")
    Set dictSheetTemplates = CreateObject("Scripting.Dictionary")
    Set dictReviewerFiles = CreateObject("Scripting.Dictionary")

    ' STREAMING_CHUNK:Configuring system settings and preparing log sheet...
    ' --- [ดัชนี 2]: เริ่มต้นการทำงานของ FileSystemObject และการตั้งค่า Application ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    StartTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    With frmProgress
        .lblProgress.Caption = "Calculating worksheets..."
        .lblTime.Caption = "Elapsed: 00:00:00"
        progressFormShown = True
        .Show vbModeless
    End With
    Call RefreshProgressForm(frmProgress, StartTime, True)

    For Each wsCalculation In ThisWorkbook.Worksheets
        frmProgress.lblProgress.Caption = "Calculating: " & wsCalculation.Name
        Call RefreshProgressForm(frmProgress, StartTime, True)
        wsCalculation.Calculate
    Next wsCalculation
    Call RefreshProgressForm(frmProgress, StartTime, True)

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    frmProgress.lblProgress.Caption = "Preparing data..."
    Call RefreshProgressForm(frmProgress, StartTime, True)

    ' --- บันทึก LOG: เริ่มต้นการทำงานของแมโคร ---
    Call WriteLog("INFO", "Macro started: genfile_ByReviewer_V10_WithLog")

    ' --- จัดเตรียมชีท Log ---
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0
    On Error GoTo ErrorHandler

    If Not wsLog Is Nothing Then
        ' ตรวจสอบว่ามีข้อมูลต่อจากหัวข้อก่อนทำการล้างข้อมูลหรือไม่
        If wsLog.Cells(Rows.Count, "A").End(xlUp).Row > 1 Then
            wsLog.Range("A2:E" & wsLog.Cells(Rows.Count, "A").End(xlUp).Row).ClearContents
            Call WriteLog("INFO", "Cleared existing data in Log sheet.")
        End If
    End If

    Dim wsActionPage As Worksheet
    Dim sheetName As String

    ' กำหนดค่า wsActionPage จากชีท "Action_Page"
    Set wsActionPage = ThisWorkbook.Sheets("Action_Page")

    ' ดึงชื่อชีทจาก Cell A6 ของ Action Page
    sheetName = wsActionPage.Range("A6").Value

    ' กำหนดค่า wsMasterMapping ตามชื่อชีทที่ระบุใน sheetName
    Set wsMasterMapping = ThisWorkbook.Sheets(sheetName)
    'Set wsMasterMapping = ThisWorkbook.Sheets("MasterMapping")
    
    ' --- สิ้นสุด [ดัชนี 2] ---

    ' STREAMING_CHUNK:Loading MasterMapping data into memory...
    ' --- [ดัชนี 3]: โหลดข้อมูล MasterMapping ---
    Dim lastRowMasterMapping As Long
    lastRowMasterMapping = wsMasterMapping.Cells(Rows.Count, "A").End(xlUp).Row
    If lastRowMasterMapping < 5 Then ' ปรับแก้ไขเพื่อตรวจสอบตั้งแต่วัดจากแถวที่ 5
        MsgBox "MasterMapping sheet is empty or has only headers.", vbExclamation
        Call WriteLog("WARNING", "MasterMapping sheet is empty or has only headers.", "lastRowMasterMapping", lastRowMasterMapping)
        GoTo CleanUp
    End If
    arrMasterMapping = wsMasterMapping.Range("A5:AR" & lastRowMasterMapping).Value
    Call WriteLog("INFO", "MasterMapping data loaded into array.", "Rows Loaded", UBound(arrMasterMapping, 1))
    ' --- สิ้นสุด [ดัชนี 3] ---

    ' STREAMING_CHUNK:Identifying sheet templates from MasterMapping...
    ' --- [ดัชนี 4]: ระบุเทมเพลตของชีทจาก MasterMapping ---
    Dim lastColSheets As Long
    Dim masterSourceSheetName As String
    Dim templateSheetName As String
    Dim createdSheetName As String
    Dim reportSheetName As String
    Dim hasStandaloneSheets As Boolean

    On Error Resume Next
    lastColSheets = Application.Max( _
        wsMasterMapping.Cells(2, wsMasterMapping.Columns.Count).End(xlToLeft).Column, _
        wsMasterMapping.Cells(3, wsMasterMapping.Columns.Count).End(xlToLeft).Column, _
        wsMasterMapping.Cells(4, wsMasterMapping.Columns.Count).End(xlToLeft).Column)
    If lastColSheets < 4 Then lastColSheets = 3 ' ตรวจสอบให้แน่ใจว่าพิจารณาอย่างน้อยคอลัมน์ D (ดัชนี 4 ในฐาน 1) สำหรับการจับคู่ชีท
    On Error GoTo 0
    On Error GoTo ErrorHandler

    For j = 4 To lastColSheets
        masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, j).Value))
        templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, j).Value))
        reportSheetName = Trim(CStr(wsMasterMapping.Cells(4, j).Value))

        If Len(masterSourceSheetName) > 0 Then
            createdSheetName = masterSourceSheetName
            If InStr(1, createdSheetName, " (M)", vbTextCompare) = 0 Then
                createdSheetName = createdSheetName & " (M)"
            End If

            If Len(createdSheetName) > 0 Then
                If Not dictSheetTemplates.Exists(createdSheetName) Then
                    ' ใช้ชีท (M) เป็นแหล่งหลัก; ใช้ชีท (T) เฉพาะเมื่อไม่พบชีท (M)
                    Dim targetTemplate As String
                    targetTemplate = ""
                    Set wsOriginalTemplate = Nothing
                    On Error Resume Next
                    Set wsOriginalTemplate = ThisWorkbook.Sheets(createdSheetName)
                    On Error GoTo 0
                    On Error GoTo ErrorHandler

                    If Not wsOriginalTemplate Is Nothing Then
                        targetTemplate = createdSheetName
                        Call WriteLog("INFO", "Master sheet selected as output template and formula source.", "Master Sheet", createdSheetName)
                    ElseIf Len(templateSheetName) > 0 Then
                        On Error Resume Next
                        Set wsOriginalTemplate = ThisWorkbook.Sheets(templateSheetName)
                        On Error GoTo 0
                        On Error GoTo ErrorHandler
                        If Not wsOriginalTemplate Is Nothing Then
                            targetTemplate = templateSheetName
                            Call WriteLog("INFO", "Template sheet selected because the Master sheet was not found.", "Template Sheet", templateSheetName)
                        End If
                    End If

                    ' คง fallback เดิมเมื่อไม่พบทั้งชีท (M) และ (T)
                    If Len(targetTemplate) = 0 Then
                        targetTemplate = createdSheetName
                    End If
                    Set wsOriginalTemplate = Nothing

                    dictSheetTemplates.Add createdSheetName, targetTemplate
                End If
            End If
        ElseIf Len(templateSheetName) > 0 Or Len(reportSheetName) > 0 Then
            hasStandaloneSheets = True
            Call WriteLog("INFO", "Standalone (T)/(R) sheet mapping found.", "Template / Report", templateSheetName & " / " & reportSheetName)
        End If
    Next j

    If dictSheetTemplates.Count = 0 And Not hasStandaloneSheets Then
        MsgBox "No valid (M), (T)-only, or (R)-only sheet mappings found in MasterMapping (Rows 2-4, Columns D-AR).", vbExclamation
        Call WriteLog("WARNING", "No valid sheet mappings found in MasterMapping.", "dictSheetTemplates.Count", dictSheetTemplates.Count)
        GoTo CleanUp
    End If
    Call WriteLog("INFO", "Identified sheet mappings.", "Master Mappings / Standalone Found", dictSheetTemplates.Count & " / " & hasStandaloneSheets)
    ' --- สิ้นสุด [ดัชนี 4] ---

    ' STREAMING_CHUNK:Pre-populating dictionary with Master data...
    ' --- [ดัชนี 5]: โหลดข้อมูล Master (M) เตรียมไว้ใน Dictionary ---
    Dim wsMasterSource As Worksheet
    Dim lastRowMasterSource As Long
    Dim lastColMasterSource As Long
    Dim masterKey As String
    Dim masterRowIndex As Variant
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
                    progressCheckCounter = progressCheckCounter + 1
                    If progressCheckCounter >= 500 Then
                        Call RefreshProgressForm(frmProgress, StartTime)
                        progressCheckCounter = 0
                    End If
                    masterKey = Trim(CStr(arrMasterData(k, 2))) ' สมมติว่า Profit Center ID อยู่ในคอลัมน์ B ของชีท Master
                    If Not dictSingleMasterGrouped.Exists(masterKey) Then
                        Set masterRowsForProfitCenter = New Collection
                        dictSingleMasterGrouped.Add masterKey, masterRowsForProfitCenter
                    Else
                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(masterKey)
                    End If
                    masterRowsForProfitCenter.Add k
                Next k
                dictAllMasterData.Add wsMasterSource.Name, dictSingleMasterGrouped
                Erase arrMasterData
            End If
        End If
    Next wsMasterSource

    If dictAllMasterData.Count = 0 And Not hasStandaloneSheets Then
        MsgBox "No Master sheets with names containing ' (M)' and data found in this workbook.", vbExclamation
        Call WriteLog("ERROR", "No Master (M) sheets found or loaded with data.", "dictAllMasterData.Count", dictAllMasterData.Count)
        GoTo CleanUp
    ElseIf dictAllMasterData.Count = 0 Then
        Call WriteLog("INFO", "No Master (M) data found; continuing with standalone (T)/(R) sheets only.")
    Else
        Call WriteLog("INFO", "Successfully pre-populated dictionary with all Master (M) data.", "Total Master Sheets", dictAllMasterData.Count)
    End If
    ' --- สิ้นสุด [ดัชนี 5] ---

    ' STREAMING_CHUNK:Grouping profit centers by Reviewer...
    ' --- [ดัชนี 6]: โหลดข้อมูลเข้าสู่ Dictionary สำหรับไฟล์ของ Reviewer ---
    Dim masterMappingRowData_1D As Variant
    Dim collProfitCentersForReviewer As Collection

    For i = 1 To UBound(arrMasterMapping, 1)
        progressCheckCounter = progressCheckCounter + 1
        If progressCheckCounter >= 500 Then
            Call RefreshProgressForm(frmProgress, StartTime)
            progressCheckCounter = 0
        End If
        ' คอลัมน์สำหรับชื่อ Reviewer คือ AE (คอลัมน์ที่ 31 ในอาร์เรย์ฐาน 1)
        ' คอลัมน์สำหรับกล่องกาถูก TRUE/FALSE คือ C (คอลัมน์ที่ 3)
        ' ตรวจสอบว่ามีคอลัมน์ AE อยู่หรือไม่ (คอลัมน์ที่ 31)
        If UBound(arrMasterMapping, 2) >= 31 Then
            Dim reviewerName As String
            reviewerName = Trim(CStr(arrMasterMapping(i, 31))) ' ชื่อ Reviewer จากคอลัมน์ AE
            If arrMasterMapping(i, 3) = True And Len(reviewerName) > 0 Then ' กล่องกาถูกในคอลัมน์ C
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
    ' --- สิ้นสุด [ดัชนี 6] ---

    ' STREAMING_CHUNK:Displaying progress form...
    ' --- [ดัชนี 7]: เริ่มต้นการทำงานของหน้าต่าง Progress ---
    totalFilesToProcess = dictReviewerFiles.Count
    With frmProgress
        .lblProgress.Caption = "Processing: 0 / " & totalFilesToProcess
    End With
    filesProcessed = 0
    Call RefreshProgressForm(frmProgress, StartTime, True)
    Call WriteLog("INFO", "Progress UserForm initialized and displayed.")
    ' --- สิ้นสุด [ดัชนี 7] ---

    ' STREAMING_CHUNK:Executing main loop for file generation by Reviewer...
        ' --- [ดัชนี 8.1]: วนลูปตามชีทที่เลือกสำหรับ Reviewer ปัจจุบัน ---
        Dim selectedSheetCol As Long
        Dim tempCollectionForSheetData As Collection
        Dim currentMasterDataCols As Long

        For Each currentReviewerName_str In dictReviewerFiles.Keys
            Set collProfitCentersForReviewer = dictReviewerFiles.Item(currentReviewerName_str)
            Set newWorkbook = Application.Workbooks.Add(xlWBATWorksheet)
            filesProcessed = filesProcessed + 1
            frmProgress.lblProgress.Caption = "Processing: " & filesProcessed & " / " & totalFilesToProcess
            Call RefreshProgressForm(frmProgress, StartTime, True)

        For selectedSheetCol = 4 To lastColSheets ' วนลูปตั้งแต่คอลัมน์ D เป็นต้นไปสำหรับการจับคู่ชีท
            If UBound(arrMasterMapping, 2) >= selectedSheetCol Then
                Dim isSheetEnabledForReviewer As Boolean: isSheetEnabledForReviewer = False
                Dim pcMappingRow As Variant
                For Each pcMappingRow In collProfitCentersForReviewer
                    progressCheckCounter = progressCheckCounter + 1
                    If progressCheckCounter >= 500 Then
                        Call RefreshProgressForm(frmProgress, StartTime)
                        progressCheckCounter = 0
                    End If
                    ' 1. ตรวจสอบว่าดัชนีคอลัมน์ไม่เกินขนาดอาร์เรย์
                    If selectedSheetCol <= UBound(pcMappingRow) Then
                        ' 2. ตรวจสอบว่าเซลล์นั้นไม่ใช่ค่า Error (เช่น #N/A, #VALUE!)
                        If Not IsError(pcMappingRow(selectedSheetCol)) Then
                            ' 3. ตรวจสอบค่า True (รองรับทั้ง Boolean และ ข้อความ "TRUE")
                            If pcMappingRow(selectedSheetCol) = True Or UCase(Trim(CStr(pcMappingRow(selectedSheetCol)))) = "TRUE" Then
                                isSheetEnabledForReviewer = True
                                Exit For
                            End If
                        End If
                    End If
                Next pcMappingRow
                If isSheetEnabledForReviewer Then
                    Set tempCollectionForSheetData = New Collection
                    currentMasterDataCols = 0
                    masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, selectedSheetCol).Value))
                    templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, selectedSheetCol).Value))
                    reportSheetName = Trim(CStr(wsMasterMapping.Cells(4, selectedSheetCol).Value))
                    createdSheetName = masterSourceSheetName
                    If InStr(1, createdSheetName, " (M)", vbTextCompare) = 0 Then
                        createdSheetName = createdSheetName & " (M)"
                    End If

                    If dictSheetTemplates.Exists(createdSheetName) Then
                        ' ดึงชื่อ Template จาก Dictionary หรือใช้ชื่อชีท (M) ถ้าไม่มี Template
                        templateSheetName = dictSheetTemplates.Item(createdSheetName)
                        
                        Call WriteLog("INFO", "Processing sheet '" & createdSheetName & "' using source '" & templateSheetName & "' for Reviewer '" & currentReviewerName_str & "'")

                        Set wsOriginalTemplate = Nothing
                        On Error Resume Next
                        Set wsOriginalTemplate = ThisWorkbook.Sheets(templateSheetName)
                        On Error GoTo 0
                        On Error GoTo ErrorHandler

                        ' หากหาชีท Template ไม่เจอ ให้ถอยมาใช้ชีท (M) ต้นฉบับแทน
                        If wsOriginalTemplate Is Nothing Then
                            On Error Resume Next
                            Set wsOriginalTemplate = ThisWorkbook.Sheets(createdSheetName)
                            On Error GoTo 0
                            On Error GoTo ErrorHandler
                        End If

                        ' หากยังไม่พบชีทใดๆ ให้ข้ามไป
                        If wsOriginalTemplate Is Nothing Then
                            Call WriteLog("WARNING", "Source/Template sheet not found, skipping. Sheet: '" & createdSheetName & "', Reviewer: '" & currentReviewerName_str & "'")
                            GoTo NextSheetLoopInner
                        End If

                        Set dictTemplateShapeText = CreateObject("Scripting.Dictionary")
                        For Each wsTemplateShape In wsOriginalTemplate.Shapes
                            On Error Resume Next
                            If wsTemplateShape.TextFrame2.HasText Then
                                dictTemplateShapeText(wsTemplateShape.Name) = wsTemplateShape.TextFrame2.TextRange.Text
                            End If
                            Err.Clear
                            On Error GoTo 0
                            On Error GoTo ErrorHandler
                        Next wsTemplateShape

                        ' --- [ดัชนี 8.1.1]: คัดลอกชีท (M) ไปยัง Workbook ใหม่ และล้างข้อมูลตั้งแต่แถว 2 ลงไป แต่คงรูปแบบชีทไว้ ---
                        On Error Resume Next
                        wsOriginalTemplate.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                        Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                        On Error GoTo 0
                        On Error GoTo ErrorHandler

                        If wsNewWorkbookSheet Is Nothing Then
                            Call WriteLog("ERROR", "Failed to copy sheet. Sheet: '" & wsOriginalTemplate.Name & "', Reviewer: '" & currentReviewerName_str & "'")
                            GoTo NextSheetLoopInner
                        End If

                        Dim lastRowFromColumnB As Long
                        lastRowFromColumnB = wsOriginalTemplate.Cells(wsOriginalTemplate.Rows.Count, "B").End(xlUp).Row

                        ' คงชื่อชีท (M) ไว้เพื่อให้สูตรอ้างอิงใน Workbook ตรงกัน
                        On Error Resume Next
                        wsNewWorkbookSheet.Name = createdSheetName
                        On Error GoTo 0
                        On Error GoTo ErrorHandler

                        ' ปลดล็อกสำเนาและล้างเนื้อหาตั้งแต่แถว 2 โดยคงรูปแบบเซลล์ไว้
                        On Error Resume Next
                        wsNewWorkbookSheet.Unprotect Password:=SHEET_PASSWORD
                        If lastRowFromColumnB >= 2 Then
                            wsNewWorkbookSheet.Rows("2:" & lastRowFromColumnB).ClearContents
                        End If
                        If Err.Number <> 0 Then
                            Call WriteLog("ERROR", "Failed to clear contents starting from row 2.", "Error Description", Err.Description)
                            Err.Clear
                        End If
                        On Error GoTo 0
                        On Error GoTo ErrorHandler

                        For Each wsCopiedShape In wsNewWorkbookSheet.Shapes
                            If dictTemplateShapeText.Exists(wsCopiedShape.Name) Then
                                On Error Resume Next
                                wsCopiedShape.TextFrame2.TextRange.Text = dictTemplateShapeText(wsCopiedShape.Name)
                                If Err.Number <> 0 Then
                                    Call WriteLog("WARNING", "Could not restore text on copied shape.", "Shape / Error", wsCopiedShape.Name & " / " & CStr(Err.Number) & ": " & Err.Description)
                                    Err.Clear
                                End If
                                On Error GoTo 0
                                On Error GoTo ErrorHandler
                            End If
                        Next wsCopiedShape
                        Set dictTemplateShapeText = Nothing
                        Set wsTemplateShape = Nothing
                        Set wsCopiedShape = Nothing
                        ' --- สิ้นสุด [ดัชนี 8.1.1] ---

                        ' --- [ดัชนี 8.1.2]: ดึงสูตรมาจากชีท (M) ต้นฉบับ ---
                        Dim maxColsInTemplateFormulas As Long
                        maxColsInTemplateFormulas = wsOriginalTemplate.Cells(2, wsOriginalTemplate.Columns.Count).End(xlToLeft).Column
                        If maxColsInTemplateFormulas = 0 Then maxColsInTemplateFormulas = 1
                        arrTemplateFormulas = wsOriginalTemplate.Range("A2").Resize(1, maxColsInTemplateFormulas).FormulaR1C1
                        Call WriteLog("DEBUG", "Formulas retrieved from source sheet '" & wsOriginalTemplate.Name & "'. Columns: " & maxColsInTemplateFormulas)
                        ' --- สิ้นสุด [ดัชนี 8.1.2] ---

                        ' STREAMING_CHUNK:Processing data rows for the copied sheet...
                        ' --- [ดัชนี 8.1.3]: การประมวลผลข้อมูลสำหรับชีทที่สร้างขึ้นใหม่ ---
                        If dictAllMasterData.Exists(createdSheetName) Then ' ใช้ createdSheetName (เช่น "Sheet1 (M)") ในการค้นหา
                            Set dictSingleMasterGrouped = dictAllMasterData.item(createdSheetName)
                            Dim currentProfitCenterID As String, currentProfitCenterName As String
                            Set wsMasterSource = ThisWorkbook.Sheets(createdSheetName)
                            lastRowMasterSource = wsMasterSource.Cells(wsMasterSource.Rows.Count, "B").End(xlUp).Row
                            lastColMasterSource = wsMasterSource.Cells(1, wsMasterSource.Columns.Count).End(xlToLeft).Column
                            arrMasterData = wsMasterSource.Range(wsMasterSource.Cells(2, "A"), wsMasterSource.Cells(lastRowMasterSource, lastColMasterSource)).Value
                            currentMasterDataCols = UBound(arrMasterData, 2)

                            For Each pcMappingRow In collProfitCentersForReviewer
                                progressCheckCounter = progressCheckCounter + 1
                                If progressCheckCounter >= 500 Then
                                    Call RefreshProgressForm(frmProgress, StartTime)
                                    progressCheckCounter = 0
                                End If
                                If pcMappingRow(selectedSheetCol) = True Then ' ตรวจสอบว่า PC นี้ถูกเปิดใช้งานสำหรับชีทปัจจุบันหรือไม่
                                    currentProfitCenterID = Trim(CStr(pcMappingRow(1))) ' Profit Center ID จากคอลัมน์ A
                                    currentProfitCenterName = Trim(CStr(pcMappingRow(2))) ' Profit Center Name จากคอลัมน์ B
                                    If dictSingleMasterGrouped.Exists(currentProfitCenterID) Then
                                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(currentProfitCenterID)
                                        If masterRowsForProfitCenter.Count > 0 Then
                                            For Each masterRowIndex In masterRowsForProfitCenter
                                                progressCheckCounter = progressCheckCounter + 1
                                                If progressCheckCounter >= 500 Then
                                                    Call RefreshProgressForm(frmProgress, StartTime)
                                                    progressCheckCounter = 0
                                                End If
                                                Dim rowToAdd As Variant
                                                ReDim rowToAdd(1 To currentMasterDataCols)
                                                For k = 1 To currentMasterDataCols
                                                    rowToAdd(k) = arrMasterData(CLng(masterRowIndex), k)
                                                Next k
                                                ' คอลัมน์ AF (ตำแหน่งที่ 32 ในอาร์เรย์ MasterMapping) ไปยังคอลัมน์ A ของชีทใหม่ (ถ้ามี)
                                                If UBound(pcMappingRow) >= 32 Then
                                                    If UBound(rowToAdd) >= 1 Then rowToAdd(1) = pcMappingRow(32) ' คอลัมน์ AF
                                                End If
                                                ' Profit Center ID ไปยังคอลัมน์ B ของชีทใหม่
                                                If UBound(rowToAdd) >= 2 Then rowToAdd(2) = currentProfitCenterID
                                                ' Profit Center Name ไปยังคอลัมน์ C ของชีทใหม่
                                                If UBound(rowToAdd) >= 3 Then rowToAdd(3) = currentProfitCenterName
                                                
                                                tempCollectionForSheetData.Add rowToAdd
                                               
                                            Next masterRowIndex
                                        Else
                                            Call WriteLog("WARNING", "No master data found for Profit Center. PC: '" & currentProfitCenterID & " (" & currentProfitCenterName & ")', Sheet: '" & wsNewWorkbookSheet.Name & "'")
                                        End If
                                    Else
                                        Call WriteLog("WARNING", "Profit Center ID not found in Master (M) data. PC ID: '" & currentProfitCenterID & "', Master Sheet: '" & createdSheetName & "'")
                                    End If
                                End If
                            Next pcMappingRow

                            If tempCollectionForSheetData.Count > 0 And currentMasterDataCols > 0 Then
                                If maxColsInTemplateFormulas > 0 Then
                                    arrTemplateFormulas = wsOriginalTemplate.Range("A2").Resize(tempCollectionForSheetData.Count, maxColsInTemplateFormulas).FormulaR1C1
                                End If

                                ReDim arrCurrentSheetFinalData(1 To tempCollectionForSheetData.Count, 1 To currentMasterDataCols)
                                dataRowCounter = 0
                                For Each rowToAdd In tempCollectionForSheetData
                                    dataRowCounter = dataRowCounter + 1
                                    For k = 1 To Application.Min(UBound(rowToAdd), currentMasterDataCols)
                                        arrCurrentSheetFinalData(dataRowCounter, k) = rowToAdd(k)
                                    Next k
                                Next rowToAdd

                                ' เขียนค่าก่อน แล้วจึงเขียนสูตรทีละเซลล์เหมือน Gen_by_Cost_Center
                                wsNewWorkbookSheet.Range("A2").Resize(tempCollectionForSheetData.Count, currentMasterDataCols).Value = arrCurrentSheetFinalData
                                Call RefreshProgressForm(frmProgress, StartTime)
                                Call WriteLog("INFO", "Data values written to new sheet.", "Sheet Name / Rows", wsNewWorkbookSheet.Name & " / " & tempCollectionForSheetData.Count)

                                If IsArray(arrTemplateFormulas) Then
                                    Dim rIdx As Long
                                    Dim cellFormula As String
                                    Dim formulaStartRow As Long
                                    Dim formulaRunLength As Long
                                    Dim formulaOffset As Long
                                    Dim formulaBlock As Variant
                                    Dim formulaBlockError As String
                                    Dim formulaFallbackErrors As Long
                                    Dim formulaFallbackFirstFailure As String
                                    Dim formulaRange As Range

                                    For k = 1 To currentMasterDataCols
                                        rIdx = 1
                                        Do While rIdx <= tempCollectionForSheetData.Count
                                            progressCheckCounter = progressCheckCounter + 1
                                            If progressCheckCounter >= 500 Then
                                                Call RefreshProgressForm(frmProgress, StartTime)
                                                progressCheckCounter = 0
                                            End If
                                            cellFormula = ""
                                            If rIdx <= UBound(arrTemplateFormulas, 1) And k <= UBound(arrTemplateFormulas, 2) Then
                                                cellFormula = CStr(arrTemplateFormulas(rIdx, k))
                                            End If

                                            If Left(cellFormula, 1) = "=" And Len(cellFormula) > 1 Then
                                                formulaStartRow = rIdx
                                                Do While rIdx <= tempCollectionForSheetData.Count
                                                    If rIdx > UBound(arrTemplateFormulas, 1) Or k > UBound(arrTemplateFormulas, 2) Then Exit Do
                                                    cellFormula = CStr(arrTemplateFormulas(rIdx, k))
                                                    If Left(cellFormula, 1) <> "=" Or Len(cellFormula) <= 1 Then Exit Do
                                                    rIdx = rIdx + 1
                                                Loop

                                                formulaRunLength = rIdx - formulaStartRow
                                                ReDim formulaBlock(1 To formulaRunLength, 1 To 1)
                                                For formulaOffset = 1 To formulaRunLength
                                                    formulaBlock(formulaOffset, 1) = arrTemplateFormulas(formulaStartRow + formulaOffset - 1, k)
                                                Next formulaOffset

                                                Set formulaRange = wsNewWorkbookSheet.Cells(formulaStartRow + 1, k).Resize(formulaRunLength, 1)
                                                On Error Resume Next
                                                formulaRange.FormulaR1C1 = formulaBlock
                                                If Err.Number <> 0 Then
                                                    formulaBlockError = CStr(Err.Number) & ": " & Err.Description
                                                    Err.Clear
                                                    formulaFallbackErrors = 0
                                                    formulaFallbackFirstFailure = ""

                                                    For formulaOffset = 1 To formulaRunLength
                                                        wsNewWorkbookSheet.Cells(formulaStartRow + formulaOffset, k).FormulaR1C1 = formulaBlock(formulaOffset, 1)
                                                        If Err.Number <> 0 Then
                                                            formulaFallbackErrors = formulaFallbackErrors + 1
                                                            If Len(formulaFallbackFirstFailure) = 0 Then
                                                                formulaFallbackFirstFailure = wsNewWorkbookSheet.Cells(formulaStartRow + formulaOffset, k).Address(False, False) & " / " & CStr(formulaBlock(formulaOffset, 1)) & " / " & CStr(Err.Number) & ": " & Err.Description
                                                            End If
                                                            Err.Clear
                                                        End If
                                                    Next formulaOffset

                                                    If formulaFallbackErrors > 0 Then
                                                        formulaErrorsCount = formulaErrorsCount + formulaFallbackErrors
                                                        Call WriteLog("WARNING", "Formula block and cell-by-cell fallback both failed.", "Range / Errors", formulaRange.Address(False, False) & " / Block: " & formulaBlockError & " / Cells failed: " & formulaFallbackErrors & " / First failure: " & formulaFallbackFirstFailure)
                                                    Else
                                                        Call WriteLog("INFO", "Formula block failed; cell-by-cell fallback succeeded.", "Range / Original Error", formulaRange.Address(False, False) & " / " & formulaBlockError)
                                                    End If
                                                End If
                                                On Error GoTo 0
                                                On Error GoTo ErrorHandler
                                                Set formulaRange = Nothing
                                            Else
                                                rIdx = rIdx + 1
                                            End If
                                        Loop
                                    Next k
                                    Call WriteLog("INFO", "Formulas applied in contiguous blocks for formula cells only.", "Sheet Name", wsNewWorkbookSheet.Name)
                                End If

                                Erase arrCurrentSheetFinalData
                            Else
                                Call WriteLog("INFO", "No data to write to sheet after filtering. Sheet: '" & wsNewWorkbookSheet.Name & "', Reviewer: '" & currentReviewerName_str & "'")
                            End If
                        Else
                            Call WriteLog("WARNING", "Master (M) sheet data not found in dictAllMasterData. Master Sheet: '" & createdSheetName & "'")
                        End If
                        If IsArray(arrMasterData) Then Erase arrMasterData
                        Set wsNewWorkbookSheet = Nothing
                        Set wsOriginalTemplate = Nothing
                    ElseIf Len(masterSourceSheetName) = 0 Then
                        If Len(reportSheetName) > 0 Then
                            standaloneSheetName = reportSheetName
                            Set wsStandaloneSource = Nothing
                            On Error Resume Next
                            Set wsStandaloneSource = ThisWorkbook.Sheets(standaloneSheetName)
                            On Error GoTo 0
                            On Error GoTo ErrorHandler

                            If Not wsStandaloneSource Is Nothing Then
                                wsStandaloneSource.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                On Error Resume Next
                                wsNewWorkbookSheet.Name = standaloneSheetName
                                wsNewWorkbookSheet.Unprotect Password:=SHEET_PASSWORD
                                wsNewWorkbookSheet.UsedRange.Value = wsNewWorkbookSheet.UsedRange.Value
                                If Err.Number <> 0 Then
                                    Call WriteLog("ERROR", "Failed to convert standalone (R) sheet formulas to values.", "Sheet / Error", standaloneSheetName & " / " & Err.Description)
                                    Err.Clear
                                End If
                                On Error GoTo 0
                                On Error GoTo ErrorHandler

                                copiedTableCount = wsNewWorkbookSheet.ListObjects.Count
                                Set wsOutputFilterRange = wsNewWorkbookSheet.UsedRange

                                On Error Resume Next
                                For tableIndex = copiedTableCount To 1 Step -1
                                    Set wsOutputTable = wsNewWorkbookSheet.ListObjects(tableIndex)
                                    wsOutputTable.Unlist
                                    If Err.Number <> 0 Then
                                        Call WriteLog("WARNING", "Could not convert copied (R) table to a range.", "Table / Error", wsOutputTable.Name & " / " & CStr(Err.Number) & ": " & Err.Description)
                                        Err.Clear
                                    End If
                                Next tableIndex

                                If wsNewWorkbookSheet.AutoFilterMode Then wsNewWorkbookSheet.AutoFilterMode = False
                                wsOutputFilterRange.AutoFilter
                                If Err.Number <> 0 Then
                                    Call WriteLog("WARNING", "Could not apply AutoFilter to copied (R) data range.", "Range / Error", wsOutputFilterRange.Address(False, False) & " / " & CStr(Err.Number) & ": " & Err.Description)
                                    Err.Clear
                                Else
                                    Call WriteLog("INFO", "Applied AutoFilter to copied (R) data range.", "Sheet / Range / Tables Converted", standaloneSheetName & " / " & wsOutputFilterRange.Address(False, False) & " / " & copiedTableCount)
                                End If
                                On Error GoTo 0
                                On Error GoTo ErrorHandler
                                Set wsOutputTable = Nothing
                                Set wsOutputFilterRange = Nothing

                                Call WriteLog("INFO", "Copied standalone (R) sheet as values while preserving its formatting.", "Sheet Name", standaloneSheetName)
                                Set wsNewWorkbookSheet = Nothing
                            Else
                                Call WriteLog("WARNING", "Standalone (R) sheet not found; skipping.", "Sheet Name", standaloneSheetName)
                            End If
                            Set wsStandaloneSource = Nothing
                        End If

                        If Len(templateSheetName) > 0 Then
                            standaloneSheetName = templateSheetName
                            Set wsStandaloneSource = Nothing
                            On Error Resume Next
                            Set wsStandaloneSource = ThisWorkbook.Sheets(standaloneSheetName)
                            On Error GoTo 0
                            On Error GoTo ErrorHandler

                            If Not wsStandaloneSource Is Nothing Then
                                wsStandaloneSource.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                On Error Resume Next
                                wsNewWorkbookSheet.Name = standaloneSheetName
                                If Err.Number <> 0 Then
                                    Call WriteLog("ERROR", "Failed to name standalone (T) sheet.", "Sheet / Error", standaloneSheetName & " / " & CStr(Err.Number) & ": " & Err.Description)
                                    Err.Clear
                                End If
                                On Error GoTo 0
                                On Error GoTo ErrorHandler
                                Call WriteLog("INFO", "Copied standalone (T) sheet with values, formulas, and formatting intact.", "Sheet Name", standaloneSheetName)
                                Set wsNewWorkbookSheet = Nothing
                            Else
                                Call WriteLog("WARNING", "Standalone (T) sheet not found; skipping.", "Sheet Name", standaloneSheetName)
                            End If
                            Set wsStandaloneSource = Nothing
                        End If
                    End If
                End If
            End If
NextSheetLoopInner:
        Next selectedSheetCol
        ' --- สิ้นสุด [ดัชนี 8.1] ---

        ' Restore standalone (T) formulas after all mapped sheets have been copied,
        ' so same-workbook sheet and table references can resolve in the output.
        For Each wsOutputSheet In newWorkbook.Worksheets
            If Len(wsOutputSheet.Name) >= 4 And Right$(wsOutputSheet.Name, 4) = " (T)" Then
                standaloneSheetName = wsOutputSheet.Name
                Set wsStandaloneSource = Nothing
                Set wsStandaloneFormulaCells = Nothing
                standaloneFormulaErrors = 0
                standaloneFormulaCellCount = 0
                firstStandaloneFormulaError = ""

                On Error Resume Next
                Set wsStandaloneSource = ThisWorkbook.Sheets(standaloneSheetName)
                wsOutputSheet.Unprotect Password:=SHEET_PASSWORD
                If Err.Number <> 0 Then
                    Call WriteLog("WARNING", "Could not unprotect standalone (T) sheet before restoring formulas.", "Sheet / Error", standaloneSheetName & " / " & CStr(Err.Number) & ": " & Err.Description)
                    Err.Clear
                End If
                If Not wsStandaloneSource Is Nothing Then
                    Set wsStandaloneFormulaCells = wsStandaloneSource.UsedRange.SpecialCells(xlCellTypeFormulas)
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set wsStandaloneFormulaCells = Nothing
                    End If
                End If
                On Error GoTo 0
                On Error GoTo ErrorHandler

                If Not wsStandaloneFormulaCells Is Nothing Then
                    standaloneFormulaCellCount = wsStandaloneFormulaCells.Cells.Count
                    For Each wsFormulaCell In wsStandaloneFormulaCells.Cells
                        progressCheckCounter = progressCheckCounter + 1
                        If progressCheckCounter >= 500 Then
                            Call RefreshProgressForm(frmProgress, StartTime)
                            progressCheckCounter = 0
                        End If
                        formulaCellText = CStr(wsFormulaCell.FormulaR1C1)
                        formulaCellAddress = wsFormulaCell.Address(False, False)
                        Set wsTargetFormulaCell = wsOutputSheet.Range(formulaCellAddress)

                        On Error Resume Next
                        wsTargetFormulaCell.FormulaR1C1 = formulaCellText
                        If Err.Number <> 0 Then
                            standaloneFormulaErrors = standaloneFormulaErrors + 1
                            If Len(firstStandaloneFormulaError) = 0 Then
                                firstStandaloneFormulaError = formulaCellAddress & " / " & formulaCellText & " / " & CStr(Err.Number) & ": " & Err.Description
                            End If
                            Err.Clear
                        End If
                        On Error GoTo 0
                        On Error GoTo ErrorHandler
                    Next wsFormulaCell
                End If

                If standaloneFormulaErrors > 0 Then
                    Call WriteLog("WARNING", "Some standalone (T) formulas could not be reapplied from the source sheet.", "Sheet / Count / First Error", standaloneSheetName & " / " & standaloneFormulaErrors & " / " & firstStandaloneFormulaError)
                Else
                    Call WriteLog("INFO", "Standalone (T) formulas reapplied from source after all sheets were copied.", "Sheet Name / Formula Cells", standaloneSheetName & " / " & standaloneFormulaCellCount)
                End If
                Set wsStandaloneFormulaCells = Nothing
                Set wsFormulaCell = Nothing
                Set wsTargetFormulaCell = Nothing
                Set wsStandaloneSource = Nothing
            End If
        Next wsOutputSheet

        ' STREAMING_CHUNK:Cleaning up default sheets, clearing unused rows, and setting sheet protection...
        ' --- [ดัชนี 8.2]: ลบชีทเริ่มต้นในเวิร์กบุ๊กใหม่ ---
        On Error Resume Next
        For Each wsTemp In newWorkbook.Sheets
            If Left(wsTemp.Name, 5) = "Sheet" And IsNumeric(Mid(wsTemp.Name, 6)) Then
                wsTemp.Delete
            End If
        Next wsTemp
        On Error GoTo 0
        On Error GoTo ErrorHandler
        Call WriteLog("INFO", "Deleted default sheets in new workbook. Workbook: '" & newWorkbook.Name & "'")
        ' --- สิ้นสุด [ดัชนี 8.2] ---
        
        ' --- [ดัชนี 8.3]: ล้างแถวที่ไม่ต้องการในแต่ละชีทของเวิร์กบุ๊กใหม่ ---
        Dim ws As Worksheet, lastRowInB As Long
        For Each ws In newWorkbook.Sheets
            If InStr(1, ws.Name, " (R)", vbTextCompare) = 0 And InStr(1, ws.Name, " (T)", vbTextCompare) = 0 Then
                lastRowInB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                If lastRowInB < ws.Rows.Count Then
                    ws.Rows(lastRowInB + 1 & ":" & ws.Rows.Count).Clear
                    Call WriteLog("DEBUG", "Cleared unused rows in sheet. Sheet: '" & ws.Name & "', Last Row with Data: " & lastRowInB)
                End If
            End If
        Next ws
        ' --- สิ้นสุด [ดัชนี 8.3] ---
        
        ' [8.4] ล็อกชีททั้งหมดด้วยรหัสผ่าน (ตามคำขอของผู้ใช้)
        '--------------------------------------------------------------------------------------------------------------------
        Dim ws2 As Worksheet
        For Each ws2 In newWorkbook.Worksheets
            
            ws2.Unprotect Password:=SHEET_PASSWORD
        
            ws2.Protect Password:=SHEET_PASSWORD, _
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

        ' STREAMING_CHUNK:Saving output workbook to folder path...
        ' --- [ดัชนี 8.5]: การดำเนินการบันทึกไฟล์ ---
        ' คอลัมน์ AD (คอลัมน์ที่ 30 ในอาร์เรย์ MasterMapping ฐาน 1) ประกอบด้วยเส้นทางไฟล์ใหม่
        If collProfitCentersForReviewer.Count > 0 Then
            newFilePath = CStr(collProfitCentersForReviewer.item(1)(30)) ' คอลัมน์ AD
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
                On Error GoTo ErrorHandler
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
            On Error GoTo ErrorHandler
SkipFileSave:
            On Error GoTo ErrorHandler
            newWorkbook.Close SaveChanges:=False
            Set newWorkbook = Nothing
        Else
            Call WriteLog("WARNING", "No file path provided, workbook closed without saving. Reviewer Name: '" & currentReviewerName_str & "'")
            newWorkbook.Close SaveChanges:=False
            Set newWorkbook = Nothing
        End If
    Next currentReviewerName_str
    ' --- สิ้นสุด [ดัชนี 8] ---

    ' STREAMING_CHUNK:Cleaning up and restoring Excel application settings...
    ' --- [ดัชนี 9]: คืนค่าการตั้งค่า Excel และล้างตัวแปร ---
CleanUp:
    On Error Resume Next
    If applicationStateCaptured Then
        Application.ScreenUpdating = savedScreenUpdating
        Application.Calculation = savedCalculation
        Application.EnableEvents = savedEnableEvents
        Application.DisplayAlerts = savedDisplayAlerts
    End If
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
    ' ตรวจสอบว่า arrMasterData ถูกกำหนดค่าและไม่ว่างเปล่าก่อนล้างค่า
    If Not IsEmpty(arrMasterData) And IsArray(arrMasterData) Then Erase arrMasterData
    ' ตรวจสอบว่า arrCurrentSheetFinalData ถูกกำหนดค่าและไม่ว่างเปล่าก่อนล้างค่า
    If Not IsEmpty(arrCurrentSheetFinalData) And IsArray(arrCurrentSheetFinalData) Then Erase arrCurrentSheetFinalData
    ' ตรวจสอบว่า arrTemplateFormulas ถูกกำหนดค่าและไม่ว่างเปล่าก่อนล้างค่า
    If IsArray(arrTemplateFormulas) Then Erase arrTemplateFormulas
    Call WriteLog("INFO", "All array variables cleared.")

    endTime = Timer
    runTime = endTime - StartTime
    ' --- สิ้นสุด [ดัชนี 9] ---

    ' STREAMING_CHUNK:Updating UserForm and displaying final completion message...
    ' --- [ดัชนี 10]: อัปเดต UserForm และแสดงข้อความสิ้นสุดการทำงาน ---
    If runFailed Then
        If progressFormShown Then
            Unload frmProgress
            progressFormShown = False
        End If
        MsgBox "Macro stopped due to runtime error " & runtimeErrorNumber & _
            " (" & runtimeErrorSource & "): " & runtimeErrorDescription, vbCritical
    Else
        If progressFormShown Then
            With frmProgress
                .lblProgress.Caption = "Processing Complete!"
                minutes = Int(runTime / 60)
                seconds = Int(runTime Mod 60)
                .lblTime.Caption = "Total Time: " & Format(minutes, "00") & ":" & Format(seconds, "00")
                If .Visible Then Application.Wait Now + TimeValue("00:00:03")
                Unload frmProgress
            End With
            progressFormShown = False
            Set frmProgress = Nothing
            Call WriteLog("INFO", "Progress UserForm closed.")
        End If

        minutes = Int(runTime / 60)
        seconds = Int(runTime Mod 60)
        MsgBox "Completed Run Time: " & Format(minutes, "00") & " minutes and " & Format(seconds, "00") & " seconds"
        Call WriteLog("INFO", "Macro finished.", "Total Run Time", Format(minutes, "00") & ":" & Format(seconds, "00"))
    End If
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET_NAME)
    If Not wsLog Is Nothing Then wsLog.Columns("A:E").AutoFit
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    runtimeErrorNumber = Err.Number
    runtimeErrorDescription = Err.Description
    runtimeErrorSource = Err.Source
    runFailed = True
    On Error Resume Next
    Call WriteLog("ERROR", "Unexpected runtime error.", "Number / Source / Description", CStr(runtimeErrorNumber) & " / " & runtimeErrorSource & " / " & runtimeErrorDescription)
    Resume CleanUp

End Sub

Private Sub RefreshProgressForm(ByVal progressForm As Object, ByVal startTime As Double, Optional ByVal forceUpdate As Boolean = False)
    Static lastUpdate As Double
    Dim currentTimer As Double
    Dim elapsedSeconds As Double
    Dim hours As Long
    Dim elapsedMinutes As Long
    Dim elapsedRemainderSeconds As Long

    currentTimer = Timer
    If Not forceUpdate Then
        If currentTimer >= lastUpdate Then
            If currentTimer - lastUpdate < 5 Then Exit Sub
        ElseIf currentTimer + 86400 - lastUpdate < 5 Then
            Exit Sub
        End If
    End If

    elapsedSeconds = currentTimer - startTime
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400
    hours = Int(elapsedSeconds / 3600)
    elapsedMinutes = Int((elapsedSeconds Mod 3600) / 60)
    elapsedRemainderSeconds = Int(elapsedSeconds Mod 60)

    progressForm.lblTime.Caption = "Elapsed: " & Format(hours, "00") & ":" & Format(elapsedMinutes, "00") & ":" & Format(elapsedRemainderSeconds, "00")
    progressForm.Repaint
    DoEvents
    lastUpdate = currentTimer
End Sub

' STREAMING_CHUNK:Writing helper function for logging...
' --- ฟังก์ชันผู้ช่วยสำหรับการบันทึก Log (จากต้นแบบ) ---
Sub WriteLog(logType As String, message As String, Optional varName As String = "", Optional varValue As Variant)
    ' วัตถุประสงค์: บันทึกข้อมูล Log ลงในชีทที่กำหนด
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim logDetail As String
    Dim currentWorkbook As Workbook
    Set currentWorkbook = ThisWorkbook

    On Error GoTo ErrorHandler

    ' ตรวจสอบว่ามีชีท Log อยู่หรือไม่ หากไม่มีให้สร้างใหม่
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
        End With
    End If

    ' ค้นหาแถวถัดไปที่ว่างอยู่ในชีท Log
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    ' สร้างข้อความรายละเอียดตัวแปรหากมีการระบุ varName
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

    ' เขียนรายการ Log ลงในชีท Log
    With wsLog
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = logType
        .Cells(nextRow, 3).Value = message
        .Cells(nextRow, 4).Value = varName
        .Cells(nextRow, 5).Value = logDetail
    End With

    Exit Sub

ErrorHandler:
    Debug.Print "Error in WriteLog function: " & Err.Description & " (Log Type: " & logType & ", Message: " & message & ")"
    On Error GoTo 0
End Sub
