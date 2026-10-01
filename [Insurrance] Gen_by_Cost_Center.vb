' STREAMING_CHUNK:Declaring global constants and subroutines...
' โค้ด VBA สำหรับ Standard Module (เช่น Module1)

' กำหนดค่าคงที่แบบ Public สำหรับชื่อ Sheet ที่ใช้เก็บ Log
Public Const LOG_SHEET_NAME As String = "Log"
Public Const SHEET_PASSWORD As String = "MTIEPBCS"

Sub genfile_ByProfitCenter_V1_9() ' ทำงานร่วมกับ Log ได้อย่างถูกต้อง

    ' STREAMING_CHUNK:Setting up error handling...
    ' [0.1] การจัดการข้อผิดพลาดระดับ Global Error Handling (การเปลี่ยนแปลงใน V1.2)
    ' ช่วยให้มั่นใจว่าหากเกิด Runtime Error ที่ไม่ได้รับการจัดการ แมโครจะข้ามไปยังส่วน CleanUp
    ' เพื่อคืนค่าการตั้งค่าต่างๆ ของโปรแกรม และปิด UserForm
    On Error GoTo CleanUp

    ' วัตถุประสงค์: แมโครนี้ช่วยสร้างไฟล์ Excel ตามเงื่อนไขสำหรับแต่ละ Profit Center โดยอัตโนมัติ
    ' โดยจะอ่านค่าการตั้งค่าจากแผ่นงาน "MasterMapping" รวบรวมข้อมูลจากแผ่นงาน "Master (M)" ต่างๆ
    ' และใช้แผ่นงาน "Template (T)" รวมถึง "Report (R)" เพื่อสร้างไฟล์ .xlsx ใหม่ที่มีข้อมูลพร้อมสำหรับแต่ละ Profit Center ที่เลือก
    ' เวอร์ชันนี้ได้รับการปรับปรุงประสิทธิภาพโดยประมวลผลข้อมูลในหน่วยความจำผ่าน Array และ Dictionary
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.1] การประกาศตัวแปร - เวลาและ Object
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Declaring timing and sheet variables...
    Dim StartTime As Double           ' เก็บเวลาเริ่มทำงานของแมโครเพื่อวัดประสิทธิภาพ
    Dim endTime As Double             ' เก็บเวลาสิ้นสุดการทำงานของแมโคร
    Dim runTime As Double             ' คำนวณเวลาการทำงานทั้งหมด
    Dim wsMasterMapping As Worksheet    ' Object สำหรับแผ่นงาน "MasterMapping"
    Dim wsTemp As Worksheet             ' Worksheet Object ชั่วคราวที่ใช้ระหว่างเคลียร์แผ่นงาน (ลบ Sheet เริ่มต้น)
    Dim wsNewWorkbookSheet As Worksheet ' Object สำหรับแผ่นงานภายใน Workbook ที่สร้างขึ้นใหม่
    Dim wsOriginalTemplate As Worksheet ' Object สำหรับแผ่นงาน Template จาก Workbook ดั้งเดิม
    Dim wsMasterSource As Worksheet     ' Object สำหรับวนลูปอ่านแผ่นงาน Master (M)
    Dim wsLog As Worksheet              ' Object สำหรับแผ่นงาน Log ซึ่งใช้บันทึกรายละเอียดการทำงานของแมโคร

    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.2] การประกาศตัวแปร - การจัดเก็บข้อมูล (Arrays และ Dictionaries)
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Declaring data storage arrays and dictionaries...
    Dim arrMasterMapping As Variant        ' Array สำหรับเก็บข้อมูลจากแผ่นงาน "MasterMapping" เพื่อให้เข้าถึงได้รวดเร็ว
                                           ' บรรจุข้อมูลการตั้งค่าสำหรับ Profit Center และการเลือกแผ่นงาน
    Dim arrMasterData As Variant           ' Array ชั่วคราวสำหรับโหลดข้อมูลจากแผ่นงาน "Master (M)" แต่ละแผ่น
                                           ' ใช้ถ่ายโอนข้อมูลจากแผ่นงานเข้าสู่หน่วยความจำเพื่อประมวลผล
    Dim arrCurrentSheetFinalData As Variant ' Array สำหรับสร้างข้อมูลและสูตรขั้นสุดท้ายของแผ่นงานเฉพาะก่อนเขียนลงแผ่นงาน
                                           ' Array นี้จะเก็บเนื้อหาที่จะเขียนลงในแผ่นงานของ Workbook ใหม่
    Dim arrTemplateFormulas As Variant     ' Array สำหรับเก็บสูตรรูปแบบ R1C1 จากแผ่นงาน Template ดั้งเดิม
                                           ' ใช้รักษาโครงสร้างสูตรเมื่อทำการคัดลอกข้อมูล
    Dim dictAllMasterData As Object        ' Dictionary หลักสำหรับเก็บข้อมูลแผ่นงาน Master (M) ทั้งหมด จัดกลุ่มตามชื่อแผ่นงาน
    Set dictAllMasterData = CreateObject("Scripting.Dictionary") ' Key: ชื่อแผ่นงาน Master (M), Value: Dictionary ย่อย
                                                                ' Dictionary ย่อยจะจัดกลุ่มข้อมูลตาม Profit Center ID
    ' Key ของ Dictionary ย่อย: Profit Center ID, Value: Collection ของแถวข้อมูล (แต่ละแถวคือ Array 1 มิติ)
    Dim dictSheetTemplates As Object       ' Dictionary สำหรับจับคู่ชื่อแผ่นงานที่สร้างขึ้นกับชื่อแผ่นงาน Template
    Set dictSheetTemplates = CreateObject("Scripting.Dictionary") ' Key: ชื่อแผ่นงานที่สร้าง (เช่น "Admin" จาก "Admin (M)"),
                                                                ' Value: ชื่อแผ่นงาน Template (เช่น "Admin (T)")
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.3] การประกาศตัวแปร - ตัวนับการวนลูปและการจัดการข้อมูล
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Declaring loop counters and data handling variables...
    Dim i As Long, j As Long, k As Long    ' ตัวนับการวนลูปทั่วไปสำหรับการท่องไปใน Array และ Collection
    Dim dataRowCounter As Long             ' ตัวนับสำหรับแถวใน arrCurrentSheetFinalData โดยเฉพาะขณะเติมข้อมูล
    Dim col As Long                        ' ตัวนับการวนลูปสำหรับคอลัมน์ภายใน Array ข้อมูล
    Dim masterSourceSheetName As String    ' ชื่อของแผ่นงาน Master (M) (เช่น "Admin (M)")
    Dim templateSheetName As String        ' ชื่อของแผ่นงาน Template (T) (เช่น "Admin (T)")
    Dim createdSheetName As String         ' ชื่อสุดท้ายสำหรับแผ่นงานใน Workbook ใหม่ (เช่น "Admin" หรือ "Admin (M)")
    Dim masterKey As String                ' Key สำหรับ dictSingleMasterGrouped (ปกติคือ Profit Center ID)
    Dim masterRowData As Variant           ' Array 1 มิติที่เป็นตัวแทนของข้อมูล Master หนึ่งแถว
    Dim masterRowsForProfitCenter As Collection ' Collection ของแถวข้อมูล Master สำหรับ Profit Center ID เฉพาะ
    Dim dictSingleMasterGrouped As Object  ' Dictionary ชั่วคราวเพื่อจัดกลุ่มข้อมูลตาม Profit Center ID สำหรับแผ่นงาน Master (M) หนึ่งแผ่น
    Dim profitCenterID As String           ' Profit Center ID จาก "MasterMapping" (คอลัมน์ A)
    Dim profitCenterName As String         ' ชื่อ Profit Center จาก "MasterMapping" (คอลัมน์ B)
    Dim r1_condition_met As Boolean        ' ตัวแปรสถานะสำหรับเงื่อนไขเฉพาะจาก "MasterMapping" คอลัมน์ T (ไม่ได้ใช้งานในเวอร์ชันนี้แต่ประกาศไว้)
    Dim templateCellContent As Variant     ' เนื้อหาเซลล์จาก Template ซึ่งอาจเป็นสูตรหรือค่าข้อมูล

    Dim lastRowMasterMapping As Long       ' แถวสุดท้ายที่มีข้อมูลใน "MasterMapping"
    Dim lastColMasterMapping As Long       ' คอลัมน์สุดท้ายที่มีข้อมูลในแถวที่ 1 ของแผ่นงาน "MasterMapping"
    Dim lastColSheets As Long              ' คอลัมน์สุดท้ายที่มีการเลือกแผ่นงานใน "MasterMapping" (คอลัมน์ D-AC)
    Dim lastRowMasterSource As Long        ' แถวสุดท้ายที่มีข้อมูลในแผ่นงาน "Master (M)"
    Dim lastColMasterSource As Long        ' คอลัมน์สุดท้ายที่มีข้อมูลในแผ่นงาน "Master (M)"
    Dim numMasterDataCols As Long          ' จำนวนคอลัมน์ในข้อมูล Master
    Dim numRowsData As Long                ' จำนวนแถวข้อมูลสำหรับ Profit Center เฉพาะ
    Dim selectedSheetCol As Long           ' ตัวนับการวนลูปสำหรับคอลัมน์ที่เป็นตัวแทนของแผ่นงานที่จะประมวลผล
    Dim effectiveLastColForArray As Long   ' คอลัมน์สุดท้ายจริงที่จะโหลดเข้า arrMasterMapping โดยพิจารณาจากคอลัมน์ที่เกี่ยวข้องทั้งหมด
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.4] การประกาศตัวแปร - การดำเนินการกับไฟล์
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Declaring file operation variables...
    Dim newFilePath As String              ' พาธโฟลเดอร์หลักจาก "MasterMapping" คอลัมน์ AD (ดัชนี 30)
    Dim folderPath As String               ' พาธเต็มไปยังโฟลเดอร์ผลลัพธ์สำหรับ Profit Center ปัจจุบัน
    Dim fileName As String                 ' ชื่อของไฟล์ Excel ใหม่ที่จะบันทึก
    Dim newWorkbook As Workbook            ' Object สำหรับ Excel Workbook ที่สร้างขึ้นใหม่
    Dim baseFileName As String             ' ชื่อหลักสำหรับไฟล์ที่สร้างขึ้น (เช่น "PCID - PCName")
    Dim fso As Object                      ' FileSystemObject สำหรับการจัดการโฟลเดอร์/ไฟล์ (Late binding)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim copyNum As Long                    ' ตัวนับสำหรับชื่อไฟล์ที่ซ้ำกัน เพื่อให้ออกชื่อไฟล์ไม่ซ้ำกัน
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.5] การประกาศตัวแปร - ความคืบหน้าของ UserForm
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Declaring progress form variables...
    Dim frmProgress As New frmProgress     ' Instance ของ UserForm แสดงความคืบหน้าเพื่อแสดงสถานะของแมโคร
    Dim totalFilesToProcess As Long        ' จำนวนไฟล์ทั้งหมดที่จะสร้างขึ้นตามการเลือกใน "MasterMapping"
    Dim filesProcessed As Long             ' ตัวนับจำนวนไฟล์ที่ประมวลผลเสร็จแล้ว
    Dim currentTime As Double              ' เวลาที่ผ่านไปสำหรับอัปเดต UserForm
    Dim minutes As Long                    ' ส่วนประกอบนาทีของเวลาที่ผ่านไป
    Dim seconds As Long                    ' ส่วนประกอบวินาทีของเวลาที่ผ่านไป
    Dim formulaErrorsCount As Long         ' จำนวนสูตรที่ไม่สามารถเขียนลงชีทปลายทางได้
    ' --------------------------------------------------------------------------------------------------------------------
    ' [2.1] การเริ่มต้นการทำงานของแมโคร - การปรับแต่งประสิทธิภาพและเวลาเริ่มต้น
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Initializing macro performance settings...
    StartTime = Timer                       ' บันทึกเวลาเริ่มต้นเพื่อการวัดประสิทธิภาพ
    Application.ScreenUpdating = False      ' ปิดการอัปเดตหน้าจอเพื่อเพิ่มความเร็วในการทำงานของแมโคร
    Application.Calculation = xlCalculationManual ' ตั้งค่าการคำนวณเป็น Manual เพื่อป้องกันไม่ให้ Excel คำนวณซ้ำบ่อยๆ
    Application.EnableEvents = False        ' ปิด Event ต่างๆ เพื่อป้องกันไม่ให้ทำงานระหว่างแมโครรัน
    Application.DisplayAlerts = False       ' ซ่อนข้อความแจ้งเตือนของระบบและ Message Box
    formulaErrorsCount = 0

    ' --- LOGGING: เริ่มต้นการทำงานของแมโคร ---
    Call WriteLog("INFO", "Macro started: genfile_ByProfitCenter_V1_8")

    ' --- [2.2] เตรียมแผ่นงาน Log ---
    ' พยายามตั้งค่าอ้างอิงไปยังแผ่นงาน Log
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0

    ' หากแผ่นงาน Log มีอยู่และมีข้อมูลเกินแถวส่วนหัว ให้ทำการล้างข้อมูล
    If Not wsLog Is Nothing Then
        If wsLog.Cells(Rows.Count, "A").End(xlUp).Row > 1 Then
            wsLog.Range("A2:E" & wsLog.Cells(Rows.Count, "A").End(xlUp).Row).ClearContents
            Call WriteLog("INFO", "Cleared existing data in Log sheet.")
        End If
    ' หากยังไม่มีแผ่นงาน Log ฟังก์ชัน WriteLog จะสร้างขึ้นให้อัตโนมัติในการเรียกใช้งานครั้งแรก
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [3.1] โหลดข้อมูล MasterMapping
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Loading MasterMapping sheet data...
    
            Dim wsActionPage As Worksheet
            Dim sheetName As String

        ' กำหนดตัวแปร wsActionPage ให้เรียกดึงแผ่นงาน "Action_Page"
            Set wsActionPage = ThisWorkbook.Sheets("Action_Page")

        ' ดึงชื่อแผ่นงานจาก Cell A6 ของแผ่นงาน Action Page
        sheetName = wsActionPage.Range("A6").Value

        ' กำหนดตัวแปร wsMasterMapping จากชื่อแผ่นงานที่ได้จาก sheetName
        Set wsMasterMapping = ThisWorkbook.Sheets(sheetName)
        
        'Set wsMasterMapping = ThisWorkbook.Sheets("MasterMapping")

    ' ค้นหาแถวสุดท้ายที่มีข้อมูลในคอลัมน์ A โดยพิจารณาข้อมูล Profit Center เริ่มตั้งแต่แถวที่ 5
    lastRowMasterMapping = wsMasterMapping.Cells(Rows.Count, "A").End(xlUp).Row
    ' ตรวจสอบว่ามีข้อมูลเพียงพอสำหรับ Profit Center เริ่มตั้งแต่แถวที่ 5 หรือไม่
    If lastRowMasterMapping < 5 Then
        MsgBox "MasterMapping sheet is empty or has only headers up to row 4, or no profit center data found.", vbExclamation, "Data Error"
        Call WriteLog("WARNING", "MasterMapping sheet is empty or has only headers up to row 4, or no profit center data found.", "lastRowMasterMapping", lastRowMasterMapping)
        GoTo CleanUp ' ออกจากแมโครหากพบว่าไม่มีข้อมูลเพียงพอ
    End If
    
    ' ค้นหาคอลัมน์สุดท้ายที่มีข้อมูลในแถวที่ 1 ของแผ่นงาน "MasterMapping"
    lastColMasterMapping = wsMasterMapping.Cells(1, Columns.Count).End(xlToLeft).Column
    If lastColMasterMapping < 1 Then lastColMasterMapping = 1 ' รับประกันว่ามีอย่างน้อย 1 คอลัมน์เพื่อความปลอดภัย
    
    ' กำหนดคอลัมน์สุดท้ายสำหรับการเลือกแผ่นงาน คอลัมน์ AC คือคอลัมน์ที่ 30
    lastColSheets = 29 ' กำหนดค่าคงที่ตามตรรกะเดิม
    
    ' กำหนดคอลัมน์สุดท้ายจริงสำหรับการโหลดข้อมูล "MasterMapping" เข้าสู่ Array
    ' โดยใช้ค่าที่มากกว่าระหว่างคอลัมน์สุดท้ายจริงกับคอลัมน์สุดท้ายของการเลือกแผ่นงาน (AC)
    effectiveLastColForArray = Application.WorksheetFunction.Max(lastColMasterMapping, lastColSheets)
    
    ' โหลดข้อมูลจากขอบเขตที่กำหนดของ "MasterMapping" เข้าสู่ Variant Array เพื่อการประมวลผลที่รวดเร็วขึ้น
    ' ขอบเขตในปัจจุบันรวมทุกแถวเริ่มตั้งแต่แถวที่ 5 (ข้ามส่วนหัวและการกำหนดแผ่นงาน (R))
    ' และทุกคอลัมน์จนถึงคอลัมน์สุดท้ายจริง
    arrMasterMapping = wsMasterMapping.Range(wsMasterMapping.Cells(5, "A"), wsMasterMapping.Cells(lastRowMasterMapping, effectiveLastColForArray)).Value
    Call WriteLog("INFO", "MasterMapping data for Profit Centers loaded into arrMasterMapping array (starting from row 5).", "Rows Loaded", UBound(arrMasterMapping, 1))

    ' [3.2] แสดงค่า Master Mapping (วัตถุประสงค์เพื่อการตรวจสอบ/Debugging)
    ' STREAMING_CHUNK:Debugging MasterMapping content...
    Debug.Print "---------------------------------------"
    Debug.Print "Content of arrMasterMapping array (Profit Center data only):"
    ' --- LOGGING: เนื้อหาใน MasterMapping Array (สำหรับการตรวจสอบแบบละเอียด) ---
    Call WriteLog("DEBUG", "Starting to log content of arrMasterMapping array (Profit Center data only) to Immediate Window for brevity.")
    
    ' วนลูปผ่านแต่ละแถวของ MasterMapping Array
    For i = LBound(arrMasterMapping, 1) To UBound(arrMasterMapping, 1)
        Dim rowData As String
        rowData = "Row " & i & ": "
    
        ' วนลูปผ่านแต่ละคอลัมน์เพื่อต่อค่าเข้าด้วยกันเป็น String เดียว
        For j = LBound(arrMasterMapping, 2) To UBound(arrMasterMapping, 2)
            rowData = rowData & " | Col " & j & ": " & arrMasterMapping(i, j)
        Next j
    
        ' พิมพ์ String ของแถวที่รวมแล้วไปยัง Immediate Window
        Debug.Print rowData
        ' ทางเลือก: บันทึกข้อมูลแต่ละแถวลงแผ่นงาน Log หากจำเป็น แต่ข้อมูลอาจมีปริมาณมาก
        ' Call WriteLog("DEBUG", "MasterMapping Row Data", "Row " & i, rowData)
    Next i
    
    Debug.Print "---------------------------------------"
    Call WriteLog("DEBUG", "Finished logging content of arrMasterMapping array.")


    '--------------------------------------------------------------------------------------------------------------------
    ' [4.1] ระบุ Template และการจับคู่แผ่นงาน
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Mapping sheet templates...
    ' วนลูปผ่านคอลัมน์ต่างๆ ของ "MasterMapping" (ตั้งแต่ D ถึง AC หรือคอลัมน์ 4 ถึง lastColSheets)
    ' เพื่อระบุและจับคู่ชื่อแผ่นงาน Master (จากแถวที่ 2) กับชื่อแผ่นงาน Template ที่ตรงกัน (จากแถวที่ 3)
    For j = 4 To lastColSheets
        ' ดึงชื่อแผ่นงาน Master จากแถวที่ 2 ของ MasterMapping
        masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, j).Value))
        ' ดึงชื่อแผ่นงาน Template จากแถวที่ 3 ของ MasterMapping
        templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, j).Value))

        ' ดำเนินการต่อเมื่อระบุชื่อแผ่นงาน Master (M) (ไม่จำเป็นต้องมี Template (T))
        If Len(masterSourceSheetName) > 0 Then
            ' สรุปชื่อแผ่นงานสุดท้ายสำหรับ Workbook ใหม่โดยการลบ " (M)" ออกจากชื่อแผ่นงาน Master
            createdSheetName = Replace(masterSourceSheetName, " (M)", "")
            ' ระบบสำรองเพื่อดึงชื่อจากชื่อ Template หากชื่อ Master ไม่มี " (M)"
            If createdSheetName = masterSourceSheetName And InStr(1, masterSourceSheetName, " (M)", vbTextCompare) > 0 Then
                If Len(templateSheetName) > 0 Then
                    createdSheetName = Replace(templateSheetName, " (T)", "")
                End If
            End If

            ' หากได้ชื่อ createdSheetName ที่ถูกต้อง ให้เพิ่มลงใน Dictionary
            If Len(createdSheetName) > 0 Then
                ' เพิ่มการจับคู่ลงใน Dictionary หากยังไม่มีอยู่
                If Not dictSheetTemplates.Exists(createdSheetName) Then
                    dictSheetTemplates.Add createdSheetName, IIf(Len(templateSheetName) > 0, templateSheetName, masterSourceSheetName)
                    Call WriteLog("INFO", "Added sheet mapping.", "Created Sheet -> Source Sheet", createdSheetName & " -> " & IIf(Len(templateSheetName) > 0, templateSheetName, masterSourceSheetName))
                End If
            End If
        End If
    Next j

    ' ตรวจสอบว่าพบการจับคู่แผ่นงานที่ถูกต้องบ้างหรือไม่
    If dictSheetTemplates.Count = 0 Then
        MsgBox "No valid sheet mappings found in MasterMapping (Row 2 & 3, Columns D-AC). Please ensure Row 2 contains master sheet names like 'Sheet (M)'.", vbExclamation, "Configuration Error"
        Call WriteLog("WARNING", "No valid sheet mappings found.", "dictSheetTemplates.Count", dictSheetTemplates.Count)
        GoTo CleanUp ' ออกจากแมโครหากไม่พบการจับคู่
    Else
        Call WriteLog("INFO", "Identified sheet templates and mappings (from MasterMapping rows 2 & 3).", "Total Mappings", dictSheetTemplates.Count)
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [5.1] โหลดข้อมูล Master (M) ทั้งหมดเข้าสู่ Dictionary ล่วงหน้า
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Pre-populating master data into memory...
    ' วนลูปผ่านทุกแผ่นงานใน Workbook ปัจจุบันเพื่อค้นหาแผ่นงาน Master (M)
    For Each wsMasterSource In ThisWorkbook.Sheets
        ' ตรวจสอบว่าชื่อแผ่นงานมี " (M)" หรือไม่ ซึ่งระบุว่าเป็นแผ่นงานข้อมูล Master
        If InStr(1, wsMasterSource.Name, " (M)", vbTextCompare) > 0 Then
            ' ค้นหาแถวสุดท้ายที่มีข้อมูลในคอลัมน์ B (สมมติว่า Profit Center ID อยู่ในคอลัมน์ B) เพื่อกำหนดขอบเขตข้อมูล
            lastRowMasterSource = wsMasterSource.Cells(Rows.Count, "B").End(xlUp).Row
            If lastRowMasterSource >= 2 Then ' ตรวจสอบให้แน่ใจว่ามีข้อมูลเกินกว่าแถวส่วนหัว
                ' ค้นหาคอลัมน์สุดท้ายที่มีข้อมูลในแผ่นงาน Master ต้นทาง
                lastColMasterSource = wsMasterSource.Cells(1, wsMasterSource.Columns.Count).End(xlToLeft).Column
                If lastColMasterSource < 1 Then lastColMasterSource = 1 ' รับประกันว่ามีอย่างน้อย 1 คอลัมน์เพื่อความปลอดภัย

                ' โหลดขอบเขตข้อมูลทั้งหมดของแผ่นงาน Master เข้าสู่ Array ชั่วคราวเพื่อประสิทธิภาพ
                arrMasterData = wsMasterSource.Range(wsMasterSource.Cells(2, "A"), wsMasterSource.Cells(lastRowMasterSource, lastColMasterSource)).Value
                Call WriteLog("INFO", "Loaded Master (M) sheet data.", "Sheet Name", wsMasterSource.Name & " (Rows: " & UBound(arrMasterData, 1) & ", Cols: " & UBound(arrMasterData, 2) & ")")

                ' สร้าง Dictionary ชั่วคราวเพื่อจัดกลุ่มข้อมูลตาม Profit Center ID สำหรับแผ่นงาน Master ปัจจุบัน
                Set dictSingleMasterGrouped = CreateObject("Scripting.Dictionary")

                ' วนลูปผ่านแต่ละแถวของ Array ข้อมูล Master
                For k = LBound(arrMasterData, 1) To UBound(arrMasterData, 1)
                    ' ดึง Profit Center ID จากคอลัมน์ที่สอง (ดัชนี 2)
                    masterKey = Trim(CStr(arrMasterData(k, 2)))
                    ' สร้าง Array 1 มิติจากแถวปัจจุบันของ Array ข้อมูล Master 2 มิติ
                    ReDim masterRowData(1 To UBound(arrMasterData, 2))
                    For col = LBound(arrMasterData, 2) To UBound(arrMasterData, 2)
                        masterRowData(col) = arrMasterData(k, col)
                    Next col

                    ' เพิ่มแถวข้อมูลปัจจุบันเข้าสู่ Collection ที่เชื่อมโยงกับ Profit Center ID
                    If Not dictSingleMasterGrouped.Exists(masterKey) Then
                        Set masterRowsForProfitCenter = New Collection
                        dictSingleMasterGrouped.Add masterKey, masterRowsForProfitCenter
                    Else
                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(masterKey)
                    End If
                    masterRowsForProfitCenter.Add masterRowData
                Next k
                ' เพิ่ม Dictionary ชั่วคราว (ซึ่งเก็บข้อมูลจัดกลุ่มตาม PC ID) ลงใน Dictionary หลัก
                ' Key ของ Dictionary หลักคือชื่อแผ่นงาน Master (เช่น "Admin (M)")
                dictAllMasterData.Add wsMasterSource.Name, dictSingleMasterGrouped
                Call WriteLog("INFO", "Grouped master data by Profit Center ID.", "Master Sheet", wsMasterSource.Name & " (Unique PC IDs: " & dictSingleMasterGrouped.Count & ")")
            End If
        End If
    Next wsMasterSource

    ' ล้างข้อมูลใน Array ชั่วคราวของ Master เพื่อคืนพื้นที่หน่วยความจำหลังจากประมวลผล
    If Not IsEmpty(arrMasterData) Then Erase arrMasterData

    ' ตรวจสอบว่ามีข้อมูล Master ถูกโหลดสำเร็จจากทุกแผ่นงานหรือไม่
    If dictAllMasterData.Count = 0 Then
        MsgBox "No Master sheets with names containing ' (M)' and data found in this workbook.", vbExclamation, "Data Error"
        Call WriteLog("ERROR", "No Master (M) sheets found or loaded with data.", "dictAllMasterData.Count", dictAllMasterData.Count)
        GoTo CleanUp ' ออกจากแมโครหากไม่พบข้อมูล Master
    Else
        Call WriteLog("INFO", "Successfully pre-populated dictionary with all Master (M) data.", "Total Master Sheets", dictAllMasterData.Count)
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [6.1] คำนวณจำนวนไฟล์ทั้งหมดที่จะประมวลผล
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Counting total target files...
    totalFilesToProcess = 0
    ' นำนวนแถวใน MasterMapping Array ที่คอลัมน์ C (ดัชนี 3) ถูกทำเครื่องหมายเป็น TRUE
    ' ตัวชี้นี้ระบุว่า Profit Center ใดถูกเลือกสำหรับการสร้างไฟล์
    For i = LBound(arrMasterMapping, 1) To UBound(arrMasterMapping, 1)
        ' เมื่อ arrMasterMapping เริ่มจากแถวที่ 5 สมาชิกแรก (i=1) จะตรงกับ Profit Center แรก
        ' เราสมมติว่าคอลัมน์ที่ 3 (C) ภายใน Array นี้ยังคงเก็บค่า TRUE/FALSE สำหรับการเลือก
        If UBound(arrMasterMapping, 2) >= 3 Then ' ตรวจสอบให้แน่ใจว่ามีคอลัมน์ C ใน Array
            If arrMasterMapping(i, 3) = True Then ' ตรวจสอบว่า Flag "Generate" (คอลัมน์ C) ถูกตั้งค่าเป็น TRUE หรือไม่
                totalFilesToProcess = totalFilesToProcess + 1
            End If
        Else
            ' คำเตือนนี้สำคัญขึ้นเนื่องจาก arrMasterMapping เริ่มต้นจากข้อมูล Profit Center
            Call WriteLog("ERROR", "MasterMapping column C (selection column) not found within the profit center data range. Check MasterMapping sheet structure.", "MasterMapping Array Row Index", i)
        End If
    Next i
    Call WriteLog("INFO", "Calculated total files to process.", "Total Files to Generate", totalFilesToProcess)

    ' ตรวจสอบว่ามีไฟล์ถูกเลือกเพื่อประมวลผลหรือไม่
    If totalFilesToProcess = 0 Then
        MsgBox "No files selected for processing in MasterMapping Column C. Please mark 'TRUE' in Column C for files you wish to generate.", vbExclamation, "Selection Error"
        Call WriteLog("WARNING", "No files selected for processing in MasterMapping Column C.")
        GoTo CleanUp ' ออกจากแมโครหากไม่มีไฟล์ถูกเลือก
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [7.1] เริ่มต้นและแสดง UserForm ความคืบหน้า
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Displaying progress form...
    With frmProgress
        .lblProgress.Caption = "Processing: 0 / " & totalFilesToProcess
        .lblTime.Caption = "Time Elapsed: 00:00"
        .Show vbModeless ' แสดง UserForm แบบ non-modal เพื่อให้แมโครทำงานต่อไปได้
    End With
    filesProcessed = 0 ' เริ่มต้นตัวนับสำหรับไฟล์ที่ประมวลผลแล้ว
    Call WriteLog("INFO", "Progress UserForm initialized and displayed.")

    ' --------------------------------------------------------------------------------------------------------------------
    ' [8.1] ลูปหลัก - สร้างไฟล์สำหรับแต่ละ Profit Center ที่เลือก
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Processing main generation loop...
    ' วนลูปผ่านแต่ละแถวใน Array ข้อมูล "MasterMapping" โดยแต่ละแถวเป็นตัวแทนของ Profit Center
    For i = LBound(arrMasterMapping, 1) To UBound(arrMasterMapping, 1)

        If arrMasterMapping(i, 3) = True Then ' ตรวจสอบว่า Profit Center ปัจจุบัน (แถว) ถูกเลือกเพื่อประมวลผลหรือไม่ (คอลัมน์ C เป็น TRUE)
            filesProcessed = filesProcessed + 1
            Call WriteLog("INFO", "Starting file generation for a new Profit Center.", "Processed Count", filesProcessed & " / " & totalFilesToProcess)

            ' อัปเดตความคืบหน้าและเวลาที่ใช้ไปบน UserForm
            frmProgress.lblProgress.Caption = "Processing: " & filesProcessed & " / " & totalFilesToProcess
            currentTime = Timer - StartTime
            minutes = Int(currentTime / 60)
            seconds = Int(currentTime Mod 60)
            frmProgress.lblTime.Caption = "Time Elapsed: " & Format(minutes, "00") & ":" & Format(seconds, "00")
            DoEvents ' อนุญาตให้หน้าจอ UI รีเฟรชและอัปเดต UserForm เพื่อป้องกันไม่ให้โปรแกรมค้าง

            Set newWorkbook = Application.Workbooks.Add(xlWBATWorksheet) ' สร้าง Workbook เปล่าใหม่สำหรับ Profit Center ปัจจุบัน
            ' ดึง Profit Center ID และชื่อ จาก Array "MasterMapping" (คอลัมน์ A และ B)
            profitCenterID = Trim(CStr(arrMasterMapping(i, 1)))
            profitCenterName = Trim(CStr(arrMasterMapping(i, 2)))
            Call WriteLog("INFO", "New workbook created for Profit Center.", "Profit Center ID/Name", profitCenterID & " - " & profitCenterName)

            ' --------------------------------------------------------------------------------------------------------------------
            ' [8.1.1] ตรรกะการคัดลอกแผ่นงานที่ลงท้ายด้วย (R)
            ' ส่วนนี้จัดการคัดลอกแผ่นงานทั้งหมดที่ระบุด้วย "(R)" จาก MasterMapping (แถวที่ 4)
            ' ไปยัง Workbook ใหม่ โดยแปลงเซลล์ทั้งหมดเป็นค่าข้อมูล (Value) พร้อมทั้งคงรูปแบบเดิมไว้
            ' --------------------------------------------------------------------------------------------------------------------
            ' STREAMING_CHUNK:Copying report sheets ending with (R)...
            For selectedSheetCol = 4 To lastColSheets ' วนลูปผ่านคอลัมน์ D ถึง AC สำหรับข้อกำหนดแผ่นงาน
                Dim reportSheetName As String
                ' ดึงชื่อแผ่นงานจากแถวที่ 4 ของ MasterMapping สำหรับคอลัมน์ปัจจุบัน
                reportSheetName = Trim(CStr(wsMasterMapping.Cells(4, selectedSheetCol).Value))

                ' ตรวจสอบว่าชื่อแผ่นงานที่ลงท้ายด้วย " (R)" ถูกระบุไว้อย่างถูกต้องและมีแผ่นงานนั้นอยู่ใน ThisWorkbook หรือไม่
                If Len(reportSheetName) > 0 And InStr(1, reportSheetName, " (R)", vbTextCompare) > 0 Then
                    Dim wsOriginalReport As Worksheet
                    Set wsOriginalReport = Nothing ' กำหนดค่า Worksheet Object ให้เป็น Nothing
                    On Error Resume Next           ' ข้ามข้อผิดพลาดชั่วคราวหากไม่พบแผ่นงาน Report
                    Set wsOriginalReport = ThisWorkbook.Sheets(reportSheetName)
                    On Error GoTo 0                ' เปิดการจัดการข้อผิดพลาดกลับมาตามปกติ

                    If Not wsOriginalReport Is Nothing Then
                        ' คัดลอกแผ่นงาน (R) ทั้งแผ่นไปยัง Workbook ที่สร้างขึ้นใหม่
                        wsOriginalReport.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                        Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count) ' อ้างอิงไปยังแผ่นงานที่เพิ่งคัดลอกใหม่

                        ' เปลี่ยนชื่อแผ่นงานที่คัดลอกมาให้เป็นชื่อเดิมรวมถึง "(R)"
                        wsNewWorkbookSheet.Name = reportSheetName

                        ' แปลงทุกเซลล์ในแผ่นงานใหม่เป็นค่าข้อมูล (Value) พร้อมคงรูปแบบการจัดวาง
                        If Not wsNewWorkbookSheet.UsedRange Is Nothing Then
                            On Error Resume Next
                            ' คัดลอกเนื้อหาทั้งหมด (รวมถึงรูปแบบ) ของขอบเขตที่มีการใช้งาน
                            wsNewWorkbookSheet.UsedRange.Copy

                            ' Paste special: วางแบบค่าข้อมูล (Values) เพื่อแปลงสูตรให้เป็นผลลัพธ์คำนวณ
                            wsNewWorkbookSheet.UsedRange.PasteSpecial xlPasteValues
                            
                            ' Paste special: วางแบบรูปแบบ (Formats) เพื่อนำรูปแบบเดิมกลับมาใช้อีกครั้ง
                            wsNewWorkbookSheet.UsedRange.PasteSpecial xlPasteFormats

                            Application.CutCopyMode = False ' เคลียร์ Clipboard หลังจากการวาง
                            If Err.Number <> 0 Then
                                Call WriteLog("ERROR", "Failed to convert (R) sheet cells to values while preserving format.", "Error Description", Err.Description)
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                        Call WriteLog("INFO", "Copied (R) sheet and converted all cells to values (preserved format).", "Sheet Name", reportSheetName)
                        Set wsNewWorkbookSheet = Nothing ' ยกเลิกการอ้างอิง Worksheet Object เพื่อคืนพื้นที่หน่วยความจำ
                    Else
                        Call WriteLog("WARNING", "Report (R) sheet not found in the current workbook, skipping copy operation.", "Sheet Name", reportSheetName)
                    End If
                End If
            Next selectedSheetCol
            Call WriteLog("INFO", "Finished copying all selected (R) sheets for current Profit Center.")

            ' --- สิ้นสุดตรรกะการคัดลอกแผ่นงานที่ลงท้ายด้วย (R) ---


            ' --------------------------------------------------------------------------------------------------------------------
            ' [8.2] ลูปย่อย - คัดลอกและเติมข้อมูลลงในแผ่นงานของ Workbook ใหม่
            ' ส่วนนี้จัดการการคัดลอกแผ่นงาน Template (T) (ไม่ว่าจะคู่กับแผ่นงาน Master (M) หรือไม่ก็ตาม)
            ' และเติมข้อมูลลงไป
            ' --------------------------------------------------------------------------------------------------------------------
            ' STREAMING_CHUNK:Populating sheets for profit center...
            For selectedSheetCol = 4 To lastColSheets
                If UBound(arrMasterMapping, 2) >= selectedSheetCol Then ' ตรวจสอบให้แน่ใจว่าคอลัมน์มีอยู่ใน MasterMapping Array
                    If arrMasterMapping(i, selectedSheetCol) = True Then ' ตรวจสอบว่าแผ่นงานเฉพาะนี้ถูกเลือกสำหรับ Profit Center ปัจจุบันหรือไม่

                        Dim currentMasterSourceSheetName As String
                        Dim currentTemplateSheetName As String
                        ' ดึงชื่อแผ่นงาน Master (M) และ Template (T) จากแถวหัวข้อของ MasterMapping (แถวที่ 2 และ 3)
                        currentMasterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, selectedSheetCol).Value)) ' จากแถวที่ 2 (Master)
                        currentTemplateSheetName = Trim(CStr(wsMasterMapping.Cells(3, selectedSheetCol).Value))   ' จากแถวที่ 3 (Template)

                        ' --- เริ่มตรรกะใหม่: [8.2.1] จัดการแผ่นงาน Template (T) ที่ไม่มีข้อมูล Master (M) (V1.4) ---
                        ' บล็อกนี้ตรวจสอบว่ามีการกำหนดแผ่นงาน Template (แถว 3) แต่ไม่มีแผ่นงาน Master (แถว 2 ว่างเปล่า)
                        If Len(currentTemplateSheetName) > 0 And Len(currentMasterSourceSheetName) = 0 Then
                            Call WriteLog("INFO", "Processing a (T) sheet without a corresponding (M) sheet.", "Template Name", currentTemplateSheetName)

                            Dim wsOriginalTSheet As Worksheet
                            Set wsOriginalTSheet = Nothing
                            On Error Resume Next ' ข้ามข้อผิดพลาดชั่วคราวหากไม่พบแผ่นงาน Template
                            Set wsOriginalTSheet = ThisWorkbook.Sheets(currentTemplateSheetName)
                            On Error GoTo 0      ' เปิดการจัดการข้อผิดพลาดกลับมาตามปกติ

                            If Not wsOriginalTSheet Is Nothing Then
                                ' คัดลอกแผ่นงาน (T) ไปยัง Workbook ใหม่
                                wsOriginalTSheet.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                wsNewWorkbookSheet.Name = currentTemplateSheetName ' คงนามสกุล (T) สำหรับชื่อแผ่นงานใน Workbook ใหม่

                                Dim lastRowUsed As Long, lastColUsed As Long
                                ' ค้นหาแถวและคอลัมน์สุดท้ายที่มีการใช้งานในแผ่นงานที่คัดลอกมา
                                ' ใช้ LookIn:=xlValues เพื่อให้แน่ใจว่าเซลล์ที่มีค่าจริงถูกพิจารณาเป็น 'UsedRange'
                                On Error Resume Next ' จัดการกรณีที่แผ่นงานว่างเปล่าสิ้นเชิง (วิธี Find จะเกิด Error)
                                With wsNewWorkbookSheet
                                    ' ค้นหาแถวสุดท้ายที่มีเนื้อหา (ค่าข้อมูลหรือสูตร)
                                    If Not .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
                                        lastRowUsed = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
                                    Else
                                        lastRowUsed = 1 ' สมมติว่ามีอย่างน้อยแถวส่วนหัวหากไม่มีข้อมูล หรือหากแผ่นงานว่างเปล่าจริง
                                    End If
                                    ' ค้นหาคอลัมน์สุดท้ายที่มีเนื้อหา (ค่าข้อมูลหรือสูตร)
                                    If Not .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
                                        lastColUsed = .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
                                    Else
                                        lastColUsed = 1 ' สมมติว่ามีอย่างน้อย 1 คอลัมน์
                                    End If
                                End With
                                On Error GoTo 0
                                
                                
                                ' [8.2.1.1] แปลงเป็นค่าข้อมูลก่อน จากนั้นใช้สูตรซ้ำเพื่อลบ External Link ออก
                                ' กระบวนการ 2 ขั้นตอนนี้ช่วยให้มั่นใจว่าการอ้างอิงภายนอกทั้งหมดถูกตัดออก ก่อนฝังสูตรภายในไฟล์ใหม่
                                If lastRowUsed >= 1 And lastColUsed >= 1 Then
                                    Dim targetRange As Range
                                    Dim sourceRange As Range

                                    ' กำหนดขอบเขตเป้าหมายใน Workbook ใหม่และขอบเขตต้นทางใน Workbook ดั้งเดิม
                                    Set targetRange = wsNewWorkbookSheet.Range(wsNewWorkbookSheet.Cells(1, 1), wsNewWorkbookSheet.Cells(lastRowUsed, lastColUsed))
                                    Set sourceRange = wsOriginalTSheet.Range(wsOriginalTSheet.Cells(1, 1), wsOriginalTSheet.Cells(lastRowUsed, lastColUsed))

                                    ' ขั้นตอนที่ 1: แปลงเซลล์ทั้งหมดใน UsedRange ของแผ่นงานใหม่เป็นค่าข้อมูลเพื่อตัด External Link ทั้งหมด
                                    targetRange.Value = targetRange.Value
                                    Call WriteLog("INFO", "Converted (T) sheet to values to break external links.", "Sheet Name", currentTemplateSheetName)

                                    ' ขั้นตอนที่ 2: คัดลอกสูตรจากแผ่นงาน Template ดั้งเดิมแล้วนำมาปรับใช้กับแผ่นงานใหม่
                                    ' ซึ่งเป็นการฝังสูตรกลับเข้าไปใหม่ โดยทำให้สูตรอ้างอิงภายใน Workbook ใหม่เท่านั้น
                                    targetRange.FormulaR1C1 = sourceRange.FormulaR1C1
                                    Call WriteLog("INFO", "Re-applied formulas from original (T) sheet.", "Sheet Name", currentTemplateSheetName)

                                    ' เคลียร์ Range Object
                                    Set targetRange = Nothing
                                    Set sourceRange = Nothing
                                End If

                                ' [8.2.1.2] เติม Profit Center ID และชื่อ ลงในคอลัมน์ B และ C
                                ' ข้อมูลควรเริ่มต้นตั้งแต่แถวที่ 2
                                If lastRowUsed >= 2 Then ' เติมข้อมูลเฉพาะเมื่อมีแถวให้เติมถัดจากแถวส่วนหัว
                                    wsNewWorkbookSheet.Range("B2:B" & lastRowUsed).Value = profitCenterID
                                    wsNewWorkbookSheet.Range("C2:C" & lastRowUsed).Value = profitCenterName
                                    Call WriteLog("INFO", "Populated Profit Center ID/Name in (T) sheet (no M).", "Sheet Name", currentTemplateSheetName & " (Rows 2-" & lastRowUsed & ")")

                                    ' --- เริ่มตรรกะใหม่ (V1.7) ---
                                    ' เติมคอลัมน์ A ด้วยค่าจาก MasterMapping คอลัมน์ AF (ดัชนี 32)
                                    ' โดยสมมติว่าคอลัมน์ AF คือตำแหน่งที่เก็บค่าที่ต้องการใน MasterMapping สำหรับ Profit Center ปัจจุบัน
                                    If UBound(arrMasterMapping, 2) >= 32 Then ' ตรวจสอบให้แน่ใจว่าคอลัมน์ AF (ดัชนี 32) มีอยู่ใน MasterMapping Array
                                        wsNewWorkbookSheet.Range("A2:A" & lastRowUsed).Value = Trim(CStr(arrMasterMapping(i, 32)))
                                        
                                        ' --- เริ่มการแก้ไข (V1.8) ---
                                        ' รวมรายละเอียดตัวแปรสำหรับ WriteLog ให้ตรงกับจำนวนพารามิเตอร์
                                        Dim logVarName As String
                                        Dim logVarValue As String
                                        logVarName = "Sheet Name / Value"
                                        logVarValue = currentTemplateSheetName & " (Rows 2-" & lastRowUsed & ") / " & Trim(CStr(arrMasterMapping(i, 32)))
                                        Call WriteLog("INFO", "Populated Column A in (T) sheet with MasterMapping AF value.", logVarName, logVarValue)
                                        ' --- สิ้นสุดการแก้ไข (V1.8) ---
                                    Else
                                        Call WriteLog("WARNING", "MasterMapping Column AF (index 32) not found, skipping population of Column A in (T) sheet.", "Sheet Name", currentTemplateSheetName)
                                    End If
                                    ' --- สิ้นสุดตรรกะใหม่ (V1.7) ---
                                End If



                                Call WriteLog("INFO", "Successfully copied and processed (T) sheet (without M).", "Sheet Name", currentTemplateSheetName)
                                Set wsNewWorkbookSheet = Nothing ' ยกเลิกการอ้างอิงหลังการใช้งาน
                                Set wsOriginalTSheet = Nothing
                            Else
                                Call WriteLog("WARNING", "Template (T) sheet not found in current workbook for (T) without (M) configuration, skipping.", "Template Name", currentTemplateSheetName)
                            End If
                        ' --- สิ้นสุดตรรกะใหม่ (V1.4) ---

                        ' --- เริ่มตรรกะใหม่: [8.2.2] จัดการแผ่นงานที่มี (M) (ไม่จำเป็นต้องมี (T)) ---
                        ' ส่วนนี้จะทำงานเมื่อระบุชื่อแผ่นงาน Master (M) ในแถวที่ 2
                        ElseIf Len(currentMasterSourceSheetName) > 0 Then
                            
                            masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, selectedSheetCol).Value))
                            templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, selectedSheetCol).Value))
                            createdSheetName = Replace(masterSourceSheetName, " (M)", "")
                            Call WriteLog("INFO", "Processing sheet for new workbook.", "Sheet Info", masterSourceSheetName & " (Template: " & templateSheetName & ")")

                            If dictSheetTemplates.Exists(createdSheetName) Then
                                Dim wsOriginalMaster As Worksheet
                                Set wsOriginalMaster = Nothing
                                On Error Resume Next
                                Set wsOriginalMaster = ThisWorkbook.Sheets(masterSourceSheetName)
                                On Error GoTo 0

                                Set wsOriginalTemplate = Nothing
                                If Len(templateSheetName) > 0 Then
                                    On Error Resume Next
                                    Set wsOriginalTemplate = ThisWorkbook.Sheets(templateSheetName)
                                    On Error GoTo 0
                                End If

                                Dim wsFormulaSource As Worksheet
                                Set wsFormulaSource = wsOriginalMaster

                                If wsOriginalMaster Is Nothing Then
                                    Call WriteLog("WARNING", "Master sheet (M) not found, skipping sheet for current profit center.", "Master Name", masterSourceSheetName)
                                ElseIf dictAllMasterData.Exists(masterSourceSheetName) Then
                                    Set dictSingleMasterGrouped = dictAllMasterData.item(masterSourceSheetName)

                                    If dictSingleMasterGrouped.Exists(profitCenterID) Then
                                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(profitCenterID)

                                        If masterRowsForProfitCenter.Count > 0 Then
                                            If IsArray(masterRowsForProfitCenter.item(1)) Then
                                                numMasterDataCols = UBound(masterRowsForProfitCenter.item(1))
                                            Else
                                                Call WriteLog("WARNING", "Master data for sheet is not in expected array format, skipping sheet.", "Master Sheet", masterSourceSheetName)
                                            End If
                                            numRowsData = masterRowsForProfitCenter.Count

                                            ' จำกัดขอบเขตคอลัมน์ไม่ให้เกินคอลัมน์ที่มีข้อมูลจริง
                                            Dim lastColMasterActual As Long
                                            lastColMasterActual = wsOriginalMaster.Cells(1, wsOriginalMaster.Columns.Count).End(xlToLeft).Column
                                            If numMasterDataCols > lastColMasterActual Then
                                                numMasterDataCols = lastColMasterActual
                                            End If

                                            Call WriteLog("INFO", "Master data found for sheet and profit center.", "Rows/Cols", numRowsData & " rows, " & numMasterDataCols & " columns")

                                            ' โหลดสูตรจากแผ่นงานต้นแบบ
                                            If Not wsFormulaSource Is Nothing Then
                                                On Error Resume Next
                                                arrTemplateFormulas = wsFormulaSource.Range("A2").Resize(numRowsData, numMasterDataCols).FormulaR1C1
                                                On Error GoTo 0
                                                Call WriteLog("INFO", "Loaded formulas from source sheet: " & wsFormulaSource.Name)
                                            End If

                                            ' คัดลอกแผ่นงาน Master (M) ไปยัง Workbook ใหม่
                                            On Error Resume Next
                                            wsOriginalMaster.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                            Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                            wsNewWorkbookSheet.Name = masterSourceSheetName
                                            If Err.Number <> 0 Then
                                                Call WriteLog("ERROR", "Failed to copy or rename sheet in new workbook.", "Error Description", Err.Description)
                                                Err.Clear
                                            End If
                                            On Error GoTo 0

                                            If wsNewWorkbookSheet Is Nothing Then
                                                Call WriteLog("ERROR", "Sheet copy failed, skipping sheet for current profit center.", "Master Name", masterSourceSheetName)
                                            Else
                                                ' ปลดการป้องกันที่ติดมากับชีทต้นแบบก่อนแก้ไขข้อมูลและสูตร
                                                On Error Resume Next
                                                wsNewWorkbookSheet.Unprotect Password:=SHEET_PASSWORD
                                                On Error GoTo 0

                                                ' เคลียร์เนื้อหาตั้งแต่แถว 2 เป็นต้นไป โดยคงรูปแบบเดิมไว้
                                                If wsNewWorkbookSheet.Rows.Count >= 2 Then
                                                    On Error Resume Next
                                                    wsNewWorkbookSheet.Rows("2:" & wsNewWorkbookSheet.Rows.Count).ClearContents
                                                    If Err.Number <> 0 Then
                                                        Call WriteLog("ERROR", "Failed to clear contents starting from row 2.", "Error Description", Err.Description)
                                                        Err.Clear
                                                    End If
                                                    On Error GoTo 0
                                                End If
                                                Call WriteLog("INFO", "Master sheet copied and contents cleared from row 2 onwards.", "New Sheet Name", masterSourceSheetName)

                                                ReDim arrCurrentSheetFinalData(1 To numRowsData, 1 To numMasterDataCols)
                                                dataRowCounter = 0

                                                For Each masterRowData In masterRowsForProfitCenter
                                                    dataRowCounter = dataRowCounter + 1
                                                    For k = LBound(masterRowData) To UBound(masterRowData)
                                                        If k <= numMasterDataCols Then
                                                            arrCurrentSheetFinalData(dataRowCounter, k) = masterRowData(k)
                                                        End If
                                                    Next k

                                                    If UBound(arrMasterMapping, 2) >= 32 Then
                                                        arrCurrentSheetFinalData(dataRowCounter, 1) = Trim(CStr(arrMasterMapping(i, 32)))
                                                    Else
                                                        arrCurrentSheetFinalData(dataRowCounter, 1) = ""
                                                        Call WriteLog("WARNING", "Column 32 (AF) in MasterMapping not found for row.", "MasterMapping Row Index", i)
                                                    End If

                                                    arrCurrentSheetFinalData(dataRowCounter, 2) = profitCenterID
                                                    arrCurrentSheetFinalData(dataRowCounter, 3) = profitCenterName
                                                Next masterRowData

                                                ' ขั้นตอนที่ 1: เขียนค่าข้อมูลดิบลงในชีท
                                                wsNewWorkbookSheet.Range("A2").Resize(numRowsData, numMasterDataCols).Value = arrCurrentSheetFinalData
                                                Call WriteLog("INFO", "Data values written to new sheet.", "Sheet Name", masterSourceSheetName)

                                                ' ขั้นตอนที่ 2: วนลูปเขียนสูตรเฉพาะเซลล์ที่มีสูตรจริง ๆ (ป้องกัน Out of Memory)
                                                If IsArray(arrTemplateFormulas) Then
                                                    Dim rIdx As Long
                                                    Dim cellFormula As String

                                                    For k = 1 To numMasterDataCols
                                                        For rIdx = 1 To numRowsData
                                                            If rIdx <= UBound(arrTemplateFormulas, 1) And k <= UBound(arrTemplateFormulas, 2) Then
                                                                cellFormula = CStr(arrTemplateFormulas(rIdx, k))
                                                                If Left(cellFormula, 1) = "=" And Len(cellFormula) > 1 Then
                                                                    On Error Resume Next
                                                                    wsNewWorkbookSheet.Cells(rIdx + 1, k).FormulaR1C1 = cellFormula
                                                                    If Err.Number <> 0 Then
                                                                        formulaErrorsCount = formulaErrorsCount + 1
                                                                        Call WriteLog("WARNING", "Could not apply formula to generated sheet.", "Cell / Formula", wsNewWorkbookSheet.Cells(rIdx + 1, k).Address(False, False) & " / " & cellFormula)
                                                                        Err.Clear
                                                                    End If
                                                                    On Error GoTo 0
                                                                End If
                                                            End If
                                                        Next rIdx
                                                    Next k
                                                    Call WriteLog("INFO", "Formulas applied cell by cell for formula cells only.", "Sheet Name", masterSourceSheetName)
                                                End If
                                                ' --- คืนหน่วยความจำ Array หลังใช้งานเสร็จ ---
                                                Erase arrCurrentSheetFinalData
                                                If IsArray(arrTemplateFormulas) Then Erase arrTemplateFormulas

                                            End If ' ปิด If wsNewWorkbookSheet
                                        Else
                                            Call WriteLog("WARNING", "No master data rows found for the specified Profit Center ID.", "Profit Center ID", profitCenterID)
                                        End If
                                    Else
                                        Call WriteLog("WARNING", "Profit Center ID not found in grouped master data for sheet.", "Profit Center ID", profitCenterID & " - Master Sheet: " & masterSourceSheetName)
                                    End If
                                Else
                                    Call WriteLog("WARNING", "Master data dictionary for sheet is empty or not found.", "Master Sheet", masterSourceSheetName)
                                meEndIf: ' (ลบหรือย้ายไว้ตำแหน่งที่ถูกต้อง)
                                End If
                            Else
                                Call WriteLog("ERROR", "Template mapping for created sheet name not found, unexpected condition.", "Created Sheet Name", createdSheetName)
                            End If

                        ' --- สิ้นสุดตรรกะใหม่: [8.2.2] จัดการแผ่นงานที่มี (M) ---

                        Else
                            ' [8.2.3] บันทึก Log สำหรับคอลัมน์ที่เลือกแต่นอกเหนือเงื่อนไข (เช่น มีเฉพาะชื่อ (M) หรือรูปแบบไม่ถูกต้อง)
                            Call WriteLog("DEBUG", "MasterMapping column selected but no valid (T)-only or (M)/(T) pair configuration found.", _
                                          "Column Index", selectedSheetCol & " - Master (M): " & currentMasterSourceSheetName & ", Template (T): " & currentTemplateSheetName)
                        End If
                    End If
                End If
NextSheetLoop: ' Label นี้ส่วนใหญ่มีไว้สำหรับอ้างอิงย้อนหลัง หรือเผื่อกรณีนำ GoTo กลับมาใช้ในบล็อกย่อย
            Next selectedSheetCol
            Call WriteLog("INFO", "Finished processing all selected sheets for current Profit Center.")

            
            ' ลบแผ่นงานเริ่มต้น (เช่น "Sheet1", "Sheet2") ออกจาก Workbook ใหม่
            On Error Resume Next ' ข้ามข้อผิดพลาดชั่วคราวระหว่างการลบแผ่นงาน
            For Each wsTemp In newWorkbook.Sheets
                If Left(wsTemp.Name, 5) = "Sheet" And IsNumeric(Mid(wsTemp.Name, 6)) Then
                    wsTemp.Delete
                    Call WriteLog("INFO", "Deleted default sheet from new workbook.", "Sheet Name", wsTemp.Name)
                End If
            Next wsTemp
            On Error GoTo 0 ' เปิดการจัดการข้อผิดพลาดกลับมาตามปกติ
            
                    ' --- เริ่มตรรกะใหม่: ล้างข้อมูลใต้แถวที่มีการใช้งานล่าสุดในคอลัมน์ B สำหรับทุกแผ่นงาน ---
                    ' [8.3] การเคลียร์ข้อมูลขั้นสุดท้ายก่อนบันทึก
                    ' ลูปนี้จะวนผ่านทุกแผ่นงานใน Workbook ที่สร้างขึ้นใหม่ เพื่อเคลียร์ข้อมูล สูตร
                    ' และรูปแบบที่ไม่จำเป็นออก ทั้งหมดที่อยู่ใต้แถวสุดท้ายที่มีข้อมูลในคอลัมน์ B
                    ' --------------------------------------------------------------------------------------------------------------------
                    ' STREAMING_CHUNK:Cleaning unused rows below column B...
                    Dim ws As Worksheet
                    Dim lastRowInB As Long
                    Call WriteLog("INFO", "Starting final cleanup for all sheets in new workbook.", "Profit Center ID", profitCenterID)

                    For Each ws In newWorkbook.Sheets
                        ' ค้นหาแถวสุดท้ายที่มีข้อมูลเฉพาะในคอลัมน์ B ของแผ่นงานปัจจุบัน
                        lastRowInB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

                        ' ตรวจสอบว่าแถวสุดท้ายที่มีข้อมูลไม่ใช่แถวสุดท้ายของแผ่นงาน
                        If lastRowInB < ws.Rows.Count Then
                            ' กำหนดขอบเขตตั้งแต่แถวถัดจากแถวข้อมูลสุดท้ายไปจนถึงท้ายแผ่นงาน แล้วล้างข้อมูลทั้งหมด
                            ws.Rows(lastRowInB + 1 & ":" & ws.Rows.Count).Clear
                            Call WriteLog("INFO", "Cleared excess rows below the last data row in Column B.", "Sheet Name / Last Row", ws.Name & " / " & lastRowInB)
                        Else
                            Call WriteLog("DEBUG", "No rows to clear; data extends to the end of the sheet or sheet is empty.", "Sheet Name", ws.Name)
                        End If
                    Next ws
                    Call WriteLog("INFO", "Finished final cleanup for all sheets.")
                    ' --- สิ้นสุดตรรกะใหม่ ---
                      '--------------------------------------------------------------------------------------------------------------------

            '--------------------------------------------------------------------------------------------------------------------
            ' [8.4] ล็อกทุกแผ่นงานด้วยรหัสผ่าน (ตามความต้องการของผู้ใช้)
            '--------------------------------------------------------------------------------------------------------------------
            ' STREAMING_CHUNK:Protecting sheets with password...
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
            ' --------------------------------------------------------------------------------------------------------------------
            ' [9.1] การดำเนินการบันทึกไฟล์
            ' --------------------------------------------------------------------------------------------------------------------
            ' STREAMING_CHUNK:Saving generated file...
            ' ดึงพาธไฟล์หลักจากคอลัมน์ AD (ดัชนี 30) ของ MasterMapping Array
            If UBound(arrMasterMapping, 2) >= 30 Then
                newFilePath = CStr(arrMasterMapping(i, 30))
            Else
                newFilePath = ""
            End If
            Call WriteLog("INFO", "Retrieved base file path from MasterMapping.", "newFilePath", newFilePath)

            If Len(newFilePath) > 0 Then ' ตรวจสอบว่ามีการระบุพาธไฟล์หรือไม่
                ' สร้างชื่อไฟล์หลักโดยใช้ข้อมูลจากคอลัมน์ AE
                baseFileName = CStr(arrMasterMapping(i, 31))

                ' แทนที่ตัวอักษรที่ไม่ได้รับอนุญาตในชื่อไฟล์ เพื่อให้แน่ใจว่าพาธไฟล์ถูกต้อง
                baseFileName = Replace(baseFileName, "/", "_")
                baseFileName = Replace(baseFileName, "\", "_")
                baseFileName = Replace(baseFileName, ":", "_")
                baseFileName = Replace(baseFileName, "*", "_")
                baseFileName = Replace(baseFileName, "?", "_")
                baseFileName = Replace(baseFileName, Chr(34), "_") ' ตัวอักษรอัญประกาศ (ฟันหนู)
                baseFileName = Replace(baseFileName, "<", "_")
                baseFileName = Replace(baseFileName, ">", "_")
                baseFileName = Replace(baseFileName, "|", "_")
                Call WriteLog("INFO", "Constructed base filename and sanitized invalid characters.", "baseFileName", baseFileName)

                ' สร้างพาธเต็มสำหรับโฟลเดอร์ผลลัพธ์ของ Profit Center ปัจจุบัน
                folderPath = newFilePath & Application.PathSeparator & baseFileName

                ' สร้างโฟลเดอร์หากยังไม่มีอยู่
                If Not fso.FolderExists(folderPath) Then
                    On Error Resume Next ' ข้ามข้อผิดพลาดชั่วคราวระหว่างการสร้างโฟลเดอร์
                    fso.CreateFolder folderPath
                    If Err.Number <> 0 Then
                        Call WriteLog("ERROR", "Failed to create folder.", "Error Description", Err.Description & " - Folder Path: " & folderPath)
                        Err.Clear
                        GoTo SkipFileSave ' ข้ามการบันทึกไฟล์หากสร้างโฟลเดอร์ไม่สำเร็จ
                    Else
                        Call WriteLog("INFO", "Created new output folder.", "folderPath", folderPath)
                    End If
                    On Error GoTo 0 ' เปิดการจัดการข้อผิดพลาดกลับมาตามปกติ
                Else
                    Call WriteLog("INFO", "Output folder already exists.", "folderPath", folderPath)
                End If

                fileName = baseFileName & ".xlsx"
                copyNum = 1

                ' ตรวจสอบไฟล์ที่มีชื่อเดียวกันซ้ำ และเพิ่มหมายเลขเพื่อให้ชื่อไฟล์ไม่ซ้ำกัน
                Do While fso.fileExists(folderPath & Application.PathSeparator & fileName)
                    copyNum = copyNum + 1
                    fileName = baseFileName & " (" & copyNum & ").xlsx"
                    Call WriteLog("INFO", "Duplicate filename detected, generating unique name.", "New Filename", fileName)
                Loop

                ' บันทึก Workbook ใหม่ด้วยชื่อไฟล์ที่ไม่ซ้ำ
                On Error Resume Next ' ข้ามข้อผิดพลาดชั่วคราวระหว่างบันทึกไฟล์
                newWorkbook.SaveAs fileName:=folderPath & Application.PathSeparator & fileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                If Err.Number <> 0 Then
                    Call WriteLog("ERROR", "Failed to save workbook.", "Error Description", Err.Description & " - Filename: " & fileName)
                    Err.Clear
                Else
                    Call WriteLog("INFO", "Workbook saved successfully.", "Full Path", folderPath & Application.PathSeparator & fileName)
                End If
                On Error GoTo 0 ' เปิดการจัดการข้อผิดพลาดกลับมาตามปกติ

SkipFileSave: ' Label สำหรับข้ามขั้นตอนการบันทึกหากเกิดข้อผิดพลาดในการสร้างโฟลเดอร์หรือปัญหาอื่นๆ
                newWorkbook.Close SaveChanges:=False ' ปิด Workbook โดยไม่บันทึกหากเกิดข้อผิดพลาด หรือพาธไม่ถูกต้อง
                Set newWorkbook = Nothing ' ยกเลิกการอ้างอิง Workbook Object เพื่อคืนพื้นที่หน่วยความจำ
                Call WriteLog("INFO", "Workbook closed.", "Profit Center ID", profitCenterID)
            Else
                newWorkbook.Close SaveChanges:=False ' ปิด Workbook โดยไม่บันทึกหากไม่ได้ระบุพาธไฟล์
                Set newWorkbook = Nothing
                Call WriteLog("WARNING", "No file path provided, workbook closed without saving.", "Profit Center ID", profitCenterID)
            End If
        End If
    Next i
    Call WriteLog("INFO", "Main loop completed. All selected files processed.")

    ' --------------------------------------------------------------------------------------------------------------------
    ' [10.1] การเคลียร์ระบบ - คืนค่าการตั้งค่าและคืนพื้นที่หน่วยความจำ
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Cleaning up memory and restoring application settings...
CleanUp: ' Label สำหรับการเคลียร์ระบบ เพื่อให้แน่ใจว่าการตั้งค่าต่างๆ ถูกคืนค่าแม้เกิดข้อผิดพลาด
    Application.ScreenUpdating = True        ' เปิดการอัปเดตหน้าจอกลับมา
    Application.EnableEvents = True          ' เปิดการใช้งาน Event กลับมา
    Application.DisplayAlerts = True         ' เปิดการแสดงข้อความแจ้งเตือนกลับมา
    Call WriteLog("INFO", "Application settings restored (ScreenUpdating, Calculation, Events, Alerts).")

    ' เคลียร์ตัวแปร Object ทั้งหมดเพื่อคืนพื้นที่หน่วยความจำอย่างมีประสิทธิภาพ ยกเว้น frmProgress ในตอนนี้
    Set wsMasterMapping = Nothing
    Set wsTemp = Nothing
    Set wsNewWorkbookSheet = Nothing
    Set wsOriginalTemplate = Nothing
    Set dictAllMasterData = Nothing
    Set dictSheetTemplates = Nothing
    Set dictSingleMasterGrouped = Nothing
    Set masterRowsForProfitCenter = Nothing
    Set fso = Nothing
    Set wsLog = Nothing
    ' --- เริ่มการแก้ไข (V1.3 - ลบการคืนค่า Set frmProgress = Nothing ก่อนเวลาอันควร) ---
    ' ลบออก: Set frmProgress = Nothing
    ' --- สิ้นสุดการแก้ไข ---
    Call WriteLog("INFO", "All object variables cleared (excluding UserForm for final update).")

    ' เคลียร์ตัวแปร Array ทั้งหมดเพื่อคืนพื้นที่หน่วยความจำ
    If Not IsEmpty(arrMasterMapping) Then Erase arrMasterMapping
    If Not IsEmpty(arrMasterData) Then Erase arrMasterData
    If Not IsEmpty(arrCurrentSheetFinalData) Then Erase arrCurrentSheetFinalData
    If Not IsEmpty(arrTemplateFormulas) Then Erase arrTemplateFormulas
    Call WriteLog("INFO", "All array variables cleared.")
    If formulaErrorsCount > 0 Then
        Call WriteLog("WARNING", "Some formulas could not be applied and require review in the Log sheet.", "Formula Error Count", formulaErrorsCount)
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [11.1] การคำนวณเวลารวมขั้นสุดท้ายและการอัปเดต UserForm
    ' --------------------------------------------------------------------------------------------------------------------
    ' STREAMING_CHUNK:Finalizing timing and unloading UserForm...
    endTime = Timer                          ' บันทึกเวลาสิ้นสุด
    runTime = endTime - StartTime            ' คำนวณเวลาประมวลผลทั้งหมด
    
    ' อัปเดตและปิด UserForm ความคืบหน้าหากยังทำงานอยู่
    If Not frmProgress Is Nothing Then
        With frmProgress
            If formulaErrorsCount > 0 Then
                .lblProgress.Caption = "Complete: " & formulaErrorsCount & " formula(s) need review in Log"
            Else
                .lblProgress.Caption = "Processing Complete!"
            End If
            minutes = Int(runTime / 60)
            seconds = Int(runTime Mod 60)
            .lblTime.Caption = "Total Time: " & Format(minutes, "00") & ":" & Format(seconds, "00")
            If .Visible Then Application.Wait Now + TimeValue("00:00:01") ' แสดงหน้าต่างค้างไว้ครู่หนึ่ง
            Unload frmProgress ' โหลด UserForm ออกจากหน่วยความจำ
        End With
        ' --- เริ่มการแก้ไข (V1.3 - ย้าย Set frmProgress = Nothing มาไว้ที่นี่) ---
        ' ยกเลิกการอ้างอิง UserForm Object หลังจาก Unload แล้ว
        Set frmProgress = Nothing
        ' --- สิ้นสุดการแก้ไข ---
        Call WriteLog("INFO", "Progress UserForm closed.")
    End If

    minutes = Int(runTime / 60)
    seconds = Int(runTime Mod 60)
    If formulaErrorsCount > 0 Then
        MsgBox "Completed Run Time: " & Format(minutes, "00") & " minutes and " & Format(seconds, "00") & " seconds" & vbCrLf & vbCrLf & _
               "Warning: " & formulaErrorsCount & " formula(s) need review in Log.", vbExclamation, "Macro Complete"
    Else
        MsgBox "Completed Run Time: " & Format(minutes, "00") & " minutes and " & Format(seconds, "00") & " seconds", vbInformation, "Macro Complete"
    End If
    Call WriteLog("INFO", "Macro finished.", "Total Run Time", Format(minutes, "00") & ":" & Format(seconds, "00"))

End Sub


' --------------------------------------------------------------------------------------------------------------------
' [12.1] ฟังก์ชัน WriteLog (ฟังก์ชันช่วยสำหรับการบันทึก Log)
' --------------------------------------------------------------------------------------------------------------------
' STREAMING_CHUNK:Defining WriteLog helper procedure...
Sub WriteLog(logType As String, message As String, Optional varName As String = "", Optional varValue As Variant)
    ' วัตถุประสงค์: บันทึกรายการ Log ลงในแผ่นงาน Log ที่กำหนด
    ' พารามิเตอร์:
    '   logType: ประเภทของรายการ Log (เช่น "INFO", "WARNING", "ERROR", "DEBUG")
    '   message: ข้อความหลักของรายการ Log
    '   varName: (ไม่บังคับ) ชื่อของตัวแปรที่จะบันทึก Log
    '   varValue: (ไม่บังคับ) ค่าของตัวแปรที่จะบันทึก Log

    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim logDetail As String
    Dim currentWorkbook As Workbook ' ประกาศ Workbook ปัจจุบันเพื่อหลีกเลี่ยงปัญหาจากการใช้ ActiveWorkbook

    Set currentWorkbook = ThisWorkbook ' รับประกันว่าทำงานกับ Workbook ที่บรรจุแมโครนี้อยู่

    On Error GoTo ErrorHandler ' เปิดการจัดการข้อผิดพลาดสำหรับฟังก์ชันนี้

    ' [12.2] ตรวจสอบว่าแผ่นงาน Log มีอยู่หรือไม่ หากไม่มีให้สร้างขึ้นใหม่
    On Error Resume Next ' ปิดการจัดการข้อผิดพลาดชั่วคราวสำหรับการเข้าถึงแผ่นงาน
    Set wsLog = currentWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0      ' เปิดการจัดการข้อผิดพลาดกลับมาตามปกติ

    If wsLog Is Nothing Then
        ' สร้างแผ่นงาน Log หากยังไม่มีอยู่
        Set wsLog = currentWorkbook.Sheets.Add(After:=currentWorkbook.Sheets(currentWorkbook.Sheets.Count))
        wsLog.Name = LOG_SHEET_NAME
        ' เพิ่มแถวหัวข้อในแผ่นงาน Log ที่สร้างขึ้นใหม่
        With wsLog
            .Cells(1, 1).Value = "Timestamp"
            .Cells(1, 2).Value = "Type"
            .Cells(1, 3).Value = "Message"
            .Cells(1, 4).Value = "Variable Name"
            .Cells(1, 5).Value = "Variable Value"
            .Rows(1).Font.Bold = True ' กำหนดแถวหัวข้อให้เป็นตัวหนา
            .Columns("A:E").AutoFit   ' ปรับขนาดคอลัมน์ให้อัตโนมัติเพื่อให้อ่านง่าย
        End With
    End If

    ' [12.3] ค้นหาแถวถัดไปที่ว่างในแผ่นงาน Log
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    ' [12.4] สร้าง String รายละเอียดตัวแปรหากมีการระบุ varName
    If varName <> "" Then
        ' พยายามแปลง varValue ให้เป็น String และจัดการ Object
        If IsObject(varValue) And Not IsEmpty(varValue) Then
            On Error Resume Next ' พยายามดึง Property ทั่วไป เช่น .Name หรือ .Value
            logDetail = varValue.Name
            If Err.Number <> 0 Then
                logDetail = "Object" ' ค่าเริ่มต้นเป็น "Object" หากไม่มี Property Name
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf IsError(varValue) Then
            logDetail = "Error: " & CStr(varValue)
        Else
            logDetail = CStr(varValue)
        End If
    Else
        logDetail = "" ' ไม่มีค่าตัวแปรให้บันทึก Log
    End If

    ' [12.5] บันทึกรายการ Log ลงในแผ่นงาน Log
    With wsLog
        .Cells(nextRow, 1).Value = Now      ' วันที่และเวลาปัจจุบัน
        .Cells(nextRow, 2).Value = logType  ' ประเภทของ Log (INFO, WARNING, ERROR, DEBUG)
        .Cells(nextRow, 3).Value = message  ' ข้อความหลักของ Log
        .Cells(nextRow, 4).Value = varName  ' ชื่อตัวแปร (ถ้ามี)
        .Cells(nextRow, 5).Value = logDetail ' รายละเอียด/ค่าตัวแปร (ถ้ามี)
        .Columns("A:E").AutoFit ' ปรับความกว้างคอลัมน์หลังจากบันทึกเพื่อให้ระเบียบเรียบร้อย
    End With

    Exit Sub ' ออกจากฟังก์ชันเมื่อทำงานสำเร็จ

ErrorHandler:
    ' [12.6] ระบบสำรองสำหรับข้อผิดพลาดของ Log โดยพิมพ์ไปยัง Immediate Window หาก WriteLog เองทำงานล้มเหลว
    Debug.Print "Error in WriteLog function: " & Err.Description & " (Log Type: " & logType & ", Message: " & message & ")"
    On Error GoTo 0 ' ตรวจสอบให้แน่ใจว่าได้รีเซ็ตการจัดการข้อผิดพลาด
End Sub