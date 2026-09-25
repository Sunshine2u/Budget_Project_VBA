' VBA Code for a Standard Module (e.g., Module1)

' Define a Public Constant for the Log Sheet Name
Public Const LOG_SHEET_NAME As String = "Log"

Sub genfile_ByProfitCenter_V1_9() ' OK with Log

    ' [0.1] Global Error Handling (V1.2 Change)
    ' This ensures that in case of any unhandled runtime error, the macro will
    ' jump to the CleanUp section to restore application settings and close the UserForm.
    On Error GoTo CleanUp

    ' Purpose: This macro automates the generation of customized Excel files for different profit centers.
    ' It reads configuration from a "MasterMapping" sheet, collects data from various "Master (M)" sheets,
    ' and uses "Template (T)" sheets and "Report (R)" sheets to create new, populated .xlsx files for each selected profit center.
    ' This version is optimized for performance using in-memory data processing with arrays and dictionaries.
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.1] Variable Declarations - Timing and Objects
    ' --------------------------------------------------------------------------------------------------------------------
    Dim StartTime As Double           ' Stores the macro's start time for performance tracking.
    Dim endTime As Double             ' Stores the macro's end time.
    Dim runTime As Double             ' Calculates total execution time.
    Dim wsMasterMapping As Worksheet    ' Object for the "MasterMapping" sheet.
    Dim wsTemp As Worksheet             ' Temporary worksheet object used during cleanup (deleting default sheets).
    Dim wsNewWorkbookSheet As Worksheet ' Object for a sheet within the newly created workbook.
    Dim wsOriginalTemplate As Worksheet ' Object for a template sheet from the original workbook.
    Dim wsMasterSource As Worksheet     ' Object for iterating through Master (M) sheets.
    Dim wsLog As Worksheet              ' Object for the Log sheet, where macro execution details are recorded.

    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.2] Variable Declarations - Data Storage (Arrays and Dictionaries)
    ' --------------------------------------------------------------------------------------------------------------------
    Dim arrMasterMapping As Variant        ' Array to hold data from the "MasterMapping" sheet for fast access.
                                           ' Contains configuration for profit centers and sheet selections.
    Dim arrMasterData As Variant           ' Temporary array to load data from individual "Master (M)" sheets.
                                           ' Used to transfer data from sheets to memory for processing.
    Dim arrCurrentSheetFinalData As Variant ' Array to build the final data and formulas for a specific sheet before writing.
                                           ' This array stores the content to be written to the new workbook's sheet.
    Dim arrTemplateFormulas As Variant     ' Array to hold R1C1 formulas from the original template sheet.
                                           ' Used to preserve formulas when copying data.
    Dim dictAllMasterData As Object        ' Main dictionary to store all Master (M) sheet data, grouped by sheet name.
    Set dictAllMasterData = CreateObject("Scripting.Dictionary") ' Key: Master (M) sheet name, Value: Another dictionary.
                                                                ' The inner dictionary groups data by Profit Center ID.
    ' Inner Dictionary Key: Profit Center ID, Value: Collection of rows (each row is a 1D array).
    Dim dictSheetTemplates As Object       ' Dictionary to map created sheet names to their template sheet names.
    Set dictSheetTemplates = CreateObject("Scripting.Dictionary") ' Key: Created sheet name (e.g., "Admin" from "Admin (M)"),
                                                                ' Value: Template sheet name (e.g., "Admin (T)").
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.3] Variable Declarations - Loop Counters and Data Handling
    ' --------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long    ' General loop counters for iterating through arrays and collections.
    Dim dataRowCounter As Long             ' Counter specifically for rows in `arrCurrentSheetFinalData` when populating.
    Dim col As Long                        ' Loop counter for columns within data arrays.
    Dim masterSourceSheetName As String    ' Name of the Master (M) sheet (e.g., "Admin (M)").
    Dim templateSheetName As String        ' Name of the Template (T) sheet (e.g., "Admin (T)").
    Dim createdSheetName As String         ' The final name for the sheet in the new workbook (e.g., "Admin" or "Admin (M)").
    Dim masterKey As String                ' Key for `dictSingleMasterGrouped` (typically Profit Center ID).
    Dim masterRowData As Variant           ' A 1D array representing a single row of master data.
    Dim masterRowsForProfitCenter As Collection ' Collection of master data rows for a specific Profit Center ID.
    Dim dictSingleMasterGrouped As Object  ' Temporary dictionary to group data by Profit Center ID for one Master (M) sheet.
    Dim profitCenterID As String           ' Profit Center ID from "MasterMapping" (Column A).
    Dim profitCenterName As String         ' Profit Center Name from "MasterMapping" (Column B).
    Dim r1_condition_met As Boolean        ' Flag for a specific condition from "MasterMapping" Column T (not used in this version but declared).
    Dim templateCellContent As Variant     ' Content of a cell from the template, potentially a formula or value.

    Dim lastRowMasterMapping As Long       ' Last row with data in "MasterMapping".
    Dim lastColMasterMapping As Long       ' Last column with data in row 1 of the "MasterMapping" sheet.
    Dim lastColSheets As Long              ' Last column with sheet selection in "MasterMapping" (Columns D-AC).
    Dim lastRowMasterSource As Long        ' Last row with data in a "Master (M)" sheet.
    Dim lastColMasterSource As Long        ' Last column with data in a "Master (M)" sheet.
    Dim numMasterDataCols As Long          ' Number of columns in the master data.
    Dim numRowsData As Long                ' Number of data rows for a specific profit center.
    Dim selectedSheetCol As Long           ' Loop counter for columns representing sheets to be processed.
    Dim effectiveLastColForArray As Long   ' Effective last column to load into arrMasterMapping, considering all relevant columns.
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.4] Variable Declarations - File Operations
    ' --------------------------------------------------------------------------------------------------------------------
    Dim newFilePath As String              ' Base directory path from "MasterMapping" Column AD (index 30).
    Dim folderPath As String               ' Full path to the output folder for the current profit center.
    Dim fileName As String                 ' Name of the new Excel file to be saved.
    Dim newWorkbook As Workbook            ' Object for the newly created Excel workbook.
    Dim baseFileName As String             ' Base name for the generated file (e.g., "PCID - PCName").
    Dim fso As Object                      ' FileSystemObject for folder/file management (Late binding).
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim copyNum As Long                    ' Counter for duplicate filenames, ensuring unique file names.
    ' --------------------------------------------------------------------------------------------------------------------
    ' [1.5] Variable Declarations - UserForm Progress
    ' --------------------------------------------------------------------------------------------------------------------
    Dim frmProgress As New frmProgress     ' Instance of the progress UserForm to display macro status.
    Dim totalFilesToProcess As Long        ' Total count of files to generate based on "MasterMapping" selection.
    Dim filesProcessed As Long             ' Counter for files already processed.
    Dim currentTime As Double              ' Elapsed time for updating the UserForm.
    Dim minutes As Long                    ' Minutes component of elapsed time.
    Dim seconds As Long                    ' Seconds component of elapsed time.
    ' --------------------------------------------------------------------------------------------------------------------
    ' [2.1] Macro Initialization - Performance Optimization & Start Time
    ' --------------------------------------------------------------------------------------------------------------------
    StartTime = Timer                       ' Records the start time for performance measurement.
    Application.ScreenUpdating = False      ' Turns off screen updates to improve macro speed.
    Application.Calculation = xlCalculationManual ' Sets calculation to manual to prevent Excel from recalculating frequently.
    Application.EnableEvents = False        ' Disables events to prevent them from firing during macro execution.
    Application.DisplayAlerts = False       ' Suppresses system alerts and message boxes.

    ' --- LOGGING: Start of Macro ---
    Call WriteLog("INFO", "Macro started: genfile_ByProfitCenter_V1_8")

    ' --- [2.2] Prepare Log Sheet ---
    ' Attempt to set a reference to the log sheet.
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0

    ' If the log sheet exists and contains data beyond the header, clear it.
    If Not wsLog Is Nothing Then
        If wsLog.Cells(Rows.Count, "A").End(xlUp).Row > 1 Then
            wsLog.Range("A2:E" & wsLog.Cells(Rows.Count, "A").End(xlUp).Row).ClearContents
            Call WriteLog("INFO", "Cleared existing data in Log sheet.")
        End If
    ' If the log sheet doesn't exist, the WriteLog function will create it on its first call.
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [3.1] Load MasterMapping Data
    ' --------------------------------------------------------------------------------------------------------------------
    
    
            Dim wsActionPage As Worksheet
            Dim sheetName As String

        ' ???????? wsActionPage ??????? "Action Page"
            Set wsActionPage = ThisWorkbook.Sheets("Action_Page")

        ' ????????????? Cell A6 ?????? Action Page
        sheetName = wsActionPage.Range("A6").Value

        ' ????? wsMasterMapping ???????????????????????????? sheetName
        Set wsMasterMapping = ThisWorkbook.Sheets(sheetName)
        
        'Set wsMasterMapping = ThisWorkbook.Sheets("MasterMapping")

    ' Find the last row with data in column A. This now considers profit center data starting from row 5.
    lastRowMasterMapping = wsMasterMapping.Cells(Rows.Count, "A").End(xlUp).Row
    ' Check if there is sufficient data for profit centers starting from row 5.
    If lastRowMasterMapping < 5 Then
        MsgBox "MasterMapping sheet is empty or has only headers up to row 4, or no profit center data found.", vbExclamation, "Data Error"
        Call WriteLog("WARNING", "MasterMapping sheet is empty or has only headers up to row 4, or no profit center data found.", "lastRowMasterMapping", lastRowMasterMapping)
        GoTo CleanUp ' Exits the macro if no sufficient data is found.
    End If
    
    ' Finds the last column with data in row 1 of the "MasterMapping" sheet.
    lastColMasterMapping = wsMasterMapping.Cells(1, Columns.Count).End(xlToLeft).Column
    If lastColMasterMapping < 1 Then lastColMasterMapping = 1 ' Ensure at least one column for safety.
    
    ' Defines the last column for sheet selection. Column AC is the 30th column.
    lastColSheets = 29 ' Hardcoded as per original logic.
    
    ' Determines the effective last column for loading the "MasterMapping" data into an array.
    ' It uses the greater of the actual last column and the hard-coded last column for sheets (AC).
    effectiveLastColForArray = Application.WorksheetFunction.Max(lastColMasterMapping, lastColSheets)
    
    ' Loads the data from the defined range of "MasterMapping" into a variant array for faster processing.
    ' The range now includes all rows from the fifth row (skipping headers and (R) sheet definitions)
    ' and all columns up to the effective last column.
    arrMasterMapping = wsMasterMapping.Range(wsMasterMapping.Cells(5, "A"), wsMasterMapping.Cells(lastRowMasterMapping, effectiveLastColForArray)).Value
    Call WriteLog("INFO", "MasterMapping data for Profit Centers loaded into arrMasterMapping array (starting from row 5).", "Rows Loaded", UBound(arrMasterMapping, 1))

    ' [3.2] View Master Mapping Value (for debugging purposes)
    Debug.Print "---------------------------------------"
    Debug.Print "Content of arrMasterMapping array (Profit Center data only):"
    ' --- LOGGING: MasterMapping Array Content (for detailed debugging) ---
    Call WriteLog("DEBUG", "Starting to log content of arrMasterMapping array (Profit Center data only) to Immediate Window for brevity.")
    
    ' Loops through each row of the MasterMapping array.
    For i = LBound(arrMasterMapping, 1) To UBound(arrMasterMapping, 1)
        Dim rowData As String
        rowData = "Row " & i & ": "
    
        ' Loops through each column to concatenate values into a single string.
        For j = LBound(arrMasterMapping, 2) To UBound(arrMasterMapping, 2)
            rowData = rowData & " | Col " & j & ": " & arrMasterMapping(i, j)
        Next j
    
        ' Prints the concatenated row string to the Immediate Window.
        Debug.Print rowData
        ' Optional: Log each row to the log sheet if absolutely necessary, but can be very verbose.
        ' Call WriteLog("DEBUG", "MasterMapping Row Data", "Row " & i, rowData)
    Next i
    
    Debug.Print "---------------------------------------"
    Call WriteLog("DEBUG", "Finished logging content of arrMasterMapping array.")


    ' --------------------------------------------------------------------------------------------------------------------
    ' [4.1] Identify Sheet Templates and Mappings
    ' --------------------------------------------------------------------------------------------------------------------
    ' Loops through the columns of "MasterMapping" (from D to AC, i.e., column 4 to lastColSheets)
    ' to identify and map master sheet names (from row 2) to their corresponding template sheet names (from row 3).
    For j = 4 To lastColSheets
        ' Retrieves the master sheet name from row 2 of MasterMapping.
        masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, j).Value))
        ' Retrieves the template sheet name from row 3 of MasterMapping.
        templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, j).Value))

        ' Proceed only if both master and template sheet names are provided.
        If Len(masterSourceSheetName) > 0 And Len(templateSheetName) > 0 Then
            ' Derives the final sheet name for the new workbook by removing " (M)" from the master sheet name.
            ' This is used as the key for dictSheetTemplates.
            createdSheetName = Replace(masterSourceSheetName, " (M)", "")
            ' Provides a fallback mechanism to derive the name from the template name
            ' if the master name doesn't contain " (M)" but the masterSourceSheetName still has " (M)".
            If createdSheetName = masterSourceSheetName And InStr(1, masterSourceSheetName, " (M)", vbTextCompare) > 0 Then
                createdSheetName = Replace(templateSheetName, " (T)", "")
            End If

            ' If a valid createdSheetName is determined, add it to the dictionary.
            If Len(createdSheetName) > 0 Then
                ' Adds the mapping to a dictionary if it doesn't already exist.
                ' The dictionary stores the final sheet name (without suffix) as the key
                ' and the template sheet name (with (T) suffix) as the value.
                If Not dictSheetTemplates.Exists(createdSheetName) Then
                    dictSheetTemplates.Add createdSheetName, templateSheetName
                    Call WriteLog("INFO", "Added sheet mapping.", "Created Sheet -> Template Sheet", createdSheetName & " -> " & templateSheetName)
                End If
            End If
        End If
    Next j

    ' Checks if any valid sheet mappings were found.
    If dictSheetTemplates.Count = 0 Then
        MsgBox "No valid sheet mappings found in MasterMapping (Row 2 & 3, Columns D-AC). Please ensure Row 2 contains master sheet names like 'Sheet (M)' and Row 3 contains template sheet names like 'Sheet (T)'.", vbExclamation, "Configuration Error"
        Call WriteLog("WARNING", "No valid sheet mappings found.", "dictSheetTemplates.Count", dictSheetTemplates.Count)
        GoTo CleanUp ' Exits the macro if no mappings are found.
    Else
        Call WriteLog("INFO", "Identified sheet templates and mappings (from MasterMapping rows 2 & 3).", "Total Mappings", dictSheetTemplates.Count)
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [5.1] Pre-populate Dictionary with All Master (M) Data
    ' --------------------------------------------------------------------------------------------------------------------
    ' Iterates through all sheets in the current workbook to find Master (M) sheets.
    For Each wsMasterSource In ThisWorkbook.Sheets
        ' Checks if the sheet name contains " (M)", indicating it's a master data sheet.
        If InStr(1, wsMasterSource.Name, " (M)", vbTextCompare) > 0 Then
            ' Finds the last row with data in column B (assuming Profit Center ID is in B) to define the data range.
            lastRowMasterSource = wsMasterSource.Cells(Rows.Count, "B").End(xlUp).Row
            If lastRowMasterSource >= 2 Then ' Ensures there is data beyond the header row.
                ' Finds the last column with data in the master source sheet.
                lastColMasterSource = wsMasterSource.Cells(1, wsMasterSource.Columns.Count).End(xlToLeft).Column
                If lastColMasterSource < 1 Then lastColMasterSource = 1 ' Ensure at least one column for safety.

                ' Loads the entire data range of the master sheet into a temporary array for efficiency.
                arrMasterData = wsMasterSource.Range(wsMasterSource.Cells(2, "A"), wsMasterSource.Cells(lastRowMasterSource, lastColMasterSource)).Value
                Call WriteLog("INFO", "Loaded Master (M) sheet data.", "Sheet Name", wsMasterSource.Name & " (Rows: " & UBound(arrMasterData, 1) & ", Cols: " & UBound(arrMasterData, 2) & ")")

                ' Creates a temporary dictionary to group data by profit center ID for the current master sheet.
                Set dictSingleMasterGrouped = CreateObject("Scripting.Dictionary")

                ' Loops through each row of the master data array.
                For k = LBound(arrMasterData, 1) To UBound(arrMasterData, 1)
                    ' Retrieves the Profit Center ID from the second column (index 2).
                    masterKey = Trim(CStr(arrMasterData(k, 2)))
                    ' Creates a 1D array from the current row of the 2D master data array.
                    ReDim masterRowData(1 To UBound(arrMasterData, 2))
                    For col = LBound(arrMasterData, 2) To UBound(arrMasterData, 2)
                        masterRowData(col) = arrMasterData(k, col)
                    Next col

                    ' Adds the current row of data to a collection associated with its Profit Center ID.
                    If Not dictSingleMasterGrouped.Exists(masterKey) Then
                        Set masterRowsForProfitCenter = New Collection
                        dictSingleMasterGrouped.Add masterKey, masterRowsForProfitCenter
                    Else
                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(masterKey)
                    End If
                    masterRowsForProfitCenter.Add masterRowData
                Next k
                ' Adds the temporary dictionary (containing grouped data by PC ID) to the main dictionary.
                ' The main dictionary's key is the master sheet name (e.g., "Admin (M)").
                dictAllMasterData.Add wsMasterSource.Name, dictSingleMasterGrouped
                Call WriteLog("INFO", "Grouped master data by Profit Center ID.", "Master Sheet", wsMasterSource.Name & " (Unique PC IDs: " & dictSingleMasterGrouped.Count & ")")
            End If
        End If
    Next wsMasterSource

    ' Clears the temporary master data array to free up memory after processing.
    If Not IsEmpty(arrMasterData) Then Erase arrMasterData

    ' Checks if any master data was successfully loaded across all sheets.
    If dictAllMasterData.Count = 0 Then
        MsgBox "No Master sheets with names containing ' (M)' and data found in this workbook.", vbExclamation, "Data Error"
        Call WriteLog("ERROR", "No Master (M) sheets found or loaded with data.", "dictAllMasterData.Count", dictAllMasterData.Count)
        GoTo CleanUp ' Exits the macro if no master data is found.
    Else
        Call WriteLog("INFO", "Successfully pre-populated dictionary with all Master (M) data.", "Total Master Sheets", dictAllMasterData.Count)
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [6.1] Calculate Total Files to Process
    ' --------------------------------------------------------------------------------------------------------------------
    totalFilesToProcess = 0
    ' Counts the number of rows in the "MasterMapping" array where column C (index 3) is marked as TRUE.
    ' This indicates which profit centers are selected for file generation.
    For i = LBound(arrMasterMapping, 1) To UBound(arrMasterMapping, 1)
        ' When arrMasterMapping starts from row 5, the first element (i=1) corresponds to the first profit center.
        ' We assume column 3 (C) within this array still holds the TRUE/FALSE for selection.
        If UBound(arrMasterMapping, 2) >= 3 Then ' Ensure column C exists in the array.
            If arrMasterMapping(i, 3) = True Then ' Check if the "Generate" flag (Column C) is set to TRUE.
                totalFilesToProcess = totalFilesToProcess + 1
            End If
        Else
            ' This warning is more critical now as arrMasterMapping starts from profit center data.
            Call WriteLog("ERROR", "MasterMapping column C (selection column) not found within the profit center data range. Check MasterMapping sheet structure.", "MasterMapping Array Row Index", i)
        End If
    Next i
    Call WriteLog("INFO", "Calculated total files to process.", "Total Files to Generate", totalFilesToProcess)

    ' Checks if any files were selected for processing.
    If totalFilesToProcess = 0 Then
        MsgBox "No files selected for processing in MasterMapping Column C. Please mark 'TRUE' in Column C for files you wish to generate.", vbExclamation, "Selection Error"
        Call WriteLog("WARNING", "No files selected for processing in MasterMapping Column C.")
        GoTo CleanUp ' Exits the macro if no files are selected.
    End If

    ' --------------------------------------------------------------------------------------------------------------------
    ' [7.1] Initialize and Display Progress UserForm
    ' --------------------------------------------------------------------------------------------------------------------
    With frmProgress
        .lblProgress.Caption = "Processing: 0 / " & totalFilesToProcess
        .lblTime.Caption = "Time Elapsed: 00:00"
        .Show vbModeless ' Displays the UserForm non-modally so the macro can continue execution.
    End With
    filesProcessed = 0 ' Initialize the counter for files processed.
    Call WriteLog("INFO", "Progress UserForm initialized and displayed.")

    ' --------------------------------------------------------------------------------------------------------------------
    ' [8.1] Main Loop - Generate Files for Each Selected Profit Center
    ' --------------------------------------------------------------------------------------------------------------------
    ' Iterates through each row in the "MasterMapping" data array. Each row represents a profit center.
    For i = LBound(arrMasterMapping, 1) To UBound(arrMasterMapping, 1)

        If arrMasterMapping(i, 3) = True Then ' Checks if the current profit center (row) is selected for processing (Column C is TRUE).
            filesProcessed = filesProcessed + 1
            Call WriteLog("INFO", "Starting file generation for a new Profit Center.", "Processed Count", filesProcessed & " / " & totalFilesToProcess)

            ' Updates the progress and elapsed time on the UserForm.
            frmProgress.lblProgress.Caption = "Processing: " & filesProcessed & " / " & totalFilesToProcess
            currentTime = Timer - StartTime
            minutes = Int(currentTime / 60)
            seconds = Int(currentTime Mod 60)
            frmProgress.lblTime.Caption = "Time Elapsed: " & Format(minutes, "00") & ":" & Format(seconds, "00")
            DoEvents ' Allows the UI to refresh and the UserForm to update, preventing unresponsiveness.

            Set newWorkbook = Application.Workbooks.Add(xlWBATWorksheet) ' Creates a new blank workbook for the current profit center.
            ' Retrieves Profit Center ID and Name from the "MasterMapping" array (Columns A and B).
            profitCenterID = Trim(CStr(arrMasterMapping(i, 1)))
            profitCenterName = Trim(CStr(arrMasterMapping(i, 2)))
            Call WriteLog("INFO", "New workbook created for Profit Center.", "Profit Center ID/Name", profitCenterID & " - " & profitCenterName)

            ' --------------------------------------------------------------------------------------------------------------------
            ' [8.1.1] Logic to Copy Sheets ending with (R)
            ' This section handles copying entire sheets marked with "(R)" from the MasterMapping (row 4)
            ' to the new workbook, converting all cells to values while preserving their formatting.
            ' --------------------------------------------------------------------------------------------------------------------
            For selectedSheetCol = 4 To lastColSheets ' Loops through columns D to AC for sheet definitions.
                Dim reportSheetName As String
                ' Get the sheet name from row 4 of the MasterMapping for the current column.
                reportSheetName = Trim(CStr(wsMasterMapping.Cells(4, selectedSheetCol).Value))

                ' Check if a valid sheet name ending with " (R)" is defined and if that sheet exists in ThisWorkbook.
                If Len(reportSheetName) > 0 And InStr(1, reportSheetName, " (R)", vbTextCompare) > 0 Then
                    Dim wsOriginalReport As Worksheet
                    Set wsOriginalReport = Nothing ' Initialize the worksheet object.
                    On Error Resume Next           ' Temporarily ignore errors if the report sheet is not found.
                    Set wsOriginalReport = ThisWorkbook.Sheets(reportSheetName)
                    On Error GoTo 0                ' Re-enable error handling.

                    If Not wsOriginalReport Is Nothing Then
                        ' Copy the entire (R) sheet to the newly created workbook.
                        wsOriginalReport.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                        Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count) ' Reference the newly copied sheet.

                        ' Rename the copied sheet to its original name including "(R)".
                        wsNewWorkbookSheet.Name = reportSheetName

                        ' Convert all cells in the new sheet to values while preserving formatting.
                        If Not wsNewWorkbookSheet.UsedRange Is Nothing Then
                            On Error Resume Next
                            ' Copy all contents (including formatting) of the used range.
                            wsNewWorkbookSheet.UsedRange.Copy

                            ' Paste special: Values to convert formulas to their calculated results.
                            wsNewWorkbookSheet.UsedRange.PasteSpecial xlPasteValues
                            
                            ' Paste special: Formats to re-apply any original formatting that might have been affected.
                            wsNewWorkbookSheet.UsedRange.PasteSpecial xlPasteFormats

                            Application.CutCopyMode = False ' Clear the clipboard after paste operations.
                            If Err.Number <> 0 Then
                                Call WriteLog("ERROR", "Failed to convert (R) sheet cells to values while preserving format.", "Error Description", Err.Description)
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                        Call WriteLog("INFO", "Copied (R) sheet and converted all cells to values (preserved format).", "Sheet Name", reportSheetName)
                        Set wsNewWorkbookSheet = Nothing ' Dereference the worksheet object to free memory.
                    Else
                        Call WriteLog("WARNING", "Report (R) sheet not found in the current workbook, skipping copy operation.", "Sheet Name", reportSheetName)
                    End If
                End If
            Next selectedSheetCol
            Call WriteLog("INFO", "Finished copying all selected (R) sheets for current Profit Center.")

            ' --- End of Logic to Copy Sheets ending with (R) ---


            ' --------------------------------------------------------------------------------------------------------------------
            ' [8.2] Sub-Loop - Copy and Populate Sheets in New Workbook
            ' This section handles copying Template (T) sheets (either with or without a paired Master (M) sheet)
            ' and populating them with data.
            ' --------------------------------------------------------------------------------------------------------------------
            For selectedSheetCol = 4 To lastColSheets
                If UBound(arrMasterMapping, 2) >= selectedSheetCol Then ' Ensure the column exists in the MasterMapping array.
                    If arrMasterMapping(i, selectedSheetCol) = True Then ' Checks if the specific sheet is selected for the current profit center.

                        Dim currentMasterSourceSheetName As String
                        Dim currentTemplateSheetName As String
                        ' Retrieves Master (M) and Template (T) sheet names from MasterMapping header rows (row 2 and 3).
                        currentMasterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, selectedSheetCol).Value)) ' From row 2 (Master)
                        currentTemplateSheetName = Trim(CStr(wsMasterMapping.Cells(3, selectedSheetCol).Value))   ' From row 3 (Template)

                        ' --- NEW LOGIC START: [8.2.1] Handle Template (T) sheets without Master (M) data (V1.4) ---
                        ' This block checks if there's a Template sheet defined (row 3) but NO Master sheet (row 2 is empty).
                        If Len(currentTemplateSheetName) > 0 And Len(currentMasterSourceSheetName) = 0 Then
                            Call WriteLog("INFO", "Processing a (T) sheet without a corresponding (M) sheet.", "Template Name", currentTemplateSheetName)

                            Dim wsOriginalTSheet As Worksheet
                            Set wsOriginalTSheet = Nothing
                            On Error Resume Next ' Temporarily ignore errors if the template sheet is not found.
                            Set wsOriginalTSheet = ThisWorkbook.Sheets(currentTemplateSheetName)
                            On Error GoTo 0      ' Re-enable error handling.

                            If Not wsOriginalTSheet Is Nothing Then
                                ' Copy the (T) sheet to the new workbook.
                                wsOriginalTSheet.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                wsNewWorkbookSheet.Name = currentTemplateSheetName ' Keep (T) suffix for sheet name in new workbook.

                                Dim lastRowUsed As Long, lastColUsed As Long
                                ' Find the last used row and column in the copied sheet.
                                ' Use LookIn:=xlValues to ensure cells with actual values are considered for 'UsedRange'.
                                On Error Resume Next ' Handle case of completely empty sheet (Find method would error)
                                With wsNewWorkbookSheet
                                    ' Find last row with any content (value or formula)
                                    If Not .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
                                        lastRowUsed = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
                                    Else
                                        lastRowUsed = 1 ' Assume at least header if no data, or if sheet is truly empty.
                                    End If
                                    ' Find last column with any content (value or formula)
                                    If Not .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues) Is Nothing Then
                                        lastColUsed = .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
                                    Else
                                        lastColUsed = 1 ' Assume at least one column.
                                    End If
                                End With
                                On Error GoTo 0
                                
                                
                                ' [8.2.1.1] Convert to values first, then re-apply formulas to remove external links.
                                ' This two-step process ensures all external references are broken before re-embedding local formulas.
                                If lastRowUsed >= 1 And lastColUsed >= 1 Then
                                    Dim targetRange As Range
                                    Dim sourceRange As Range

                                    ' Define the target range in the new workbook and source range in the original workbook.
                                    Set targetRange = wsNewWorkbookSheet.Range(wsNewWorkbookSheet.Cells(1, 1), wsNewWorkbookSheet.Cells(lastRowUsed, lastColUsed))
                                    Set sourceRange = wsOriginalTSheet.Range(wsOriginalTSheet.Cells(1, 1), wsOriginalTSheet.Cells(lastRowUsed, lastColUsed))

                                    ' Step 1: Convert all cells in the new sheet's used range to values to break all external links.
                                    targetRange.Value = targetRange.Value
                                    Call WriteLog("INFO", "Converted (T) sheet to values to break external links.", "Sheet Name", currentTemplateSheetName)

                                    ' Step 2: Copy formulas from the original template sheet and apply them to the new sheet.
                                    ' This re-embeds the formulas, making them local to the new workbook.
                                    targetRange.FormulaR1C1 = sourceRange.FormulaR1C1
                                    Call WriteLog("INFO", "Re-applied formulas from original (T) sheet.", "Sheet Name", currentTemplateSheetName)

                                    ' Clean up range objects
                                    Set targetRange = Nothing
                                    Set sourceRange = Nothing
                                End If

                                ' [8.2.1.2] Populate Profit Center ID and Name into columns B and C
                                ' Data should start from row 2.
                                If lastRowUsed >= 2 Then ' Only populate if there are rows to fill beyond header.
                                    wsNewWorkbookSheet.Range("B2:B" & lastRowUsed).Value = profitCenterID
                                    wsNewWorkbookSheet.Range("C2:C" & lastRowUsed).Value = profitCenterName
                                    Call WriteLog("INFO", "Populated Profit Center ID/Name in (T) sheet (no M).", "Sheet Name", currentTemplateSheetName & " (Rows 2-" & lastRowUsed & ")")

                                    ' --- NEW LOGIC START (V1.7) ---
                                    ' Populate Column A with the value from MasterMapping Column AF (index 32).
                                    ' This assumes Column AF is where the desired value is located in MasterMapping for the current Profit Center.
                                    If UBound(arrMasterMapping, 2) >= 32 Then ' Ensure Column AF (index 32) exists in MasterMapping array
                                        wsNewWorkbookSheet.Range("A2:A" & lastRowUsed).Value = Trim(CStr(arrMasterMapping(i, 32)))
                                        
                                        ' --- MODIFICATION START (V1.8) ---
                                        ' Consolidated variable details for WriteLog to match argument count.
                                        Dim logVarName As String
                                        Dim logVarValue As String
                                        logVarName = "Sheet Name / Value"
                                        logVarValue = currentTemplateSheetName & " (Rows 2-" & lastRowUsed & ") / " & Trim(CStr(arrMasterMapping(i, 32)))
                                        Call WriteLog("INFO", "Populated Column A in (T) sheet with MasterMapping AF value.", logVarName, logVarValue)
                                        ' --- MODIFICATION END (V1.8) ---
                                    Else
                                        Call WriteLog("WARNING", "MasterMapping Column AF (index 32) not found, skipping population of Column A in (T) sheet.", "Sheet Name", currentTemplateSheetName)
                                    End If
                                    ' --- NEW LOGIC END (V1.7) ---
                                End If



                                Call WriteLog("INFO", "Successfully copied and processed (T) sheet (without M).", "Sheet Name", currentTemplateSheetName)
                                Set wsNewWorkbookSheet = Nothing ' Dereference after use.
                                Set wsOriginalTSheet = Nothing
                            Else
                                Call WriteLog("WARNING", "Template (T) sheet not found in current workbook for (T) without (M) configuration, skipping.", "Template Name", currentTemplateSheetName)
                            End If
                        ' --- NEW LOGIC END (V1.4) ---

                        ' --- EXISTING LOGIC START: [8.2.2] Handle paired (M) and (T) sheets (Original [8.2] content) ---
                        ' This part will only run if both Master (M) and Template (T) names are present.
                        ElseIf Len(currentMasterSourceSheetName) > 0 And Len(currentTemplateSheetName) > 0 Then
                            ' The following block is the original content of the [8.2] loop for M/T paired sheets.
                            ' It's now nested within an ElseIf to ensure mutual exclusivity with the (T)-only logic.
                            
                            masterSourceSheetName = Trim(CStr(wsMasterMapping.Cells(2, selectedSheetCol).Value))
                            templateSheetName = Trim(CStr(wsMasterMapping.Cells(3, selectedSheetCol).Value))
                            createdSheetName = Replace(masterSourceSheetName, " (M)", "")
                            Call WriteLog("INFO", "Processing sheet for new workbook.", "Sheet Info", masterSourceSheetName & " (Template: " & templateSheetName & ")")

                            If dictSheetTemplates.Exists(createdSheetName) Then
                                Set wsOriginalTemplate = Nothing
                                On Error Resume Next
                                Set wsOriginalTemplate = ThisWorkbook.Sheets(templateSheetName)
                                On Error GoTo 0

                                If wsOriginalTemplate Is Nothing Then
                                    Call WriteLog("WARNING", "Template sheet not found, skipping sheet for current profit center.", "Template Name", templateSheetName)
                                    ' No GoTo needed here; the ElseIf structure handles flow.
                                ElseIf dictAllMasterData.Exists(masterSourceSheetName) Then
                                    Set dictSingleMasterGrouped = dictAllMasterData.item(masterSourceSheetName)

                                    If dictSingleMasterGrouped.Exists(profitCenterID) Then
                                        Set masterRowsForProfitCenter = dictSingleMasterGrouped.item(profitCenterID)

                                        If masterRowsForProfitCenter.Count > 0 Then
                                            If IsArray(masterRowsForProfitCenter.item(1)) Then
                                                numMasterDataCols = UBound(masterRowsForProfitCenter.item(1))
                                            Else
                                                Call WriteLog("WARNING", "Master data for sheet is not in expected array format, skipping sheet.", "Master Sheet", masterSourceSheetName)
                                                ' No GoTo needed.
                                            End If
                                            numRowsData = masterRowsForProfitCenter.Count
                                            Call WriteLog("INFO", "Master data found for sheet and profit center.", "Rows/Cols", numRowsData & " rows, " & numMasterDataCols & " columns")

                                            arrTemplateFormulas = wsOriginalTemplate.Range("A2").Resize(numRowsData, numMasterDataCols).FormulaR1C1
                                            Call WriteLog("INFO", "Loaded template formulas from original sheet.")

                                            On Error Resume Next
                                            ThisWorkbook.Sheets(templateSheetName).Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                            Set wsNewWorkbookSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
                                            wsNewWorkbookSheet.Name = masterSourceSheetName
                                            If Err.Number <> 0 Then
                                                Call WriteLog("ERROR", "Failed to copy or rename sheet in new workbook.", "Error Description", Err.Description)
                                                Err.Clear
                                            End If
                                            On Error GoTo 0

                                            If wsNewWorkbookSheet Is Nothing Then
                                                Call WriteLog("ERROR", "Sheet copy failed, skipping sheet for current profit center.", "Template Name", templateSheetName)
                                                ' No GoTo needed.
                                            Else
                                                If Not wsNewWorkbookSheet.UsedRange Is Nothing Then
                                                    On Error Resume Next
                                                    wsNewWorkbookSheet.UsedRange.Value = wsNewWorkbookSheet.UsedRange.Value
                                                    If Err.Number <> 0 Then
                                                        Call WriteLog("ERROR", "Failed to convert formulas to values in new sheet.", "Error Description", Err.Description)
                                                        Err.Clear
                                                    End If
                                                    On Error GoTo 0
                                                End If
                                                Call WriteLog("INFO", "Template sheet copied and formulas converted to values.", "New Sheet Name", masterSourceSheetName)

                                                ReDim arrCurrentSheetFinalData(1 To numRowsData, 1 To numMasterDataCols)
                                                dataRowCounter = 0

                                                For Each masterRowData In masterRowsForProfitCenter
                                                    dataRowCounter = dataRowCounter + 1
                                                    For k = LBound(masterRowData) To UBound(masterRowData)
                                                        templateCellContent = CStr(arrTemplateFormulas(dataRowCounter, k))

                                                        If Left(templateCellContent, 1) = "=" And Len(templateCellContent) > 1 Then
                                                            arrCurrentSheetFinalData(dataRowCounter, k) = templateCellContent
                                                        Else
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

                                                wsNewWorkbookSheet.Range("A2").Resize(numRowsData, numMasterDataCols).FormulaR1C1 = arrCurrentSheetFinalData
                                                Call WriteLog("INFO", "Data written to new sheet.", "Sheet Name", masterSourceSheetName)

                                                Erase arrCurrentSheetFinalData
                                            End If
                                        Else
                                            Call WriteLog("WARNING", "No master data rows found for the specified Profit Center ID.", "Profit Center ID", profitCenterID)
                                        End If
                                    Else
                                        Call WriteLog("WARNING", "Profit Center ID not found in grouped master data for sheet.", "Profit Center ID", profitCenterID & " - Master Sheet: " & masterSourceSheetName)
                                    End If
                                Else
                                    Call WriteLog("WARNING", "Master data dictionary for sheet is empty or not found.", "Master Sheet", masterSourceSheetName)
                                End If
                            Else
                                Call WriteLog("ERROR", "Template mapping for created sheet name not found, unexpected condition.", "Created Sheet Name", createdSheetName)
                            End If

                        ' --- EXISTING LOGIC END: [8.2.2] Handle paired (M) and (T) sheets ---

                        Else
                            ' [8.2.3] Log for unhandled selected columns (e.g., only (M) name present, or invalid combination)
                            Call WriteLog("DEBUG", "MasterMapping column selected but no valid (T)-only or (M)/(T) pair configuration found.", _
                                          "Column Index", selectedSheetCol & " - Master (M): " & currentMasterSourceSheetName & ", Template (T): " & currentTemplateSheetName)
                        End If
                    End If
                End If
NextSheetLoop: ' This label is now mostly for historical context or if you reintroduce GoTo in specific sub-blocks.
            Next selectedSheetCol
            Call WriteLog("INFO", "Finished processing all selected sheets for current Profit Center.")

            
            ' Deletes the default sheets (e.g., "Sheet1", "Sheet2") from the new workbook.
            On Error Resume Next ' Temporarily ignore errors during sheet deletion.
            For Each wsTemp In newWorkbook.Sheets
                If Left(wsTemp.Name, 5) = "Sheet" And IsNumeric(Mid(wsTemp.Name, 6)) Then
                    wsTemp.Delete
                    Call WriteLog("INFO", "Deleted default sheet from new workbook.", "Sheet Name", wsTemp.Name)
                End If
            Next wsTemp
            On Error GoTo 0 ' Re-enable error handling.
            
                    ' --- NEW LOGIC START: Clear data below the last used row in Column B for all sheets ---
                    ' [8.3] Final Cleanup Before Saving
                    ' This loop iterates through every sheet in the newly created workbook to clear any
                    ' unnecessary data, formulas, and formatting below the last row containing data in Column B.
                    ' --------------------------------------------------------------------------------------------------------------------
                    Dim ws As Worksheet
                    Dim lastRowInB As Long
                    Call WriteLog("INFO", "Starting final cleanup for all sheets in new workbook.", "Profit Center ID", profitCenterID)

                    For Each ws In newWorkbook.Sheets
                        ' Find the last row with data specifically in Column B for the current sheet.
                        lastRowInB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

                        ' Check if the last data row is not the very last row of the worksheet.
                        If lastRowInB < ws.Rows.Count Then
                            ' Define the range from the row after the last data row to the end of the sheet and clear everything.
                            ws.Rows(lastRowInB + 1 & ":" & ws.Rows.Count).Clear
                            Call WriteLog("INFO", "Cleared excess rows below the last data row in Column B.", "Sheet Name / Last Row", ws.Name & " / " & lastRowInB)
                        Else
                            Call WriteLog("DEBUG", "No rows to clear; data extends to the end of the sheet or sheet is empty.", "Sheet Name", ws.Name)
                        End If
                    Next ws
                    Call WriteLog("INFO", "Finished final cleanup for all sheets.")
                    ' --- NEW LOGIC END ---
                      '--------------------------------------------------------------------------------------------------------------------

            '--------------------------------------------------------------------------------------------------------------------
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
                        ' --------------------------------------------------------------------------------------------------------------------
            ' [9.1] File Saving Operations
            ' --------------------------------------------------------------------------------------------------------------------
            ' Retrieves the base file path from column AD (index 30) of the MasterMapping array.
            If UBound(arrMasterMapping, 2) >= 30 Then
                newFilePath = CStr(arrMasterMapping(i, 30))
            Else
                newFilePath = ""
            End If
            Call WriteLog("INFO", "Retrieved base file path from MasterMapping.", "newFilePath", newFilePath)

            If Len(newFilePath) > 0 Then ' Checks if a file path was provided.
                ' Constructs the base filename using Data from column AE
                baseFileName = CStr(arrMasterMapping(i, 31))

                ' Replaces invalid characters in the filename to ensure a valid file path.
                baseFileName = Replace(baseFileName, "/", "_")
                baseFileName = Replace(baseFileName, "\", "_")
                baseFileName = Replace(baseFileName, ":", "_")
                baseFileName = Replace(baseFileName, "*", "_")
                baseFileName = Replace(baseFileName, "?", "_")
                baseFileName = Replace(baseFileName, Chr(34), "_") ' Double quote character
                baseFileName = Replace(baseFileName, "<", "_")
                baseFileName = Replace(baseFileName, ">", "_")
                baseFileName = Replace(baseFileName, "|", "_")
                Call WriteLog("INFO", "Constructed base filename and sanitized invalid characters.", "baseFileName", baseFileName)

                ' Constructs the full path for the output folder for the current profit center.
                folderPath = newFilePath & Application.PathSeparator & baseFileName

                ' Creates the folder if it doesn't already exist.
                If Not fso.FolderExists(folderPath) Then
                    On Error Resume Next ' Temporarily ignore errors during folder creation.
                    fso.CreateFolder folderPath
                    If Err.Number <> 0 Then
                        Call WriteLog("ERROR", "Failed to create folder.", "Error Description", Err.Description & " - Folder Path: " & folderPath)
                        Err.Clear
                        GoTo SkipFileSave ' Skip saving the file if folder creation failed.
                    Else
                        Call WriteLog("INFO", "Created new output folder.", "folderPath", folderPath)
                    End If
                    On Error GoTo 0 ' Re-enable error handling.
                Else
                    Call WriteLog("INFO", "Output folder already exists.", "folderPath", folderPath)
                End If

                fileName = baseFileName & ".xlsx"
                copyNum = 1

                ' Checks for existing files with the same name and adds a number to make the filename unique.
                Do While fso.fileExists(folderPath & Application.PathSeparator & fileName)
                    copyNum = copyNum + 1
                    fileName = baseFileName & " (" & copyNum & ").xlsx"
                    Call WriteLog("INFO", "Duplicate filename detected, generating unique name.", "New Filename", fileName)
                Loop

                ' Saves the new workbook with the unique filename.
                On Error Resume Next ' Temporarily ignore errors during file saving.
                newWorkbook.SaveAs fileName:=folderPath & Application.PathSeparator & fileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                If Err.Number <> 0 Then
                    Call WriteLog("ERROR", "Failed to save workbook.", "Error Description", Err.Description & " - Filename: " & fileName)
                    Err.Clear
                Else
                    Call WriteLog("INFO", "Workbook saved successfully.", "Full Path", folderPath & Application.PathSeparator & fileName)
                End If
                On Error GoTo 0 ' Re-enable error handling.

SkipFileSave: ' Label to skip the saving process if an error occurred during folder creation or other issues.
                newWorkbook.Close SaveChanges:=False ' Closes the workbook without saving if an error occurred or path was invalid.
                Set newWorkbook = Nothing ' Dereferences the workbook object to free memory.
                Call WriteLog("INFO", "Workbook closed.", "Profit Center ID", profitCenterID)
            Else
                newWorkbook.Close SaveChanges:=False ' Closes the workbook without saving if no file path is provided.
                Set newWorkbook = Nothing
                Call WriteLog("WARNING", "No file path provided, workbook closed without saving.", "Profit Center ID", profitCenterID)
            End If
        End If
    Next i
    Call WriteLog("INFO", "Main loop completed. All selected files processed.")

    ' --------------------------------------------------------------------------------------------------------------------
    ' [10.1] Cleanup - Restore Settings and Free Memory
    ' --------------------------------------------------------------------------------------------------------------------
CleanUp: ' Label for cleanup operations, ensuring settings are restored even if an error occurs.
    Application.ScreenUpdating = True        ' Re-enables screen updates.
    Application.Calculation = xlCalculationAutomatic ' Restores automatic calculation mode.
    Application.EnableEvents = True          ' Re-enables events.
    Application.DisplayAlerts = True         ' Re-enables alerts.
    Call WriteLog("INFO", "Application settings restored (ScreenUpdating, Calculation, Events, Alerts).")

    ' Clears all object variables to release memory effectively, EXCEPT frmProgress for now.
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
    ' --- MODIFICATION START (V1.3 - Removed premature Set frmProgress = Nothing) ---
    ' Removed: Set frmProgress = Nothing
    ' --- MODIFICATION END ---
    Call WriteLog("INFO", "All object variables cleared (excluding UserForm for final update).")

    ' Clears all array variables to release memory.
    If Not IsEmpty(arrMasterMapping) Then Erase arrMasterMapping
    If Not IsEmpty(arrMasterData) Then Erase arrMasterData
    If Not IsEmpty(arrCurrentSheetFinalData) Then Erase arrCurrentSheetFinalData
    If Not IsEmpty(arrTemplateFormulas) Then Erase arrTemplateFormulas
    Call WriteLog("INFO", "All array variables cleared.")

    ' --------------------------------------------------------------------------------------------------------------------
    ' [11.1] Final Timing and UserForm Update
    ' --------------------------------------------------------------------------------------------------------------------
    endTime = Timer                          ' Records the end time.
    runTime = endTime - StartTime            ' Calculates the total execution time.
    
    ' Updates and closes the progress UserForm if it is still active.
    If Not frmProgress Is Nothing Then
        With frmProgress
            .lblProgress.Caption = "Processing Complete!"
            minutes = Int(runTime / 60)
            seconds = Int(runTime Mod 60)
            .lblTime.Caption = "Total Time: " & Format(minutes, "00") & ":" & Format(seconds, "00")
            If .Visible Then Application.Wait Now + TimeValue("00:00:01") ' Keeps the form visible for a brief moment.
            Unload frmProgress ' Unloads the UserForm from memory.
        End With
        ' --- MODIFICATION START (V1.3 - Moved Set frmProgress = Nothing here) ---
        ' Dereferences the UserForm object AFTER unloading it.
        Set frmProgress = Nothing
        ' --- MODIFICATION END ---
        Call WriteLog("INFO", "Progress UserForm closed.")
    End If

    minutes = Int(runTime / 60)
    seconds = Int(runTime Mod 60)
    MsgBox "Completed Run Time: " & Format(minutes, "00") & " minutes and " & Format(seconds, "00") & " seconds", vbInformation, "Macro Complete"
    Call WriteLog("INFO", "Macro finished.", "Total Run Time", Format(minutes, "00") & ":" & Format(seconds, "00"))

End Sub


' --------------------------------------------------------------------------------------------------------------------
' [12.1] WriteLog Function (Helper Function for Logging)
' --------------------------------------------------------------------------------------------------------------------
Sub WriteLog(logType As String, message As String, Optional varName As String = "", Optional varValue As Variant)
    ' Purpose: Writes a log entry to the specified log sheet.
    ' Arguments:
    '   logType: Type of log entry (e.g., "INFO", "WARNING", "ERROR", "DEBUG").
    '   message: The main message for the log entry.
    '   varName: (Optional) The name of a variable to log.
    '   varValue: (Optional) The value of the variable to log.

    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim logDetail As String
    Dim currentWorkbook As Workbook ' Declare current workbook to avoid issues with activeworkbook.

    Set currentWorkbook = ThisWorkbook ' Ensure we are working with the workbook containing the macro.

    On Error GoTo ErrorHandler ' Enable error handling for this function.

    ' [12.2] Check if the log sheet exists, create it if not.
    On Error Resume Next ' Temporarily disable error handling for sheet access.
    Set wsLog = currentWorkbook.Sheets(LOG_SHEET_NAME)
    On Error GoTo 0      ' Re-enable error handling.

    If wsLog Is Nothing Then
        ' Create the log sheet if it does not exist.
        Set wsLog = currentWorkbook.Sheets.Add(After:=currentWorkbook.Sheets(currentWorkbook.Sheets.Count))
        wsLog.Name = LOG_SHEET_NAME
        ' Add headers to the newly created log sheet.
        With wsLog
            .Cells(1, 1).Value = "Timestamp"
            .Cells(1, 2).Value = "Type"
            .Cells(1, 3).Value = "Message"
            .Cells(1, 4).Value = "Variable Name"
            .Cells(1, 5).Value = "Variable Value"
            .Rows(1).Font.Bold = True ' Make header row bold.
            .Columns("A:E").AutoFit   ' Auto-fit columns for readability.
        End With
    End If

    ' [12.3] Find the next available row in the log sheet.
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    ' [12.4] Construct the variable detail string if varName is provided.
    If varName <> "" Then
        ' Attempt to convert varValue to string; handle objects.
        If IsObject(varValue) And Not IsEmpty(varValue) Then
            On Error Resume Next ' Attempt to get a common property like .Name or .Value.
            logDetail = varValue.Name
            If Err.Number <> 0 Then
                logDetail = "Object" ' Default to "Object" if Name property isn't available.
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf IsError(varValue) Then
            logDetail = "Error: " & CStr(varValue)
        Else
            logDetail = CStr(varValue)
        End If
    Else
        logDetail = "" ' No variable value to log.
    End If

    ' [12.5] Write the log entry to the log sheet.
    With wsLog
        .Cells(nextRow, 1).Value = Now      ' Current date and time.
        .Cells(nextRow, 2).Value = logType  ' Type of log (INFO, WARNING, ERROR', 'DEBUG').
        .Cells(nextRow, 3).Value = message  ' Main log message.
        .Cells(nextRow, 4).Value = varName  ' Optional variable name.
        .Cells(nextRow, 5).Value = logDetail ' Optional variable value/detail.
        .Columns("A:E").AutoFit ' Adjust column width after writing to keep readable.
    End With

    Exit Sub ' Exit the function successfully.

ErrorHandler:
    ' [12.6] Fallback for logging errors, prints to Immediate Window if WriteLog itself fails.
    Debug.Print "Error in WriteLog function: " & Err.Description & " (Log Type: " & logType & ", Message: " & message & ")"
    On Error GoTo 0 ' Ensure error handling is reset.
End Sub





