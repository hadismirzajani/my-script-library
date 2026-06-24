' =============================================================
' BulkExportMPP_v3.bas
' Bulk export MS Project files from Project Online (PWA)
' VERSION 3 - Uses correct Enterprise Project open method
# THIS WAS NEVER CHECKED AS I DID NOT HAVE REQUIRED API ENDPOINT ACCESS DUE TO WHICH IT BECAME A SEMI AUTOMATED JOB USING VBS (COULD NOT CONNECT TO PWA TO OPEN A FILE)
'=============================================================

Option Explicit

'-------------------------------------------------------------
' CONFIG
'-------------------------------------------------------------
Const PWA_URL As String = "https://accessmbrc.sharepoint.com/sites/PWA-MBRCPMOOffice"
Const EXCEL_PATH As String = "C:\Users\hadismirzajani\...\page_source\VBS Feed.xlsx"
Const OUTPUT_FOLDER As String = "C:\Users\hadismirzajani\...\page_source\VBSBackupTest\"
Const PROJECT_COL As Integer = 1 ' Column A
Const HEADER_ROWS As Integer = 1 ' Skip 1 header row
'-------------------------------------------------------------

Sub BulkExportMPP()

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim lastRow As Long
    Dim i As Long
    Dim projectName As String
    Dim safeName As String
    Dim outputPath As String
    Dim successCount As Integer
    Dim failCount As Integer
    Dim logLines As String
    Dim logPath As String

    logPath = OUTPUT_FOLDER & "ExportLog_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"

    ' --- Ensure output folder exists ---
    If Dir(OUTPUT_FOLDER, vbDirectory) = "" Then
        MkDir OUTPUT_FOLDER
    End If

    ' --- Open Excel and read project list ---
    On Error GoTo ErrorHandler

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWb = xlApp.Workbooks.Open(EXCEL_PATH)
    Set xlWs = xlWb.Sheets(1)

    lastRow = xlWs.Cells(xlWs.Rows.Count, PROJECT_COL).End(-4162).Row

    If lastRow <= HEADER_ROWS Then
        MsgBox "No project names found in the Excel file.", vbExclamation
        GoTo Cleanup
    End If

    successCount = 0
    failCount = 0
    logLines = "Bulk MPP Export Log - " & Now & vbCrLf
    logLines = logLines & "PWA: " & PWA_URL & vbCrLf
    logLines = logLines & "Total rows found: " & (lastRow - HEADER_ROWS) & vbCrLf
    logLines = logLines & String(60, "-") & vbCrLf

    ' --- Loop through each project ---
    For i = HEADER_ROWS + 1 To lastRow

        projectName = Trim(CStr(xlWs.Cells(i, PROJECT_COL).Value))

        If projectName = "" Then
            logLines = logLines & "Row " & i & ": SKIPPED (empty name)" & vbCrLf
            GoTo NextProject
        End If

        ' Build safe filename
        safeName = projectName
        safeName = ReplaceChar(safeName, "/", "-")
        safeName = ReplaceChar(safeName, "\", "-")
        safeName = ReplaceChar(safeName, ":", "-")
        safeName = ReplaceChar(safeName, "*", "-")
        safeName = ReplaceChar(safeName, "?", "-")
        safeName = ReplaceChar(safeName, """", "-")
        safeName = ReplaceChar(safeName, "<", "-")
        safeName = ReplaceChar(safeName, ">", "-")
        safeName = ReplaceChar(safeName, "|", "-")

        outputPath = OUTPUT_FOLDER & safeName & ".mpp"

        ' Skip if already exported (safe to re-run)
        If Dir(outputPath) <> "" Then
            logLines = logLines & "Row " & i & " [" & projectName & "]: SKIPPED (already exists)" & vbCrLf
            GoTo NextProject
        End If

        ' --- Open project from PWA Enterprise Projects ---
        On Error Resume Next
        Err.Clear
        
        ' This mimics exactly what File > Open does when you double-click
        ' an enterprise project in MS Project Desktop
        Application.Open PWA_URL & "/" & projectName, _
                         ReadOnly:=False, _
                         Merge:=0, _
                         openpool:=False

        Wait 4 ' Give it time to fully load from server

        If Err.Number <> 0 Then
            logLines = logLines & "Row " & i & " [" & projectName & "]: FAILED to open - " & Err.Description & vbCrLf
            failCount = failCount + 1
            Err.Clear
            GoTo NextProject
        End If

        ' Double check a project is actually open
        If Application.Projects.Count = 0 Then
            logLines = logLines & "Row " & i & " [" & projectName & "]: FAILED - no project loaded" & vbCrLf
            failCount = failCount + 1
            GoTo NextProject
        End If

        On Error GoTo ErrorHandler

        ' --- Save as local .mpp ---
        On Error Resume Next
        ActiveProject.SaveAs outputPath, pjMPP

        If Err.Number <> 0 Then
            logLines = logLines & "Row " & i & " [" & projectName & "]: FAILED to save - " & Err.Description & vbCrLf
            failCount = failCount + 1
            Err.Clear
        Else
            logLines = logLines & "Row " & i & " [" & projectName & "]: SUCCESS -> " & outputPath & vbCrLf
            successCount = successCount + 1
        End If
        On Error GoTo ErrorHandler

        ' --- Close without saving back to PWA ---
        On Error Resume Next
        ActiveProject.Close SaveChanges:=False
        On Error GoTo ErrorHandler

        Wait 2

NextProject:
    Next i

    ' --- Summary ---
    logLines = logLines & String(60, "-") & vbCrLf
    logLines = logLines & "COMPLETED: " & successCount & " exported, " & failCount & " failed." & vbCrLf

    WriteLog logPath, logLines

    MsgBox "Export complete!" & vbCrLf & vbCrLf & _
           "Exported: " & successCount & vbCrLf & _
           "Failed: " & failCount & vbCrLf & vbCrLf & _
           "Log saved to:" & vbCrLf & logPath, vbInformation, "Bulk MPP Export"

Cleanup:
    On Error Resume Next
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Unexpected error on row " & i & ": " & Err.Description, vbCritical
    logLines = logLines & "UNEXPECTED ERROR on row " & i & ": " & Err.Description & vbCrLf
    WriteLog logPath, logLines
    Resume Cleanup

End Sub

'-------------------------------------------------------------
' Helper: Wait n seconds
'-------------------------------------------------------------
Private Sub Wait(seconds As Integer)
    Dim t As Date
    t = Now + TimeSerial(0, 0, seconds)
    Do While Now < t
        DoEvents
    Loop
End Sub

'-------------------------------------------------------------
' Helper: Replace a character in a string
'-------------------------------------------------------------
Private Function ReplaceChar(str As String, findChar As String, replaceWith As String) As String
    Dim pos As Integer
    pos = InStr(str, findChar)
    Do While pos > 0
        str = Left(str, pos - 1) & replaceWith & Mid(str, pos + 1)
        pos = InStr(pos + 1, str, findChar)
    Loop
    ReplaceChar = str
End Function

'-------------------------------------------------------------
' Helper: Write log file
'-------------------------------------------------------------
Private Sub WriteLog(filePath As String, content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub
