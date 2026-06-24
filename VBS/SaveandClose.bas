Attribute VB_Name = "Module1"
'=============================================================
' SaveAndClose.bas
' One-click save of currently open PWA project to local .mpp
' HOW TO USE:
'   1. Open a project from PWA via File > Open as normal
'   2. Press Alt+F8, select "SaveCurrentProject", click Run
'   OR assign it to a toolbar button (instructions below)
' This module was saved in Global
'=============================================================
 
Option Explicit
 
Const OUTPUT_FOLDER As String = "C:\Users\hadismirzajani\OneDrive - City of Moreton Bay\Desktop\page_source\VBSBackupTest\"
Const LOG_FILE      As String = "C:\Users\hadismirzajani\OneDrive - City of Moreton Bay\Desktop\page_source\VBSBackupTest\ExportLog.txt"
 
'-------------------------------------------------------------
' MAIN MACRO - Run this after each project is open
'-------------------------------------------------------------
Sub SaveCurrentProject()
 
    Dim projectName As String
    Dim safeName    As String
    Dim outputPath  As String
 
    ' Check a project is actually open
    If Application.Projects.Count = 0 Then
        MsgBox "No project is open! Please open a project from PWA first.", vbExclamation
        Exit Sub
    End If
 
    ' Get the current project name
    projectName = ActiveProject.Name
 
    ' Confirm with user before saving
    Dim confirm As Integer
    confirm = MsgBox("Save this project as local .mpp?" & vbCrLf & vbCrLf & _
                     "Project: " & projectName & vbCrLf & _
                     "Folder:  " & OUTPUT_FOLDER, _
                     vbYesNo + vbQuestion, "Save Project")
 
    If confirm = vbNo Then
        Exit Sub
    End If
 
    ' Ensure output folder exists
    If Dir(OUTPUT_FOLDER, vbDirectory) = "" Then
        MkDir OUTPUT_FOLDER
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
 
    ' Warn if file already exists
    If Dir(outputPath) <> "" Then
        Dim overwrite As Integer
        overwrite = MsgBox("File already exists! Overwrite?" & vbCrLf & outputPath, _
                           vbYesNo + vbExclamation, "File Exists")
        If overwrite = vbNo Then
            Exit Sub
        End If
    End If
 
    ' Save as local .mpp
    On Error GoTo SaveError
    ActiveProject.SaveAs outputPath, pjMPP
    On Error GoTo 0
 
    ' Log the success
    AppendLog LOG_FILE, Now & " | SUCCESS | " & projectName & " -> " & outputPath
 
    ' Close without saving changes back to PWA
    FileClose pjDoNotSave
 
    MsgBox "Saved successfully!" & vbCrLf & vbCrLf & outputPath, vbInformation, "Done"
    Exit Sub
 
SaveError:
    MsgBox "Failed to save: " & Err.Description, vbCritical
    AppendLog LOG_FILE, Now & " | FAILED  | " & projectName & " - " & Err.Description
 
End Sub
 
'-------------------------------------------------------------
' Helper: Append a line to the log file
'-------------------------------------------------------------
Private Sub AppendLog(filePath As String, content As String)
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, content
    Close #fileNum
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

