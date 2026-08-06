' There are 3 Spreadsheets with 3 sheets ("Troubled", "On Watch", "On Track") and filtered columns, each have "Project ID" and "Meeting Minutes" , ... columns
' for each one, we need to read the visible "Project ID", find its "Meeting Minutes", 
' find the project in the destination spreadsheet and copy the minutes in the destination spreadsheet in column "Meeting Minutes"
' "Project ID" is the Key and we will recieve a MM_Log sheet too
-----------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub ConsolidateMeetingMinutes()
    Dim masterWB As Workbook, masterWS As Worksheet, logWS As Worksheet
    Dim srcWB As Workbook, ws As Worksheet
    Dim srcFiles As Variant, sheetNames As Variant
    Dim dict As Object, seenCount As Object, dupList As Object, noMatchList As Object
    Dim i As Long, j As Long, r As Long, logRow As Long
    Dim idCol As Long, mmCol As Long, lastRow As Long
    Dim pid As String, sourceLabel As String

    Set masterWB = ThisWorkbook ' run this macro from the master file
    Set masterWS = masterWB.Sheets(1) ' change if master sheet isn't the first
    Set dict = CreateObject("Scripting.Dictionary") ' pid -> minutes (last value wins)
    Set seenCount = CreateObject("Scripting.Dictionary") ' pid -> count of occurrences
    Set dupList = CreateObject("Scripting.Dictionary") ' pid -> list of "File|Sheet" locations
    Set noMatchList = CreateObject("Scripting.Dictionary")

    ' Update these paths to your 3 source files
    srcFiles = Array("C:\path\File1.xlsx", "C:\path\File2.xlsx", "C:\path\File3.xlsx")
    sheetNames = Array("Troubled", "On Watch", "On Track")

    For i = LBound(srcFiles) To UBound(srcFiles)
        Set srcWB = Workbooks.Open(srcFiles(i), ReadOnly:=True)
        For j = LBound(sheetNames) To UBound(sheetNames)
            Set ws = srcWB.Sheets(sheetNames(j))
            idCol = FindColumn(ws, "Project ID")
            mmCol = FindColumn(ws, "Meeting Minutes")
            lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
            sourceLabel = Dir(srcFiles(i)) & " | " & ws.Name

            For r = 2 To lastRow
                If ws.Rows(r).Hidden = False Then ' skips filtered-out rows
                    pid = Trim(ws.Cells(r, idCol).Value)
                    If pid <> "" Then
                        ' Track occurrence count + locations for duplicate detection
                        If seenCount.Exists(pid) Then
                            seenCount(pid) = seenCount(pid) + 1
                            If dupList.Exists(pid) Then
                                dupList(pid) = dupList(pid) & "; " & sourceLabel
                            Else
                                dupList(pid) = dict(pid + "_firstloc") ' placeholder, fixed below
                            End If
                        Else
                            seenCount(pid) = 1
                        End If
                        dict(pid & "_loc" & seenCount(pid)) = sourceLabel ' store each location
                        dict(pid) = ws.Cells(r, mmCol).Value ' last value wins
                    End If
                End If
            Next r
        Next j
        srcWB.Close SaveChanges:=False
    Next i

    ' Build clean duplicate location strings
    Dim dupPids As Object
    Set dupPids = CreateObject("Scripting.Dictionary")
    For Each pidKey In seenCount.Keys
        If seenCount(pidKey) > 1 Then
            Dim locs As String, k As Long
            locs = ""
            For k = 1 To seenCount(pidKey)
                If dict.Exists(pidKey & "_loc" & k) Then
                    If locs <> "" Then locs = locs & "; "
                    locs = locs & dict(pidKey & "_loc" & k)
                End If
            Next k
            dupPids(pidKey) = locs
        End If
    Next pidKey

    ' Write to master
    idCol = FindColumn(masterWS, "Project ID")
    mmCol = FindColumn(masterWS, "Meeting Minutes")
    lastRow = masterWS.Cells(masterWS.Rows.Count, idCol).End(xlUp).Row

    For r = 2 To lastRow
        pid = Trim(masterWS.Cells(r, idCol).Value)
        If pid <> "" Then
            If dict.Exists(pid) Then
                masterWS.Cells(r, mmCol).Value = dict(pid)
            Else
                noMatchList(pid) = True
            End If
        End If
    Next r

    ' --- Write log sheet ---
    On Error Resume Next
    Application.DisplayAlerts = False
    masterWB.Sheets("MM_Log").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set logWS = masterWB.Sheets.Add(After:=masterWB.Sheets(masterWB.Sheets.Count))
    logWS.Name = "MM_Log"
    logWS.Range("A1").Value = "Type"
    logWS.Range("B1").Value = "Project ID"
    logWS.Range("C1").Value = "Details"
    logWS.Range("A1:C1").Font.Bold = True

    logRow = 2
    For Each pidKey In dupPids.Keys
        logWS.Cells(logRow, 1).Value = "Duplicate"
        logWS.Cells(logRow, 2).Value = pidKey
        logWS.Cells(logRow, 3).Value = dupPids(pidKey)
        logRow = logRow + 1
    Next pidKey

    For Each pidKey In noMatchList.Keys
        logWS.Cells(logRow, 1).Value = "No Match in Source"
        logWS.Cells(logRow, 2).Value = pidKey
        logWS.Cells(logRow, 3).Value = "Present in master, not found (visible) in any source sheet"
        logRow = logRow + 1
    Next pidKey

    logWS.Columns("A:C").AutoFit

    MsgBox "Done." & vbCrLf & _
           "Matched & updated: " & (dict.Count) & " unique IDs (see dict for raw count)." & vbCrLf & _
           "Duplicates found: " & dupPids.Count & vbCrLf & _
           "No match in source: " & noMatchList.Count & vbCrLf & _
           "See 'MM_Log' sheet for details.", vbInformation
End Sub

Function FindColumn(ws As Worksheet, headerName As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then FindColumn = c.Column Else MsgBox "Header '" & headerName & "' not found on " & ws.Name
End Function
