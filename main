Sub SheetToDiscreet()
    Dim lastRow As Long, i As Long
    Dim key As String, val As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim operativeNote As String
    Dim firstName As String
    Dim lastName As String
    Dim dateOfSurgery As String
    Dim fileName As String
    Dim folderPath As String
    
    ' Prompt user to select a folder to save the files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select a folder to save the files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        key = Cells(i, "B").Value & "|" & Cells(i, "C").Value & "|" & Format(Cells(i, "D").Value, "mm/dd/yy")
        val = Cells(i, "A").Value
        If dict.exists(key) Then
            dict(key) = dict(key) & Chr(10) & val
        Else
            dict(key) = val
        End If
    Next i
    
    Range("A2:D" & lastRow).ClearContents
    
    For i = 0 To dict.Count - 1
        key = dict.keys()(i)
        Cells(i + 2, "B").Value = Split(key, "|")(0)
        Cells(i + 2, "C").Value = Split(key, "|")(1)
        Cells(i + 2, "D").Value = Split(key, "|")(2)
        Cells(i + 2, "A").Value = dict(key)
        
        operativeNote = dict(key)
        firstName = Split(key, "|")(0)
        lastName = Split(key, "|")(1)
        dateOfSurgery = Split(key, "|")(2)
        fileName = lastName & ", " & firstName & " - " & Format(dateOfSurgery, "mm-dd-yyyy") & ".rtf"
        
        ' Create a new text file and write the operative note to it
        Open folderPath & fileName For Output As #1
        Print #1, operativeNote
        Close #1
    Next i
    
    ' Open the selected folder in Windows File Explorer
    Shell "explorer.exe """ & folderPath & """"
End Sub
