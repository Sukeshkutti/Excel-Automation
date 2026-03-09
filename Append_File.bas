Attribute VB_Name = "Module1"
Sub Append_MR_To_Master()
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim MR As String, MR1 As String
    Dim MR_Date As Date, MR1_Date As Date
    Dim wbMR As Workbook, wbMaster As Workbook
    Dim wsMR As Worksheet, wsMaster As Worksheet
    Dim dict As Object
    Dim i As Long, j As Long
    Dim lastRowMR As Long, lastRowMaster As Long
    Dim lastColMR As Long, lastColMaster As Long
    Dim headerName As String
    Dim targetCol As Long
    'Select Folder
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If fDialog.Show = -1 Then
        folderPath = fDialog.SelectedItems(1)
    Else
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    'Find MR and MR1
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "xls*" Then
            If file.DateLastModified > MR_Date Then
                MR1 = MR
                MR1_Date = MR_Date
                MR = file.Path
                MR_Date = file.DateLastModified
            ElseIf file.DateLastModified > MR1_Date Then
                MR1 = file.Path
                MR1_Date = file.DateLastModified
            End If
        End If
    Next file
    'Open files
    Set wbMR = Workbooks.Open(MR)
    Set wbMaster = Workbooks.Open(MR1)
    Set wsMR = wbMR.Sheets(1)
    Set wsMaster = wbMaster.Sheets(1)
    lastRowMR = wsMR.Cells(Rows.Count, 1).End(xlUp).Row
    lastColMR = wsMR.Cells(1, Columns.Count).End(xlToLeft).Column
    lastRowMaster = wsMaster.Cells(Rows.Count, 1).End(xlUp).Row
    lastColMaster = wsMaster.Cells(1, Columns.Count).End(xlToLeft).Column
    'Create dictionary for master headers
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To lastColMaster
        headerName = Trim(wsMaster.Cells(1, i).Value)
        If headerName <> "" Then dict(headerName) = i
    Next i
    'Append data
    For i = 2 To lastRowMR
        lastRowMaster = lastRowMaster + 1
        For j = 1 To lastColMR
            headerName = Trim(wsMR.Cells(1, j).Value)
            If dict.exists(headerName) Then
                targetCol = dict(headerName)
                wsMaster.Cells(lastRowMaster, targetCol).Value = wsMR.Cells(i, j).Value
            End If
        Next j
    Next i
    wbMaster.Save
    wbMR.Close False
    
    MsgBox "Data appended successfully!"
End Sub
