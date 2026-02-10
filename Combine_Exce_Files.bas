Attribute VB_Name = "Module2"
Sub Merge_All_Excel_Files_With_Header_Alignment()

    Dim fDialog As FileDialog
    Dim FolderPath As String, FileName As String
    Dim srcWB As Workbook, destWB As Workbook
    Dim srcWS As Worksheet, destWS As Worksheet
    Dim lastRow As Long, destLastRow As Long
    Dim dictSheets As Object, dictHeaders As Object
    Dim combinedPath As String
    Dim colMap As Object
    Dim i As Long, j As Long
    Dim header As Variant
    Dim srcLastCol As Long, destLastCol As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Select folder
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If fDialog.Show = 0 Then Exit Sub
    FolderPath = fDialog.SelectedItems(1) & "\"

    combinedPath = FolderPath & "Combined_Excel.xlsx"
    If Dir(combinedPath) <> "" Then Kill combinedPath

    Set destWB = Workbooks.Add
    Set dictSheets = CreateObject("Scripting.Dictionary")
    Set dictHeaders = CreateObject("Scripting.Dictionary")

    FileName = Dir(FolderPath & "*.xls*")

    Do While FileName <> ""
        Set srcWB = Workbooks.Open(FolderPath & FileName)

        For Each srcWS In srcWB.Worksheets

            ' Create or get combined sheet
            If Not dictSheets.Exists(srcWS.Name) Then
                Set destWS = destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
                destWS.Name = srcWS.Name & "_Combined"
                dictSheets.Add srcWS.Name, destWS

                ' Capture header structure from first file
                srcLastCol = srcWS.Cells(1, srcWS.Columns.Count).End(xlToLeft).Column
                For i = 1 To srcLastCol
                    dictHeaders(srcWS.Name & "|" & srcWS.Cells(1, i).Value) = i
                    destWS.Cells(1, i).Value = srcWS.Cells(1, i).Value
                Next i
            Else
                Set destWS = dictSheets(srcWS.Name)
            End If

            ' Build column mapping
            Set colMap = CreateObject("Scripting.Dictionary")
            srcLastCol = srcWS.Cells(1, srcWS.Columns.Count).End(xlToLeft).Column
            destLastCol = destWS.Cells(1, destWS.Columns.Count).End(xlToLeft).Column

            For i = 1 To srcLastCol
                header = srcWS.Cells(1, i).Value
                For j = 1 To destLastCol
                    If destWS.Cells(1, j).Value = header Then
                        colMap(i) = j
                        Exit For
                    End If
                Next j
            Next i

            lastRow = srcWS.Cells(srcWS.Rows.Count, 1).End(xlUp).row
            destLastRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).row + 1

            ' Append data with aligned columns
            For i = 2 To lastRow
                For Each header In colMap.Keys
                    destWS.Cells(destLastRow, colMap(header)).Value = srcWS.Cells(i, header).Value
                Next header
                destLastRow = destLastRow + 1
            Next i

        Next srcWS

        srcWB.Close False
        FileName = Dir
    Loop

    destWB.SaveAs combinedPath

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "All files merged with header alignment successfully!", vbInformation

End Sub


