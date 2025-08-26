' ============================================================================
' Microsoft Excel VBA Image Insertion - Multiple Pictures, Multi-Row
' This macro inserts images horizontally with automatic row wrapping
' First row: B5 to BJ5, Second row: Q29 onwards, etc
' ============================================================================

Sub InsertAndFit_Multiple_Images_MultiRow()
    ' --- CONFIGURATION ---
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pictures") ' Target worksheet
    
    Dim folderPath As String
    folderPath = "C:\Images\" ' !!! IMPORTANT: UPDATE this to your folder path !!!
    
    Dim sortMethod As String
    sortMethod = "DateCreated" ' Options: "DateCreated", "FileName", "DateTaken"
    
    ' Row and column settings
    Dim firstRowStartCell As Range
    Set firstRowStartCell = ws.Range("B5") ' First row starts at B5
    
    Dim lastColumnInRow As Integer
    lastColumnInRow = 62 ' BJ column (BJ5 is the last cell in first row)
    
    Dim secondRowStartCell As Range
    Set secondRowStartCell = ws.Range("Q29") ' Second row starts at Q29
    
    Dim colOffset As Integer
    colOffset = 15 ' Number of columns to jump between images
    
    Dim rowOffset As Integer
    rowOffset = 24 ' Number of rows to jump when moving to next row (29-5=24)
    ' --- END CONFIGURATION ---

    ' Declare variables
    Dim fileArray() As String
    Dim fileDate() As Date
    Dim i As Integer, j As Integer
    Dim currentCell As Range
    Dim mergedArea As Range
    Dim pic As Object
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim fileCount As Integer
    Dim successCount As Integer
    Dim originalZoom As Integer
    Dim currentRow As Integer
    Dim currentCol As Integer

    On Error GoTo MacroError

    ' For consistent positioning, temporarily set zoom to 100%
    originalZoom = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100

    ' --- File Collection and Sorting ---
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "ERROR: Folder not found: " & folderPath, vbCritical
        GoTo MacroCleanup
    End If

    Set folder = fso.GetFolder(folderPath)
    fileCount = 0
    For Each file In folder.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "bmp" Then
            ReDim Preserve fileArray(fileCount)
            ReDim Preserve fileDate(fileCount)
            fileArray(fileCount) = file.Name
            
            ' Sort by selected method
            Select Case UCase(sortMethod)
                Case "FILENAME"
                    fileDate(fileCount) = CDate("1900-01-01") ' Dummy date for filename sorting
                Case "DATETAKEN"
                    fileDate(fileCount) = GetDateTaken(folderPath & file.Name)
                Case Else ' "DATECREATED"
                    fileDate(fileCount) = file.DateCreated
            End Select
            
            fileCount = fileCount + 1
        End If
    Next file

    If fileCount = 0 Then
        MsgBox "No image files were found in " & folderPath, vbInformation
        GoTo MacroCleanup
    End If

    ' Sort files based on selected method
    If UCase(sortMethod) = "FILENAME" Then
        ' Sort alphabetically by filename
        For i = 0 To fileCount - 2
            For j = i + 1 To fileCount - 1
                If UCase(fileArray(i)) > UCase(fileArray(j)) Then
                    Dim tempFile As String: tempFile = fileArray(i): fileArray(i) = fileArray(j): fileArray(j) = tempFile
                    Dim tempDate As Date: tempDate = fileDate(i): fileDate(i) = fileDate(j): fileDate(j) = tempDate
                End If
            Next j
        Next i
    Else
        ' Sort by date (DateCreated or DateTaken)
        For i = 0 To fileCount - 2
            For j = i + 1 To fileCount - 1
                If fileDate(i) > fileDate(j) Then
                    tempDate = fileDate(i): fileDate(i) = fileDate(j): fileDate(j) = tempDate
                    tempFile = fileArray(i): fileArray(i) = fileArray(j): fileArray(j) = tempFile
                End If
            Next j
        Next i
    End If
    
    ' --- Image Insertion Loop with Row Wrapping ---
    MsgBox "Found " & fileCount & " images. Sorting by " & sortMethod & ". Now inserting...", vbInformation
    Set currentCell = firstRowStartCell
    successCount = 0
    currentRow = firstRowStartCell.Row
    currentCol = firstRowStartCell.Column

    For i = 0 To fileCount - 1
        If currentCell Is Nothing Then
            MsgBox "Error: The target cell became invalid. Aborting.", vbCritical
            Exit For
        End If
        
        ' Get the full merged area from the current cell's position
        If currentCell.MergeCells Then
            Set mergedArea = currentCell.MergeArea
        Else
            Set mergedArea = currentCell
        End If

        ' Clear any old pictures in the target area first
        Dim shp As Shape
        For Each shp In ws.Shapes
            If Not Intersect(shp.TopLeftCell, mergedArea) Is Nothing Then
                shp.Delete
            End If
        Next shp

        ' Insert, position, and size the new picture
        On Error Resume Next
        Set pic = ws.Shapes.AddPicture(folderPath & fileArray(i), msoFalse, msoTrue, mergedArea.Left, mergedArea.Top, mergedArea.Width, mergedArea.Height)
        
        If Err.Number <> 0 Then
            Debug.Print "Could not insert image: " & fileArray(i) & ". Error: " & Err.Description
            Err.Clear
        Else
            pic.LockAspectRatio = msoFalse
            pic.Placement = xlMoveAndSize
            successCount = successCount + 1
        End If
        On Error GoTo MacroError

        ' Calculate next position with row wrapping
        currentCol = currentCol + colOffset
        
        ' Check if we've exceeded the last column in the current row
        If currentCol > lastColumnInRow Then
            ' Move to next row
            currentRow = currentRow + rowOffset
            currentCol = secondRowStartCell.Column ' Reset to starting column (Q)
        End If
        
        ' Set the next cell position
        Set currentCell = ws.Cells(currentRow, currentCol)
    Next i

MacroCleanup:
    If originalZoom > 0 Then ActiveWindow.Zoom = originalZoom
    Set fso = Nothing
    Set file = Nothing
    Set folder = Nothing
    Set pic = Nothing
    MsgBox "Insertion complete. Successfully inserted " & successCount & " of " & fileCount & " images." & vbCrLf & _
           "Sorted by: " & sortMethod, vbInformation
    Exit Sub

MacroError:
    MsgBox "A critical error occurred: " & vbCrLf & Err.Description, vbCritical
    Resume MacroCleanup
End Sub

' Function to extract Date Taken from image metadata
Function GetDateTaken(imagePath As String) As Date
    On Error Resume Next
    
    Dim shell As Object
    Dim folder As Object
    Dim file As Object
    Dim dateTaken As String
    
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.Namespace(Left(imagePath, InStrRev(imagePath, "\")))
    Set file = folder.ParseName(Mid(imagePath, InStrRev(imagePath, "\") + 1))
    
    ' Property 12 is "Date taken" in Windows Explorer
    dateTaken = folder.GetDetailsOf(file, 12)
    
    If dateTaken <> "" And IsDate(dateTaken) Then
        GetDateTaken = CDate(dateTaken)
    Else
        ' Fallback to file creation date if Date Taken is not available
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        GetDateTaken = fso.GetFile(imagePath).DateCreated
    End If
    
    On Error GoTo 0
End Function
