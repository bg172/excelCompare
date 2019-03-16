' Description:
'   Reads a Microsoft Excel file.
' Parameters:
'   strFile - [in] The name of the Excel file to read.
' Returns:
'   A two-dimension array of cell values, if successful.
'   Null on error
' needs full filename including path to read file succesfully
Option Explicit

Function ReadExcelFile(ByVal strFile)

  ' Local variable declarations
  Dim objExcel, objSheet, objCells
  Dim nUsedRows, nUsedCols, nTop, nLeft, nRow, nCol
  Dim arrSheet()

  ' Default return value
  ReadExcelFile = Null

  ' Create the Excel object
  On Error Resume Next
  Set objExcel = CreateObject("Excel.Application")
  If (Err.Number <> 0) Then
    Exit Function
  End If

  ' Don't display any alert messages
  objExcel.DisplayAlerts = 0  

  ' Open the document as read-only
  On Error Resume Next
  Call objExcel.Workbooks.Open(strFile, False, True)
  If (Err.Number <> 0) Then
    Exit Function
  End If

  ' If you wanted to read all sheets, you could call
  ' objExcel.Worksheets.Count to get the number of sheets
  ' and the loop through each one. But in this example, we
  ' will just read the first sheet.
  Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

  ' Get the number of used rows
  nUsedRows = objSheet.UsedRange.Rows.Count

  ' Get the number of used columns
  nUsedCols = objSheet.UsedRange.Columns.Count

  ' Get the topmost row that has data
  nTop = objSheet.UsedRange.Row

  ' Get leftmost column that has data
  nLeft = objSheet.UsedRange.Column

  ' Get the used cells
  Set objCells = objSheet.Cells

  ' Dimension the sheet array
  ReDim arrSheet(nUsedRows - 1, nUsedCols - 1)

  ' Loop through each row
  For nRow = 0 To (nUsedRows - 1)
    ' Loop through each column
    For nCol = 0 To (nUsedCols - 1)
  ' Add the cell value to the sheet array
  arrSheet(nRow, nCol) = objCells(nRow + nTop, nCol + nLeft).Value
    Next
  Next

  ' Close the workbook without saving
  Call objExcel.ActiveWorkbook.Close(False)

  ' Quit Excel
  objExcel.Application.Quit

  ' Return the sheet data to the caller
  ReadExcelFile = arrSheet

End Function




' Local variable declarations
Dim strFile, objFSO, arrSheet, i, j, varCell, strFormat, str, cols

' Prompt for the Excel file to read  
strFile = WScript.Arguments(0)
'Wscript.Echo strFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
strFile = objFSO.GetAbsolutePathName(strFile)
'Wscript.Echo "Absolute path: " & strFile


' Read the Excel file
arrSheet = ReadExcelFile(strFile)
    
' Dump the worksheet to the command line
cols = UBound(arrSheet, 2)
For i = 0 To UBound(arrSheet, 1)
    str = ""
    For j = 0 To cols
      varCell = arrSheet(i, j)
      str = str & CStr(varCell)
      If (j < cols - 1) Then str = str & "|"
    Next
    Wscript.Echo str
Next

