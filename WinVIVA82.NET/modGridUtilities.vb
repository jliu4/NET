Option Strict Off
Option Explicit On

Module modGridUtilities

    Public Sub CopyToExcelFile(ByRef dataGridView1 As DataGridView)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim dlgSave As System.Windows.Forms.SaveFileDialog = _
        New System.Windows.Forms.SaveFileDialog

        'should the user cancel the dialog box, exit
        On Error GoTo ErrHandler

        dlgSave.OverwritePrompt = True
        dlgSave.Filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
        dlgSave.FileName = "*.xls"

        dlgSave.ShowDialog()

        Dim i As Int16, j As Int16

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        For i = 0 To dataGridView1.RowCount - 2
            For j = 0 To dataGridView1.ColumnCount - 1
                If dataGridView1(j, i).Value IsNot Nothing Then
                    xlWorkSheet.Cells(i + 1, j + 1) = dataGridView1(j, i).Value.ToString()
                End If
            Next
        Next
        If dlgSave.FileName IsNot Nothing Then

            xlWorkBook.SaveAs(dlgSave.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
             Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
        End If

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MessageBox.Show("Saved to the Excel File")
        Exit Sub

ErrHandler:
        'user pressed Cancel button
        Exit Sub

    End Sub


    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
            MessageBox.Show("Exception Occured while releasing object " + ex.ToString())
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub AddRow(ByRef dataGridView1 As DataGridView, ByRef CheckingGrid As Boolean, Optional ByRef MaxRows As Short = 0)

        Dim CurRow As Short

        CheckingGrid = True
        With dataGridView1
            If .Rows.Count = 1 Then
                CurRow = .Rows.Add
            Else
                CurRow = .CurrentCell.RowIndex
            End If
            Try
                .Rows.Insert(CurRow, New String() {})
            Catch ex As Exception
            End Try


        End With
        NumberAllRows(dataGridView1)
        CheckingGrid = False

    End Sub


    Public Sub DeleteRow(ByRef dataGridView1 As DataGridView, ByRef CheckingGrid As Boolean, Optional ByRef MinRows As Short = 0)
        CheckingGrid = True
        With dataGridView1
            Dim CurRow As Short = .CurrentCell.RowIndex
            If CurRow < .RowCount - 1 Then
                .Rows.RemoveAt(CurRow)

            End If

        End With
        NumberAllRows(dataGridView1)

        CheckingGrid = False

    End Sub
    Public Sub copyToClipBoard(ByRef dataGridView1 As DataGridView)

        'dataGridView1.SelectAll()
        Dim dataObj As DataObject = dataGridView1.GetClipboardContent()

        Clipboard.SetDataObject(dataObj, True)
        dataGridView1.ClearSelection()

    End Sub

    Public Sub PastefromClipBoardToDataGridView(ByRef dataGridView1 As DataGridView, ByRef CheckingGrid As Boolean, Optional ByRef MaxRows As Short = 0)

        Dim rowSplitter As Char() = {vbCr, vbLf}
        Dim columnSplitter As Char() = {vbTab}
        Dim maxAllowableRows As Short
        CheckingGrid = True
        'get the text from clipboard

        Dim dataInClipboard As IDataObject = Clipboard.GetDataObject()
        Dim stringInClipboard As String = CStr(dataInClipboard.GetData(DataFormats.Text))
        If stringInClipboard Is Nothing Then Exit Sub
        'split it into lines
        Dim rowsInClipboard As String() = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries)

        'get the row and column of selected cell in grid
        Dim r As Integer = dataGridView1.SelectedCells(0).RowIndex
        Dim c As Integer = dataGridView1.SelectedCells(0).ColumnIndex

        'add rows into grid to fit clipboard lines
        maxAllowableRows = r + rowsInClipboard.Length
        If MaxRows > 0 And r + rowsInClipboard.Length > MaxRows Then
            maxAllowableRows = MaxRows
        End If
        'copy part of columns, rows can not be added.
        If c > 0 And maxAllowableRows > dataGridView1.Rows.Count - 1 Then
            maxAllowableRows = dataGridView1.Rows.Count - 1
        End If
        If (dataGridView1.Rows.Count - 1 < maxAllowableRows) Then
            dataGridView1.Rows.Add(maxAllowableRows - dataGridView1.Rows.Count + 1)
        End If
        ' loop through the lines, split them into cells and place the values in the corresponding cell.
        Dim iRow As Integer = 0
        While iRow < maxAllowableRows - r
            'split row into cell values
            Dim valuesInRow As String() = rowsInClipboard(iRow).Split(columnSplitter)
            'cycle through cell values
            Dim iCol As Integer = 0
            While iCol < valuesInRow.Length
                'assign cell value, only if it within columns of the grid
                If (dataGridView1.ColumnCount - 1 >= c + iCol) Then
                    'if empty, default to 1
                    'If Not IsNumeric(valuesInRow(iCol)) Then
                    'valuesInRow(iCol) = 1
                    'End If
                    dataGridView1.Rows(r + iRow).Cells(c + iCol).Value = valuesInRow(iCol)
                End If

                iCol += 1
            End While
            iRow += 1
        End While
        CheckingGrid = True
    End Sub
    Public Sub NumberAllRows(ByRef dataGridView1 As DataGridView)
        ' Add row headers.
        For i As Integer = 0 To dataGridView1.Rows.Count - 1
            dataGridView1.Rows(i).HeaderCell.Value = (i + 1).ToString()
        Next i
    End Sub

End Module