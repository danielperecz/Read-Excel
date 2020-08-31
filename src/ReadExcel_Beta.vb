Imports Excel = Microsoft.Office.Interop.Excel

Public Class ReadExcel

    Public FilePath As String
    Public SheetDimensions As String
    Public Rows As Integer
    Public Columns As Integer
    Public Sheets As Object

    ' test feature
    Public DTable As DataTable

    Public Sub New(ByVal filePath As String)
        ' Class constructor. Takes a path to an Excel file.

        Me.FilePath = filePath
        Sheets = SheetNames()

    End Sub

    Public Sub CreateDataTable(ByVal arrayFromExcel As Object, ByVal header As String())
        ' Updates DTable variable with a DataTable object of the Excel file.

        Dim dt As DataTable = New DataTable()

        ' Add column names to dt
        For Each name In header
            dt.Columns.Add(New DataColumn(name))
        Next

        ' Add rows to dt
        Dim cell As Object
        For y As Long = 1 To Rows
            Dim row(Columns - 1) As Object
            For x As Integer = 1 To Columns
                cell = arrayFromExcel(y, x)
                row(x - 1) = cell
            Next
            dt.Rows.Add(row)
        Next
        DTable = dt                                                                                     ' Update DTable with new DataTable
    End Sub

    Public Function ReadSheet(ByVal sheetName As String)
        ' Reads a single sheet from an Excel file and returns it as an array of rows and columns.

        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(FilePath)                               ' Open file
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets(sheetName)                           ' Select specific sheet
        Dim xlRange As Excel.Range = xlWorkSheet.UsedRange

        Dim array(,) As Object = xlRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)          ' Insert values into array
        'Dim array(,) As Object = DirectCast(xlRange.Value, Object(,))

        Try
            ' Try is necessary because the array may be empty, and the variable Rows will throw a NullReferenceException.

            ' Update global variables for use in other subroutines and functions.
            Rows = array.GetUpperBound(0)
            Columns = array.GetUpperBound(1)
            SheetDimensions = "[" & Rows & " rows x " & Columns & " columns]"

        Catch ex As System.NullReferenceException
            ' Exception occurs when array is empty. Set Rows and Columns to 0 since the array is empty, and return the array.
            Rows = 0
            Columns = 0
            SheetDimensions = "[" & Rows & " rows x " & Columns & " columns]"
            CloseExcel(xlApp, xlWorkBook)
            Return array
        End Try

        CloseExcel(xlApp, xlWorkBook)

        Return array

    End Function

    ' UTILITY FUNCTIONS AND SUBROUTINES

    Private Function SheetNames()
        ' Returns a list of strings which are the Excel sheet names.

        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(FilePath)
        Dim sheets As New List(Of String)

        ' Loop through the sheets and add them to the list.
        For Each xlWorkSheet As Excel.Worksheet In xlWorkBook.Sheets
            sheets.Add(xlWorkSheet.Name.ToString())
        Next

        CloseExcel(xlApp, xlWorkBook)
        Return sheets

    End Function

    Public Function GetHeader(ByVal arrayFromExcel As Object, ByVal row As Long)
        ' Returns an array of strings.
        Dim header(Columns - 1) As String
        For x As Integer = 1 To Columns
            header(x - 1) = arrayFromExcel(row, x)
        Next
        Return header
    End Function

    Public Function FindRowByKeyword(ByVal keyword As String)
        ' Returns row index of a string if a match is found. If no match is found: return -1.
        Try
            Dim y As Long = 1
            For Each row As DataRow In DTable.Rows
                If row.ItemArray.Contains(keyword) Then
                    Return y
                End If
                y += 1
            Next
            Return -1
        Catch ex As System.NullReferenceException
            ' DTable hasn't been initialized yet.
            Return "Call CreateDataTable() first."
        End Try
    End Function

    Public Function FindColumnByKeyword(ByVal keyword As String)
        ' Returns column index of a string if a match is found. If no match is found: return -1.
        Try
            Dim x As Long
            For Each row As DataRow In DTable.Rows
                If row.ItemArray.Contains(keyword) Then
                    x = Array.IndexOf(row.ItemArray.ToArray, keyword) + 1
                    Return x
                End If
            Next
            Return -1
        Catch ex As System.NullReferenceException
            ' DTable hasn't been initialized yet.
            Return "Call CreateDataTable() first."
        End Try
    End Function

    Public Function FindCoordinatesByKeyword(ByVal keyword As String)
        ' Returns coordinates of a string if a match is found. If no match is found: return -1.
        Try
            Dim y As Long = 1
            Dim x As Long
            For Each row As DataRow In DTable.Rows
                If row.ItemArray.Contains(keyword) Then
                    x = Array.IndexOf(row.ItemArray.ToArray, keyword) + 1
                    Return {y, x}
                End If
                y += 1
            Next
            Return -1
        Catch ex As System.NullReferenceException
            ' DTable hasn't been initialized yet.
            Return "Call CreateDataTable() first."
        End Try
    End Function

    Private Sub releaseObject(ByVal obj As Object)
        ' Release objects from memory (so background processes don't stay open in Task Manager).
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        obj = Nothing
    End Sub

    Public Sub CloseExcel(ByVal xlApp As Object, ByVal xlWorkBook As Object)
        ' Close and quit Excel.Application, and release objects.
        xlApp.Workbooks.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
    End Sub

End Class
