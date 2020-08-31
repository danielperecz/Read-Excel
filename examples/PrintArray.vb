Module PrintArray

    Sub Main()

        Dim path As String = "G:\Daniel\Split Excel Sheet By Column\Generic_UploadReport_14112019_1047.xlsx"
        Dim sheetName As String = "Upload Report"

        ' Create new ReadExcel object, and read the specified sheet.
        Dim obj As New ReadExcel(path)
        Dim arrayFromExcel As Object = obj.ReadSheet(sheetName)

        ' Loop through all rows of arrayFromExcel. We start at index 1 because arrayFromExcel starts at coordinates (1, 1).
        ' y represents rows, and x represents columns.
        Dim row(obj.Columns - 1) As Object
        For y As Integer = 1 To obj.Rows
            For x As Integer = 1 To obj.Columns
                'row(x - 1) = arrayFromExcel(y, x)                       ' row(x - 1) because row array starts at index 0 (and x starts at 1)
                Console.Write(arrayFromExcel(y, x) & " ")
            Next
            Console.WriteLine()
        Next
        
        Console.ReadLine()

    End Sub

End Module
