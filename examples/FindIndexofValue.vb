Module FindIndexOfValue

    Sub Main()

        Dim path As String = "G:\Daniel\Split Excel Sheet By Column\Generic_UploadReport_14112019_1047.xlsx"
        Dim sheetName As String = "Upload Report"

        ' Create new ReadExcel object, and read the specified sheet.
        Dim obj As New ReadExcel(path)
        Dim arrayFromExcel As Object = obj.ReadSheet(sheetName)
        
        ' Initialize obj.DTable using the CreateDataTable method.
        Dim header As String() = obj.GetHeader(arrayFromExcel, 5)
        obj.CreateDataTable(arrayFromExcel, header)
        
        ' Find the row index of specificed value (has to be in string format). Counting starts at 1 (not 0).
        Console.WriteLine(obj.FindRowByKeyword("3463831728"))
        
        ' Find column index.
        Console.WriteLine(obj.FindColumnByKeyword("3463831728"))
        
        ' Find coordinates (row, column).
        Dim coordinates As Object = obj.FindCoordinatesByKeyword("3463831728")
        Console.WriteLine("(" & coordinates(0) & ", " & coordinates(1) & ")")

        Console.ReadLine()

    End Sub

End Module
