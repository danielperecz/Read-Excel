Module PrintDataTable

    Sub Main()

        Dim path As String = "G:\Daniel\Split Excel Sheet By Column\Generic_UploadReport_14112019_1047.xlsx"
        Dim sheetName As String = "Upload Report"

        ' Create new ReadExcel object, and read the specified sheet.
        Dim obj As New ReadExcel(path)
        Dim arrayFromExcel As Object = obj.ReadSheet(sheetName)
        
        ' Initialize obj.DTable using the CreateDataTable method.
        Dim header As String() = obj.GetHeader(arrayFromExcel, 5)
        obj.CreateDataTable(arrayFromExcel, header)
        
        ' Loop through DataRows of DTable and print.
        For Each row As DataRow In obj.DTable.Rows
            Console.WriteLine(String.Join(" ", row.ItemArray))
        Next
        
        Console.ReadLine()

    End Sub

End Module
