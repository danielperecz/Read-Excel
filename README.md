# Documentation

## Purpose
To simlify general Excel tasks and avoid unnecessary code duplication, with efficient execution speeds.

## `ReadExcel` class explained
The object has a single paremeter - the path to an Excel file, so object initialization may look like: `Dim obj As New ReadExcel(path)`. Currently the most useful function is `ReadSheet`, which takes a sheet name as its parameter, and loads the sheet into memory in a 2 dimensional array format. To call the function: `Dim array As Object = obj.ReadSheet(sheetName)`. The first cell, i.e. `A1`, can be obtained from the new `array` variable simply by writing `array(1, 1)`.

## Usage
Initialize object
```
Dim obj As New ReadExcel(path)
```

Read sheet
```
Dim array As Object = obj.ReadSheet(sheetName)
```

Get basic information about the sheet
```
Dim numberOfRows As Long = obj.Rows
Dim numberOfColumns as Long = obj.Columns
Dim sheetDimensions as String = obj.SheetDimensions
Dim sheetNames As Object = obj.Sheets
```

## Beta features
Create a DataTable from the array for faster operations
```
Dim header As String() = obj.GetHeader(array, 5)   ' The 5 is the row number of the header in the Excel file
obj.CreateDataTable(array, header)
```

Find the row index of a value (first match)
```
Dim i As Long = obj.FindRowByKeyword("3463831728")
```
