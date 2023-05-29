Imports System.IO
Imports OfficeOpenXml

Module Utilities
    Public Sub ExportToExcel(dataGrid As DataGridView)
        ' Set the LicenseContext to properly handle licensing
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Using package As New ExcelPackage()
            Dim worksheet = package.Workbook.Worksheets.Add("Sheet1")

            ' Write column headers
            For i As Integer = 0 To dataGrid.Columns.Count - 1
                worksheet.Cells(1, i + 1).Value = dataGrid.Columns(i).HeaderText
            Next

            ' Write data rows
            For i As Integer = 0 To dataGrid.Rows.Count - 1
                For j As Integer = 0 To dataGrid.Columns.Count - 1
                    Dim cellValue = If(dataGrid.Rows(i).Cells(j).Value IsNot Nothing, dataGrid.Rows(i).Cells(j).Value.ToString(), "")
                    worksheet.Cells(i + 2, j + 1).Value = cellValue
                Next
            Next

            ' Save the Excel file
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath As String = saveFileDialog.FileName
                Dim fileInfo As New FileInfo(filePath)
                package.SaveAs(fileInfo)
            End If
        End Using
    End Sub
End Module
