Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports OfficeOpenXml

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles UploadButton_Click.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "CSV files (*.csv)|*.csv"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName

            ' Load data from CSV file
            Dim dataTable As DataTable = LoadDataFromCSV(filePath)

            ' Display the data in the DataGridView
            DataGridView1.DataSource = dataTable
        End If
    End Sub
    Private Function LoadDataFromCSV(filePath As String) As DataTable
        Dim dataTable As New DataTable()

        Using parser As New TextFieldParser(filePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            ' Read the header row
            Dim headers As String() = parser.ReadFields()
            For Each header In headers
                dataTable.Columns.Add(header)
            Next

            ' Read the data rows
            While Not parser.EndOfData
                Dim fields As String() = parser.ReadFields()
                dataTable.Rows.Add(fields)
            End While
        End Using

        Return dataTable
    End Function

    Private Sub PrintButton_Click(sender As Object, e As EventArgs) Handles PrintButton.Click
        Utilities.ExportToExcel(DataGridView1)
    End Sub

End Class
