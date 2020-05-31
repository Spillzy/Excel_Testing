Imports System.ComponentModel
Imports System.Drawing.Text
Imports Microsoft.Office.Interop


Public Class Form1

    Dim objExcel As New Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet

    Private Sub btnContinue_Click(sender As Object, e As EventArgs) Handles btnContinue.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        objWorkbook = objExcel.Workbooks.Add
        objWorksheet = objExcel.Worksheets.Add
        objExcel.Visible = True

        'MessageBox.Show("Opening Excel", "Test", MessageBoxButtons.OK)
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        'Close this Workbook and quit this instance of Excel
        objWorkbook.Close()
        objExcel.Quit()
        releaseObject(objWorksheet)
        releaseObject(objWorkbook)
        releaseObject(objExcel)

    End Sub

    Shared Sub releaseObject(ByVal obj As Object)
        'Handles any errors caused when attempted to close Excel.
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim rangeTest As Excel.Range
        rangeTest = objWorksheet.Range("A1")
        rangeTest.Value = TextBox1.Text
    End Sub

End Class
