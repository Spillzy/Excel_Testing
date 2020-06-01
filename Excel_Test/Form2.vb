'Option Strict On
Option Explicit On
Imports Microsoft.Office.Interop

Public Class Form2

    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        'Need to somehow make objWorkbook be the Workbook that is currently open and being used.
        objExcel = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application) 'Grabs the first opened instance of Excel, this is definitly not the best method, but only one I found that works so far.
        objWorkbook = CType(objExcel.ActiveWorkbook, Excel.Workbook)
        objWorksheet = CType(objExcel.ActiveSheet, Excel.Worksheet)
        'objWorkbook = objExcel.ActiveWorkbook
        'objWorksheet = objExcel.ActiveSheet

        If objExcel Is Nothing Then
            MessageBox.Show("No object for Excel", "No Object Found", MessageBoxButtons.OK)
        ElseIf objWorkbook Is Nothing Then
            MessageBox.Show("No object for Workbook", "No Object Found", MessageBoxButtons.OK)
        ElseIf objWorksheet Is Nothing Then
            MessageBox.Show("No object for Worksheet", "No Object Found", MessageBoxButtons.OK)
        Else
            Dim rangeTest As Excel.Range
            rangeTest = objWorksheet.Range("B1")
            rangeTest.Value = TextBox1.Text
        End If

        objExcel.Visible = True

    End Sub

    'Dim objWorkbook As Excel.Workbook = CType(objExcel.ActiveWorkbook, Excel.Workbook)
    'Dim objWorksheet As Excel.Worksheet = CType(objExcel.ActiveSheet, Excel.Worksheet)

End Class