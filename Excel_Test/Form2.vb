﻿Option Strict On
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

        If objWorkbook Is Nothing Then
            MessageBox.Show("No object for Wokrbook", "Test", MessageBoxButtons.OK)
        End If

        'Dim rangeTest As Excel.Range
        'rangeTest = objWorksheet.Range("B1")
        'rangeTest.Value = TextBox1.Text
    End Sub

    'Dim objWorkbook As Excel.Workbook = CType(objExcel.ActiveWorkbook, Excel.Workbook)
    'Dim objWorksheet As Excel.Worksheet = CType(objExcel.ActiveSheet, Excel.Worksheet)

End Class