﻿
Public Class Form1B
    Public Shared Unit_Flag As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1B.Click
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet1 As Excel.Worksheet = Nothing
        Dim activeWorksheet2 As Excel.Worksheet = Nothing
        Dim Range1 As String
        Dim myribbon As New GeoXRibbon

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet2 = activeWorkbook.Sheets.Add
        activeWorksheet2.Name = "Input Data"
        activeWorksheet1 = activeWorkbook.Sheets("Sheet1")

        activeWorksheet2.Range("A:A").Clear()
        activeWorksheet2.Range("B:B").Clear()
        activeWorksheet2.Range("C:C").Clear()
        activeWorksheet2.Range("D:D").Clear()

        Range1 = TextBoro.Text + ":" + TextBoro.Text
        activeWorksheet2.Range("A:A").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextAddrNo.Text + ":" + TextAddrNo.Text
        activeWorksheet2.Range("B:B").NumberFormat = "@"
        activeWorksheet2.Range("B:B").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextStName.Text + ":" + TextStName.Text
        activeWorksheet2.Range("C:C").Value = activeWorksheet1.Range(Range1).Value

        If Unit_Flag = True Then
            Range1 = TextBox_Unit.Text + ":" + TextBox_Unit.Text
            activeWorksheet2.Range("D:D").Value = activeWorksheet1.Range(Range1).Value
        End If

        If HasHeadersCheckboxFunction1B.Checked Then
            activeWorksheet2.Cells(1, 1).entirerow.delete
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "Address Number"
            activeWorksheet2.Range("C1").Value = "Street or Place Name"
            activeWorksheet2.Range("D1").Value = "Unit No."
        Else
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "Address Number"
            activeWorksheet2.Range("C1").Value = "Street or Place Name"
            activeWorksheet2.Range("D1").Value = "Unit No."
        End If

        Me.Visible = False
        myribbon.Process1BCall()
        activeWorksheet2.Name = "Input Data"
        activeWorksheet2.Visible = False
    End Sub

    Private Sub Form1B_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoro.Focus()
        If Unit_Flag = True Then
            TextBox_Unit.Enabled = True
        Else
            TextBox_Unit.Enabled = False
        End If
    End Sub

End Class