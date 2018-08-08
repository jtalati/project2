Public Class FormBN
    Private Sub FormBN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBIN.Focus()
    End Sub

    Private Sub ButtonBL_Click(sender As Object, e As EventArgs) Handles ButtonBN.Click
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

        Range1 = TextBIN.Text + ":" + TextBIN.Text
        activeWorksheet2.Range("A:A").Value = activeWorksheet1.Range(Range1).Value

        If HasHeadersCheckBoxFunctionBN.Checked Then
            activeWorksheet2.Cells(1, 1).entirerow.delete
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "BIN"
        Else
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "BIN"
        End If


        Me.Visible = False
        myribbon.ProcessBNCall()
        activeWorksheet2.Name = "Input Data"
        activeWorksheet2.Visible = False
    End Sub

    Public Sub FormBL_Close(sender As Object, e As EventArgs) Handles MyBase.Closed


    End Sub

End Class