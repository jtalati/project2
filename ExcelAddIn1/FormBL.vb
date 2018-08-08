Public Class FormBL
    Private Sub FormBL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoro.Focus()

    End Sub
    Private Sub ButtonBL_Click(sender As Object, e As EventArgs) Handles ButtonBL.Click
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

        Range1 = TextBoro.Text + ":" + TextBoro.Text
        activeWorksheet2.Range("A:A").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextBlock.Text + ":" + TextBlock.Text
        activeWorksheet2.Range("B:B").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextLot.Text + ":" + TextLot.Text
        activeWorksheet2.Range("C:C").Value = activeWorksheet1.Range(Range1).Value

        If HasHeadersCheckBoxFunctionBL.Checked Then
            activeWorksheet2.Cells(1, 1).entirerow.delete
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "Block "
            activeWorksheet2.Range("C1").Value = "Lot  "
        Else
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "Block "
            activeWorksheet2.Range("C1").Value = "Lot  "
        End If

        Me.Visible = False
        myribbon.ProcessBLCall()
        activeWorksheet2.Name = "Input Data"
        activeWorksheet2.Visible = False
    End Sub


End Class
