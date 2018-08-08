Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoro1.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
        activeWorksheet2.Range("E:E").Clear()

        Range1 = TextBoro1.Text + ":" + TextBoro1.Text
        activeWorksheet2.Range("A:A").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextAddrNo.Text + ":" + TextAddrNo.Text
        activeWorksheet2.Range("B:B").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextStName.Text + ":" + TextStName.Text
        activeWorksheet2.Range("D:D").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextBoro2.Text + ":" + TextBoro2.Text
        activeWorksheet2.Range("C:C").Value = activeWorksheet1.Range(Range1).Value

        If TextCompass.Text = vbNullString Then
        Else
            Range1 = TextCompass.Text + ":" + TextCompass.Text
            activeWorksheet2.Range("E:E").Value = activeWorksheet1.Range(Range1).Value
        End If

        If HasHeadersCheckboxFunction2.Checked Then
            activeWorksheet2.Cells(1, 1).entirerow.delete
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "First Cross Street"
            activeWorksheet2.Range("C1").Value = "Select a Borough"
            activeWorksheet2.Range("D1").Value = "Second Cross Street"
            activeWorksheet2.Range("E1").Value = "Compass Direction"
        Else
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "First Cross Street"
            activeWorksheet2.Range("C1").Value = "Select a Borough"
            activeWorksheet2.Range("D1").Value = "Second Cross Street"
            activeWorksheet2.Range("E1").Value = "Compass Direction"
        End If

        Me.Visible = False
        myribbon.Process2Call()
        activeWorksheet2.Name = "Input Data"
        activeWorksheet2.Visible = False
    End Sub

    Private Sub TextStName_TextChanged(sender As Object, e As EventArgs) Handles TextStName.TextChanged

    End Sub
End Class