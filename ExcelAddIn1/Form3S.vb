Public Class Form3S
    Private Sub Form3S_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoro.Focus()
    End Sub

    Private Sub Button3S_Click(sender As Object, e As EventArgs) Handles Button3S.Click
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet1 As Excel.Worksheet = Nothing
        Dim activeWorksheet2 As Excel.Worksheet = Nothing
        Dim Range1 As String
        Dim myribbon As New GeoXRibbon

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet2 = activeWorkbook.Sheets.Add
        activeWorksheet2.Name = "Input Data 1"
        activeWorksheet1 = activeWorkbook.Sheets("Sheet1")

        activeWorksheet2.Range("A:A").Clear()
        activeWorksheet2.Range("B:B").Clear()
        activeWorksheet2.Range("C:C").Clear()
        activeWorksheet2.Range("D:D").Clear()

        Range1 = TextBoro.Text + ":" + TextBoro.Text
        activeWorksheet2.Range("A:A").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextOnStreet.Text + ":" + TextOnStreet.Text
        activeWorksheet2.Range("B:B").Value = activeWorksheet1.Range(Range1).Value
        If TextCompassDirection1.Text = String.Empty Then
        Else
            Range1 = TextCompassDirection1.Text + ":" + TextCompassDirection1.Text
            activeWorksheet2.Range("C:C").Value = activeWorksheet1.Range(Range1).Value
        End If
        If TextFirstCrossSt.Text = String.Empty Then
        Else
            Range1 = TextFirstCrossSt.Text + ":" + TextFirstCrossSt.Text
            activeWorksheet2.Range("D:D").Value = activeWorksheet1.Range(Range1).Value
        End If
        If TextCompassDirection1.Text = String.Empty Then
        Else
            Range1 = TextCompassDirection2.Text + ":" + TextCompassDirection2.Text
            activeWorksheet2.Range("E:E").Value = activeWorksheet1.Range(Range1).Value
        End If
        If TextSecondCrossSt.Text = String.Empty Then
        Else
            Range1 = TextSecondCrossSt.Text + ":" + TextSecondCrossSt.Text
            activeWorksheet2.Range("F:F").Value = activeWorksheet1.Range(Range1).Value
        End If
        If HasHeadersCheckboxFunction3S.Checked Then
            activeWorksheet2.Cells(1, 1).entirerow.delete
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "On Street"
            activeWorksheet2.Range("C1").Value = "Compass Direction 1"
            activeWorksheet2.Range("D1").Value = "First Cross Street"
            activeWorksheet2.Range("C1").Value = "Compass Direction 2"
            activeWorksheet2.Range("D1").Value = "Second Cross Street"
        Else
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "On Street"
            activeWorksheet2.Range("C1").Value = "Compass Direction 1"
            activeWorksheet2.Range("D1").Value = "First Cross Street"
            activeWorksheet2.Range("C1").Value = "Compass Direction 2"
            activeWorksheet2.Range("D1").Value = "Second Cross Street"
        End If

        Me.Visible = False
        myribbon.Process3SCall()
        activeWorksheet2.Name = "Input Data 1"
        activeWorksheet2.Visible = False
    End Sub




End Class