Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBoro.Focus()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
        'here
        activeWorksheet2.Range("E:E").Clear()

        Range1 = TextBoro.Text + ":" + TextBoro.Text
        activeWorksheet2.Range("A:A").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextOnStreet.Text + ":" + TextFirstCrossSt.Text
        activeWorksheet2.Range("B:B").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextFirstCrossSt.Text + ":" + TextFirstCrossSt.Text
        activeWorksheet2.Range("C:C").Value = activeWorksheet1.Range(Range1).Value

        Range1 = TextSecondCrossSt.Text + ":" + TextSecondCrossSt.Text
        activeWorksheet2.Range("D:D").Value = activeWorksheet1.Range(Range1).Value


        'here
        If String.IsNullOrEmpty(TextSideOfStreet.Text) Then
            activeWorksheet2.Range("E:E").Value = String.Empty
        Else
            Range1 = TextSideOfStreet.Text + ":" + TextSideOfStreet.Text
            activeWorksheet2.Range("E:E").Value = activeWorksheet1.Range(Range1).Value
        End If



        If HasHeadersCheckboxFunction3.Checked Then
            activeWorksheet2.Cells(1, 1).entirerow.delete
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "On Street"
            activeWorksheet2.Range("C1").Value = "First Cross Street"
            activeWorksheet2.Range("D1").Value = "Second Cross Street"
            'here
            activeWorksheet2.Range("E1").Value = "Side of the Street"
        Else
            activeWorksheet2.Cells(1, 1).entirerow.insert
            activeWorksheet2.Range("A1").Value = "Select a Borough"
            activeWorksheet2.Range("B1").Value = "On Street"
            activeWorksheet2.Range("C1").Value = "First Cross Street"
            activeWorksheet2.Range("D1").Value = "Second Cross Street"
            'here
            activeWorksheet2.Range("E1").Value = "Side of the Street"
        End If
        Me.Visible = False
        myribbon.Process3Call()
        activeWorksheet2.Name = "Input Data"
        activeWorksheet2.Visible = False
    End Sub

    Private Sub TextSecondCrossSt_TextChanged(sender As Object, e As EventArgs) Handles TextSecondCrossSt.TextChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles HasHeadersCheckboxFunction3.CheckedChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
End Class