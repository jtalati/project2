Public Class Input_Load_FormBN
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If RadioButton1.Checked Then
            Dim FormBN As Object

            Me.Hide()
            FormBN = New FormBN
            FormBN.Show()

        End If

        If RadioButton2.Checked Then
            Me.Hide()
            Dim activeExcel As Excel.Application
            Dim activeWorkbook As Excel.Workbook = Nothing
            Dim activeWorksheet As Excel.Worksheet = Nothing

            Dim AppXL As Object
            AppXL = CreateObject("Word.Application")
            AppXL.Visible = True
            AppXL.Activate

tryagain:
            Try
                activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
            Catch ex As Exception
                MsgBox("ROT Issue")
                MsgBox(ex.Message)
            End Try

            activeExcel.Visible = True
            AppXL.Visible = False
            activeWorkbook = activeExcel.ActiveWorkbook
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Input Data"

            activeWorksheet.Range("A1", "A100").Clear()

            activeWorksheet.Range("A1").Value = "BIN Number"

            activeWorksheet.Range("A1").Font.Bold = True

            activeWorksheet.Range("A1:A1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:A1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()

            activeWorksheet.Range("A1:A2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            activeWorksheet.Range("A1:A2").Borders.Weight = 2.0
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub
End Class