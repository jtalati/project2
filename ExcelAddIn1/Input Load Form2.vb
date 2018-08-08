Public Class Input_Load_Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If RadioButton1.Checked Then
            Dim Form2 As Object
            Me.Hide()
            Form2 = New Form2
            Form2.Show()

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
            activeWorksheet.Range("B1", "B100").Clear()
            activeWorksheet.Range("C1", "C100").Clear()
            activeWorksheet.Range("D1", "D100").Clear()
            activeWorksheet.Range("E1", "E100").Clear()

            activeWorksheet.Range("A1").Value = "Select a Borough"
            activeWorksheet.Range("B1").Value = "First Cross Street"
            activeWorksheet.Range("C1").Value = "Select a Borough"
            activeWorksheet.Range("D1").Value = "Second Cross Street"
            activeWorksheet.Range("E1").Value = "Compass Direction"

            activeWorksheet.Range("A1").Font.Bold = True
            activeWorksheet.Range("B1").Font.Bold = True
            activeWorksheet.Range("C1").Font.Bold = True
            activeWorksheet.Range("D1").Font.Bold = True
            activeWorksheet.Range("E1").Font.Bold = True

            activeWorksheet.Range("A1:E1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:E1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()

            activeWorksheet.Range("A1:A2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            activeWorksheet.Range("A1:A2").Borders.Weight = 2.0
            activeWorksheet.Range("B1:B2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            activeWorksheet.Range("B1:B2").Borders.Weight = 2.0
            activeWorksheet.Range("C1:C2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            activeWorksheet.Range("C1:C2").Borders.Weight = 2.0
            activeWorksheet.Range("D1:D2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            activeWorksheet.Range("D1:D2").Borders.Weight = 2.0
            activeWorksheet.Range("E1:E2").Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            activeWorksheet.Range("E1:E2").Borders.Weight = 2.0


            With activeWorksheet.Range("A2", "A2").Validation
                .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, AlertStyle:=Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween, Formula1:="Manhattan, Bronx, Brooklyn, Queens, Staten Island")
            End With

            With activeWorksheet.Range("C2", "C2").Validation
                .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, AlertStyle:=Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween, Formula1:="Manhattan, Bronx, Brooklyn, Queens, Staten Island")
            End With

            With activeWorksheet.Range("E2", "E2").Validation
                .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, AlertStyle:=Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween, Formula1:="N, E, S, W, N/A")
            End With
        End If

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

End Class