Imports System
Imports System.Text
Imports System.Xml
Imports System.IO
Imports DCP.Geosupport.DotNet.GeoX
Imports DCP.Geosupport.DotNet.fld_def_lib
Imports System.Configuration
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Tools.Ribbon
Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml.Linq
Imports System.Windows.Forms
Imports Excel_Int = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports Office = Microsoft.Office.Tools.Excel
Imports System.Data
Imports System.Threading
Imports System.Diagnostics
Imports Microsoft.Win32

'Release 16.2 Change - Jigar Talati 
'Separated X/Y coordinates, From/To Node, Latitude/Longitude, From X/Y Coordinates, To X/Y Coordinates, etc columns into separate columns where needed for all Functions
'All columns were shifted one cell to the right
'Changed all BOROUGH inputs to convert all types of strings as an upper case for all functions.


Public Class GeoXRibbon
    Public Statusflag As String



    Private Sub Process_Click(sender As Object, e As RibbonControlEventArgs) Handles Process.Click

        If Func1B.Checked = True Then
            Process1BCall()
        ElseIf Func2.Checked = True Then
            Process2Call()
        ElseIf Func3.Checked = True Then
            Process3Call()
        ElseIf Func3S.Checked = True Then
            Process3SCall()
        ElseIf FuncBL.Checked = True Then
            ProcessBLCall()
        ElseIf FuncBN.Checked = True Then
            ProcessBNCall()
        End If

    End Sub
    Public Function Process1BCall() As Integer

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim fdwa1 As New Wa1
        Dim fdwa2f1b As New Wa2F1b
        Dim i, N As Integer
        Dim TempStr As String

        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        fdgeo = New geo(fdconns)

        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Dim Tempborough As String
        Dim Boroughflag As Boolean = False
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet = activeWorkbook.ActiveSheet
        activeWorksheet = activeWorkbook.Sheets("Input Data")

        fdwa1.Clear()
        fdwa2f1b.Clear()
        fdwa1.in_func_code = "1B"
        fdwa1.in_platform_ind = "C"
        fdwa1.in_tpad_switch = "Y"
        'fdwa1.in_mode_switch = "X"

        N = activeWorksheet.UsedRange.Rows.Count

        For i = 2 To N

            Boroughflag = False
            activeWorksheet = activeWorkbook.Sheets("Input Data")
            If activeWorksheet.Cells(i, 1).Value = vbNullString Then
                fdwa1.in_b10sc1.boro = " "
            Else
                Tempborough = activeWorksheet.Cells(i, 1).Value.ToString()
                Boroughflag = True
            End If


            If Boroughflag Then

                If Tempborough = "1" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf Tempborough = "2" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf Tempborough = "3" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf Tempborough = "4" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf Tempborough = "5" Then
                    fdwa1.in_b10sc1.boro = "5"
                ElseIf Tempborough.ToUpper() = "MANHATTAN" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf Tempborough.ToUpper() = "BRONX" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf Tempborough.ToUpper() = "BROOKLYN" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf Tempborough.ToUpper() = "QUEENS" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf Tempborough.ToUpper() = "STATEN ISLAND" Then
                    fdwa1.in_b10sc1.boro = "5"
                End If

            End If

            If activeWorksheet.Cells(i, 2).Value = vbNullString Then
                fdwa1.in_hnd = " "
            Else
                fdwa1.in_hnd = activeWorksheet.Cells(i, 2).Value
            End If

            If activeWorksheet.Cells(i, 3).Value = vbNullString Then
                fdwa1.in_stname1 = " "
            Else
                fdwa1.in_stname1 = Trim(activeWorksheet.Cells(i, 3).Value)
            End If

            If activeWorksheet.Cells(i, 4).Value = vbNullString Then
                fdwa1.in_unit = " "
            Else
                fdwa1.in_unit = Trim(activeWorksheet.Cells(i, 4).Value)
            End If

            TempStr = fdwa1.in_b10sc1.boro + " " + fdwa1.in_hnd + " " + fdwa1.in_stname1

            Try
                Call fdgeo.GeoCall(fdwa1, fdwa2f1b)
            Catch ex As Exception
                Return 1
                MsgBox("Error Occured at " + TempStr)
            End Try

            Call WriteData1B(fdwa1, fdwa2f1b, i)
        Next i

        Return 0
    End Function
    Public Function Process2Call() As Integer

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim fdwa1 As New Wa1
        Dim fdwa2f2 As New Wa2F2
        Dim i, N As Integer
        Dim Tempborough1 As String
        Dim Boroughflag1 As Boolean = False
        Dim Tempborough2 As String
        Dim Boroughflag2 As Boolean = False


        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        fdgeo = New geo(fdconns)
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet = activeWorkbook.ActiveSheet
        activeWorksheet = activeWorkbook.Sheets("Input Data")

        fdwa1.Clear()
        fdwa2f2.Clear()
        fdwa1.in_func_code = " 2"
        fdwa1.in_platform_ind = "C"
        fdwa1.in_xstreet_names_flag = "E"

        N = activeWorksheet.UsedRange.Rows.Count

        For i = 2 To N
            Boroughflag1 = False
            Boroughflag2 = False
            activeWorksheet = activeWorkbook.Sheets("Input Data")

            If activeWorksheet.Cells(i, 1).Value = vbNullString Then
                fdwa1.in_boro1 = " "
            Else
                Tempborough1 = activeWorksheet.Cells(i, 1).Value.ToString()
                Boroughflag1 = True
            End If


            If Boroughflag1 Then

                If Tempborough1 = "1" Then
                    fdwa1.in_boro1 = "1"
                ElseIf Tempborough1 = "2" Then
                    fdwa1.in_boro1 = "2"
                ElseIf Tempborough1 = "3" Then
                    fdwa1.in_boro1 = "3"
                ElseIf Tempborough1 = "4" Then
                    fdwa1.in_boro1 = "4"
                ElseIf Tempborough1 = "5" Then
                    fdwa1.in_boro1 = "5"
                ElseIf Tempborough1.ToUpper() = "MANHATTAN" Then
                    fdwa1.in_boro1 = "1"
                ElseIf Tempborough1.ToUpper() = "BRONX" Then
                    fdwa1.in_boro1 = "2"
                ElseIf Tempborough1.ToUpper() = "BROOKLYN" Then
                    fdwa1.in_boro1 = "3"
                ElseIf Tempborough1.ToUpper() = "QUEENS" Then
                    fdwa1.in_boro1 = "4"
                ElseIf Tempborough1.ToUpper() = "STATEN ISLAND" Then
                    fdwa1.in_boro1 = "5"
                End If

            End If

            If activeWorksheet.Cells(i, 3).Value = vbNullString Then
                fdwa1.in_boro2 = " "
            Else
                Tempborough2 = activeWorksheet.Cells(i, 3).Value.ToString()
                Boroughflag2 = True
            End If


            If Boroughflag2 Then

                If Tempborough2 = "1" Then
                    fdwa1.in_boro2 = "1"
                ElseIf Tempborough2 = "2" Then
                    fdwa1.in_boro2 = "2"
                ElseIf Tempborough2 = "3" Then
                    fdwa1.in_boro2 = "3"
                ElseIf Tempborough2 = "4" Then
                    fdwa1.in_boro2 = "4"
                ElseIf Tempborough2 = "5" Then
                    fdwa1.in_boro2 = "5"
                ElseIf Tempborough2.ToUpper() = "MANHATTAN" Then
                    fdwa1.in_boro2 = "1"
                ElseIf Tempborough2.ToUpper() = "BRONX" Then
                    fdwa1.in_boro2 = "2"
                ElseIf Tempborough2.ToUpper() = "BROOKLYN" Then
                    fdwa1.in_boro2 = "3"
                ElseIf Tempborough2.ToUpper() = "QUEENS" Then
                    fdwa1.in_boro2 = "4"
                ElseIf Tempborough2.ToUpper() = "STATEN ISLAND" Then
                    fdwa1.in_boro2 = "5"
                End If

            End If


            If activeWorksheet.Cells(i, 2).Value = vbNullString Then
                fdwa1.in_stname1 = " "
            Else
                fdwa1.in_stname1 = activeWorksheet.Cells(i, 2).Value
            End If

            If activeWorksheet.Cells(i, 4).Value = vbNullString Then
                fdwa1.in_stname2 = " "
            Else
                fdwa1.in_stname2 = activeWorksheet.Cells(i, 4).Value
            End If

            If activeWorksheet.Cells(i, 5).Value = vbNullString Then
                fdwa1.in_compass_dir = " "
            Else
                fdwa1.in_compass_dir = activeWorksheet.Cells(i, 5).Value
            End If


            Call fdgeo.GeoCall(fdwa1, fdwa2f2)
            Call WriteData2(fdwa1, fdwa2f2, i)
        Next i

        Return 0
    End Function
    Public Function Process3Call() As Integer

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim fdwa1 As New Wa1
        Dim fdwa2f3 As New Wa2F3xas
        Dim fdwa2f3C As New Wa2F3cxas
        Dim i, N As Integer
        Dim Tempborough3 As String
        Dim Boroughflag3 As Boolean = False

        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        fdgeo = New geo(fdconns)
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet = activeWorkbook.ActiveSheet
        activeWorksheet = activeWorkbook.Sheets("Input Data")

        fdwa1.Clear()
        fdwa2f3.Clear()
        fdwa2f3C.Clear()
        fdwa1.in_func_code = " 3"
        fdwa1.in_platform_ind = "C"
        fdwa1.in_xstreet_names_flag = "E"
        fdwa1.in_auxseg_switch = "Y"
        fdwa1.in_mode_switch = "X"

        N = activeWorksheet.UsedRange.Rows.Count
        For i = 2 To N
            Boroughflag3 = False
            activeWorksheet = activeWorkbook.Sheets("Input Data")

            If activeWorksheet.Cells(i, 1).Value = vbNullString Then
                fdwa1.in_boro1 = " "
            Else
                Tempborough3 = activeWorksheet.Cells(i, 1).Value.ToString()
                Boroughflag3 = True
            End If

            If Boroughflag3 Then

                If Tempborough3 = "1" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf Tempborough3 = "2" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf Tempborough3 = "3" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf Tempborough3 = "4" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf Tempborough3 = "5" Then
                    fdwa1.in_b10sc1.boro = "5"
                ElseIf Tempborough3.ToUpper() = "MANHATTAN" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf Tempborough3.ToUpper() = "BRONX" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf Tempborough3.ToUpper() = "BROOKLYN" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf Tempborough3.ToUpper() = "QUEENS" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf Tempborough3.ToUpper() = "STATEN ISLAND" Then
                    fdwa1.in_b10sc1.boro = "5"
                End If

            End If
            'For i = 2 To N
            '    activeWorksheet = activeWorkbook.Sheets("Input Data")
            '    If activeWorksheet.Cells(i, 1).Value = vbNullString Then
            '        fdwa1.in_b10sc1.boro = " "
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Manhattan" Then
            '        fdwa1.in_b10sc1.boro = "1"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Bronx" Then
            '        fdwa1.in_b10sc1.boro = "2"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Brooklyn" Then
            '        fdwa1.in_b10sc1.boro = "3"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Queens" Then
            '        fdwa1.in_b10sc1.boro = "4"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Staten Island" Then
            '        fdwa1.in_b10sc1.boro = "5"
            '    End If

            If activeWorksheet.Cells(i, 2).Value = vbNullString Then
                fdwa1.in_stname1 = " "
            Else
                fdwa1.in_stname1 = activeWorksheet.Cells(i, 2).Value
            End If

            If activeWorksheet.Cells(i, 3).Value = vbNullString Then
                fdwa1.in_stname2 = " "
            Else
                fdwa1.in_stname2 = activeWorksheet.Cells(i, 3).Value
            End If

            If activeWorksheet.Cells(i, 4).Value = vbNullString Then
                fdwa1.in_stname3 = " "
            Else
                fdwa1.in_stname3 = activeWorksheet.Cells(i, 4).Value
            End If
            'here
            If activeWorksheet.Cells(i, 5).Value = vbNullString Then
                fdwa1.in_compass_dir = String.Empty
            Else
                fdwa1.in_compass_dir = String.Empty
                fdwa1.in_compass_dir = activeWorksheet.Cells(i, 5).Value
                fdwa1.in_func_code = "3C"
            End If

            If Not activeWorksheet.Cells(i, 5).Value = vbNullString Then
                Call fdgeo.GeoCall(fdwa1, fdwa2f3C)
                Call WriteData3C(fdwa1, fdwa2f3C, i)
            Else
                Call fdgeo.GeoCall(fdwa1, fdwa2f3)
                Call WriteData3(fdwa1, fdwa2f3, i)
            End If
        Next i

        Return 0
    End Function
    Public Function Process3SCall() As Integer

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim fdwa1 As New Wa1
        Dim fdwa2f3s As New Wa2F3s
        Dim i, N As Integer
        Dim Tempborough3S As String
        Dim Boroughflag3S As Boolean = False

        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        'fdconns = New GeoConnCollection("C:\temp\GeoConns.xml")
        fdgeo = New geo(fdconns)
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet = activeWorkbook.ActiveSheet
        activeWorksheet = activeWorkbook.Sheets("Input Data")

        fdwa1.Clear()
        fdwa2f3s.Clear()
        fdwa1.in_func_code = "3S"
        fdwa1.in_platform_ind = "C"
        fdwa1.in_real_street_only = "R"

        N = activeWorksheet.UsedRange.Rows.Count
        For i = 2 To N
            Boroughflag3S = False
            activeWorksheet = activeWorkbook.Sheets("Input Data")
            fdwa1.Clear()
            fdwa2f3s.Clear()
            fdwa1.in_func_code = "3S"
            fdwa1.in_platform_ind = "C"
            fdwa1.in_real_street_only = "R"

            If activeWorksheet.Cells(i, 1).Value = vbNullString Then
                fdwa1.in_boro1 = " "
            Else
                Tempborough3S = activeWorksheet.Cells(i, 1).Value.ToString()
                Boroughflag3S = True
            End If

            If Boroughflag3S Then

                If Tempborough3S = "1" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf Tempborough3S = "2" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf Tempborough3S = "3" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf Tempborough3S = "4" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf Tempborough3S = "5" Then
                    fdwa1.in_b10sc1.boro = "5"
                ElseIf Tempborough3S.ToUpper() = "MANHATTAN" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf Tempborough3S.ToUpper() = "BRONX" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf Tempborough3S.ToUpper() = "BROOKLYN" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf Tempborough3S.ToUpper() = "QUEENS" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf Tempborough3S.ToUpper() = "STATEN ISLAND" Then
                    fdwa1.in_b10sc1.boro = "5"
                End If

            End If

            'For i = 2 To N
            '    activeWorksheet = activeWorkbook.Sheets("Input Data")
            '    If activeWorksheet.Cells(i, 1).Value = vbNullString Then
            '        fdwa1.in_b10sc1.boro = " "
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Manhattan" Then
            '        fdwa1.in_b10sc1.boro = "1"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Bronx" Then
            '        fdwa1.in_b10sc1.boro = "2"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Brooklyn" Then
            '        fdwa1.in_b10sc1.boro = "3"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Queens" Then
            '        fdwa1.in_b10sc1.boro = "4"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Staten Island" Then
            '        fdwa1.in_b10sc1.boro = "5"
            '    End If

            If activeWorksheet.Cells(i, 2).Value = vbNullString Then
                fdwa1.in_stname1 = String.Empty
            Else
                fdwa1.in_stname1 = activeWorksheet.Cells(i, 2).Value
            End If
            'here for com 1
            If activeWorksheet.Cells(i, 3).Value = vbNullString Then
                fdwa1.in_compass_dir = " "
            Else
                fdwa1.in_compass_dir = activeWorksheet.Cells(i, 3).Value
            End If

            If activeWorksheet.Cells(i, 4).Value = vbNullString Then
                fdwa1.in_stname2 = String.Empty
            Else
                fdwa1.in_stname2 = activeWorksheet.Cells(i, 4).Value
            End If

            'here for com 2
            If activeWorksheet.Cells(i, 5).Value = vbNullString Then
                fdwa1.in_compass_dir2 = " "
            Else
                fdwa1.in_compass_dir2 = activeWorksheet.Cells(i, 5).Value
            End If

            If activeWorksheet.Cells(i, 6).Value = vbNullString Then
                fdwa1.in_stname3 = String.Empty
            Else
                fdwa1.in_stname3 = activeWorksheet.Cells(i, 6).Value
            End If


            Call fdgeo.GeoCall(fdwa1, fdwa2f3s)
            Call WriteData3S(fdwa1, fdwa2f3s, i)
        Next i

        Return 0
    End Function
    Public Function ProcessBLCall() As Integer

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim fdwa1 As New Wa1
        Dim fdwa2f1a As New Wa2F1a
        Dim i, N As Integer
        Dim TempboroughBL As String
        Dim BoroughflagBL As Boolean = False

        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        fdgeo = New geo(fdconns)
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet = activeWorkbook.ActiveSheet
        activeWorksheet = activeWorkbook.Sheets("Input Data")

        fdwa1.Clear()
        fdwa2f1a.Clear()
        fdwa1.in_func_code = "BL"
        fdwa1.in_platform_ind = "C"
        fdwa1.in_tpad_switch = "Y"
        fdwa1.in_mode_switch = "X"

        N = activeWorksheet.UsedRange.Rows.Count
        For i = 2 To N
            BoroughflagBL = False
            activeWorksheet = activeWorkbook.Sheets("Input Data")

            If activeWorksheet.Cells(i, 1).Value = vbNullString Then
                fdwa1.in_boro1 = " "
            Else
                TempboroughBL = activeWorksheet.Cells(i, 1).Value.ToString()
                BoroughflagBL = True
            End If

            If BoroughflagBL Then

                If TempboroughBL = "1" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf TempboroughBL = "2" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf TempboroughBL = "3" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf TempboroughBL = "4" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf TempboroughBL = "5" Then
                    fdwa1.in_b10sc1.boro = "5"
                ElseIf TempboroughBL.ToUpper() = "MANHATTAN" Then
                    fdwa1.in_b10sc1.boro = "1"
                ElseIf TempboroughBL.ToUpper() = "BRONX" Then
                    fdwa1.in_b10sc1.boro = "2"
                ElseIf TempboroughBL.ToUpper() = "BROOKLYN" Then
                    fdwa1.in_b10sc1.boro = "3"
                ElseIf TempboroughBL.ToUpper() = "QUEENS" Then
                    fdwa1.in_b10sc1.boro = "4"
                ElseIf TempboroughBL.ToUpper() = "STATEN ISLAND" Then
                    fdwa1.in_b10sc1.boro = "5"
                End If

            End If
            'For i = 2 To N
            '    activeWorksheet = activeWorkbook.Sheets("Input Data")
            '    If activeWorksheet.Cells(i, 1).Value = vbNullString Then
            '        fdwa1.in_b10sc1.boro = " "
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Manhattan" Then
            '        fdwa1.in_b10sc1.boro = "1"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Bronx" Then
            '        fdwa1.in_b10sc1.boro = "2"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Brooklyn" Then
            '        fdwa1.in_b10sc1.boro = "3"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Queens" Then
            '        fdwa1.in_b10sc1.boro = "4"
            '    ElseIf activeWorksheet.Cells(i, 1).Value = "Staten Island" Then
            '        fdwa1.in_b10sc1.boro = "5"
            '    End If

            If activeWorksheet.Cells(i, 2).Value = vbNullString Then
                fdwa1.in_bbl.block = " "
            Else
                fdwa1.in_bbl.block = activeWorksheet.Cells(i, 2).Value
            End If

            If activeWorksheet.Cells(i, 3).Value = vbNullString Then
                fdwa1.in_bbl.lot = " "
            Else
                fdwa1.in_bbl.lot = activeWorksheet.Cells(i, 3).Value
            End If

            Call fdgeo.GeoCall(fdwa1, fdwa2f1a)
            Call WriteDataBL(fdwa1, fdwa2f1a, i)
        Next i

        Return 0
    End Function
    Public Function ProcessBNCall() As Integer

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim fdwa1 As New Wa1
        Dim fdwa2f1al As New Wa2F1ax
        Dim i, N As Integer

        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        fdgeo = New geo(fdconns)
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        activeWorksheet = activeWorkbook.ActiveSheet
        activeWorksheet = activeWorkbook.Sheets("Input Data")

        N = activeWorksheet.UsedRange.Rows.Count

        fdwa1.Clear()
        fdwa2f1al.Clear()
        fdwa1.in_func_code = "BN"
        fdwa1.in_platform_ind = "C"
        fdwa1.in_tpad_switch = "Y"
        fdwa1.in_mode_switch = "X"

        For i = 2 To N
            activeWorksheet = activeWorkbook.Sheets("Input Data")
            If activeWorksheet.Cells(i, 1).Value = vbNullString Then
                fdwa1.in_bin_string = " "
            Else
                fdwa1.in_bin_string = activeWorksheet.Cells(i, 1).Value
            End If

            Call fdgeo.GeoCall(fdwa1, fdwa2f1al)
            Call WriteDataBN(fdwa1, fdwa2f1al, i)
        Next i

        Return 0
    End Function
    Public Function WriteData1B(ByRef fdwa1 As Wa1, ByRef fdwa2f1b As Wa2F1b, i As Integer) As Integer
        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 2
        Static Dim k As Integer = 1
        Dim xlCenter As Long = -4108
        Dim gotw_fld_dict As New fld_dict

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-1B Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "Borough"
            activeWorksheet.Cells(1, 4).Value = "Address Number"
            activeWorksheet.Cells(1, 5).Value = "Street Name"
            activeWorksheet.Cells(1, 6).Value = "Unit Number"
            activeWorksheet.Range("A1:F1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-1B Output"

            AddFlag = True
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("2:2").Font.Bold = True

            activeWorksheet.Cells(1, 2).value = "Input Data"
            activeWorksheet.Range("B1:D1").Merge()
            activeWorksheet.Range("B1:D1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 6).value = "Geographic Information"
            activeWorksheet.Range("E1:AP1").Merge()
            activeWorksheet.Range("E1:AP1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 43).value = "City Service Information"
            activeWorksheet.Range("AQ1:BG1").Merge()
            activeWorksheet.Range("AQ1:BG1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 60).value = "Political Information"
            activeWorksheet.Range("BH1:BN1").Merge()
            activeWorksheet.Range("BH1:BN1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 67).value = "Property Level Information"
            activeWorksheet.Range("BO1:CR1").Merge()
            activeWorksheet.Range("BO1:CR1").HorizontalAlignment = xlCenter

            activeWorksheet.Range("A1:CR1").Interior.Color = RGB(204, 204, 255)
            activeWorksheet.Range("A1:CR1").Font.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:CR2").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:CR2").Font.Color = RGB(255, 255, 255)

            activeWorksheet.Cells(2, 1).Value = "Return Code/Reason Code"
            activeWorksheet.Cells(2, 2).Value = "Borough"
            activeWorksheet.Cells(2, 3).Value = "Address Number"
            activeWorksheet.Cells(2, 3).NumberFormat = "@"
            activeWorksheet.Cells(2, 4).Value = "Street Name"
            activeWorksheet.Cells(2, 5).Value = "Unit Number"

            activeWorksheet.Cells(2, 6).Value = "X Coordinate"
            activeWorksheet.Cells(2, 7).Value = "Y Coordinate"
            activeWorksheet.Cells(2, 8).Value = "From Node"
            activeWorksheet.Cells(2, 9).Value = "To Node"
            activeWorksheet.Cells(2, 10).Value = "Latitude"
            activeWorksheet.Cells(2, 11).Value = "Longitude"
            activeWorksheet.Cells(2, 12).value = "From X Coordinate"
            activeWorksheet.Cells(2, 13).value = "From Y Coordinate"
            activeWorksheet.Cells(2, 14).value = "To X Coordinate"
            activeWorksheet.Cells(2, 15).value = "To Y Coordinate"
            activeWorksheet.Cells(2, 16).value = "Community District"
            activeWorksheet.Cells(2, 17).Value = "LION Facecode"
            activeWorksheet.Cells(2, 18).Value = "LION Sequence No"
            activeWorksheet.Cells(2, 19).Value = "Coincident Segment count"
            activeWorksheet.Cells(2, 20).Value = "Street Code B10SC"
            activeWorksheet.Cells(2, 21).Value = "Segment ID /Length"
            activeWorksheet.Cells(2, 22).Value = "Alley/Cross Street Flag"
            activeWorksheet.Cells(2, 23).Value = "Segment Type"
            activeWorksheet.Cells(2, 24).Value = "Traffic Direction"
            activeWorksheet.Cells(2, 25).Value = "Feature Type"
            activeWorksheet.Cells(2, 26).Value = "Roadway Type"
            activeWorksheet.Cells(2, 27).Value = "Right of Way Type"
            activeWorksheet.Cells(2, 28).Value = "Physical ID"
            activeWorksheet.Cells(2, 29).Value = "Generic ID"
            activeWorksheet.Cells(2, 30).Value = "Bike Lane"
            'njp (2017-01-04 - 17.1 changes to add Bike Traffic Direction)
            activeWorksheet.Cells(2, 31).Value = "Bike Traffic Direction"
            activeWorksheet.Cells(2, 32).Value = "Special Address"
            activeWorksheet.Cells(2, 33).Value = "Low House Number"
            activeWorksheet.Cells(2, 34).Value = "High House Number"
            activeWorksheet.Cells(2, 35).Value = "2010 Census Tract"
            activeWorksheet.Cells(2, 36).Value = "2010 Census Block"
            activeWorksheet.Cells(2, 37).Value = "2000 Census Tract"
            activeWorksheet.Cells(2, 38).Value = "2000 Census Block"
            activeWorksheet.Cells(2, 39).Value = "Street Width Min"
            activeWorksheet.Cells(2, 40).Value = "Street Width Max"
            activeWorksheet.Cells(2, 41).Value = "Street Width Irregular"
            'speed limit goes here
            activeWorksheet.Cells(2, 42).Value = "Speed Limit"


            activeWorksheet.Cells(2, 43).Value = "Police Patrol Borough"
            activeWorksheet.Cells(2, 44).Value = "Police Precinct"

            activeWorksheet.Cells(2, 45).Value = "Fire Division"
            activeWorksheet.Cells(2, 46).Value = "Fire Battalion"
            activeWorksheet.Cells(2, 47).Value = "Fire Company"
            activeWorksheet.Cells(2, 48).Value = "Health Area"
            activeWorksheet.Cells(2, 49).Value = "Health Center District"
            activeWorksheet.Cells(2, 50).Value = "DOT Street Light Area"
            activeWorksheet.Cells(2, 51).Value = "Sanitation District/Section"
            activeWorksheet.Cells(2, 52).Value = "Sanitation Subsection:"
            activeWorksheet.Cells(2, 53).Value = "Regular Sanitation Pickup"
            activeWorksheet.Cells(2, 54).Value = "Recycling Sanitation Pickup"
            activeWorksheet.Cells(2, 55).Value = "Organics Recylcing Pickup"
            activeWorksheet.Cells(2, 56).Value = "School District"
            activeWorksheet.Cells(2, 57).Value = "DSNY Snow Priority"
            activeWorksheet.Cells(2, 58).Value = "Sanitation Bulk Pickup"
            activeWorksheet.Cells(2, 59).Value = "Hurricane Zone"


            activeWorksheet.Cells(2, 60).Value = "City Council District"
            activeWorksheet.Cells(2, 61).Value = "Assembly District"
            activeWorksheet.Cells(2, 62).Value = "Congressional District"
            activeWorksheet.Cells(2, 63).Value = "Municipal Court District"
            activeWorksheet.Cells(2, 64).Value = "Election District"
            activeWorksheet.Cells(2, 65).Value = "State Senate District"
            activeWorksheet.Cells(2, 66).Value = "BOE Preferred B7SC/Street Name"

            activeWorksheet.Cells(2, 67).Value = "Tax Block"
            activeWorksheet.Cells(2, 68).Value = "Tax Lot"
            activeWorksheet.Cells(2, 69).Value = "BBL"
            activeWorksheet.Cells(2, 70).Value = "Block Faces"
            activeWorksheet.Cells(2, 71).Value = "Sanborn Boro/Vol/Page"
            activeWorksheet.Cells(2, 72).Value = "RPAD SCC"
            activeWorksheet.Cells(2, 73).Value = "RPAD Building Class"
            activeWorksheet.Cells(2, 74).Value = "RPAD Interior Lot"
            activeWorksheet.Cells(2, 75).Value = "RPAD Irreg. Shaped Lot"
            activeWorksheet.Cells(2, 76).Value = "RPAD Condo Number"
            activeWorksheet.Cells(2, 77).Value = "RPAD Co-op Number"
            activeWorksheet.Cells(2, 78).Value = "Vacant Lot"
            activeWorksheet.Cells(2, 79).Value = "Condo Lot"
            activeWorksheet.Cells(2, 80).Value = "Low BBL of Condo"
            activeWorksheet.Cells(2, 81).Value = "High BBL of Condo"
            activeWorksheet.Cells(2, 82).Value = "Tax Map/Section/Volume"
            activeWorksheet.Cells(2, 83).Value = "BIN"
            activeWorksheet.Cells(2, 84).Value = "BIN Status"
            activeWorksheet.Cells(2, 85).Value = "TPAD BIN"
            activeWorksheet.Cells(2, 86).Value = "TPAD BIN Status"
            activeWorksheet.Cells(2, 87).Value = "TPAD Conflict Flag"
            activeWorksheet.Cells(2, 88).Value = "Corner Code"
            activeWorksheet.Cells(2, 89).Value = "Business Improvement District"
            activeWorksheet.Cells(2, 90).Value = "X/Y Coordinates"
            activeWorksheet.Cells(2, 91).Value = "Blockface ID"
            activeWorksheet.Cells(2, 92).Value = "No. of Traveling Lanes"
            activeWorksheet.Cells(2, 93).Value = "No. of Parking Lanes"
            activeWorksheet.Cells(2, 94).Value = "Total No. of Lanes"
            'dcp zoning map goes here
            'Jigar Talati New Field Added for F1B
            'activeWorksheet.Cells(2, 94).Value = "Speed Limit"
            activeWorksheet.Cells(2, 95).Value = "DCP Zoning Map"
            activeWorksheet.Cells(2, 96).Value = "PUMA"
            'police sector
            'activeWorksheet.Cells(2, 97).Value = "Police Sector"
            'police service area
            'activeWorksheet.Cells(2, 98).Value = "Police Service Area"

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

        Else
            activeWorksheet = activeWorkbook.Sheets("Func-1B Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet = activeWorkbook.Sheets("Func-1B Output")
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "-" + fdwa2f1b.wa2f1ax.grc + "/" + fdwa1.out_reason_code + "-" + fdwa2f1b.wa2f1ax.reason_code
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 3).Value = "'" + fdwa1.in_hnd
            activeWorksheet.Cells(j, 4).Value = fdwa1.in_stname1
            activeWorksheet.Cells(j, 5).Value = fdwa1.out_unit

            activeWorksheet.Cells(j, 6).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.x_coord))
            activeWorksheet.Cells(j, 7).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.y_coord))
            activeWorksheet.Cells(j, 8).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.from_node))
            activeWorksheet.Cells(j, 9).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.to_node))
            activeWorksheet.Cells(j, 10).Value = fdwa2f1b.wa2f1ex.latitude
            activeWorksheet.Cells(j, 11).Value = fdwa2f1b.wa2f1ex.longitude
            activeWorksheet.Cells(j, 12).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.lo_x_coord))
            activeWorksheet.Cells(j, 13).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.lo_y_coord))
            activeWorksheet.Cells(j, 14).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.hi_x_coord))
            activeWorksheet.Cells(j, 15).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.hi_y_coord))
            activeWorksheet.Cells(j, 16).Value = fdwa2f1b.wa2f1ex.com_dist.boro + fdwa2f1b.wa2f1ex.com_dist.district_number
            activeWorksheet.Cells(j, 17).Value = fdwa2f1b.wa2f1ex.lion_key.face_code
            activeWorksheet.Cells(j, 18).Value = fdwa2f1b.wa2f1ex.lion_key.sequence_number
            activeWorksheet.Cells(j, 18).NumberFormat = "00000"
            activeWorksheet.Cells(j, 19).Value = fdwa2f1b.wa2f1ex.coincident_seg_cnt
            activeWorksheet.Cells(j, 20).Value = fdwa1.out_b10sc1.B10scToString()
            activeWorksheet.Cells(j, 21).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ex.segment_id)) + " / " + Convert.ToString(Val(fdwa2f1b.wa2f1ex.segment_len))
            If fdwa2f1b.wa2f1ex.alx = " " Then
                activeWorksheet.Cells(j, 22).Value = "None"
            Else
                activeWorksheet.Cells(j, 22).Value = fdwa2f1b.wa2f1ex.alx
            End If

            activeWorksheet.Cells(j, 23).Value = gotw_fld_dict.get_short_def("segment_type", fdwa2f1b.wa2f1ex.segment_type.Trim())

            activeWorksheet.Cells(j, 24).Value = gotw_fld_dict.get_short_def("traffic_direction", fdwa2f1b.wa2f1ex.traffic_dir)
            activeWorksheet.Cells(j, 25).Value = gotw_fld_dict.get_short_def("feature_type", fdwa2f1b.wa2f1ex.feature_type)
            activeWorksheet.Cells(j, 26).Value = gotw_fld_dict.get_short_def("roadway_type", fdwa2f1b.wa2f1ex.roadway_type.Trim())
            activeWorksheet.Cells(j, 27).Value = gotw_fld_dict.get_short_def("right_of_way_type", fdwa2f1b.wa2f1ex.right_of_way_type)
            activeWorksheet.Cells(j, 28).Value = fdwa2f1b.wa2f1ex.physical_id
            activeWorksheet.Cells(j, 28).NumberFormat = "0000000"
            activeWorksheet.Cells(j, 29).Value = fdwa2f1b.wa2f1ex.generic_id
            activeWorksheet.Cells(j, 29).NumberFormat = "0000000"
            activeWorksheet.Cells(j, 30).Value = gotw_fld_dict.get_short_def("bike_lane2", fdwa2f1b.wa2f1ex.bike_lane2)
            'njp (2017-01-04 -- 17.1 Changes for Bike Traffic Direction)
            activeWorksheet.Cells(j, 31).Value = gotw_fld_dict.get_short_def("bike_traffic_direction", fdwa2f1b.wa2f1ex.bike_traffic_direction)

            activeWorksheet.Cells(j, 32).Value = gotw_fld_dict.get_short_def("spec_addr_flag", fdwa2f1b.wa2f1ex.spec_addr_flag)

            activeWorksheet.Cells(j, 33).Value = ExtractHouseNumberFromString(fdwa2f1b.wa2f1ex.lo_hns)
            activeWorksheet.Cells(j, 34).Value = ExtractHouseNumberFromString(fdwa2f1b.wa2f1ex.hi_hns)

            activeWorksheet.Cells(j, 35).Value = fdwa2f1b.wa2f1ex.census_tract_2010
            activeWorksheet.Cells(j, 36).Value = fdwa2f1b.wa2f1ex.census_block_2010
            activeWorksheet.Cells(j, 37).Value = fdwa2f1b.wa2f1ex.census_tract_2000
            activeWorksheet.Cells(j, 38).Value = fdwa2f1b.wa2f1ex.census_block_2000
            activeWorksheet.Cells(j, 39).Value = fdwa2f1b.wa2f1ex.street_width
            activeWorksheet.Cells(j, 40).Value = fdwa2f1b.wa2f1ex.st_width_max
            activeWorksheet.Cells(j, 41).Value = fdwa2f1b.wa2f1ex.street_width_irregular
            activeWorksheet.Cells(j, 42).Value = fdwa2f1b.wa2f1ex.speed_limit

            activeWorksheet.Cells(j, 43).Value = gotw_fld_dict.get_short_def("police_patrol_boro", fdwa2f1b.wa2f1ex.police_patrol_boro)
            activeWorksheet.Cells(j, 44).Value = fdwa2f1b.wa2f1ex.police_pct
            'police sector
            'police service area
            activeWorksheet.Cells(j, 45).Value = fdwa2f1b.wa2f1ex.fire_div
            activeWorksheet.Cells(j, 46).Value = fdwa2f1b.wa2f1ex.fire_bat
            activeWorksheet.Cells(j, 47).Value = gotw_fld_dict.get_short_def("fire_co_type", fdwa2f1b.wa2f1ex.fire_co_type + " " + fdwa2f1b.wa2f1ex.fire_co_num)

            activeWorksheet.Cells(j, 48).Value = fdwa2f1b.wa2f1ex.health_area.Substring(0, 2) + "." + fdwa2f1b.wa2f1ex.health_area.Substring(2, 2)
            activeWorksheet.Cells(j, 48).NumberFormat = "00.00"

            activeWorksheet.Cells(j, 49).Value = fdwa2f1b.wa2f1ex.health_center_dist
            activeWorksheet.Cells(j, 50).Value = fdwa2f1b.wa2f1ex.dot_st_light_contract_area
            activeWorksheet.Cells(j, 51).Value = fdwa2f1b.wa2f1ex.san_dist + " / " + fdwa2f1b.wa2f1ex.san_dist.Remove(0, 1) + fdwa2f1b.wa2f1ex.san_sched.Remove(1)
            activeWorksheet.Cells(j, 52).Value = fdwa2f1b.wa2f1ex.san_sched
            activeWorksheet.Cells(j, 53).Value = fdwa2f1b.wa2f1ex.san_reg
            activeWorksheet.Cells(j, 54).Value = fdwa2f1b.wa2f1ex.san_recycle
            activeWorksheet.Cells(j, 55).Value = fdwa2f1b.wa2f1ex.san_org_pick_up
            activeWorksheet.Cells(j, 56).Value = fdwa2f1b.wa2f1ex.school_dist
            activeWorksheet.Cells(j, 57).Value = gotw_fld_dict.get_short_def("dsny_snow_priority", fdwa2f1b.wa2f1ex.dsny_snow_priority)
            activeWorksheet.Cells(j, 58).Value = fdwa2f1b.wa2f1ex.san_bulk
            activeWorksheet.Cells(j, 59).Value = fdwa2f1b.wa2f1ex.hurricane_zone


            activeWorksheet.Cells(j, 60).Value = fdwa2f1b.wa2f1ex.co
            activeWorksheet.Cells(j, 61).Value = fdwa2f1b.wa2f1ex.ad
            activeWorksheet.Cells(j, 62).Value = fdwa2f1b.wa2f1ex.cd
            activeWorksheet.Cells(j, 63).Value = fdwa2f1b.wa2f1ex.mc
            activeWorksheet.Cells(j, 64).Value = fdwa2f1b.wa2f1ex.ed
            activeWorksheet.Cells(j, 65).Value = fdwa2f1b.wa2f1ex.sd
            activeWorksheet.Cells(j, 66).Value = fdwa2f1b.wa2f1ex.boe_preferred_b7sc.ToString() + " / " + fdwa2f1b.wa2f1ex.boe_preferred_stname.ToString()


            activeWorksheet.Cells(j, 67).Value = fdwa2f1b.wa2f1ax.bbl.block
            activeWorksheet.Cells(j, 68).Value = fdwa2f1b.wa2f1ax.bbl.lot
            activeWorksheet.Cells(j, 69).Value = fdwa2f1b.wa2f1ax.bbl.ToString
            activeWorksheet.Cells(j, 70).Value = fdwa2f1b.wa2f1ax.num_of_blockfaces
            activeWorksheet.Cells(j, 71).Value = fdwa2f1b.wa2f1ax.sanborn.boro + "/" + fdwa2f1b.wa2f1ax.sanborn.volume + fdwa2f1b.wa2f1ax.sanborn.volume_suffix + "/" + fdwa2f1b.wa2f1ax.sanborn.page + fdwa2f1b.wa2f1ax.sanborn.page_suffix
            activeWorksheet.Cells(j, 72).Value = fdwa2f1b.wa2f1ax.rpad_scc
            activeWorksheet.Cells(j, 73).Value = fdwa2f1b.wa2f1ax.rpad_bldg_class

            If fdwa2f1b.wa2f1ax.interior_flag = "" Then
                activeWorksheet.Cells(j, 74).Value = "No"
            Else
                activeWorksheet.Cells(j, 74).Value = fdwa2f1b.wa2f1ax.interior_flag
            End If

            'activeWorksheet.Cells(j, 61).Value = fdwa2f1b.wa2f1ax.interior_flag

            If fdwa2f1b.wa2f1ax.irreg_flag = "" Then
                activeWorksheet.Cells(j, 75).Value = "No"
            Else
                activeWorksheet.Cells(j, 75).Value = fdwa2f1b.wa2f1ax.irreg_flag
            End If


            'activeWorksheet.Cells(j, 62).Value = fdwa2f1b.wa2f1ax.irreg_flag
            'activeWorksheet.Cells(j, 63).Value = fdwa2f1b.wa2f1ax.condo_num

            If Val(fdwa2f1b.wa2f1ax.condo_num) = 0 Or fdwa2f1b.wa2f1ax.condo_num = String.Empty Then
                activeWorksheet.Cells(j, 76).Value = "N/A"
            Else
                activeWorksheet.Cells(j, 76).Value = fdwa2f1b.wa2f1ax.condo_num
            End If

            If Val(fdwa2f1b.wa2f1ax.coop_num) = 0 Or fdwa2f1b.wa2f1ax.coop_num = String.Empty Then
                activeWorksheet.Cells(j, 77).Value = "N/A"
            Else
                activeWorksheet.Cells(j, 77).Value = fdwa2f1b.wa2f1ax.coop_num
            End If
            activeWorksheet.Cells(j, 78).Value = fdwa2f1b.wa2f1ax.vacant_flag
            activeWorksheet.Cells(j, 79).Value = gotw_fld_dict.get_short_def("condo_flag", fdwa2f1b.wa2f1ax.condo_flag)
            If fdwa2f1b.wa2f1ax.condo_flag = "C" Then
                activeWorksheet.Cells(j, 80).Value = fdwa2f1b.wa2f1ax.condo_lo_bbl.boro + " - " + fdwa2f1b.wa2f1ax.condo_lo_bbl.block + " - " + fdwa2f1b.wa2f1ax.condo_lo_bbl.lot
                activeWorksheet.Cells(j, 81).Value = fdwa2f1b.wa2f1ax.condo_hi_bbl.boro + " - " + fdwa2f1b.wa2f1ax.condo_hi_bbl.block + " - " + fdwa2f1b.wa2f1ax.condo_hi_bbl.lot
            Else
                activeWorksheet.Cells(j, 80).Value = "N/A"
                activeWorksheet.Cells(j, 81).Value = "N/A"
            End If

            activeWorksheet.Cells(j, 82).Value = "'" + fdwa2f1b.wa2f1ax.dof_map.boro + " / " + fdwa2f1b.wa2f1ax.dof_map.section_volume.Remove(2, 2) + " / " + fdwa2f1b.wa2f1ax.dof_map.section_volume.Remove(0, 2)

            activeWorksheet.Cells(j, 83).Value = fdwa2f1b.wa2f1ax.bin.BINToString()
            activeWorksheet.Cells(j, 84).Value = fdwa2f1b.wa2f1ax.TPAD_bin_status
            activeWorksheet.Cells(j, 85).Value = fdwa2f1b.wa2f1ax.TPAD_new_bin.ToString()
            activeWorksheet.Cells(j, 86).Value = fdwa2f1b.wa2f1ax.TPAD_new_bin_status
            activeWorksheet.Cells(j, 87).Value = fdwa2f1b.wa2f1ax.TPAD_conflict_flag
            activeWorksheet.Cells(j, 88).Value = gotw_fld_dict.get_short_def("corner_code", fdwa2f1b.wa2f1ax.corner_code)

            If fdwa2f1b.wa2f1ax.bid_id.B5scToString().Trim() = "" Then
                activeWorksheet.Cells(j, 89).Value = ""
            Else
                activeWorksheet.Cells(j, 89).Value = getStreetName(fdwa2f1b.wa2f1ax.bid_id.boro, fdwa2f1b.wa2f1ax.bid_id.B5scToString().Remove(0, 1))
            End If


            '            activeWorksheet.Cells(j, 76).Value = fdwa2f1b.wa2f1ax.bid_id.ToString
            activeWorksheet.Cells(j, 90).Value = Convert.ToString(Val(fdwa2f1b.wa2f1ax.x_coord)) + "/" + Convert.ToString(Val(fdwa2f1b.wa2f1ax.y_coord))

            activeWorksheet.Cells(j, 91).Value = fdwa2f1b.wa2f1ex.blockface_id
            activeWorksheet.Cells(j, 92).Value = fdwa2f1b.wa2f1ex.No_Traveling_lanes
            activeWorksheet.Cells(j, 93).Value = fdwa2f1b.wa2f1ex.No_Parking_lanes
            activeWorksheet.Cells(j, 94).Value = fdwa2f1b.wa2f1ex.No_Total_Lanes
            activeWorksheet.Cells(j, 95).Value = fdwa2f1b.wa2f1ax.DCP_Zoning_Map
            activeWorksheet.Cells(j, 96).Value = fdwa2f1b.wa2f1ex.puma_code
            'police sector activeWorksheet.Cells(j, 97).Value = fdwa2f1b.wa2f1ex.police_sector
            'police service area activeWorksheet.Cells(j, 98).Value = fdwa2f1b.wa2f1ex.police_area
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-1B Errors")
            activeWorksheet.Cells(k, 1).Value = "'" + fdwa1.out_grc + "/" + fdwa1.out_reason_code
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If
            activeWorksheet.Cells(k, 4).Value = fdwa1.in_hnd
            activeWorksheet.Cells(k, 5).Value = fdwa1.in_stname1
            activeWorksheet.Cells(k, 6).Value = fdwa1.out_unit

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function
    Public Function WriteData2(ByRef fdwa1 As Wa1, ByRef fdwa2f2 As Wa2F2, i As Integer) As Integer
        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 2
        Static Dim k As Integer = 1
        Dim xlCenter As Long = -4108
        Dim gotw_fld_dict As New fld_dict

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-2 Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "Borough1/Borough2"
            activeWorksheet.Cells(1, 4).Value = "First Cross Street"
            activeWorksheet.Cells(1, 5).Value = "Second Cross Street"
            activeWorksheet.Cells(1, 6).Value = "ZIP Code"
            activeWorksheet.Range("A1:F1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-2 Output"
            AddFlag = True
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("2:2").Font.Bold = True

            activeWorksheet.Cells(1, 2).value = "Input Data"
            activeWorksheet.Range("B1:D1").Merge()
            activeWorksheet.Range("B1:D1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 6).value = "Geographic Information"
            activeWorksheet.Range("E1:Q1").Merge()
            activeWorksheet.Range("E1:Q1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 18).value = "City Service Information"
            activeWorksheet.Range("R1:AC1").Merge()
            activeWorksheet.Range("R1:AC1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 30).value = "Political Information"
            activeWorksheet.Range("AD1:AH1").Merge()
            activeWorksheet.Range("AD1:AH1").HorizontalAlignment = xlCenter

            activeWorksheet.Range("A1:AH1").Interior.Color = RGB(204, 204, 255)
            activeWorksheet.Range("A1:AH1").Font.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:AH2").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:AH2").Font.Color = RGB(255, 255, 255)

            activeWorksheet.Cells(2, 1).Value = "Return Code/Reason Code"
            activeWorksheet.Cells(2, 2).Value = "Borough1/Borough2"
            activeWorksheet.Cells(2, 3).Value = "First Cross Street"
            activeWorksheet.Cells(2, 4).Value = "Second Cross Street"
            activeWorksheet.Cells(2, 5).Value = "ZIP Code"

            activeWorksheet.Cells(2, 6).Value = "X Coordinate"
            activeWorksheet.Cells(2, 7).Value = "Y Coordinate"
            activeWorksheet.Cells(2, 8).Value = "Community District"
            activeWorksheet.Cells(2, 9).Value = "Compass Direction"
            activeWorksheet.Cells(2, 10).Value = "LION Node Number"
            activeWorksheet.Cells(2, 11).Value = "DCP Preferred B7SC / Street Name for Street 1"
            activeWorksheet.Cells(2, 12).Value = "DCP Preferred B7SC / Street Name for Street 2"
            activeWorksheet.Cells(2, 13).Value = "2010 Census Tract"
            activeWorksheet.Cells(2, 14).Value = "2000 Census Tract"
            activeWorksheet.Cells(2, 15).Value = "Sanborn 1 Boro/Vol/Page"
            activeWorksheet.Cells(2, 16).Value = "Sanborn 2 Boro/Vol/Page"
            activeWorksheet.Cells(2, 17).Value = "Atomic Polygon"

            activeWorksheet.Cells(2, 18).Value = "Police Patrol Borough"
            activeWorksheet.Cells(2, 19).Value = "Police Precinct"
            activeWorksheet.Cells(2, 20).Value = "Fire Division"
            activeWorksheet.Cells(2, 21).Value = "Fire Battalion"
            activeWorksheet.Cells(2, 22).Value = "Fire Company"
            activeWorksheet.Cells(2, 23).Value = "Health Area"
            activeWorksheet.Cells(2, 24).Value = "Health Center District"
            activeWorksheet.Cells(2, 25).Value = "DOT Street Light Area"
            activeWorksheet.Cells(2, 26).Value = "Sanitation District/Section"
            activeWorksheet.Cells(2, 27).Value = "Sanitation Subsection:"
            activeWorksheet.Cells(2, 28).Value = "School District"
            activeWorksheet.Cells(2, 29).Value = "CD Eligibility"

            activeWorksheet.Cells(2, 30).Value = "City Council District"
            activeWorksheet.Cells(2, 31).Value = "Assembly District"
            activeWorksheet.Cells(2, 32).Value = "Congressional District"
            activeWorksheet.Cells(2, 33).Value = "Municipal Court District"
            activeWorksheet.Cells(2, 34).Value = "State Senate District"
            'police Sector 
            'activeWorksheet.Cells(2, 35).Value = "Police Sector"

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

        Else
            activeWorksheet = activeWorkbook.Sheets("Func-2 Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet = activeWorkbook.Sheets("Func-2 Output")
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_reason_code

            If fdwa1.in_boro1 = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa1.in_boro1 = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa1.in_boro1 = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa1.in_boro1 = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa1.in_boro1 = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If

            If fdwa1.in_boro2 = "1" Then
                activeWorksheet.Cells(j, 2).Value = activeWorksheet.Cells(j, 2).Value + "/" + "Manhattan"
            ElseIf fdwa1.in_boro2 = "2" Then
                activeWorksheet.Cells(j, 2).Value = activeWorksheet.Cells(j, 2).Value + "/" + "Bronx"
            ElseIf fdwa1.in_boro2 = "3" Then
                activeWorksheet.Cells(j, 2).Value = activeWorksheet.Cells(j, 2).Value + "/" + "Brooklyn"
            ElseIf fdwa1.in_boro2 = "4" Then
                activeWorksheet.Cells(j, 2).Value = activeWorksheet.Cells(j, 2).Value + "/" + "Queens"
            ElseIf fdwa1.in_boro2 = "5" Then
                activeWorksheet.Cells(j, 2).Value = activeWorksheet.Cells(j, 2).Value + "/" + "Staten Island"
            End If

            activeWorksheet.Cells(j, 3).Value = fdwa1.in_stname1
            activeWorksheet.Cells(j, 4).Value = fdwa1.in_stname2
            activeWorksheet.Cells(j, 5).Value = fdwa2f2.zip_code
            activeWorksheet.Cells(j, 6).Value = fdwa2f2.x_coord
            activeWorksheet.Cells(j, 7).Value = fdwa2f2.y_coord
            activeWorksheet.Cells(j, 8).Value = fdwa2f2.com_dist.boro + fdwa2f2.com_dist.district_number
            activeWorksheet.Cells(j, 9).Value = fdwa2f2.compass
            activeWorksheet.Cells(j, 10).Value = fdwa2f2.lion_node_num
            activeWorksheet.Cells(j, 11).Value = fdwa2f2.dcp_pref_lgc1
            activeWorksheet.Cells(j, 12).Value = fdwa2f2.dcp_pref_lgc2
            activeWorksheet.Cells(j, 13).Value = fdwa2f2.census_tract_2010
            activeWorksheet.Cells(j, 14).Value = fdwa2f2.census_tract_2000
            activeWorksheet.Cells(j, 15).Value = fdwa2f2.sanborn1.boro + "/" + fdwa2f2.sanborn1.volume + fdwa2f2.sanborn1.volume_suffix + "/" + fdwa2f2.sanborn1.page + fdwa2f2.sanborn1.page_suffix
            activeWorksheet.Cells(j, 16).Value = fdwa2f2.sanborn2.boro + "/" + fdwa2f2.sanborn2.volume + fdwa2f2.sanborn2.volume_suffix + "/" + fdwa2f2.sanborn2.page + fdwa2f2.sanborn2.page_suffix
            activeWorksheet.Cells(j, 17).Value = fdwa2f2.atomic_polygon

            activeWorksheet.Cells(j, 18).Value = gotw_fld_dict.get_short_def("police_patrol_boro", fdwa2f2.police_patrol_boro)
            activeWorksheet.Cells(j, 19).Value = fdwa2f2.police_pct
            'police sector
            activeWorksheet.Cells(j, 20).Value = fdwa2f2.fire_div
            activeWorksheet.Cells(j, 21).Value = fdwa2f2.fire_bat
            activeWorksheet.Cells(j, 22).Value = gotw_fld_dict.get_short_def("fire_co_type", fdwa2f2.fire_co_type + " " + fdwa2f2.fire_co_num)
            activeWorksheet.Cells(j, 23).Value = fdwa2f2.health_area.Substring(0, 2) + "." + fdwa2f2.health_area.Substring(2, 2)
            activeWorksheet.Cells(j, 23).NumberFormat = "00.00"
            activeWorksheet.Cells(j, 24).Value = fdwa2f2.health_center_dist
            activeWorksheet.Cells(j, 25).Value = fdwa2f2.dot_st_light_contract_area
            activeWorksheet.Cells(j, 26).Value = fdwa2f2.san_dist
            activeWorksheet.Cells(j, 27).Value = fdwa2f2.san_sub_section
            activeWorksheet.Cells(j, 28).Value = fdwa2f2.school_dist
            activeWorksheet.Cells(j, 29).Value = fdwa2f2.cd_eligible

            activeWorksheet.Cells(j, 30).Value = fdwa2f2.co
            activeWorksheet.Cells(j, 31).Value = fdwa2f2.ad
            activeWorksheet.Cells(j, 32).Value = fdwa2f2.cd
            activeWorksheet.Cells(j, 33).Value = fdwa2f2.mc
            activeWorksheet.Cells(j, 34).Value = fdwa2f2.sd
            'police sector activeWorksheet.Cells(j, 35).Value = fdwa2f2.police_sector
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-2 Errors")
            activeWorksheet.Cells(k, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_grc2
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message

            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If

            If fdwa1.in_boro1 = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_boro1 = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_boro1 = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_boro1 = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_boro1 = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If

            If fdwa1.in_boro2 = "1" Then
                activeWorksheet.Cells(k, 3).Value = activeWorksheet.Cells(k, 3).Value + "/" + "Manhattan"
            ElseIf fdwa1.in_boro2 = "2" Then
                activeWorksheet.Cells(k, 3).Value = activeWorksheet.Cells(k, 3).Value + "/" + "Bronx"
            ElseIf fdwa1.in_boro2 = "3" Then
                activeWorksheet.Cells(k, 3).Value = activeWorksheet.Cells(k, 3).Value + "/" + "Brooklyn"
            ElseIf fdwa1.in_boro2 = "4" Then
                activeWorksheet.Cells(k, 3).Value = activeWorksheet.Cells(k, 3).Value + "/" + "Queens"
            ElseIf fdwa1.in_boro2 = "5" Then
                activeWorksheet.Cells(k, 3).Value = activeWorksheet.Cells(k, 3).Value + "/" + "Staten Island"
            End If

            activeWorksheet.Cells(k, 4).Value = fdwa1.in_stname1
            activeWorksheet.Cells(k, 5).Value = fdwa1.in_stname2
            activeWorksheet.Cells(k, 6).Value = fdwa2f2.zip_code
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function

    Public Function WriteData3(ByRef fdwa1 As Wa1, ByRef fdwa2f3 As Wa2F3xas, i As Integer) As Integer
        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 2
        Static Dim k As Integer = 1
        Dim xlCenter As Long = -4108
        Dim gotw_fld_dict As New fld_dict

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-3 Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "Borough"
            activeWorksheet.Cells(1, 4).Value = "On Street"
            activeWorksheet.Cells(1, 5).Value = "First Cross Street"
            activeWorksheet.Cells(1, 6).Value = "Second Cross Street"
            activeWorksheet.Range("A1:F1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-3 Output"
            AddFlag = True
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("2:2").Font.Bold = True

            activeWorksheet.Cells(1, 2).value = "Input Data"
            activeWorksheet.Range("B1:E1").Merge()
            activeWorksheet.Range("B1:E1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 6).value = "Geographic Information"
            activeWorksheet.Range("F1:AF1").Merge()
            activeWorksheet.Range("F1:AF1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 33).value = "Left side of Street Information"
            activeWorksheet.Range("AG1:AZ1").Merge()
            activeWorksheet.Range("AG1:AZ1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 53).value = "Right side of Street Information"
            activeWorksheet.Range("BA1:BT1").Merge()
            activeWorksheet.Range("BA1:BT1").HorizontalAlignment = xlCenter

            activeWorksheet.Range("A1:BT1").Interior.Color = RGB(204, 204, 255)
            activeWorksheet.Range("A1:BT1").Font.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:BT2").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:BT2").Font.Color = RGB(255, 255, 255)

            activeWorksheet.Cells(2, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(2, 2).value = "Borough"
            activeWorksheet.Cells(2, 3).value = "On Street"
            activeWorksheet.Cells(2, 4).Value = "First Cross Street"
            activeWorksheet.Cells(2, 5).Value = "Second Cross Street"
            activeWorksheet.Cells(2, 6).Value = "ZIP code"
            activeWorksheet.Cells(2, 7).Value = "From Node"
            activeWorksheet.Cells(2, 8).Value = "To Node"
            activeWorksheet.Cells(2, 9).Value = "LION Key"
            activeWorksheet.Cells(2, 10).Value = "From X Coordinate"
            activeWorksheet.Cells(2, 11).Value = "From Y Coordinate"
            activeWorksheet.Cells(2, 12).Value = "To X Coordinate"
            activeWorksheet.Cells(2, 13).Value = "To Y Coordinate"
            activeWorksheet.Cells(2, 14).Value = "DOT Street Light Area"
            activeWorksheet.Cells(2, 15).Value = "Segment ID/Length"
            activeWorksheet.Cells(2, 16).Value = "Physical ID"
            activeWorksheet.Cells(2, 17).Value = "Generic ID"
            activeWorksheet.Cells(2, 18).Value = "Location Status"
            activeWorksheet.Cells(2, 19).Value = "Bike Lane"
            activeWorksheet.Cells(2, 20).Value = "Bike Traffic Direction"
            activeWorksheet.Cells(2, 21).Value = "Traffic Direction"
            activeWorksheet.Cells(2, 22).Value = "Segment Type"
            activeWorksheet.Cells(2, 23).Value = "Feature Type"
            activeWorksheet.Cells(2, 24).Value = "Roadway Type"
            activeWorksheet.Cells(2, 25).Value = "Right of Way Type"
            activeWorksheet.Cells(2, 26).Value = "No of Traveling Lanes"
            activeWorksheet.Cells(2, 27).Value = "No of Parking Lanes"
            activeWorksheet.Cells(2, 28).Value = "Total No. of Lanes"
            activeWorksheet.Cells(2, 29).Value = "Street Width Min"
            activeWorksheet.Cells(2, 30).Value = "Street Width Max"
            activeWorksheet.Cells(2, 31).Value = "Street Width Irregular"
            activeWorksheet.Cells(2, 32).Value = "Speed Limit"


            activeWorksheet.Cells(2, 33).Value = "Borough"
            activeWorksheet.Cells(2, 34).Value = "Community District"
            activeWorksheet.Cells(2, 35).Value = "Low-/-High House Number"
            activeWorksheet.Cells(2, 36).Value = "ZIP Code"
            activeWorksheet.Cells(2, 37).Value = "2010 Census Tract"
            activeWorksheet.Cells(2, 38).Value = "2010 Census Block"
            activeWorksheet.Cells(2, 39).Value = "2000 Census Tract"
            activeWorksheet.Cells(2, 40).Value = "2000 Census Block"
            activeWorksheet.Cells(2, 41).Value = "Police Patrol Borough"
            activeWorksheet.Cells(2, 42).Value = "Police Precinct"
            activeWorksheet.Cells(2, 43).Value = "Fire Division"
            activeWorksheet.Cells(2, 44).Value = "Fire Battalion"
            activeWorksheet.Cells(2, 45).Value = "Fire Company"
            activeWorksheet.Cells(2, 46).Value = "Health Area"
            activeWorksheet.Cells(2, 47).Value = "Health Center District"
            activeWorksheet.Cells(2, 48).Value = "DOT Street Light Area"
            activeWorksheet.Cells(2, 49).Value = "School District"
            activeWorksheet.Cells(2, 50).Value = "CD Eligibility"
            activeWorksheet.Cells(2, 51).Value = "Left Blockface ID"
            activeWorksheet.Cells(2, 52).Value = "Left Side PUMA"
            'activeWorksheet.Cells(2, 53).Value = "Left Side Police Sector"

            activeWorksheet.Cells(2, 53).Value = "Borough"
            activeWorksheet.Cells(2, 54).Value = "Community District"
            activeWorksheet.Cells(2, 55).Value = "Low-/-High House Number"
            activeWorksheet.Cells(2, 56).Value = "ZIP Code"
            activeWorksheet.Cells(2, 57).Value = "2010 Census Tract"
            activeWorksheet.Cells(2, 58).Value = "2010 Census Block"
            activeWorksheet.Cells(2, 69).Value = "2000 Census Tract"
            activeWorksheet.Cells(2, 60).Value = "2000 Census Block"
            activeWorksheet.Cells(2, 61).Value = "Police Patrol Borough"
            activeWorksheet.Cells(2, 62).Value = "Police Precinct"
            activeWorksheet.Cells(2, 63).Value = "Fire Division"
            activeWorksheet.Cells(2, 64).Value = "Fire Battalion"
            activeWorksheet.Cells(2, 65).Value = "Fire Company"
            activeWorksheet.Cells(2, 66).Value = "Health Area"
            activeWorksheet.Cells(2, 67).Value = "Health Center District"
            activeWorksheet.Cells(2, 68).Value = "DOT Street Light Area"
            activeWorksheet.Cells(2, 69).Value = "School District"
            activeWorksheet.Cells(2, 70).Value = "CD Eligibility"
            activeWorksheet.Cells(2, 71).Value = "Right Blockface ID"
            activeWorksheet.Cells(2, 72).Value = "Right Side PUMA"
            'activeWorksheet.Cells(2, 74).Value = "Right Side Police Sector"
        Else
            activeWorksheet = activeWorkbook.Sheets("Func-3 Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet = activeWorkbook.Sheets("Func-3 Output")
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_grc2
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 3).Value = fdwa1.in_stname1
            activeWorksheet.Cells(j, 4).Value = fdwa1.in_stname2
            activeWorksheet.Cells(j, 5).Value = fdwa1.in_stname3

            activeWorksheet.Cells(j, 6).Value = fdwa2f3.wa2f3x.left_side.zip_code
            activeWorksheet.Cells(j, 7).Value = Convert.ToString(Val(fdwa2f3.wa2f3x.from_node))
            activeWorksheet.Cells(j, 8).Value = Convert.ToString(Val(fdwa2f3.wa2f3x.to_node))
            activeWorksheet.Cells(j, 9).Value = fdwa2f3.wa2f3x.lionkey.ToString()
            activeWorksheet.Cells(j, 10).Value = Convert.ToString(Val(fdwa2f3.wa2f3x.from_x_coord))
            activeWorksheet.Cells(j, 11).Value = Convert.ToString(Val(fdwa2f3.wa2f3x.from_y_coord))
            activeWorksheet.Cells(j, 12).Value = Convert.ToString(Val(fdwa2f3.wa2f3x.to_x_coord))
            activeWorksheet.Cells(j, 13).Value = Convert.ToString(Val(fdwa2f3.wa2f3x.to_y_coord))
            activeWorksheet.Cells(j, 14).Value = fdwa2f3.wa2f3x.dot_street_light_contract_area
            activeWorksheet.Cells(j, 15).Value = fdwa2f3.wa2f3x.segment_id + "/" + Convert.ToString(Val(fdwa2f3.wa2f3x.segment_len))
            activeWorksheet.Cells(j, 16).Value = fdwa2f3.wa2f3x.physical_id
            activeWorksheet.Cells(j, 17).Value = fdwa2f3.wa2f3x.generic_id
            activeWorksheet.Cells(j, 18).Value = fdwa2f3.wa2f3x.loc_status
            activeWorksheet.Cells(j, 19).Value = gotw_fld_dict.get_short_def("bike_lane2", fdwa2f3.wa2f3x.bike_lane2)
            'njp(2017-01-04 - 17.1 Changes to add Bike Traffic Direction)
            activeWorksheet.Cells(j, 20).Value = gotw_fld_dict.get_short_def("bike_traffic_direction", fdwa2f3.wa2f3x.bike_traffic_direction)
            activeWorksheet.Cells(j, 21).Value = gotw_fld_dict.get_short_def("traffic_direction", fdwa2f3.wa2f3x.traffic_direction)
            activeWorksheet.Cells(j, 22).Value = gotw_fld_dict.get_short_def("segment_type", fdwa2f3.wa2f3x.segment_type)
            activeWorksheet.Cells(j, 23).Value = gotw_fld_dict.get_short_def("feature_type", fdwa2f3.wa2f3x.feature_type)
            activeWorksheet.Cells(j, 24).Value = gotw_fld_dict.get_short_def("roadway_type", fdwa2f3.wa2f3x.roadway_type)
            activeWorksheet.Cells(j, 25).Value = gotw_fld_dict.get_short_def("right_of_way_type", fdwa2f3.wa2f3x.right_of_way_type)
            activeWorksheet.Cells(j, 26).Value = fdwa2f3.wa2f3x.No_Traveling_lanes
            activeWorksheet.Cells(j, 27).Value = fdwa2f3.wa2f3x.No_Parking_lanes
            activeWorksheet.Cells(j, 28).Value = fdwa2f3.wa2f3x.Total_Lanes
            activeWorksheet.Cells(j, 29).Value = fdwa2f3.wa2f3x.street_width
            activeWorksheet.Cells(j, 30).Value = fdwa2f3.wa2f3x.st_width_max
            activeWorksheet.Cells(j, 31).Value = fdwa2f3.wa2f3x.street_width_irregular
            activeWorksheet.Cells(j, 32).Value = fdwa2f3.wa2f3x.speed_limit

            activeWorksheet.Cells(j, 33).Value = fdwa2f3.wa2f3x.left_side.boro
            activeWorksheet.Cells(j, 34).Value = fdwa2f3.wa2f3x.left_side.boro + fdwa2f3.wa2f3x.left_side.comdist.district_number
            activeWorksheet.Cells(j, 35).Value = fdwa2f3.wa2f3x.left_side.lhnd.Trim().TrimStart("0") + "-/-" + fdwa2f3.wa2f3x.left_side.hhnd.Trim().TrimStart("0")
            activeWorksheet.Cells(j, 36).Value = fdwa2f3.wa2f3x.left_side.zip_code
            activeWorksheet.Cells(j, 37).Value = fdwa2f3.wa2f3x.left_side.census_tract_2010
            activeWorksheet.Cells(j, 38).Value = fdwa2f3.wa2f3x.left_side.census_block_2010
            activeWorksheet.Cells(j, 39).Value = fdwa2f3.wa2f3x.left_side.census_tract_2000
            activeWorksheet.Cells(j, 40).Value = fdwa2f3.wa2f3x.left_side.census_block_2000
            activeWorksheet.Cells(j, 41).Value = gotw_fld_dict.get_short_def("police_patrol_boro", fdwa2f3.wa2f3x.left_side.police_patrol_boro)
            activeWorksheet.Cells(j, 42).Value = fdwa2f3.wa2f3x.left_side.police_pct
            activeWorksheet.Cells(j, 43).Value = fdwa2f3.wa2f3x.left_side.fire_div
            activeWorksheet.Cells(j, 44).Value = fdwa2f3.wa2f3x.left_side.fire_bat.Trim().TrimStart("0")
            activeWorksheet.Cells(j, 45).Value = gotw_fld_dict.get_short_def("fire_co_type", fdwa2f3.wa2f3x.left_side.fire_co_type + " " + fdwa2f3.wa2f3x.left_side.fire_co_num)
            activeWorksheet.Cells(j, 46).Value = fdwa2f3.wa2f3x.left_side.health_area.Substring(0, 2) + "." + fdwa2f3.wa2f3x.left_side.health_area.Substring(2, 2)
            activeWorksheet.Cells(j, 46).NumberFormat = "00.00"
            activeWorksheet.Cells(j, 47).Value = fdwa2f3.wa2f3x.left_health_center_dist
            activeWorksheet.Cells(j, 48).Value = fdwa2f3.wa2f3x.dot_street_light_contract_area
            activeWorksheet.Cells(j, 49).Value = fdwa2f3.wa2f3x.left_side.school_dist
            activeWorksheet.Cells(j, 50).Value = gotw_fld_dict.get_short_def("cd_eligible", fdwa2f3.wa2f3x.left_side.iaei)
            activeWorksheet.Cells(j, 51).Value = fdwa2f3.wa2f3x.left_blockface_id
            activeWorksheet.Cells(j, 52).Value = fdwa2f3.wa2f3x.left_puma_code
            'police sector
            'activeWorksheet.Cells(j, 53).Value = fdwa2f3.wa2f3x.left_police_sector

            activeWorksheet.Cells(j, 53).Value = fdwa2f3.wa2f3x.right_side.boro
            activeWorksheet.Cells(j, 54).Value = fdwa2f3.wa2f3x.right_side.boro + fdwa2f3.wa2f3x.right_side.comdist.district_number
            activeWorksheet.Cells(j, 55).Value = fdwa2f3.wa2f3x.right_side.lhnd.Trim().TrimStart("0") + "-/-" + fdwa2f3.wa2f3x.right_side.hhnd.Trim().TrimStart("0")
            activeWorksheet.Cells(j, 56).Value = fdwa2f3.wa2f3x.right_side.zip_code
            activeWorksheet.Cells(j, 57).Value = fdwa2f3.wa2f3x.right_side.census_tract_2010
            activeWorksheet.Cells(j, 58).Value = fdwa2f3.wa2f3x.right_side.census_block_2010
            activeWorksheet.Cells(j, 59).Value = fdwa2f3.wa2f3x.right_side.census_tract_2000
            activeWorksheet.Cells(j, 60).Value = fdwa2f3.wa2f3x.right_side.census_block_2000
            activeWorksheet.Cells(j, 61).Value = gotw_fld_dict.get_short_def("police_patrol_boro", fdwa2f3.wa2f3x.right_side.police_patrol_boro)
            activeWorksheet.Cells(j, 62).Value = fdwa2f3.wa2f3x.right_side.police_pct
            activeWorksheet.Cells(j, 63).Value = fdwa2f3.wa2f3x.right_side.fire_div
            activeWorksheet.Cells(j, 64).Value = fdwa2f3.wa2f3x.right_side.fire_bat.Trim().TrimStart("0")
            activeWorksheet.Cells(j, 65).Value = gotw_fld_dict.get_short_def("fire_co_type", fdwa2f3.wa2f3x.right_side.fire_co_type + " " + fdwa2f3.wa2f3x.right_side.fire_co_num)
            activeWorksheet.Cells(j, 66).Value = fdwa2f3.wa2f3x.right_side.health_area.Substring(0, 2) + "." + fdwa2f3.wa2f3x.right_side.health_area.Substring(2, 2)
            activeWorksheet.Cells(j, 66).NumberFormat = "00.00"
            activeWorksheet.Cells(j, 67).Value = fdwa2f3.wa2f3x.right_health_center_dist
            activeWorksheet.Cells(j, 68).Value = fdwa2f3.wa2f3x.dot_street_light_contract_area
            activeWorksheet.Cells(j, 69).Value = fdwa2f3.wa2f3x.right_side.school_dist
            activeWorksheet.Cells(j, 70).Value = gotw_fld_dict.get_short_def("cd_eligible", fdwa2f3.wa2f3x.right_side.iaei)
            activeWorksheet.Cells(j, 71).Value = fdwa2f3.wa2f3x.right_blockface_id
            activeWorksheet.Cells(j, 72).Value = fdwa2f3.wa2f3x.right_puma_code
            'police sector
            'activeWorksheet.Cells(j, 74).Value = fdwa2f3.wa2f3x.right_police_sector

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-3 Errors")
            activeWorksheet.Cells(k, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_reason_code
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If
            activeWorksheet.Cells(k, 4).Value = fdwa1.in_stname1
            activeWorksheet.Cells(k, 5).Value = fdwa1.in_stname2
            activeWorksheet.Cells(k, 6).Value = fdwa1.in_stname3
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function

    Public Function WriteData3C(ByRef fdwa1 As Wa1, ByRef fdwa2f3C As Wa2F3cxas, i As Integer) As Integer
        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 2
        Static Dim k As Integer = 1
        Dim xlCenter As Long = -4108
        Dim gotw_fld_dict As New fld_dict

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-3 Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "Borough"
            activeWorksheet.Cells(1, 4).Value = "On Street"
            activeWorksheet.Cells(1, 5).Value = "First Cross Street"
            activeWorksheet.Cells(1, 6).Value = "Second Cross Street"
            activeWorksheet.Range("A1:F1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-3 Output"
            AddFlag = True
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("2:2").Font.Bold = True

            activeWorksheet.Cells(1, 2).value = "Input Data"
            activeWorksheet.Range("B1:F1").Merge()
            activeWorksheet.Range("B1:F1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 7).value = "Geographic Information"
            activeWorksheet.Range("G1:AF1").Merge()
            activeWorksheet.Range("G1:AF1").HorizontalAlignment = xlCenter

            activeWorksheet.Cells(1, 33).value = "Segment side of Street Information"
            'Uncomment when adding police sector
            'activeWorksheet.Range("AG1:BA1").Merge()
            'activeWorksheet.Range("AG1:BA1").HorizontalAlignment = xlCenter
            activeWorksheet.Range("AG1:AZ1").Merge()
            activeWorksheet.Range("AG1:AZ1").HorizontalAlignment = xlCenter

            'activeWorksheet.Cells(1, 53).value = "Right side of Street Information"
            'activeWorksheet.Range("BA1:BT1").Merge()
            'activeWorksheet.Range("BA1:BT1").HorizontalAlignment = xlCenter

            activeWorksheet.Range("A1:AZ1").Interior.Color = RGB(204, 204, 255)
            activeWorksheet.Range("A1:AZ1").Font.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:AZ2").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A2:AZ2").Font.Color = RGB(255, 255, 255)

            activeWorksheet.Cells(2, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(2, 2).value = "Borough"
            activeWorksheet.Cells(2, 3).value = "On Street"
            activeWorksheet.Cells(2, 4).Value = "First Cross Street"
            activeWorksheet.Cells(2, 5).Value = "Second Cross Street"
            activeWorksheet.Cells(2, 6).Value = "Side Of Street"
            activeWorksheet.Cells(2, 7).Value = "From Node"
            activeWorksheet.Cells(2, 8).Value = "To Node"
            activeWorksheet.Cells(2, 9).Value = "LION Key"
            activeWorksheet.Cells(2, 10).Value = "From X Coordinate"
            activeWorksheet.Cells(2, 11).Value = "From Y Coordinate"
            activeWorksheet.Cells(2, 12).Value = "To X Coordinate"
            activeWorksheet.Cells(2, 13).Value = "To Y Coordinate"
            activeWorksheet.Cells(2, 14).Value = "DOT Street Light Area"
            activeWorksheet.Cells(2, 15).Value = "Segment ID/Length"
            activeWorksheet.Cells(2, 16).Value = "Physical ID"
            activeWorksheet.Cells(2, 17).Value = "Generic ID"
            activeWorksheet.Cells(2, 18).Value = "Location Status"
            activeWorksheet.Cells(2, 19).Value = "Bike Lane"
            activeWorksheet.Cells(2, 20).Value = "Bike Traffic Direction"
            activeWorksheet.Cells(2, 21).Value = "Traffic Direction"
            activeWorksheet.Cells(2, 22).Value = "Segment Type"
            activeWorksheet.Cells(2, 23).Value = "Feature Type"
            activeWorksheet.Cells(2, 24).Value = "Roadway Type"
            activeWorksheet.Cells(2, 25).Value = "Right of Way Type"
            activeWorksheet.Cells(2, 26).Value = "No of Traveling Lanes"
            activeWorksheet.Cells(2, 27).Value = "No of Parking Lanes"
            activeWorksheet.Cells(2, 28).Value = "Total No. of Lanes"
            activeWorksheet.Cells(2, 29).Value = "Street Width Min"
            activeWorksheet.Cells(2, 30).Value = "Street Width Max"
            activeWorksheet.Cells(2, 31).Value = "Street Width Irregular"
            activeWorksheet.Cells(2, 32).Value = "Speed Limit"

            activeWorksheet.Cells(2, 33).Value = "Segment Side Borough"
            activeWorksheet.Cells(2, 34).Value = "Segment Side Community District"
            activeWorksheet.Cells(2, 35).Value = "Segment Side Low-/-High House Number"
            activeWorksheet.Cells(2, 36).Value = "Segment Side ZIP Code"
            activeWorksheet.Cells(2, 37).Value = "Segment Side 2010 Census Tract"
            activeWorksheet.Cells(2, 38).Value = "Segment Side 2010 Census Block"
            activeWorksheet.Cells(2, 39).Value = "Segment Side 2000 Census Tract"
            activeWorksheet.Cells(2, 40).Value = "Segment Side 2000 Census Block"
            activeWorksheet.Cells(2, 41).Value = "Segment Side Police Patrol Borough"
            activeWorksheet.Cells(2, 42).Value = "Segment Side Police Precinct"
            activeWorksheet.Cells(2, 43).Value = "Segment Side Fire Division"
            activeWorksheet.Cells(2, 44).Value = "Segment Side Fire Battalion"
            activeWorksheet.Cells(2, 45).Value = "Segment Side Fire Company"
            activeWorksheet.Cells(2, 46).Value = "Segment Side Health Area"
            activeWorksheet.Cells(2, 47).Value = "Segment Side Health Center District"
            activeWorksheet.Cells(2, 48).Value = "Segment Side DOT Street Light Area"
            activeWorksheet.Cells(2, 49).Value = "Segment Side School District"
            activeWorksheet.Cells(2, 50).Value = "Segment Side CD Eligibility"
            activeWorksheet.Cells(2, 51).Value = "Segment Side Blockface ID"
            activeWorksheet.Cells(2, 52).Value = "Segment Side PUMA"
            'police sector 
            'activeWorksheet.Cells(2, 53).Value = "Segment Side Police Sector"

            'activeWorksheet.Cells(2, 53).Value = "Borough"
            'activeWorksheet.Cells(2, 54).Value = "Community District"
            'activeWorksheet.Cells(2, 55).Value = "Low-/-High House Number"
            'activeWorksheet.Cells(2, 56).Value = "ZIP Code"
            'activeWorksheet.Cells(2, 57).Value = "2010 Census Tract"
            'activeWorksheet.Cells(2, 58).Value = "2010 Census Block"
            'activeWorksheet.Cells(2, 59).Value = "2000 Census Tract"
            'activeWorksheet.Cells(2, 60).Value = "2000 Census Block"
            'activeWorksheet.Cells(2, 61).Value = "Police Patrol Borough"
            'activeWorksheet.Cells(2, 62).Value = "Police Precinct"
            'activeWorksheet.Cells(2, 63).Value = "Fire Division"
            'activeWorksheet.Cells(2, 64).Value = "Fire Battalion"
            'activeWorksheet.Cells(2, 65).Value = "Fire Company"
            'activeWorksheet.Cells(2, 66).Value = "Health Area"
            'activeWorksheet.Cells(2, 67).Value = "Health Center District"
            'activeWorksheet.Cells(2, 68).Value = "DOT Street Light Area"
            'activeWorksheet.Cells(2, 69).Value = "School District"
            'activeWorksheet.Cells(2, 70).Value = "CD Eligibility"
            'activeWorksheet.Cells(2, 71).Value = "Right Blockface ID"
            'activeWorksheet.Cells(2, 72).Value = "Right Side PUMA"
        Else
            activeWorksheet = activeWorkbook.Sheets("Func-3 Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet = activeWorkbook.Sheets("Func-3 Output")
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_grc2
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 3).Value = fdwa1.in_stname1
            activeWorksheet.Cells(j, 4).Value = fdwa1.in_stname2
            activeWorksheet.Cells(j, 5).Value = fdwa1.in_stname3

            activeWorksheet.Cells(j, 6).Value = fdwa1.in_compass_dir
            activeWorksheet.Cells(j, 7).Value = Convert.ToString(Val(fdwa2f3C.wa2f3cx.from_node))
            activeWorksheet.Cells(j, 8).Value = Convert.ToString(Val(fdwa2f3C.wa2f3cx.to_node))
            activeWorksheet.Cells(j, 9).Value = fdwa2f3C.wa2f3cx.lionkey.ToString()
            activeWorksheet.Cells(j, 10).Value = Convert.ToString(Val(fdwa2f3C.wa2f3cx.from_x_coord))
            activeWorksheet.Cells(j, 11).Value = Convert.ToString(Val(fdwa2f3C.wa2f3cx.from_y_coord))
            activeWorksheet.Cells(j, 12).Value = Convert.ToString(Val(fdwa2f3C.wa2f3cx.to_x_coord))
            activeWorksheet.Cells(j, 13).Value = Convert.ToString(Val(fdwa2f3C.wa2f3cx.to_y_coord))
            activeWorksheet.Cells(j, 14).Value = fdwa2f3C.wa2f3cx.dot_street_light_contract_area
            activeWorksheet.Cells(j, 15).Value = fdwa2f3C.wa2f3cx.segment_id + "/" + Convert.ToString(Val(fdwa2f3C.wa2f3cx.segment_len))
            activeWorksheet.Cells(j, 16).Value = fdwa2f3C.wa2f3cx.physical_id
            activeWorksheet.Cells(j, 17).Value = fdwa2f3C.wa2f3cx.generic_id
            activeWorksheet.Cells(j, 18).Value = fdwa2f3C.wa2f3cx.loc_status
            activeWorksheet.Cells(j, 19).Value = gotw_fld_dict.get_short_def("bike_lane2", fdwa2f3C.wa2f3cx.bike_lane2)
            'njp(2017-01-04 - 17.1 Changes to add Bike Traffic Direction)
            activeWorksheet.Cells(j, 20).Value = gotw_fld_dict.get_short_def("bike_traffic_direction", fdwa2f3C.wa2f3cx.bike_traffic_direction)
            activeWorksheet.Cells(j, 21).Value = gotw_fld_dict.get_short_def("traffic_direction", fdwa2f3C.wa2f3cx.traffic_direction)
            activeWorksheet.Cells(j, 22).Value = gotw_fld_dict.get_short_def("segment_type", fdwa2f3C.wa2f3cx.segment_type)
            activeWorksheet.Cells(j, 23).Value = gotw_fld_dict.get_short_def("feature_type", fdwa2f3C.wa2f3cx.feature_type)
            activeWorksheet.Cells(j, 24).Value = gotw_fld_dict.get_short_def("roadway_type", fdwa2f3C.wa2f3cx.roadway_type)
            activeWorksheet.Cells(j, 25).Value = gotw_fld_dict.get_short_def("right_of_way_type", fdwa2f3C.wa2f3cx.right_of_way_type)
            activeWorksheet.Cells(j, 26).Value = fdwa2f3C.wa2f3cx.No_Traveling_lanes
            activeWorksheet.Cells(j, 27).Value = fdwa2f3C.wa2f3cx.No_Parking_lanes
            activeWorksheet.Cells(j, 28).Value = fdwa2f3C.wa2f3cx.Total_Lanes
            activeWorksheet.Cells(j, 29).Value = fdwa2f3C.wa2f3cx.street_width
            activeWorksheet.Cells(j, 30).Value = fdwa2f3C.wa2f3cx.st_width_max
            activeWorksheet.Cells(j, 31).Value = fdwa2f3C.wa2f3cx.street_width_irregular
            activeWorksheet.Cells(j, 32).Value = fdwa2f3C.wa2f3cx.speed_limit

            activeWorksheet.Cells(j, 33).Value = fdwa2f3C.wa2f3cx.seg_side.boro
            activeWorksheet.Cells(j, 34).Value = fdwa2f3C.wa2f3cx.seg_side.boro + fdwa2f3C.wa2f3cx.seg_side.comdist.district_number
            activeWorksheet.Cells(j, 35).Value = fdwa2f3C.wa2f3cx.seg_side.lhnd.Trim().TrimStart("0") + "-/-" + fdwa2f3C.wa2f3cx.seg_side.hhnd.Trim().TrimStart("0")
            activeWorksheet.Cells(j, 36).Value = fdwa2f3C.wa2f3cx.seg_side.zip_code
            activeWorksheet.Cells(j, 37).Value = fdwa2f3C.wa2f3cx.seg_side.census_tract_2010
            activeWorksheet.Cells(j, 38).Value = fdwa2f3C.wa2f3cx.seg_side.census_block_2010
            activeWorksheet.Cells(j, 39).Value = fdwa2f3C.wa2f3cx.seg_side.census_tract_2000
            activeWorksheet.Cells(j, 40).Value = fdwa2f3C.wa2f3cx.seg_side.census_block_2000
            activeWorksheet.Cells(j, 41).Value = gotw_fld_dict.get_short_def("police_patrol_boro", fdwa2f3C.wa2f3cx.seg_side.police_patrol_boro)
            activeWorksheet.Cells(j, 42).Value = fdwa2f3C.wa2f3cx.seg_side.police_pct
            activeWorksheet.Cells(j, 43).Value = fdwa2f3C.wa2f3cx.seg_side.fire_div
            activeWorksheet.Cells(j, 44).Value = fdwa2f3C.wa2f3cx.seg_side.fire_bat.Trim().TrimStart("0")
            activeWorksheet.Cells(j, 45).Value = gotw_fld_dict.get_short_def("fire_co_type", fdwa2f3C.wa2f3cx.seg_side.fire_co_type + " " + fdwa2f3C.wa2f3cx.seg_side.fire_co_num)
            activeWorksheet.Cells(j, 46).Value = fdwa2f3C.wa2f3cx.seg_side.health_area.Substring(0, 2) + "." + fdwa2f3C.wa2f3cx.seg_side.health_area.Substring(2, 2)
            activeWorksheet.Cells(j, 46).NumberFormat = "00.00"
            activeWorksheet.Cells(j, 47).Value = fdwa2f3C.wa2f3cx.left_health_center_dist
            activeWorksheet.Cells(j, 48).Value = fdwa2f3C.wa2f3cx.dot_street_light_contract_area
            activeWorksheet.Cells(j, 49).Value = fdwa2f3C.wa2f3cx.seg_side.school_dist
            activeWorksheet.Cells(j, 50).Value = gotw_fld_dict.get_short_def("cd_eligible", fdwa2f3C.wa2f3cx.seg_side.iaei)
            activeWorksheet.Cells(j, 51).Value = fdwa2f3C.wa2f3cx.blockface_id
            activeWorksheet.Cells(j, 52).Value = fdwa2f3C.wa2f3cx.puma_code
            'police sector
            'activeWorksheet.Cells(j, 53).Value = fdwa2f3C.wa2f3cx.police_sector 

            'activeWorksheet.Cells(j, 53).Value = fdwa2f3.wa2f3x.right_side.boro
            'activeWorksheet.Cells(j, 54).Value = fdwa2f3.wa2f3x.right_side.boro + fdwa2f3.wa2f3x.right_side.comdist.district_number
            'activeWorksheet.Cells(j, 55).Value = fdwa2f3.wa2f3x.right_side.lhnd.Trim().TrimStart("0") + "-/-" + fdwa2f3.wa2f3x.right_side.hhnd.Trim().TrimStart("0")
            'activeWorksheet.Cells(j, 56).Value = fdwa2f3.wa2f3x.right_side.zip_code
            'activeWorksheet.Cells(j, 57).Value = fdwa2f3.wa2f3x.right_side.census_tract_2010
            'activeWorksheet.Cells(j, 58).Value = fdwa2f3.wa2f3x.right_side.census_block_2010
            'activeWorksheet.Cells(j, 59).Value = fdwa2f3.wa2f3x.right_side.census_tract_2000
            'activeWorksheet.Cells(j, 60).Value = fdwa2f3.wa2f3x.right_side.census_block_2000
            'activeWorksheet.Cells(j, 61).Value = gotw_fld_dict.get_short_def("police_patrol_boro", fdwa2f3.wa2f3x.right_side.police_patrol_boro)
            'activeWorksheet.Cells(j, 62).Value = fdwa2f3.wa2f3x.right_side.police_pct
            'activeWorksheet.Cells(j, 63).Value = fdwa2f3.wa2f3x.right_side.fire_div
            'activeWorksheet.Cells(j, 64).Value = fdwa2f3.wa2f3x.right_side.fire_bat.Trim().TrimStart("0")
            'activeWorksheet.Cells(j, 65).Value = gotw_fld_dict.get_short_def("fire_co_type", fdwa2f3.wa2f3x.right_side.fire_co_type + " " + fdwa2f3.wa2f3x.right_side.fire_co_num)
            'activeWorksheet.Cells(j, 66).Value = fdwa2f3.wa2f3x.right_side.health_area.Substring(0, 2) + "." + fdwa2f3.wa2f3x.right_side.health_area.Substring(2, 2)
            'activeWorksheet.Cells(j, 66).NumberFormat = "00.00"
            'activeWorksheet.Cells(j, 67).Value = fdwa2f3.wa2f3x.right_health_center_dist
            'activeWorksheet.Cells(j, 68).Value = fdwa2f3.wa2f3x.dot_street_light_contract_area
            'activeWorksheet.Cells(j, 69).Value = fdwa2f3.wa2f3x.right_side.school_dist
            'activeWorksheet.Cells(j, 70).Value = gotw_fld_dict.get_short_def("cd_eligible", fdwa2f3.wa2f3x.right_side.iaei)
            'activeWorksheet.Cells(j, 71).Value = fdwa2f3.wa2f3x.right_blockface_id
            'activeWorksheet.Cells(j, 72).Value = fdwa2f3.wa2f3x.right_puma_code

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-3 Errors")
            activeWorksheet.Cells(k, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_reason_code
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If
            activeWorksheet.Cells(k, 4).Value = fdwa1.in_stname1
            activeWorksheet.Cells(k, 5).Value = fdwa1.in_stname2
            activeWorksheet.Cells(k, 6).Value = fdwa1.in_stname3
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function

    Public Function WriteData3S(ByRef fdwa1 As Wa1, ByRef fdwa2f3S As Wa2F3s, i As Integer) As Integer
        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection

        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 1
        Static Dim k As Integer = 1
        Dim y As Integer = 9
        Dim xlCenter As Long = -4108
        Dim mywa1_dl1 As Wa1
        Dim gotw_fld_dict As New fld_dict

        fdconns = New GeoConnCollection("C:\Program Files\GeoExcel\GeoConns.xml")
        'fdconns = New GeoConnCollection("C:\temp\GeoConns.xml")
        fdgeo = New geo(fdconns)
        mywa1_dl1 = New Wa1
        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-3S Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "Borough"
            activeWorksheet.Cells(1, 4).Value = "On Street"
            activeWorksheet.Cells(1, 5).Value = "First Cross Street"
            activeWorksheet.Cells(1, 6).Value = "Second Cross Street"
            activeWorksheet.Range("A1:F1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:F1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-3S Output"
            AddFlag = True
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("A1:CNV1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:CNV1").Font.Color = RGB(255, 255, 255)

            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Borough"
            activeWorksheet.Cells(1, 3).value = "On Street"
            activeWorksheet.Cells(1, 4).Value = "Compass Direction 1"
            activeWorksheet.Cells(1, 5).Value = "First Cross Street"
            activeWorksheet.Cells(1, 6).Value = "Compass Direction 2"
            activeWorksheet.Cells(1, 7).Value = "Second Cross Street"
            activeWorksheet.Cells(1, 8).Value = "Number of Intersections"
            For x As Integer = 9 To 2409
                activeWorksheet.Cells(1, x).Value = "Intersecting Street"
                activeWorksheet.Cells(1, (x + 1)).Value = "2nd Intersecting Street"
                activeWorksheet.Cells(1, (x + 2)).Value = "Cross Street Count"
                activeWorksheet.Cells(1, (x + 3)).Value = "Number of Ft. from Previous Intersection"
                activeWorksheet.Cells(1, (x + 4)).Value = "Gap Flag"
                activeWorksheet.Cells(1, (x + 5)).Value = "Node ID"
                x = x + 5
            Next x

        Else
            activeWorksheet = activeWorkbook.Sheets("Func-3S Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet = activeWorkbook.Sheets("Func-3S Output")
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_grc2
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 3).Value = fdwa1.in_stname1
            activeWorksheet.Cells(j, 4).Value = fdwa1.in_compass_dir
            activeWorksheet.Cells(j, 5).Value = fdwa1.in_stname2
            activeWorksheet.Cells(j, 6).Value = fdwa1.in_compass_dir2
            activeWorksheet.Cells(j, 7).Value = fdwa1.in_stname3
            activeWorksheet.Cells(j, 8).Value = fdwa2f3S.num_of_intersections
            For x As Integer = 0 To (fdwa2f3S.num_of_intersections - 1)
                mywa1_dl1.Clear()
                mywa1_dl1.in_func_code = "DL"
                mywa1_dl1.in_platform_ind = "C"
                mywa1_dl1.out_b7sc_list(0) = fdwa2f3S.xstr_list(x).xstr_b7sc_list(0)
                If fdwa2f3S.xstr_list(x).xstr_cnt > 1 Then
                    mywa1_dl1.out_b7sc_list(1) = fdwa2f3S.xstr_list(x).xstr_b7sc_list(1)
                End If
                Call fdgeo.GeoCall(mywa1_dl1)

                activeWorksheet.Cells(j, (y)).Value = mywa1_dl1.out_stname_list(0)
                activeWorksheet.Cells(j, (y + 1)).Value = mywa1_dl1.out_stname_list(1)
                activeWorksheet.Cells(j, (y + 2)).Value = fdwa2f3S.xstr_list(x).xstr_cnt
                activeWorksheet.Cells(j, (y + 3)).Value = fdwa2f3S.xstr_list(x).distance
                activeWorksheet.Cells(j, (y + 4)).Value = fdwa2f3S.xstr_list(x).gap_flag
                activeWorksheet.Cells(j, (y + 5)).Value = fdwa2f3S.xstr_list(x).node_num
                y = y + 6
            Next x

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-3S Errors")
            activeWorksheet.Cells(k, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_reason_code
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 4).Value = fdwa1.in_stname1
            activeWorksheet.Cells(j, 5).Value = fdwa1.in_compass_dir
            activeWorksheet.Cells(j, 6).Value = fdwa1.in_stname2
            activeWorksheet.Cells(j, 7).Value = fdwa1.in_compass_dir2
            activeWorksheet.Cells(j, 8).Value = fdwa1.in_stname3
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function

    Public Function WriteDataBL(ByRef fdwa1 As Wa1, ByRef fdwa2f1a As Wa2F1a, i As Integer) As Integer
        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 1
        Static Dim k As Integer = 1
        Dim xlCenter As Long = -4108
        Dim gotw_fld_dict As New fld_dict

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-BL Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "Borough"
            activeWorksheet.Cells(1, 4).Value = "Block"
            activeWorksheet.Cells(1, 5).Value = "Lot"
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("A1:E1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:E1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-BL Output"
            AddFlag = True

            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Cells(1, 1).Value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).Value = "Borough"
            activeWorksheet.Cells(1, 3).Value = "Tax Block"
            activeWorksheet.Cells(1, 4).Value = "Tax Lot"

            activeWorksheet.Cells(1, 5).Value = "BBL"
            activeWorksheet.Cells(1, 6).Value = "Block Faces"
            activeWorksheet.Cells(1, 7).Value = "Sanborn Boro/Vol/Page"
            activeWorksheet.Cells(1, 8).Value = "X Coordinate"
            activeWorksheet.Cells(1, 9).Value = "Y Coordinate"
            activeWorksheet.Cells(1, 10).Value = "Latitude"
            activeWorksheet.Cells(1, 11).Value = "Longitude"
            activeWorksheet.Cells(1, 12).Value = "Vacant Lot"
            activeWorksheet.Cells(1, 13).Value = "Condo Lot"
            activeWorksheet.Cells(1, 14).Value = "Low BBL of Condo"
            activeWorksheet.Cells(1, 15).Value = "High BBL of Condo"
            activeWorksheet.Cells(1, 16).Value = "BIN"
            activeWorksheet.Cells(1, 17).Value = "BIN Status"
            activeWorksheet.Cells(1, 18).Value = "Corner Code"
            activeWorksheet.Cells(1, 19).Value = "Structures"
            activeWorksheet.Cells(1, 20).Value = "Business Improvement District"
            activeWorksheet.Cells(1, 21).Value = "RPAD SCC"
            activeWorksheet.Cells(1, 22).Value = "RPAD Building Class"
            activeWorksheet.Cells(1, 23).Value = "RPAD Interior Lot"
            activeWorksheet.Cells(1, 24).Value = "RPAD Irreg. Shaped Lot"
            activeWorksheet.Cells(1, 25).Value = "RPAD Condo Number"
            activeWorksheet.Cells(1, 26).Value = "RPAD Co-op Number"
            activeWorksheet.Cells(1, 27).Value = "Tax Map /Section /Volume"
            activeWorksheet.Cells(1, 28).Value = "DCP Zoning Map"
            activeWorksheet.Range("A1:AB1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:AB1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

        Else
            activeWorksheet = activeWorkbook.Sheets("Func-BL Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_reason_code
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 3).Value = fdwa2f1a.bbl.block
            activeWorksheet.Cells(j, 4).Value = fdwa2f1a.bbl.lot

            activeWorksheet.Cells(j, 5).Value = fdwa2f1a.bbl.BBLToString()
            activeWorksheet.Cells(j, 6).Value = fdwa2f1a.num_of_blockfaces
            activeWorksheet.Cells(j, 7).Value = fdwa2f1a.sanborn.boro + "/" + fdwa2f1a.sanborn.volume + fdwa2f1a.sanborn.volume_suffix + "/" + fdwa2f1a.sanborn.page + fdwa2f1a.sanborn.page_suffix
            activeWorksheet.Cells(j, 8).Value = Convert.ToString(Val(fdwa2f1a.x_coord))
            activeWorksheet.Cells(j, 9).Value = Convert.ToString(Val(fdwa2f1a.y_coord))
            activeWorksheet.Cells(j, 10).Value = fdwa2f1a.latitude
            activeWorksheet.Cells(j, 11).Value = fdwa2f1a.longitude
            activeWorksheet.Cells(j, 12).Value = fdwa2f1a.vacant_flag
            activeWorksheet.Cells(j, 13).Value = gotw_fld_dict.get_short_def("condo_flag", fdwa2f1a.condo_flag)
            If fdwa2f1a.condo_flag = "C" Then
                activeWorksheet.Cells(j, 14).Value = fdwa2f1a.condo_lo_bbl.boro + fdwa2f1a.condo_lo_bbl.block + fdwa2f1a.condo_lo_bbl.lot
                activeWorksheet.Cells(j, 15).Value = fdwa2f1a.condo_hi_bbl.boro + fdwa2f1a.condo_hi_bbl.block + fdwa2f1a.condo_hi_bbl.lot
            Else
                activeWorksheet.Cells(j, 14).Value = "N/A"
                activeWorksheet.Cells(j, 15).Value = "N/A"
            End If
            activeWorksheet.Cells(j, 16).Value = fdwa2f1a.bin.ToString()
            activeWorksheet.Cells(j, 17).Value = fdwa2f1a.TPAD_bin_status
            activeWorksheet.Cells(j, 18).Value = gotw_fld_dict.get_short_def("corner_code", fdwa2f1a.corner_code)
            activeWorksheet.Cells(j, 19).Value = fdwa2f1a.num_of_bldgs

            If fdwa2f1a.bid_id.B5scToString().Trim() = "" Then
                activeWorksheet.Cells(j, 20).Value = ""
            Else
                activeWorksheet.Cells(j, 20).Value = getStreetName(fdwa2f1a.bid_id.boro, fdwa2f1a.bid_id.B5scToString().Remove(0, 1))
            End If



            '            activeWorksheet.Cells(j, 18).Value = fdwa2f1a.bid_id.B5scToString().Trim()
            activeWorksheet.Cells(j, 21).Value = fdwa2f1a.rpad_scc
            activeWorksheet.Cells(j, 22).Value = fdwa2f1a.rpad_bldg_class

            If fdwa2f1a.interior_flag = " " Then
                activeWorksheet.Cells(j, 23).Value = "No"
            Else
                activeWorksheet.Cells(j, 23).Value = fdwa2f1a.interior_flag
            End If

            'activeWorksheet.Cells(j, 21).Value = fdwa2f1a.interior_flag

            If fdwa2f1a.irreg_flag = " " Then
                activeWorksheet.Cells(j, 24).Value = "No"
            Else
                activeWorksheet.Cells(j, 24).Value = fdwa2f1a.irreg_flag
            End If

            'activeWorksheet.Cells(j, 22).Value = fdwa2f1a.irreg_flag
            'activeWorksheet.Cells(j, 23).Value = fdwa2f1a.condo_num
            'activeWorksheet.Cells(j, 24).Value = fdwa2f1a.coop_num

            If Val(fdwa2f1a.condo_num) = 0 Or fdwa2f1a.condo_num = String.Empty Then
                activeWorksheet.Cells(j, 25).Value = "N/A"
            Else
                activeWorksheet.Cells(j, 25).Value = fdwa2f1a.condo_num
            End If

            If Val(fdwa2f1a.coop_num) = 0 Or fdwa2f1a.coop_num = String.Empty Then
                activeWorksheet.Cells(j, 26).Value = "N/A"
            Else
                activeWorksheet.Cells(j, 26).Value = fdwa2f1a.coop_num
            End If

            activeWorksheet.Cells(j, 27).Value = "'" + fdwa2f1a.dof_map.boro + " / " + fdwa2f1a.dof_map.section_volume.Remove(2, 2) + " / " + fdwa2f1a.dof_map.section_volume.Remove(0, 2)
            activeWorksheet.Cells(j, 28).Value = fdwa2f1a.DCP_Zoning_Map

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-BL Errors")
            activeWorksheet.Cells(k, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_grc2
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message
            If fdwa1.in_b10sc1.boro = "1" Then
                activeWorksheet.Cells(k, 3).Value = "Manhattan"
            ElseIf fdwa1.in_b10sc1.boro = "2" Then
                activeWorksheet.Cells(k, 3).Value = "Bronx"
            ElseIf fdwa1.in_b10sc1.boro = "3" Then
                activeWorksheet.Cells(k, 3).Value = "Brooklyn"
            ElseIf fdwa1.in_b10sc1.boro = "4" Then
                activeWorksheet.Cells(k, 3).Value = "Queens"
            ElseIf fdwa1.in_b10sc1.boro = "5" Then
                activeWorksheet.Cells(k, 3).Value = "Staten Island"
            End If
            activeWorksheet.Cells(k, 4).Value = fdwa1.in_bbl.block
            activeWorksheet.Cells(k, 5).Value = fdwa1.in_bbl.lot

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function
    Public Function WriteDataBN(ByRef fdwa1 As Wa1, ByRef fdwa2f1al As Wa2F1ax, i As Integer) As Integer
        Static AddFlag As Boolean = False
        Dim activeExcel As Excel.Application
        Dim activeWorkbook As Excel.Workbook = Nothing
        Dim activeWorksheet As Excel.Worksheet = Nothing
        Static Dim j As Integer = 1
        Static Dim k As Integer = 1
        Dim xlCenter As Long = -4108
        Dim gotw_fld_dict As New fld_dict

        activeExcel = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
        activeWorkbook = activeExcel.ActiveWorkbook
        If AddFlag = False Then
            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-BN Errors"
            activeWorksheet.Cells(1, 1).value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).value = "Error Message"
            activeWorksheet.Cells(1, 3).value = "BIN Number"
            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Range("A1:C1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:C1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

            activeWorksheet = activeWorkbook.Sheets.Add
            activeWorksheet.Name = "Func-BN Output"
            AddFlag = True

            activeWorksheet.Range("1:1").Font.Bold = True
            activeWorksheet.Cells(1, 1).Value = "Return Code/Reason Code"
            activeWorksheet.Cells(1, 2).Value = "Borough"
            activeWorksheet.Cells(1, 3).Value = "Tax Block"
            activeWorksheet.Cells(1, 4).Value = "Tax Lot"

            activeWorksheet.Cells(1, 5).Value = "BBL"
            activeWorksheet.Cells(1, 6).Value = "Block Faces"
            activeWorksheet.Cells(1, 7).Value = "Sanborn Boro/Vol/Page"
            activeWorksheet.Cells(1, 8).Value = "X Coordinate"
            activeWorksheet.Cells(1, 9).Value = "Y Coordinate"
            activeWorksheet.Cells(1, 10).Value = "Latitude"
            activeWorksheet.Cells(1, 11).Value = "Longitude"
            activeWorksheet.Cells(1, 12).Value = "Vacant Lot"
            activeWorksheet.Cells(1, 13).Value = "Condo Lot"
            activeWorksheet.Cells(1, 14).Value = "Low BBL of Condo"
            activeWorksheet.Cells(1, 15).Value = "High BBL of Condo"
            activeWorksheet.Cells(1, 16).Value = "BIN"
            activeWorksheet.Cells(1, 17).Value = "BIN Status"
            activeWorksheet.Cells(1, 18).Value = "Corner Code"
            activeWorksheet.Cells(1, 19).Value = "Structures"
            activeWorksheet.Cells(1, 20).Value = "Business Improvement District"
            activeWorksheet.Cells(1, 21).Value = "RPAD SCC"
            activeWorksheet.Cells(1, 22).Value = "RPAD Building Class"
            activeWorksheet.Cells(1, 23).Value = "RPAD Interior Lot"
            activeWorksheet.Cells(1, 24).Value = "RPAD Irreg. Shaped Lot"
            activeWorksheet.Cells(1, 25).Value = "RPAD Condo Number"
            activeWorksheet.Cells(1, 26).Value = "RPAD Co-op Number"
            activeWorksheet.Cells(1, 27).Value = "Tax Map /Section /Volume"
            activeWorksheet.Cells(1, 28).Value = "DCP Zoning Map"
            activeWorksheet.Range("A1:AB1").Interior.Color = RGB(0, 0, 0)
            activeWorksheet.Range("A1:AB1").Font.Color = RGB(255, 255, 255)
            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous

        Else
            activeWorksheet = activeWorkbook.Sheets("Func-BN Output")
        End If

        If fdwa1.out_grc = "00" Or fdwa1.out_grc = "01" Then
            j = j + 1
            activeWorksheet.Cells(j, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_grc2
            If fdwa2f1al.bbl.boro = "1" Then
                activeWorksheet.Cells(j, 2).Value = "Manhattan"
            ElseIf fdwa2f1al.bbl.boro = "2" Then
                activeWorksheet.Cells(j, 2).Value = "Bronx"
            ElseIf fdwa2f1al.bbl.boro = "3" Then
                activeWorksheet.Cells(j, 2).Value = "Brooklyn"
            ElseIf fdwa2f1al.bbl.boro = "4" Then
                activeWorksheet.Cells(j, 2).Value = "Queens"
            ElseIf fdwa2f1al.bbl.boro = "5" Then
                activeWorksheet.Cells(j, 2).Value = "Staten Island"
            End If
            activeWorksheet.Cells(j, 3).Value = fdwa2f1al.bbl.block
            activeWorksheet.Cells(j, 4).Value = fdwa2f1al.bbl.lot

            activeWorksheet.Cells(j, 5).Value = fdwa2f1al.bbl.BBLToString()
            activeWorksheet.Cells(j, 6).Value = fdwa2f1al.num_of_blockfaces
            activeWorksheet.Cells(j, 7).Value = fdwa2f1al.sanborn.boro + "/" + fdwa2f1al.sanborn.volume + fdwa2f1al.sanborn.volume_suffix + "/" + fdwa2f1al.sanborn.page + fdwa2f1al.sanborn.page_suffix
            activeWorksheet.Cells(j, 8).Value = Convert.ToString(Val(fdwa2f1al.x_coord))
            activeWorksheet.Cells(j, 9).Value = Convert.ToString(Val(fdwa2f1al.y_coord))
            activeWorksheet.Cells(j, 10).Value = fdwa2f1al.latitude
            activeWorksheet.Cells(j, 11).Value = fdwa2f1al.longitude
            activeWorksheet.Cells(j, 12).Value = fdwa2f1al.vacant_flag
            activeWorksheet.Cells(j, 13).Value = gotw_fld_dict.get_short_def("condo_flag", fdwa2f1al.condo_flag)
            If fdwa2f1al.condo_flag = "C" Then
                activeWorksheet.Cells(j, 14).Value = fdwa2f1al.condo_lo_bbl.boro + fdwa2f1al.condo_lo_bbl.block + fdwa2f1al.condo_lo_bbl.lot
                activeWorksheet.Cells(j, 15).Value = fdwa2f1al.condo_hi_bbl.boro + fdwa2f1al.condo_hi_bbl.block + fdwa2f1al.condo_hi_bbl.lot
            Else
                activeWorksheet.Cells(j, 14).Value = "N/A"
                activeWorksheet.Cells(j, 15).Value = "N/A"
            End If
            activeWorksheet.Cells(j, 16).Value = fdwa2f1al.bin.ToString()
            activeWorksheet.Cells(j, 17).Value = fdwa2f1al.TPAD_bin_status
            activeWorksheet.Cells(j, 18).Value = gotw_fld_dict.get_short_def("corner_code", fdwa2f1al.corner_code)
            activeWorksheet.Cells(j, 19).Value = fdwa2f1al.num_of_bldgs
            If fdwa2f1al.bid_id.B5scToString().Trim() = "" Then
                activeWorksheet.Cells(j, 20).Value = ""
            Else
                activeWorksheet.Cells(j, 20).Value = getStreetName(fdwa2f1al.bid_id.boro, fdwa2f1al.bid_id.B5scToString().Remove(0, 1))
            End If
            'activeWorksheet.Cells(j, 18).Value = fdwa2f1al.bid_id.B5scToString().Trim()
            activeWorksheet.Cells(j, 21).Value = fdwa2f1al.rpad_scc
            activeWorksheet.Cells(j, 22).Value = fdwa2f1al.rpad_bldg_class

            If fdwa2f1al.interior_flag = "" Then
                activeWorksheet.Cells(j, 23).Value = "No"
            Else
                activeWorksheet.Cells(j, 23).Value = fdwa2f1al.interior_flag
            End If


            'activeWorksheet.Cells(j, 21).Value = fdwa2f1al.interior_flag

            If fdwa2f1al.irreg_flag = "" Then
                activeWorksheet.Cells(j, 24).Value = "No"
            Else
                activeWorksheet.Cells(j, 24).Value = fdwa2f1al.irreg_flag
            End If

            'activeWorksheet.Cells(j, 22).Value = fdwa2f1al.irreg_flag

            If Val(fdwa2f1al.condo_num) = 0 Or fdwa2f1al.condo_num = String.Empty Then
                activeWorksheet.Cells(j, 25).Value = "N/A"
            Else
                activeWorksheet.Cells(j, 25).Value = fdwa2f1al.condo_num
            End If


            If Val(fdwa2f1al.coop_num) = 0 Or fdwa2f1al.coop_num = String.Empty Then
                activeWorksheet.Cells(j, 26).Value = "N/A"
            Else
                activeWorksheet.Cells(j, 26).Value = fdwa2f1al.coop_num
            End If


            'activeWorksheet.Cells(j, 23).Value = fdwa2f1al.condo_num
            'activeWorksheet.Cells(j, 24).Value = fdwa2f1al.coop_num
            activeWorksheet.Cells(j, 27).Value = "'" + fdwa2f1al.dof_map.boro + " / " + fdwa2f1al.dof_map.section_volume.Remove(2, 2) + " / " + fdwa2f1al.dof_map.section_volume.Remove(0, 2)
            activeWorksheet.Cells(j, 28).Value = fdwa2f1al.DCP_Zoning_Map

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Else
            k = k + 1
            activeWorksheet = activeWorkbook.Sheets("Func-BN Errors")
            activeWorksheet.Cells(k, 1).Value = fdwa1.out_grc + "/" + fdwa1.out_reason_code
            activeWorksheet.Cells(k, 2).Value = fdwa1.out_error_message
            activeWorksheet.Cells(k, 3).Value = fdwa1.in_bin.ToString()

            activeWorksheet.UsedRange.EntireColumn.AutoFit()
            activeWorksheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End If

        Return 0
    End Function
    Private Sub Func1B_Click(sender As Object, e As RibbonControlEventArgs) Handles Func1B.Click
        Dim Input_Form As Object

        Statusflag = "1B"
        'Func1B.Checked = False
        Func2.Checked = False
        Func3.Checked = False
        Func3S.Checked = False
        FuncBL.Checked = False
        FuncBN.Checked = False

        Input_Form = New Input_Load_Form
        Input_Form.Show()

    End Sub

    Private Sub Func2_Click(sender As Object, e As RibbonControlEventArgs) Handles Func2.Click
        Dim Form2 As Object

        Func1B.Checked = False
        'Func2.Checked = False
        Func3.Checked = False
        Func3S.Checked = False
        FuncBL.Checked = False
        FuncBN.Checked = False

        Form2 = New Input_Load_Form2
        Form2.Show()

    End Sub

    Private Sub Func3_Click(sender As Object, e As RibbonControlEventArgs) Handles Func3.Click
        Dim Form3 As Object

        Func1B.Checked = False
        Func2.Checked = False
        'Func3.Checked = False
        Func3S.Checked = False
        FuncBL.Checked = False
        FuncBN.Checked = False

        Form3 = New Input_Load_Form3
        Form3.Show()

    End Sub

    Private Sub Func3S_Click(sender As Object, e As RibbonControlEventArgs) Handles Func3S.Click
        Dim Form3S As Object

        Func1B.Checked = False
        Func2.Checked = False
        Func3.Checked = False
        'Func3S.Checked = False
        FuncBL.Checked = False
        FuncBN.Checked = False

        Form3S = New Input_Load_Form3S
        Form3S.Show()

    End Sub

    Private Sub FuncBL_Click(sender As Object, e As RibbonControlEventArgs) Handles FuncBL.Click
        Dim FormBL As Object

        Func1B.Checked = False
        Func2.Checked = False
        Func3.Checked = False
        Func3S.Checked = False
        'FuncBL.Checked = False
        FuncBN.Checked = False

        FormBL = New Input_Load_FormBL
        FormBL.Show()

    End Sub

    Private Sub FuncBN_Click(sender As Object, e As RibbonControlEventArgs) Handles FuncBN.Click
        Dim FormBN As Object

        Func1B.Checked = False
        Func2.Checked = False
        Func3.Checked = False
        Func3S.Checked = False
        FuncBL.Checked = False
        'FuncBN.Checked = False

        FormBN = New Input_Load_FormBN
        FormBN.Show()
    End Sub


    Public Function ExtractHouseNumberFromString(ByVal input_string As String)

        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim myfdwa1_hnd As New Wa1

        myfdwa1_hnd.Clear()
        myfdwa1_hnd.in_func_code = "D"
        myfdwa1_hnd.in_platform_ind = "C"
        myfdwa1_hnd.in_hns = input_string

        Try
            Call fdgeo.GeoCall(myfdwa1_hnd)
        Catch ex As Exception
            MsgBox("Error Occured at " + "Func - D" + " " + input_string)
        End Try

        Return myfdwa1_hnd.out_hnd

    End Function

    Public Function getStreetName(ByVal borough_code As Integer, ByVal street_code As String)


        Dim fdgeo As New geo
        Dim fdconn As New GeoConn
        Dim fdconns As New GeoConnCollection
        Dim mywa1 As New Wa1

        mywa1.Clear()
        mywa1.in_func_code = "D"
        mywa1.in_platform_ind = "C"
        mywa1.in_b10sc1.boro = borough_code.ToString()
        mywa1.in_b10sc1.sc5 = street_code
        fdgeo.GeoCall(mywa1)

        Return mywa1.out_stname1

    End Function


End Class
