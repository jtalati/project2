Partial Class GeoXRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Label11 = Me.Factory.CreateRibbonLabel
        Me.Func1B = Me.Factory.CreateRibbonCheckBox
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.Label8 = Me.Factory.CreateRibbonLabel
        Me.Func2 = Me.Factory.CreateRibbonCheckBox
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Label12 = Me.Factory.CreateRibbonLabel
        Me.Func3 = Me.Factory.CreateRibbonCheckBox
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.Label9 = Me.Factory.CreateRibbonLabel
        Me.Func3S = Me.Factory.CreateRibbonCheckBox
        Me.Group8 = Me.Factory.CreateRibbonGroup
        Me.Label10 = Me.Factory.CreateRibbonLabel
        Me.FuncBL = Me.Factory.CreateRibbonCheckBox
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Label3 = Me.Factory.CreateRibbonLabel
        Me.FuncBN = Me.Factory.CreateRibbonCheckBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Label5 = Me.Factory.CreateRibbonLabel
        Me.Label6 = Me.Factory.CreateRibbonLabel
        Me.Label7 = Me.Factory.CreateRibbonLabel
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Label4 = Me.Factory.CreateRibbonLabel
        Me.Process = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.Group8.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group7)
        Me.Tab1.Groups.Add(Me.Group8)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "GEOSUPPORT"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Label11)
        Me.Group1.Items.Add(Me.Func1B)
        Me.Group1.Items.Add(Me.Label1)
        Me.Group1.Name = "Group1"
        '
        'Label11
        '
        Me.Label11.Label = "  "
        Me.Label11.Name = "Label11"
        '
        'Func1B
        '
        Me.Func1B.Label = "Function 1B"
        Me.Func1B.Name = "Func1B"
        '
        'Label1
        '
        Me.Label1.Label = "  "
        Me.Label1.Name = "Label1"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.Label8)
        Me.Group6.Items.Add(Me.Func2)
        Me.Group6.Name = "Group6"
        '
        'Label8
        '
        Me.Label8.Label = "  "
        Me.Label8.Name = "Label8"
        '
        'Func2
        '
        Me.Func2.Label = "Function 2"
        Me.Func2.Name = "Func2"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Label12)
        Me.Group2.Items.Add(Me.Func3)
        Me.Group2.Items.Add(Me.Label2)
        Me.Group2.Name = "Group2"
        '
        'Label12
        '
        Me.Label12.Label = "  "
        Me.Label12.Name = "Label12"
        '
        'Func3
        '
        Me.Func3.Label = "Function 3"
        Me.Func3.Name = "Func3"
        '
        'Label2
        '
        Me.Label2.Label = "  "
        Me.Label2.Name = "Label2"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.Label9)
        Me.Group7.Items.Add(Me.Func3S)
        Me.Group7.Name = "Group7"
        '
        'Label9
        '
        Me.Label9.Label = "  "
        Me.Label9.Name = "Label9"
        '
        'Func3S
        '
        Me.Func3S.Label = "Function 3S"
        Me.Func3S.Name = "Func3S"
        '
        'Group8
        '
        Me.Group8.Items.Add(Me.Label10)
        Me.Group8.Items.Add(Me.FuncBL)
        Me.Group8.Name = "Group8"
        '
        'Label10
        '
        Me.Label10.Label = "  "
        Me.Label10.Name = "Label10"
        '
        'FuncBL
        '
        Me.FuncBL.Label = "Function BL"
        Me.FuncBL.Name = "FuncBL"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Label3)
        Me.Group3.Items.Add(Me.FuncBN)
        Me.Group3.Name = "Group3"
        '
        'Label3
        '
        Me.Label3.Label = "  "
        Me.Label3.Name = "Label3"
        '
        'FuncBN
        '
        Me.FuncBN.Label = "Function BN"
        Me.FuncBN.Name = "FuncBN"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Label5)
        Me.Group4.Items.Add(Me.Label6)
        Me.Group4.Items.Add(Me.Label7)
        Me.Group4.Name = "Group4"
        '
        'Label5
        '
        Me.Label5.Label = "---------->             "
        Me.Label5.Name = "Label5"
        '
        'Label6
        '
        Me.Label6.Label = "---------->"
        Me.Label6.Name = "Label6"
        '
        'Label7
        '
        Me.Label7.Label = "---------->"
        Me.Label7.Name = "Label7"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Label4)
        Me.Group5.Items.Add(Me.Process)
        Me.Group5.Name = "Group5"
        '
        'Label4
        '
        Me.Label4.Label = " "
        Me.Label4.Name = "Label4"
        '
        'Process
        '
        Me.Process.Image = Global.ExcelAddIn1.My.Resources.Resources.geosupport
        Me.Process.Label = "Submit"
        Me.Process.Name = "Process"
        Me.Process.ShowImage = True
        '
        'GeoXRibbon
        '
        Me.Name = "GeoXRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.Group8.ResumeLayout(False)
        Me.Group8.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Func1B As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Func2 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Func3 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Func3S As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents FuncBL As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Label3 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents FuncBN As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label4 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Process As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Label5 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label6 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label7 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label11 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label8 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label12 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label9 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group8 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label10 As Microsoft.Office.Tools.Ribbon.RibbonLabel
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As GeoXRibbon
        Get
            Return Me.GetRibbon(Of GeoXRibbon)()
        End Get
    End Property
End Class
