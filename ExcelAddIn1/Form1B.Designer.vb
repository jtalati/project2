<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1B
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1B = New System.Windows.Forms.Button()
        Me.TextBoro = New System.Windows.Forms.TextBox()
        Me.TextStName = New System.Windows.Forms.TextBox()
        Me.TextAddrNo = New System.Windows.Forms.TextBox()
        Me.HasHeadersCheckboxFunction1B = New System.Windows.Forms.CheckBox()
        Me.TextBox_Unit = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(126, 57)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Specify the column with Borough information"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(126, 137)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(200, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Specify the column with Address Number"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(126, 218)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(254, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Specify the column with Street Name or Place Name"
        '
        'Button1B
        '
        Me.Button1B.Location = New System.Drawing.Point(224, 400)
        Me.Button1B.Name = "Button1B"
        Me.Button1B.Size = New System.Drawing.Size(178, 35)
        Me.Button1B.TabIndex = 6
        Me.Button1B.Text = "Process Data"
        Me.Button1B.UseVisualStyleBackColor = True
        '
        'TextBoro
        '
        Me.TextBoro.Location = New System.Drawing.Point(387, 54)
        Me.TextBoro.Name = "TextBoro"
        Me.TextBoro.Size = New System.Drawing.Size(100, 20)
        Me.TextBoro.TabIndex = 0
        '
        'TextStName
        '
        Me.TextStName.Location = New System.Drawing.Point(387, 211)
        Me.TextStName.Name = "TextStName"
        Me.TextStName.Size = New System.Drawing.Size(100, 20)
        Me.TextStName.TabIndex = 2
        '
        'TextAddrNo
        '
        Me.TextAddrNo.Location = New System.Drawing.Point(385, 137)
        Me.TextAddrNo.Name = "TextAddrNo"
        Me.TextAddrNo.Size = New System.Drawing.Size(100, 20)
        Me.TextAddrNo.TabIndex = 1
        '
        'HasHeadersCheckboxFunction1B
        '
        Me.HasHeadersCheckboxFunction1B.AutoSize = True
        Me.HasHeadersCheckboxFunction1B.Location = New System.Drawing.Point(271, 355)
        Me.HasHeadersCheckboxFunction1B.Name = "HasHeadersCheckboxFunction1B"
        Me.HasHeadersCheckboxFunction1B.Size = New System.Drawing.Size(179, 17)
        Me.HasHeadersCheckboxFunction1B.TabIndex = 4
        Me.HasHeadersCheckboxFunction1B.Text = "Input sheet has a Header record"
        Me.HasHeadersCheckboxFunction1B.UseVisualStyleBackColor = True
        '
        'TextBox_Unit
        '
        Me.TextBox_Unit.Enabled = False
        Me.TextBox_Unit.Location = New System.Drawing.Point(387, 287)
        Me.TextBox_Unit.Name = "TextBox_Unit"
        Me.TextBox_Unit.Size = New System.Drawing.Size(100, 20)
        Me.TextBox_Unit.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(276, 290)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(66, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Unit Number"
        '
        'Form1B
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(684, 479)
        Me.Controls.Add(Me.TextBox_Unit)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.HasHeadersCheckboxFunction1B)
        Me.Controls.Add(Me.TextAddrNo)
        Me.Controls.Add(Me.TextStName)
        Me.Controls.Add(Me.TextBoro)
        Me.Controls.Add(Me.Button1B)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1B"
        Me.Text = "Please specify the Data Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button1B As System.Windows.Forms.Button
    Friend WithEvents TextBoro As System.Windows.Forms.TextBox
    Friend WithEvents TextStName As System.Windows.Forms.TextBox
    Friend WithEvents TextAddrNo As System.Windows.Forms.TextBox
    Friend WithEvents HasHeadersCheckboxFunction1B As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox_Unit As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
