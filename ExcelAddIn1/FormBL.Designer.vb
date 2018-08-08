<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormBL
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBlock = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextLot = New System.Windows.Forms.TextBox()
        Me.TextBoro = New System.Windows.Forms.TextBox()
        Me.ButtonBL = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.HasHeadersCheckBoxFunctionBL = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'TextBlock
        '
        Me.TextBlock.Location = New System.Drawing.Point(488, 154)
        Me.TextBlock.Name = "TextBlock"
        Me.TextBlock.Size = New System.Drawing.Size(100, 20)
        Me.TextBlock.TabIndex = 32
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(229, 154)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(149, 13)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "Specify the column with Block"
        '
        'TextLot
        '
        Me.TextLot.Location = New System.Drawing.Point(488, 223)
        Me.TextLot.Name = "TextLot"
        Me.TextLot.Size = New System.Drawing.Size(100, 20)
        Me.TextLot.TabIndex = 33
        '
        'TextBoro
        '
        Me.TextBoro.Location = New System.Drawing.Point(488, 81)
        Me.TextBoro.Name = "TextBoro"
        Me.TextBoro.Size = New System.Drawing.Size(100, 20)
        Me.TextBoro.TabIndex = 31
        '
        'ButtonBL
        '
        Me.ButtonBL.Location = New System.Drawing.Point(322, 373)
        Me.ButtonBL.Name = "ButtonBL"
        Me.ButtonBL.Size = New System.Drawing.Size(178, 35)
        Me.ButtonBL.TabIndex = 38
        Me.ButtonBL.Text = "Process Data"
        Me.ButtonBL.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(229, 223)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(143, 13)
        Me.Label2.TabIndex = 36
        Me.Label2.Text = "Specify the column with LOT"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(227, 84)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Specify the column with Borough information"
        '
        'HasHeadersCheckBoxFunctionBL
        '
        Me.HasHeadersCheckBoxFunctionBL.AutoSize = True
        Me.HasHeadersCheckBoxFunctionBL.Location = New System.Drawing.Point(362, 311)
        Me.HasHeadersCheckBoxFunctionBL.Name = "HasHeadersCheckBoxFunctionBL"
        Me.HasHeadersCheckBoxFunctionBL.Size = New System.Drawing.Size(179, 17)
        Me.HasHeadersCheckBoxFunctionBL.TabIndex = 34
        Me.HasHeadersCheckBoxFunctionBL.Text = "Input sheet has a Header record"
        Me.HasHeadersCheckBoxFunctionBL.UseVisualStyleBackColor = True
        '
        'FormBL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(815, 489)
        Me.Controls.Add(Me.HasHeadersCheckBoxFunctionBL)
        Me.Controls.Add(Me.TextBlock)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextLot)
        Me.Controls.Add(Me.TextBoro)
        Me.Controls.Add(Me.ButtonBL)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FormBL"
        Me.Text = "Please specify the Data Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBlock As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextLot As System.Windows.Forms.TextBox
    Friend WithEvents TextBoro As System.Windows.Forms.TextBox
    Friend WithEvents ButtonBL As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents HasHeadersCheckBoxFunctionBL As System.Windows.Forms.CheckBox
End Class
