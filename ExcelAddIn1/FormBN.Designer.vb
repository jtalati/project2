<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormBN
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
        Me.TextBIN = New System.Windows.Forms.TextBox()
        Me.ButtonBN = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.HasHeadersCheckBoxFunctionBN = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'TextBIN
        '
        Me.TextBIN.Location = New System.Drawing.Point(406, 76)
        Me.TextBIN.Name = "TextBIN"
        Me.TextBIN.Size = New System.Drawing.Size(100, 20)
        Me.TextBIN.TabIndex = 36
        '
        'ButtonBN
        '
        Me.ButtonBN.Location = New System.Drawing.Point(240, 215)
        Me.ButtonBN.Name = "ButtonBN"
        Me.ButtonBN.Size = New System.Drawing.Size(178, 35)
        Me.ButtonBN.TabIndex = 39
        Me.ButtonBN.Text = "Process Data"
        Me.ButtonBN.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(145, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(189, 13)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "Specify the column with BIN Informatio"
        '
        'HasHeadersCheckBoxFunctionBN
        '
        Me.HasHeadersCheckBoxFunctionBN.AutoSize = True
        Me.HasHeadersCheckBoxFunctionBN.Location = New System.Drawing.Point(282, 154)
        Me.HasHeadersCheckBoxFunctionBN.Name = "HasHeadersCheckBoxFunctionBN"
        Me.HasHeadersCheckBoxFunctionBN.Size = New System.Drawing.Size(179, 17)
        Me.HasHeadersCheckBoxFunctionBN.TabIndex = 37
        Me.HasHeadersCheckBoxFunctionBN.Text = "Input sheet has a Header record"
        Me.HasHeadersCheckBoxFunctionBN.UseVisualStyleBackColor = True
        '
        'FormBN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(629, 316)
        Me.Controls.Add(Me.HasHeadersCheckBoxFunctionBN)
        Me.Controls.Add(Me.TextBIN)
        Me.Controls.Add(Me.ButtonBN)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FormBN"
        Me.Text = "Please specify the Data Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBIN As System.Windows.Forms.TextBox
    Friend WithEvents ButtonBN As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents HasHeadersCheckBoxFunctionBN As System.Windows.Forms.CheckBox
End Class
