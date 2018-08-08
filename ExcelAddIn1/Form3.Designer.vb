<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.TextFirstCrossSt = New System.Windows.Forms.TextBox()
        Me.TextSecondCrossSt = New System.Windows.Forms.TextBox()
        Me.TextBoro = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextOnStreet = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.HasHeadersCheckboxFunction3 = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextSideOfStreet = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'TextFirstCrossSt
        '
        Me.TextFirstCrossSt.Location = New System.Drawing.Point(433, 177)
        Me.TextFirstCrossSt.Name = "TextFirstCrossSt"
        Me.TextFirstCrossSt.Size = New System.Drawing.Size(100, 20)
        Me.TextFirstCrossSt.TabIndex = 2
        '
        'TextSecondCrossSt
        '
        Me.TextSecondCrossSt.Location = New System.Drawing.Point(433, 230)
        Me.TextSecondCrossSt.Name = "TextSecondCrossSt"
        Me.TextSecondCrossSt.Size = New System.Drawing.Size(100, 20)
        Me.TextSecondCrossSt.TabIndex = 3
        '
        'TextBoro
        '
        Me.TextBoro.Location = New System.Drawing.Point(433, 35)
        Me.TextBoro.Name = "TextBoro"
        Me.TextBoro.Size = New System.Drawing.Size(100, 20)
        Me.TextBoro.TabIndex = 0
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(256, 381)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(178, 35)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Process Data"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(174, 233)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(219, 13)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Specify the column with Second Cross Street"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(174, 177)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(206, 13)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Specify the column with First Crosss Street"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(172, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Specify the column with Borough information"
        '
        'TextOnStreet
        '
        Me.TextOnStreet.Location = New System.Drawing.Point(433, 108)
        Me.TextOnStreet.Name = "TextOnStreet"
        Me.TextOnStreet.Size = New System.Drawing.Size(100, 20)
        Me.TextOnStreet.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(174, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(167, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Specify the column with On Street"
        '
        'HasHeadersCheckboxFunction3
        '
        Me.HasHeadersCheckboxFunction3.AutoSize = True
        Me.HasHeadersCheckboxFunction3.Location = New System.Drawing.Point(293, 326)
        Me.HasHeadersCheckboxFunction3.Name = "HasHeadersCheckboxFunction3"
        Me.HasHeadersCheckboxFunction3.Size = New System.Drawing.Size(179, 17)
        Me.HasHeadersCheckboxFunction3.TabIndex = 5
        Me.HasHeadersCheckboxFunction3.Text = "Input sheet has a Header record"
        Me.HasHeadersCheckboxFunction3.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(174, 280)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(188, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Specify the column with Side Of Street"
        '
        'TextSideOfStreet
        '
        Me.TextSideOfStreet.Location = New System.Drawing.Point(433, 273)
        Me.TextSideOfStreet.Name = "TextSideOfStreet"
        Me.TextSideOfStreet.Size = New System.Drawing.Size(100, 20)
        Me.TextSideOfStreet.TabIndex = 23
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(704, 449)
        Me.Controls.Add(Me.TextSideOfStreet)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.HasHeadersCheckboxFunction3)
        Me.Controls.Add(Me.TextOnStreet)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextFirstCrossSt)
        Me.Controls.Add(Me.TextSecondCrossSt)
        Me.Controls.Add(Me.TextBoro)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form3"
        Me.Text = "Please specify the Data Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextFirstCrossSt As System.Windows.Forms.TextBox
    Friend WithEvents TextSecondCrossSt As System.Windows.Forms.TextBox
    Friend WithEvents TextBoro As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextOnStreet As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents HasHeadersCheckboxFunction3 As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextSideOfStreet As System.Windows.Forms.TextBox
End Class
