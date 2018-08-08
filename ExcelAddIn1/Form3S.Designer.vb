<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3S
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
        Me.TextOnStreet = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextFirstCrossSt = New System.Windows.Forms.TextBox()
        Me.TextSecondCrossSt = New System.Windows.Forms.TextBox()
        Me.TextBoro = New System.Windows.Forms.TextBox()
        Me.Button3S = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.HasHeadersCheckboxFunction3S = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextCompassDirection1 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextCompassDirection2 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'TextOnStreet
        '
        Me.TextOnStreet.Location = New System.Drawing.Point(475, 138)
        Me.TextOnStreet.Name = "TextOnStreet"
        Me.TextOnStreet.Size = New System.Drawing.Size(100, 20)
        Me.TextOnStreet.TabIndex = 23
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(214, 141)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(167, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Specify the column with On Street"
        '
        'TextFirstCrossSt
        '
        Me.TextFirstCrossSt.Location = New System.Drawing.Point(475, 238)
        Me.TextFirstCrossSt.Name = "TextFirstCrossSt"
        Me.TextFirstCrossSt.Size = New System.Drawing.Size(100, 20)
        Me.TextFirstCrossSt.TabIndex = 25
        '
        'TextSecondCrossSt
        '
        Me.TextSecondCrossSt.Location = New System.Drawing.Point(475, 334)
        Me.TextSecondCrossSt.Name = "TextSecondCrossSt"
        Me.TextSecondCrossSt.Size = New System.Drawing.Size(100, 20)
        Me.TextSecondCrossSt.TabIndex = 27
        '
        'TextBoro
        '
        Me.TextBoro.Location = New System.Drawing.Point(475, 82)
        Me.TextBoro.Name = "TextBoro"
        Me.TextBoro.Size = New System.Drawing.Size(100, 20)
        Me.TextBoro.TabIndex = 22
        '
        'Button3S
        '
        Me.Button3S.Location = New System.Drawing.Point(302, 434)
        Me.Button3S.Name = "Button3S"
        Me.Button3S.Size = New System.Drawing.Size(178, 35)
        Me.Button3S.TabIndex = 31
        Me.Button3S.Text = "Process Data"
        Me.Button3S.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(216, 341)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(219, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Specify the column with Second Cross Street"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(214, 241)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(206, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Specify the column with First Crosss Street"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(214, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Specify the column with Borough information"
        '
        'HasHeadersCheckboxFunction3S
        '
        Me.HasHeadersCheckboxFunction3S.AutoSize = True
        Me.HasHeadersCheckboxFunction3S.Location = New System.Drawing.Point(302, 391)
        Me.HasHeadersCheckboxFunction3S.Name = "HasHeadersCheckboxFunction3S"
        Me.HasHeadersCheckboxFunction3S.Size = New System.Drawing.Size(179, 17)
        Me.HasHeadersCheckboxFunction3S.TabIndex = 28
        Me.HasHeadersCheckboxFunction3S.Text = "Input sheet has a Header record"
        Me.HasHeadersCheckboxFunction3S.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(216, 190)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(150, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Compass Direction 1 (optional)"
        '
        'TextCompassDirection1
        '
        Me.TextCompassDirection1.Location = New System.Drawing.Point(475, 187)
        Me.TextCompassDirection1.Name = "TextCompassDirection1"
        Me.TextCompassDirection1.Size = New System.Drawing.Size(100, 20)
        Me.TextCompassDirection1.TabIndex = 24
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(216, 292)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(150, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Compass Direction 2 (optional)"
        '
        'TextCompassDirection2
        '
        Me.TextCompassDirection2.Location = New System.Drawing.Point(475, 285)
        Me.TextCompassDirection2.Name = "TextCompassDirection2"
        Me.TextCompassDirection2.Size = New System.Drawing.Size(100, 20)
        Me.TextCompassDirection2.TabIndex = 26
        '
        'Form3S
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 491)
        Me.Controls.Add(Me.TextCompassDirection2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextCompassDirection1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.HasHeadersCheckboxFunction3S)
        Me.Controls.Add(Me.TextOnStreet)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextFirstCrossSt)
        Me.Controls.Add(Me.TextSecondCrossSt)
        Me.Controls.Add(Me.TextBoro)
        Me.Controls.Add(Me.Button3S)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form3S"
        Me.Text = "Please specify the Data Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextOnStreet As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextFirstCrossSt As System.Windows.Forms.TextBox
    Friend WithEvents TextSecondCrossSt As System.Windows.Forms.TextBox
    Friend WithEvents TextBoro As System.Windows.Forms.TextBox
    Friend WithEvents Button3S As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents HasHeadersCheckboxFunction3S As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextCompassDirection1 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextCompassDirection2 As System.Windows.Forms.TextBox
End Class
