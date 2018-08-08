<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
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
        Me.TextAddrNo = New System.Windows.Forms.TextBox()
        Me.TextStName = New System.Windows.Forms.TextBox()
        Me.TextBoro1 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextCompass = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBoro2 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.HasHeadersCheckboxFunction2 = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'TextAddrNo
        '
        Me.TextAddrNo.Location = New System.Drawing.Point(439, 133)
        Me.TextAddrNo.Name = "TextAddrNo"
        Me.TextAddrNo.Size = New System.Drawing.Size(100, 20)
        Me.TextAddrNo.TabIndex = 8
        '
        'TextStName
        '
        Me.TextStName.Location = New System.Drawing.Point(439, 275)
        Me.TextStName.Name = "TextStName"
        Me.TextStName.Size = New System.Drawing.Size(100, 20)
        Me.TextStName.TabIndex = 10
        '
        'TextBoro1
        '
        Me.TextBoro1.Location = New System.Drawing.Point(439, 60)
        Me.TextBoro1.Name = "TextBoro1"
        Me.TextBoro1.Size = New System.Drawing.Size(100, 20)
        Me.TextBoro1.TabIndex = 6
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(273, 425)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(178, 35)
        Me.Button2.TabIndex = 17
        Me.Button2.Text = "Process Data"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(178, 282)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(219, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Specify the column with Second Cross Street"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(180, 133)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(201, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Specify the column with First Cross Street"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(178, 63)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Specify the column with Borough information"
        '
        'TextCompass
        '
        Me.TextCompass.Location = New System.Drawing.Point(439, 348)
        Me.TextCompass.Name = "TextCompass"
        Me.TextCompass.Size = New System.Drawing.Size(100, 20)
        Me.TextCompass.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(178, 355)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(210, 13)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Specify the column with Compass Direction"
        '
        'TextBoro2
        '
        Me.TextBoro2.Location = New System.Drawing.Point(439, 203)
        Me.TextBoro2.Name = "TextBoro2"
        Me.TextBoro2.Size = New System.Drawing.Size(100, 20)
        Me.TextBoro2.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(178, 206)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(216, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Specify the column with Borough information"
        '
        'HasHeadersCheckboxFunction2
        '
        Me.HasHeadersCheckboxFunction2.AutoSize = True
        Me.HasHeadersCheckboxFunction2.Location = New System.Drawing.Point(316, 391)
        Me.HasHeadersCheckboxFunction2.Name = "HasHeadersCheckboxFunction2"
        Me.HasHeadersCheckboxFunction2.Size = New System.Drawing.Size(179, 17)
        Me.HasHeadersCheckboxFunction2.TabIndex = 15
        Me.HasHeadersCheckboxFunction2.Text = "Input sheet has a Header record"
        Me.HasHeadersCheckboxFunction2.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(758, 502)
        Me.Controls.Add(Me.HasHeadersCheckboxFunction2)
        Me.Controls.Add(Me.TextBoro2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextCompass)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextAddrNo)
        Me.Controls.Add(Me.TextStName)
        Me.Controls.Add(Me.TextBoro1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form2"
        Me.Text = "Please specify the Data Columns"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextAddrNo As System.Windows.Forms.TextBox
    Friend WithEvents TextStName As System.Windows.Forms.TextBox
    Friend WithEvents TextBoro1 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextCompass As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBoro2 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents HasHeadersCheckboxFunction2 As System.Windows.Forms.CheckBox
End Class
