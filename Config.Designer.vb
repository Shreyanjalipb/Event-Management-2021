<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Config
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtdbservername = New System.Windows.Forms.TextBox()
        Me.txtdbname = New System.Windows.Forms.TextBox()
        Me.txtusername = New System.Windows.Forms.TextBox()
        Me.txtpassword = New System.Windows.Forms.TextBox()
        Me.BtChang = New System.Windows.Forms.Button()
        Me.btcancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "DB SeverName :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "DB Name :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "UserName :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 124)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Password :"
        '
        'txtdbservername
        '
        Me.txtdbservername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdbservername.Location = New System.Drawing.Point(109, 18)
        Me.txtdbservername.Name = "txtdbservername"
        Me.txtdbservername.Size = New System.Drawing.Size(193, 20)
        Me.txtdbservername.TabIndex = 5
        '
        'txtdbname
        '
        Me.txtdbname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdbname.Location = New System.Drawing.Point(109, 57)
        Me.txtdbname.Name = "txtdbname"
        Me.txtdbname.Size = New System.Drawing.Size(193, 20)
        Me.txtdbname.TabIndex = 6
        '
        'txtusername
        '
        Me.txtusername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtusername.Location = New System.Drawing.Point(109, 92)
        Me.txtusername.Name = "txtusername"
        Me.txtusername.ReadOnly = True
        Me.txtusername.Size = New System.Drawing.Size(193, 20)
        Me.txtusername.TabIndex = 7
        '
        'txtpassword
        '
        Me.txtpassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpassword.Location = New System.Drawing.Point(109, 127)
        Me.txtpassword.Name = "txtpassword"
        Me.txtpassword.ReadOnly = True
        Me.txtpassword.Size = New System.Drawing.Size(193, 20)
        Me.txtpassword.TabIndex = 8
        '
        'BtChang
        '
        Me.BtChang.Location = New System.Drawing.Point(109, 164)
        Me.BtChang.Name = "BtChang"
        Me.BtChang.Size = New System.Drawing.Size(75, 38)
        Me.BtChang.TabIndex = 9
        Me.BtChang.Text = "Chang"
        Me.BtChang.UseVisualStyleBackColor = True
        '
        'btcancel
        '
        Me.btcancel.Location = New System.Drawing.Point(227, 164)
        Me.btcancel.Name = "btcancel"
        Me.btcancel.Size = New System.Drawing.Size(75, 38)
        Me.btcancel.TabIndex = 10
        Me.btcancel.Text = "Cancel"
        Me.btcancel.UseVisualStyleBackColor = True
        '
        'Config
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(349, 249)
        Me.Controls.Add(Me.btcancel)
        Me.Controls.Add(Me.BtChang)
        Me.Controls.Add(Me.txtpassword)
        Me.Controls.Add(Me.txtusername)
        Me.Controls.Add(Me.txtdbname)
        Me.Controls.Add(Me.txtdbservername)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Config"
        Me.Text = "Config"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtdbservername As System.Windows.Forms.TextBox
    Friend WithEvents txtdbname As System.Windows.Forms.TextBox
    Friend WithEvents txtusername As System.Windows.Forms.TextBox
    Friend WithEvents txtpassword As System.Windows.Forms.TextBox
    Friend WithEvents BtChang As System.Windows.Forms.Button
    Friend WithEvents btcancel As System.Windows.Forms.Button
End Class
