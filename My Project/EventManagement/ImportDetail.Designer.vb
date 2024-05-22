<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImportDetail
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
        Me.txtDuplicate = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.BtnClose = New System.Windows.Forms.Button()
        Me.txtdontImport = New System.Windows.Forms.TextBox()
        Me.lblDuplicate = New System.Windows.Forms.Label()
        Me.lblNo = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtDuplicate
        '
        Me.txtDuplicate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDuplicate.Location = New System.Drawing.Point(8, 45)
        Me.txtDuplicate.Multiline = True
        Me.txtDuplicate.Name = "txtDuplicate"
        Me.txtDuplicate.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDuplicate.Size = New System.Drawing.Size(443, 132)
        Me.txtDuplicate.TabIndex = 2
        Me.txtDuplicate.Text = " "
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Location = New System.Drawing.Point(1, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(462, 19)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Don't Import Record"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BtnClose
        '
        Me.BtnClose.Location = New System.Drawing.Point(365, 337)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(86, 31)
        Me.BtnClose.TabIndex = 6
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'txtdontImport
        '
        Me.txtdontImport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtdontImport.Location = New System.Drawing.Point(8, 210)
        Me.txtdontImport.Multiline = True
        Me.txtdontImport.Name = "txtdontImport"
        Me.txtdontImport.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtdontImport.Size = New System.Drawing.Size(443, 120)
        Me.txtdontImport.TabIndex = 7
        '
        'lblDuplicate
        '
        Me.lblDuplicate.AutoSize = True
        Me.lblDuplicate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDuplicate.Location = New System.Drawing.Point(4, 28)
        Me.lblDuplicate.Name = "lblDuplicate"
        Me.lblDuplicate.Size = New System.Drawing.Size(90, 13)
        Me.lblDuplicate.TabIndex = 8
        Me.lblDuplicate.Text = "Duplicate Record"
        '
        'lblNo
        '
        Me.lblNo.AutoSize = True
        Me.lblNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNo.Location = New System.Drawing.Point(5, 192)
        Me.lblNo.Name = "lblNo"
        Me.lblNo.Size = New System.Drawing.Size(102, 13)
        Me.lblNo.TabIndex = 9
        Me.lblNo.Text = "Don't Import Record"
        '
        'ImportDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(463, 379)
        Me.Controls.Add(Me.lblNo)
        Me.Controls.Add(Me.lblDuplicate)
        Me.Controls.Add(Me.txtdontImport)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDuplicate)
        Me.Name = "ImportDetail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ImportDetail"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtDuplicate As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents txtdontImport As System.Windows.Forms.TextBox
    Friend WithEvents lblDuplicate As System.Windows.Forms.Label
    Friend WithEvents lblNo As System.Windows.Forms.Label
End Class
