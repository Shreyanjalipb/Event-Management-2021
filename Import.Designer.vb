<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Import
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
        Me.txtpath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Btnbrown = New System.Windows.Forms.Button()
        Me.Btnok = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSumnew = New System.Windows.Forms.Label()
        Me.lblduplicate = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblNotImportRec = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.BtnDetail1 = New System.Windows.Forms.Button()
        Me.lblstate = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtpath
        '
        Me.txtpath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpath.Location = New System.Drawing.Point(79, 20)
        Me.txtpath.Name = "txtpath"
        Me.txtpath.Size = New System.Drawing.Size(287, 20)
        Me.txtpath.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "File Name"
        '
        'Btnbrown
        '
        Me.Btnbrown.Location = New System.Drawing.Point(372, 17)
        Me.Btnbrown.Name = "Btnbrown"
        Me.Btnbrown.Size = New System.Drawing.Size(69, 24)
        Me.Btnbrown.TabIndex = 2
        Me.Btnbrown.Text = "Brown..."
        Me.Btnbrown.UseVisualStyleBackColor = True
        '
        'Btnok
        '
        Me.Btnok.Location = New System.Drawing.Point(185, 46)
        Me.Btnok.Name = "Btnok"
        Me.Btnok.Size = New System.Drawing.Size(76, 29)
        Me.Btnok.TabIndex = 3
        Me.Btnok.Text = "OK"
        Me.Btnok.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "New Record"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Duplicate Record"
        '
        'lblSumnew
        '
        Me.lblSumnew.AutoSize = True
        Me.lblSumnew.Location = New System.Drawing.Point(129, 34)
        Me.lblSumnew.Name = "lblSumnew"
        Me.lblSumnew.Size = New System.Drawing.Size(13, 13)
        Me.lblSumnew.TabIndex = 7
        Me.lblSumnew.Text = "0"
        '
        'lblduplicate
        '
        Me.lblduplicate.AutoSize = True
        Me.lblduplicate.Location = New System.Drawing.Point(129, 60)
        Me.lblduplicate.Name = "lblduplicate"
        Me.lblduplicate.Size = New System.Drawing.Size(13, 13)
        Me.lblduplicate.TabIndex = 8
        Me.lblduplicate.Text = "0"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 85)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Don't Import Record"
        '
        'lblNotImportRec
        '
        Me.lblNotImportRec.AutoSize = True
        Me.lblNotImportRec.Location = New System.Drawing.Point(129, 85)
        Me.lblNotImportRec.Name = "lblNotImportRec"
        Me.lblNotImportRec.Size = New System.Drawing.Size(13, 13)
        Me.lblNotImportRec.TabIndex = 10
        Me.lblNotImportRec.Text = "0"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.BtnDetail1)
        Me.Panel1.Controls.Add(Me.lblstate)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.lblNotImportRec)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblduplicate)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblSumnew)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(12, 91)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(429, 182)
        Me.Panel1.TabIndex = 11
        '
        'BtnDetail1
        '
        Me.BtnDetail1.Location = New System.Drawing.Point(351, 79)
        Me.BtnDetail1.Name = "BtnDetail1"
        Me.BtnDetail1.Size = New System.Drawing.Size(55, 25)
        Me.BtnDetail1.TabIndex = 17
        Me.BtnDetail1.Text = "Detail"
        Me.BtnDetail1.UseVisualStyleBackColor = True
        '
        'lblstate
        '
        Me.lblstate.AutoSize = True
        Me.lblstate.BackColor = System.Drawing.Color.AliceBlue
        Me.lblstate.Location = New System.Drawing.Point(393, 110)
        Me.lblstate.Name = "lblstate"
        Me.lblstate.Size = New System.Drawing.Size(13, 13)
        Me.lblstate.TabIndex = 15
        Me.lblstate.Text = "0"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 110)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(381, 20)
        Me.ProgressBar1.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label8.Location = New System.Drawing.Point(1, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(428, 24)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Status"
        '
        'Label4
        '
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label4.Location = New System.Drawing.Point(1, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(428, 162)
        Me.Label4.TabIndex = 16
        '
        'Import
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(469, 323)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Btnok)
        Me.Controls.Add(Me.Btnbrown)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtpath)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Import"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import Data"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtpath As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Btnbrown As System.Windows.Forms.Button
    Friend WithEvents Btnok As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblSumnew As System.Windows.Forms.Label
    Friend WithEvents lblduplicate As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblNotImportRec As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblstate As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtnDetail1 As System.Windows.Forms.Button
End Class
