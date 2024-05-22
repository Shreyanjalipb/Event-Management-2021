Imports System.Windows.Forms
Imports EventManagement.ClassMain
Public Class Main

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click, NewWindowToolStripMenuItem.Click
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteToolStripMenuItem.Click
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolBarToolStripMenuItem.Click
        Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles StatusBarToolStripMenuItem.Click
        Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub
    Private m_ChildFormNumber As Integer
    Private Sub Main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        'Dim dr As DialogResult
        'dr = Login.ShowDialog()
        'If dr = Windows.Forms.DialogResult.OK Then
        '    'set name to title of program.
        '    Me.Text = "HINO : Log in by " & LoginId
        'Else
        '    Close()
        'End If
        LangFormat = System.Threading.Thread.CurrentThread.CurrentCulture.ToString()
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")


    End Sub
    Private Sub ImgUserList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Func_Open_From(UserList, Me)
    End Sub

    Private Sub ImgUserAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Func_Open_From(UserAdd, Me)
    End Sub

    Private Sub ImgPwd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BtnMaster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnMaster.Click
        'age_kop
        Import.clearData()
        Func_Open_From(Import, Me, "Import")
    End Sub

    Private Sub BtnUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Func_Open_From(UserList, Me, "UserList")
    End Sub

    'Private Sub BtnSerial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Func_Open_From(SerialNumber, Me, "SerialNumber")
    'End Sub

    'Private Sub BtnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Func_Open_From(Order, Me, "Order")
    'End Sub

    'Private Sub BtnCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Func_Open_From(ScanSerail, Me, "ScanSerail")
    'End Sub

   
    'Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Func_Open_From(frmUpdateManual, Me, "frmUpdateManual")
    'End Sub

    Private Sub BtnConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  Func_Open_From(PrintDelivery, Me)
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Me.Close()
    End Sub

    Private Sub btconfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btconfig.Click
        Func_Open_From(Config, Me, "config")
    End Sub

    Private Sub leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusLabel1.MouseLeave

        StatusLabel1.ForeColor = Color.CadetBlue
    End Sub

    Private Sub move(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles StatusLabel1.MouseMove
        StatusLabel1.ForeColor = Color.SteelBlue
    End Sub

    Private Sub StatusLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusLabel1.Click
        System.Diagnostics.Process.Start("mailto:apiwat@royalcliff.com?subject=")
    End Sub
End Class
