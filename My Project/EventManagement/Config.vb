Public Class Config

    Private Sub btcancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btcancel.Click
        Me.Close()
    End Sub

    Private Sub Config_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If MyConfigClass.LoadConFig() Then
        txtdbservername.Text = MyConfigClass.SqlServerName
        txtdbname.Text = MyConfigClass.SqlDBName
        txtusername.Text = MyConfigClass.SqlUsername
        txtpassword.Text = MyConfigClass.SqlPassword
        'End If
    End Sub

    Private Sub BtChang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtChang.Click

        MyConfigClass.SqlServerName = txtdbservername.Text
        MyConfigClass.SqlDBName = txtdbname.Text
        MyConfigClass.SqlUsername = txtusername.Text
        MyConfigClass.SqlPassword = txtpassword.Text
        'If MyConfigClass.Save(MyConfigClass.Id) Then
        UInitialClass()
        MsgBox("Changes successfully.")
        Me.Close()
        'End If
    End Sub
End Class