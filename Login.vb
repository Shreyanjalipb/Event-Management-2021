Public Class Login

    Private Sub Btnlogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnlogin.Click
        If Trim(txtUsername.Text) = "" Or Trim(txtPass.Text) = "" Then
            MsgBox("Please Insert Username and Password")
            Exit Sub
        End If
        myUser.UserName = Trim(txtUsername.Text)
        Dim mTable As Data.DataTable = Retrive(Trim(txtUsername.Text), Trim(txtPass.Text))
        If mTable.Rows.Count > 0 Then
            LoginId = Trim(mTable.Rows(0).Item("UserName"))
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Close()
        Else
            MsgBox("Invalid Username or Password")
        End If

    End Sub
    Function Retrive(ByVal username As String, ByVal password As String) As DataTable
        Dim SQL As String
        SQL = String.Concat("select * from [User] where UserName = '" & username & "' and Password = '" & password & "' and DeleteYN = '0'")
        Return UReadSqlbase(ConnectionString, SQL, "")
    End Function
    Private intReturnValue As Integer
    Private Sub Btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btncancel.Click
        intReturnValue = MsgBox("คุณต้องการออกจากโปรแกรมใช่ไหม ?", MsgBoxStyle.OkCancel + MsgBoxStyle.Information + MsgBoxStyle.SystemModal, "Message Box")
        If (intReturnValue = MsgBoxResult.Ok) Then
            Application.Exit()
        Else
            Exit Sub
        End If
    End Sub
    Private Sub Press_Login(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btnlogin_Click(sender, e)
        End If
    End Sub

   
End Class
