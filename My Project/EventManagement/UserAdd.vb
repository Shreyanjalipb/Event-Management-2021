Imports EventManagement.ClassMain
Imports EventManagement.ClassUser
Public Class UserAdd
    Private Sub UserAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboUserType.Items.Insert(0, "Admin")
        cboUserType.Items.Insert(1, "User")
        If myUser.CheckedUser Then
            txtUsername.Enabled = False
            txtFullName.Text = Trim(myUser.FullName)
            txtUsername.Text = Trim(myUser.UserName)
            txtPassword.Text = Trim(myUser.PassWord)
            txtConfirm.Text = Trim(myUser.ConFirmPass)
            cboUserType.SelectedItem = Trim(myUser.UserType)
            myUser.CheckedUser = False
        Else
            txtUsername.Enabled = True
        End If
    End Sub

    
    Private Sub txtPass(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        If myUser.Status = "N" Then
            Dim mTable As Data.DataTable = myUser.RetriveAllUser(Trim(txtUsername.Text))
            If mTable.Rows.Count > 0 Then
                lblChecked.Text = "User นี้มีผู้ใช้แล้ว"
                myUser.UserDup = True
            Else
                lblChecked.Text = "สามารถใช้ User นี้ได้"
                myUser.UserDup = False
            End If
        End If
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Dim mErrorMsg As String = ""
        Dim mUserType As Integer
        myUser.FullName = Trim(txtFullName.Text)
        myUser.UserName = Trim(txtUsername.Text)
        myUser.PassWord = Trim(txtPassword.Text)
        myUser.ConFirmPass = Trim(txtConfirm.Text)
        mUserType = cboUserType.SelectedIndex
        If mUserType < 0 Then
            myUser.UserType = ""
        Else
            Select Case mUserType
                Case 0
                    myUser.UserType = "Admin"
                Case 1
                    myUser.UserType = "User"
            End Select
        End If
        myUser.ConFirmPass = Trim(txtConfirm.Text)
        If myUser.FullName = "" Then
            mErrorMsg = String.Concat(mErrorMsg, "Please Insert FullName", vbCrLf)
        End If
        If myUser.UserName = "" Then
            mErrorMsg = String.Concat(mErrorMsg, "Please Insert UserName", vbCrLf)
        End If
        If myUser.PassWord = "" Then
            mErrorMsg = String.Concat(mErrorMsg, "Please Insert PassWord", vbCrLf)
        End If
        If myUser.ConFirmPass = "" Then
            mErrorMsg = String.Concat(mErrorMsg, "Please Insert ConfirmPassWord", vbCrLf)
        End If

        If myUser.UserType = "" Then
            mErrorMsg = String.Concat(mErrorMsg, "Please Insert UserType", vbCrLf)
        End If
        If mErrorMsg <> "" Then
            MsgBox(mErrorMsg)
            Exit Sub
        End If
        If myUser.PassWord <> myUser.ConFirmPass Then
            MsgBox("Invalid Password")
        ElseIf Not myUser.UserDup Then
            If myUser.CheckAddUser(Trim(txtFullName.Text), Trim(txtUsername.Text), Trim(txtPassword.Text), Trim(myUser.UserType)) Then
                MsgBox("Save Complete")
                Close()
            Else
                If myUser.Save() Then
                    MsgBox("Save Complete")
                    Close()
                End If
            End If

           
        Else
            MsgBox("Dupplicate User")
            myUser.UserDup = False
        End If
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Close()
    End Sub
End Class