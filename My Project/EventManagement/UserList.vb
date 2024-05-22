Public Class UserList
    Private mDataname As String
    Private mSelect As String
    Private intReturnValue As Integer
    Private Sub BtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNew.Click
        myUser.CheckedUser = False
        Dim frm As New UserAdd
        myUser.Status = "N"
        frm.ShowDialog()
        Search()
    End Sub

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch.Click
        If Not rdoUname.Checked And Not rdoFname.Checked Then
            MsgBox("Please Select Username or Fullname")
            Exit Sub
        End If
        If rdoUname.Checked Then
            If Trim(txtSearch.Text) = "" Then
                MsgBox("Please Insert UserName")
                Exit Sub
            End If
        End If
        If rdoFname.Checked Then
            If Trim(txtSearch.Text) = "" Then
                MsgBox("Please Insert FullName")
                Exit Sub
            End If
        End If
        If rdoFname.Checked Then
            mDataname = Trim(txtSearch.Text)
            mSelect = "F"
        Else
            mDataname = Trim(txtSearch.Text)
            mSelect = "U"
        End If
        Search(mDataname, mSelect)
    End Sub
    Private Sub Search(Optional ByVal mData As String = "", Optional ByVal mselect As String = "", Optional ByVal selNdx As Integer = -1)
        Dim mTables As Data.DataTable = myUser.RetriveUser(mData, mselect)
        Dim i As Integer
        ListViewUser.Items.Clear()
        ListViewUser.View = View.Details
        ListViewUser.FullRowSelect = True
        ListViewUser.HideSelection = False
        If Not mTables Is Nothing Then
            If mTables.Rows.Count > 0 Then
                For i = 0 To mTables.Rows.Count - 1
                    ListViewUser.Items.Add(i + 1)
                    ListViewUser.Items(i).SubItems.Add(Trim(mTables.Rows(i).Item("Fullname")))
                    ListViewUser.Items(i).SubItems.Add(Trim(mTables.Rows(i).Item("Username")))
                    ListViewUser.Items(i).SubItems.Add(Trim(mTables.Rows(i).Item("Type")))
                Next
                If selNdx <> -1 Then
                    ListViewUser.Items(selNdx).Selected = True
                    ListViewUser.Items(selNdx).Focused = True
                End If
            Else
                MsgBox("ไม่มี User นี้")
            End If
        End If

    End Sub

    Private Sub BtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEdit.Click
        myUser.CheckedUser = True
        If ListViewUser.Items.Count > 0 Then
            Dim mSelectRow As Integer = ListViewUser.SelectedIndices.Count
            If mSelectRow > 0 Then
                Dim selNdx = ListViewUser.SelectedIndices(0)
                myUser.FullName = ListViewUser.SelectedItems(0).SubItems(1).Text
                myUser.UserName = ListViewUser.SelectedItems(0).SubItems(2).Text
                myUser.UserType = ListViewUser.SelectedItems(0).SubItems(3).Text
                Dim mTableUser As DataTable = myUser.RetriveUser(myUser.UserName, "U")
                If Not mTableUser Is Nothing Then
                    If mTableUser.Rows.Count > 0 Then
                        For i = 0 To mTableUser.Rows.Count - 1
                            myUser.PassWord = mTableUser.Rows(i).Item("Password")
                            myUser.ConFirmPass = myUser.PassWord
                        Next
                    End If
                End If
                myUser.Status = "E"
                Dim frm As New UserAdd
                frm.ShowDialog()
                Search("", "", selNdx)
            Else
                MsgBox("Please Select List")
            End If
        Else
            MsgBox("No Data in List")
        End If

    End Sub

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
        If ListViewUser.Items.Count > 0 Then
            Dim mSelectRow As Integer = ListViewUser.SelectedIndices.Count
            If mSelectRow > 0 Then
                intReturnValue = MsgBox("คุณต้องการลบข้อมูลในลิสใช่ไหม ?", MsgBoxStyle.OkCancel + MsgBoxStyle.Information + MsgBoxStyle.SystemModal, "Message Box")
                If (intReturnValue = MsgBoxResult.Ok) Then
                    Dim mselect As String = ListViewUser.SelectedItems(0).SubItems(2).Text
                    If myUser.UpdateYN(mselect) Then
                        MsgBox("Delete Complete")
                        ListViewUser.Items.Clear()
                        Search()
                    End If
                Else
                    Exit Sub
                End If
            Else
                MsgBox("Please Select List")
            End If
        Else
            MsgBox("No Data in List")
        End If
    End Sub
    Private Sub ShowAll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoAll.Click
        If rdoAll.Checked Then
            Search()
        End If
    End Sub

    Private Sub ShowName(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoFname.Click
        ListViewUser.Items.Clear()
        txtSearch.Text = ""
        txtSearch.Enabled = True
        BtnSearch.Enabled = True
    End Sub

    Private Sub ShowUser(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoUname.Click
        ListViewUser.Items.Clear()
        txtSearch.Text = ""
        txtSearch.Enabled = True
        BtnSearch.Enabled = True
    End Sub

End Class