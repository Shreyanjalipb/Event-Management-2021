Public Class ClassMain
#Region "Open Form"
    Public Shared Sub Func_Open_From(ByVal Frm As Form, Optional ByVal FrmMdi As Form = Nothing, Optional ByVal pFrm As String = "")
        If FrmMdi Is Nothing Then
            OpenDialog(Frm)
        Else
            OpenDialog(FrmMdi)
            OpenIn_Mdi(Frm, FrmMdi, pFrm)
        End If

    End Sub
    Shared Sub OpenDialog(ByVal Frm As Form)
        Frm.Show()
        'ผิด
        'Frm.WindowState = FormWindowState.Normal
        Frm.Focus()
    End Sub
    Shared Sub OpenIn_Mdi(ByVal Frm As Form, ByVal FrmMdi As Form, ByVal pFrm As String)
        Select Case pFrm
            Case "UserList"
                'With UserList
                '    .ListViewUser.Items.Clear()
                '    .txtSearch.Text = ""
                '    .rdoAll.Checked = False
                '    .rdoFname.Checked = False
                '    .rdoUname.Checked = False
                'End With
            Case "ScanSerail"
                'With ScanSerail 
                '    .cbxdate.Text = ""
                '    .CboOrder.Items.Clear()
                '    .CboOrder.Text = ""
                '    .txtscan.Text = ""
                'End With
            Case "KanbanList"
                'With KanbanList
                '    .ListKanban.Items.Clear()
                'End With
            Case "Order"
                'With Order
                '    .cboOrder.Text = ""
                '    .cboOrder.Items.Clear()
                '    .DateOrder.Text = ""
                '    .ListViewOrder.Items.Clear()
                'End With
            Case "SerialNumber"
                'With SerialNumber
                '    .ListKanban.Items.Clear()
                '    .Listserial.Items.Clear()
                'End With
            Case "frmUpdateManual"
                'With frmUpdateManual
                '    .CboOrder.Text = ""
                '    .CboOrder.Items.Clear()
                '    .cbxdate.Text = ""
                '    .ListOrderProcess.Items.Clear()
                '    .lvwshow.Items.Clear()
                'End With
            Case "Print"
                'With PrintPartTag
                '    .cboOrderNo.Items.Clear()
                '    .cboOrderNo.Text = ""
                '    .ListPrintPartTag.Items.Clear()
                '    .lblTotal.Text = ""
                'End With
        End Select
        Frm.MdiParent = FrmMdi
        Frm.Show()
        Frm.WindowState = FormWindowState.Maximized
    End Sub
#End Region
End Class
