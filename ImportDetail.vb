Public Class ImportDetail

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

   
    Private Sub ImportDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'If DataNotImport <> "" Then
        '    Dim Alldatas() As String = DataNotImport.Split(",")
        '    Dim mvalue As String = ""
        '    If Alldatas.Length > 0 Then
        '        For i = 0 To Alldatas.Length - 1
        '            mvalue = mvalue + "Contact : " + Alldatas(i) + " " + vbCrLf
        '        Next
        '        txtdontImport.Text = mvalue
        '    End If

        'End If
    End Sub
End Class