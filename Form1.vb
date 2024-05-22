'Imports System.Text
'Imports System.IO
'Imports System.Collections
'Imports System.Security.Permissions
'Imports Microsoft.Win32
'Imports System.Xml
Public Class Form1

End Class

'Public Class ImportMain

'    Dim togle_F As Boolean
'    Dim togle_T As Boolean
'    Dim Tend As Integer = 6
'    Dim Tendtpye As Integer = 4
'    Dim Count_T As Integer = 0
'    Dim Pstatus As Byte
'    Dim Pnum As Integer
'    Dim Pcounter As Integer = 0

'    Private Sub BtBrowse_Source_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtBrowse_Source.Click
'        FolderBrowserDialog2.ShowDialog()
'        LabelAgent.Hide()
'        If FolderBrowserDialog2.SelectedPath <> "" Then Txtsource.Text = FolderBrowserDialog2.SelectedPath



'    End Sub

'    Private Sub Btdate_from_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btdate_from.Click

'        Calendar1.Show()

'    End Sub

'    Private Sub ImportMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

'        Me.Size = New Size(477, 380)
'        Calendar1.Hide()
'        Calendar2.Hide()
'        ProgressBar1.Hide()
'        Label9.Text = ""
'        LabelAgent.Hide()
'        LabelType.Hide()
'        Button2.Hide()


'    End Sub

'    Private Sub BtBrowse_des_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtBrowse_des.Click

'        FolderBrowserDialog2.ShowDialog()
'        If FolderBrowserDialog2.SelectedPath <> "" Then Txtdes.Text = FolderBrowserDialog2.SelectedPath

'    End Sub


'    Private Sub Btdate_to_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btdate_to.Click

'        Calendar2.Show()

'    End Sub

'    Private Sub Calendar2_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Calendar2.DateChanged
'        'MsgBox("ff:" & Calendar2.SelectionRange.Start)
'        'Txttime2.Text = Format(Calendar2.SelectionRange.Start, "dd/MM/yyyy")
'    End Sub

'    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
'        Dim i, ai As Integer
'        Pstatus = 1

'        Dim filenames() As String
'        Dim filename As String
'        Dim Tstate As String
'        Dim StrTime As Date
'        Dim ENTime As Date
'        Dim pathfile As String
'        Dim fileDes As String
'        Dim filenameCoppy As String

'        If Txtsource.Text = "" Then

'            MsgBox("All search criteria must be filled-in before starting search.")
'            Txtsource.Focus()
'        ElseIf TxtTime1.Text = "" Then

'            MsgBox("All search criteria must be filled-in before starting search.")
'            TxtTime1.Focus()
'        ElseIf Txttime2.Text = "" Then

'            MsgBox("All search criteria must be filled-in before starting search.")
'            Txttime2.Focus()
'        ElseIf Txtnumagent.Text = "" Then

'            MsgBox("All search criteria must be filled-in before starting search.")
'            Txtnumagent.Focus()

'        ElseIf Txttype.Text = "" Then
'            MsgBox("All search criteria must be filled-in before starting search.")
'            Txttype.Focus()
'        ElseIf Txttype.Text.Length < 4 Then
'            MsgBox("Please fill Type courrier 4 alphanumeric character.")
'            Txttype.Focus()
'        ElseIf Txtdes.Text = "" Then

'            MsgBox("All search criteria must be filled-in before starting search.")
'            Txtdes.Focus()

'        Else
'            StrTime = FormatDateTime(Calendar1.SelectionRange.Start, DateFormat.ShortDate)
'            ENTime = FormatDateTime(Calendar2.SelectionRange.Start, DateFormat.ShortDate)

'            filenames = Directory.GetFiles(Txtsource.Text, "*.xml")



'            Dim fileTime As Date

'            ProgressBar1.Show()
'            Button2.Show()

'            Pstatus = 0
'            For ai = 0 To filenames.Length - 1
'                If Pstatus = 1 Then Exit For
'                filenameCoppy = (Path.GetFileName(filenames(ai)))
'                filename = filenames(ai)
'                fileTime = FormatDateTime(FileDateTime(filename), DateFormat.ShortDate)


'                fileDes = Txtdes.Text + "\" + filenameCoppy
'                pathfile = filename

'                If StrTime <= fileTime And fileTime <= ENTime Then
'                    CheckFileCoppy(pathfile, fileDes, filenameCoppy)
'                End If

'                ProgressBar1.Value = (ai + 1) / filenames.Length * 100
'                Tstate = ProgressBar1.Value
'                Label9.Text = Tstate + "%"
'                Application.DoEvents()
'                i += 1
'            Next ai

'            If i = filenames.Length Then

'                MsgBox("complete!")
'                ProgressBar1.Value = 0
'                ProgressBar1.Hide()
'                Button2.Hide()
'                Label9.Text = ""
'                Pcounter = 0
'            End If
'        End If




'    End Sub
'    ''' <summary>
'    ''' check+find by number agen 
'    ''' </summary>
'    ''' <param name="fileSource"> path file source</param>
'    ''' <param name="fileDes">path file Destination</param>
'    ''' <param name="filenameCoppy"> file name</param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Function CheckFileCoppy(ByVal fileSource As String, ByVal fileDes As String, ByVal filenameCoppy As String) As String

'        Dim state As Boolean
'        Dim DataX As New DataSet
'        Dim Stfile As String
'        Dim chkfile As String = filenameCoppy

'        Dim Astring() As String
'        Dim Tstring() As String
'        Dim numA As Integer
'        Dim numT As Integer
'        Dim alltype As String
'        Try


'            If Txtnumagent.TextLength = 6 Then
'                If Mid(Txtnumagent.Text, 6) <> ";" Then
'                    Txtnumagent.Text = Txtnumagent.Text + ";" '+ Txtnumagent.Text
'                End If

'            ElseIf Txtnumagent.TextLength = 5 Then
'                Txtnumagent.Text = Txtnumagent.Text + ";"

'            End If
'            If Txttype.TextLength = 4 Then
'                Txttype.Text = Txttype.Text + ";" '+ Txttype.Text

'            End If


'            Astring = Txtnumagent.Text.Split(";")
'            Tstring = Txttype.Text.Split(";")
'            numA = Astring.Length()
'            numT = Tstring.Length()

'            'If InStr(Txtnumagent.Text, ";") = 6 Then
'            '    numA = Int(Txtnumagent.TextLength / 6)

'            'ElseIf InStr(Txtnumagent.Text, ";") = 7 Then
'            '    numA = Int(Txtnumagent.TextLength / 7)

'            'End If

'            ' numT = Int(Txttype.TextLength / 5)

'            Dim a As Integer = 0
'            Dim t As Integer = 0
'            Stfile = IO.File.ReadAllText(fileSource)
'            Dim Strfind As String

'            '  MsgBox("Txtnumagent.TextLength :" & Txtnumagent.TextLength & " numA:" & numA)

'            'Dim count_Agent As Integer = 0
'            'Dim count_Type As Integer = 0
'            'If Txtnumagent.TextLength <= 6 Then
'            '    count_Agent = 1
'            'ElseIf Txtnumagent.TextLength = 7 Then
'            '    count_Agent = 1

'            'End If

'            Dim indexStr As Integer = InStr(Stfile, "<ATTRIBUTE Name=""RESERVE05"">")
'            Dim StartString As Integer = indexStr + 26
'            Dim StrKEY_find As String = Stfile.Substring(StartString, 7)
'            For a = 0 To numA - 1    ' find key of agent

'                If Pstatus = 1 Then Exit For
'                ' MsgBox("StrKEY_find :" & StrKEY_find & " Astring(" & a & "):" & Astring(a) & " numA :" & numA)
'                If Astring(a) <> "" Then 'check value
'                    If StrKEY_find.Contains(Astring(a)) Then
'                        For t = 0 To numT - 1
'                            If Pstatus = 1 Then Exit For
'                            If Tstring(t) <> "" Then ' check value
'                                If Tstring(t) = "all%" Then 'check type
'                                    alltype = ""
'                                Else
'                                    alltype = Tstring(t)
'                                End If

'                                '  MsgBox("Tstring(" & t & ")" & Tstring(t))
'                                findByType(Stfile, alltype, fileSource, fileDes, filenameCoppy)
'                            End If
'                        Next t

'                    End If

'                    Strfind = ""
'                End If
'            Next a
'        Catch ex As Exception

'        End Try
'        Return state
'    End Function
'    ''' <summary>
'    ''' find by type + coppy file 
'    ''' </summary>
'    ''' <param name="Strfile"> path file</param>
'    ''' <param name="Type"> key type</param>
'    ''' <param name="fileSource">path file source</param>
'    ''' <param name="Desfile">path file Destination</param>
'    ''' <param name="filenameCoppy"> file name</param>
'    ''' <remarks></remarks>
'    Sub findByType(ByVal Strfile As String, ByVal Type As String, ByVal fileSource As String, ByVal Desfile As String, ByVal filenameCoppy As String)
'        Try

'            Dim Nfile As String
'            Dim Strfind_type As String
'            Strfind_type = String.Concat("<ATTRIBUTE Name=""TYPEC"">", Type)

'            If Strfile.Contains(Strfind_type) Then ' find key of Type

'                Nfile = filenameCoppy.Replace(".xml", ".pdf")
'                System.IO.File.Copy(fileSource, Desfile, True)
'                SetAttr(Desfile, FileAttribute.Normal)
'                findPDF(filenameCoppy, Desfile)

'            End If

'        Catch ex As Exception

'        End Try
'    End Sub
'    ''' <summary>
'    ''' find file .pdf+coppy file
'    ''' </summary>
'    ''' <param name="filexml">file name xml</param>
'    ''' <param name="Desfile">path file Destination</param>
'    ''' <remarks></remarks>
'    Sub findPDF(ByVal filexml As String, ByVal Desfile As String)

'        Dim filenamePDF() As String
'        Dim i As Integer
'        filenamePDF = Directory.GetFiles(Txtsource.Text, "*.pdf")

'        For i = 0 To filenamePDF.Length - 1
'            If Pstatus = 1 Then Exit For

'            If filexml.Replace(".xml", ".pdf") = Path.GetFileName(filenamePDF(i)) Then
'                System.IO.File.Copy(filenamePDF(i), Desfile.Replace(".xml", ".pdf"), True)
'                SetAttr(Desfile.Replace(".xml", ".pdf"), FileAttribute.Normal)
'            End If

'        Next

'    End Sub


'    Private Sub txtInput_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Txtnumagent.KeyPress

'        'Select Case e.KeyChar
'        '    Case "0" To "9"
'        '        e.Handled = False
'        '        chkMark()
'        '    Case Chr(13), Chr(46)
'        '        e.Handled = False
'        '    Case Chr(8)
'        '        e.Handled = False
'        '        DeleteMark()
'        '    Case Else
'        '        e.Handled = True
'        'End Select
'        chkMark()


'    End Sub

'    Sub DeleteMark()

'        Dim num As Integer
'        Dim nvalue As Integer
'        num = Txtnumagent.TextLength
'        nvalue = num Mod 8
'        If Txtnumagent.TextLength = 0 Then
'            nvalue = 0
'        End If
'        If Txtnumagent.TextLength > 0 Then
'            If nvalue = 0 Then

'                If Tend > 0 Then
'                    Tend = Tend - 7
'                End If
'            End If
'        End If

'    End Sub
'    Sub DeleteMark_T()
'        Dim num As Integer
'        Dim nvalue As Integer
'        num = Txttype.TextLength
'        nvalue = num Mod 6
'        If Txttype.TextLength = 0 Then
'            nvalue = 0
'        End If
'        If Txttype.TextLength > 0 Then
'            If nvalue = 0 Then

'                If Tendtpye > 0 Then
'                    Tendtpye = Tendtpye - 5
'                End If
'            End If
'        End If

'    End Sub
'    Sub chkMark()

'        'If Txtnumagent.TextLength = Tend Then
'        '    Txtnumagent.Text = Txtnumagent.Text.Insert(Tend, ";")
'        '    Txtnumagent.SelectionStart = Txtnumagent.Text.Length
'        '    Tend = Tend + 7
'        '    Count_T = Count_T + 1


'        'End If

'    End Sub
'    Private Sub inputTxtT(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Txttype.KeyPress

'        Select Case e.KeyChar

'            Case Chr(8)
'                e.Handled = False
'                DeleteMark_T()
'            Case ";"

'                e.Handled = True
'            Case Else
'                e.Handled = False

'                If Txttype.TextLength = 6 And Tendtpye = 4 Then
'                    Tendtpye = Tendtpye + 5
'                End If
'                If Txttype.TextLength = Tendtpye Then
'                    Txttype.Text = Txttype.Text.Insert(Tendtpye, ";")
'                    Txttype.SelectionStart = Txttype.Text.Length
'                    Tendtpye = Tendtpye + 5

'                End If
'        End Select



'    End Sub

'    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
'        If MsgBox("do you want stop? Y/N", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
'            Pstatus = 1
'            ProgressBar1.Hide()
'            ProgressBar1.Value = 0
'            ProgressBar1.Hide()
'            Label9.Text = ""
'            Pcounter = 0


'        End If


'    End Sub

'    Private Sub Calendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Calendar1.DateChanged
'        TxtTime1.Text = Format(Calendar1.SelectionRange.Start, "dd/MM/yyyy")
'    End Sub


'    Private Sub showHA(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Txtnumagent.MouseMove
'        LabelAgent.Show()
'    End Sub


'    Private Sub L(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txtnumagent.MouseLeave
'        LabelAgent.Hide()
'    End Sub


'    Private Sub Thid(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txttype.MouseLeave
'        LabelType.Hide()
'    End Sub

'    Private Sub Tshow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Txttype.MouseMove
'        LabelType.Show()
'    End Sub


'    Private Sub closeT1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Calendar1.DateSelected
'        Calendar1.Hide()
'        Txttime2.Focus()
'    End Sub

'    Private Sub closeT2(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Calendar2.DateSelected
'        Txttime2.Text = Format(Calendar2.SelectionRange.Start, "dd/MM/yyyy")
'        Calendar2.Hide()
'        Txtnumagent.Focus()
'    End Sub
'End Class