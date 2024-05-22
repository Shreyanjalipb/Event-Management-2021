Imports System.Net
Public Module EmailModule

    Public Const CONST_SHADOWADMINEMAIL As String = ";kpo@frontware.co.th"


    '--------------------------------------------------------
    'send mail function
    'paramin : _mFrom - sender by
    '          _mBody - email body
    '          _mSMTPServerName - SMTP server
    '          _mSubject - email subject
    '          _mTo  - email recepient
    'return  : 1 - ok
    '          0 - unexpected error
    '         -1 - can't create mail
    '         -2 - can't Send mail
    '         [mError] - error string 
    'tko 23/07/2003 created  
    'kpo 11/08/2006 opimized and verify
    'send email to email to sender's emailaddress ,and bcc to mEmail
    '--------------------------------------------------------
    ''' <summary>
    ''' send email function
    ''' </summary>
    ''' <param name="_mFrom">Sender email address</param>
    ''' <param name="_mSubject">Subject of email</param>
    ''' <param name="_mBody">Body of email</param>
    ''' <param name="_mSMTPServerName">SMTP server name / ip</param>
    ''' <param name="_mTo">Receive email address</param>
    ''' <param name="_mBCC">BCC email address</param>
    ''' <param name="mError">Error return back variable (return)</param>
    ''' <param name="mAuthenticalUser">user to connect to smtpserver</param>
    ''' <param name="mAuthenticalPwd">user's password to connect to smtpserver</param>
    ''' <param name="mAuthenticalPort">port(default = 25,smtp.gmail.com = 587)</param>
    ''' <param name="mUseSSLYN">use SSL ?</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SendEmail(ByVal _mFrom As String, _
                       ByVal _mSubject As String, _
                       ByVal _mBody As String, _
                       ByVal _mSMTPServerName As String, _
                       ByVal _mTo As String, _
                       ByVal _mBCC As String, Optional ByRef mError As String = "", _
                               Optional ByVal mAuthenticalUser As String = "", _
                               Optional ByVal mAuthenticalPwd As String = "", _
                               Optional ByVal mAuthenticalPort As Integer = 25, _
                               Optional ByVal mUseSSLYN As Boolean = False) As Integer

        SentMailAttachMultiFiles(_mFrom, _mSubject, _mBody, False, _mSMTPServerName, _mTo, _mBCC, Nothing, , , mError, _
                                mAuthenticalUser, mAuthenticalPwd, mAuthenticalPort, mUseSSLYN)
    End Function

    ''' <summary>
    ''' send email function
    ''' </summary>
    ''' <param name="_mFrom">Sender email address</param>
    ''' <param name="_mSubject">Subject of email</param>
    ''' <param name="_mBody">Body of email</param>
    ''' <param name="_mSMTPServerName">SMTP server name / ip</param>
    ''' <param name="_mTo">Receive email address</param>
    ''' <param name="_mBCC">BCC email address</param>
    ''' <param name="mError">Error return back variable (return)</param>
    ''' <param name="mAuthenticalUser">user to connect to smtpserver</param>
    ''' <param name="mAuthenticalPwd">user's password to connect to smtpserver</param>
    ''' <param name="mAuthenticalPort">port(default = 25,smtp.gmail.com = 587)</param>
    ''' <param name="mUseSSLYN">use SSL ?</param>
    ''' <param name="file_path">array of filename fullpath</param>    
    ''' <param name="target_path">path of attachment directory </param>    
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SentMailAttachFile(ByVal _mFrom As String, _
                               ByVal _mSubject As String, _
                               ByVal _mBody As String, _
                               ByVal _mSMTPServerName As String, _
                               ByVal _mTo As String, _
                               ByVal _mBCC As String, ByVal file_path() As String, _
                               Optional ByVal target_path As String = "", Optional ByRef mError As String = "", _
                               Optional ByVal mAuthenticalUser As String = "", _
                               Optional ByVal mAuthenticalPwd As String = "", _
                               Optional ByVal mAuthenticalPort As Integer = 25, _
                               Optional ByVal mUseSSLYN As Boolean = False) As Integer

        SentMailAttachMultiFiles(_mFrom, _mSubject, _mBody, False, _mSMTPServerName, _mTo, _mBCC, file_path, target_path, , _
                     mError, mAuthenticalUser, mAuthenticalPwd, mAuthenticalPort, mUseSSLYN)

    End Function

    Private Function URemoveDuplicateEmail(ByVal mEmailList As String) As String
        Dim mArraylist As New ArrayList
        Dim mEmails() As String
        Dim mEmailtemp As String
        Dim mreturnemail As New System.Text.StringBuilder
        mEmails = mEmailList.Split(";")
        For Each mEmailtemp In mEmails
            If Not mArraylist.Contains(mEmailtemp) Then
                mArraylist.Add(mEmailtemp)
            End If
        Next
        For Each mEmailtemp In mArraylist
            If UISGoodEmailFormat(mEmailtemp) Then
                If mreturnemail.Length > 0 Then mreturnemail.Append(";")
                mreturnemail.Append(mEmailtemp)
            End If
        Next
        Return mreturnemail.ToString
    End Function

    'adding the email
    Private Sub UAddingEmailfromlist(ByVal mEmaliobj As Object, ByVal mEmalist As String)
        Dim memails() As String = mEmalist.Split(";")
        Dim memail As String
        For Each memail In memails
            If UISGoodEmailFormat(memail) Then
                mEmaliobj.add(memail)
            End If
        Next
    End Sub

    'Send email with attached many files
    'TKO : Created on 20/02/2003
    Private Function SentMailAttachMultiFiles(ByVal _mFrom As String, _
                               ByVal _mSubject As String, _
                               ByVal _mBody As String, _
                               ByVal _mUseCDOYN As Boolean, _
                               ByVal _mSMTPServerName As String, _
                               ByVal _mTo As String, _
                               ByVal _mBCC As String, ByVal file_path() As String, _
                               Optional ByVal target_path As String = "", Optional ByVal _maxamount As Integer = 50, _
                               Optional ByRef mError As String = "", _
                               Optional ByVal mAuthenticalUser As String = "", _
                               Optional ByVal mAuthenticalPwd As String = "", _
                               Optional ByVal mAuthenticalPort As Integer = 25, _
                               Optional ByVal mUseSSLYN As Boolean = False) As Integer

        Dim s_file As String = ""

        'make group of mail
        Dim _bccmails() As String = Split(_mBCC, ";")
        Dim _bccmail As String
        Dim _bccmailcollect As New ArrayList
        For Each _bccmail In _bccmails
            If UISGoodEmailFormat(_bccmail) Then
                If Not _bccmailcollect.Contains(_bccmail) Then _bccmailcollect.Add(_bccmail)
            End If
        Next
        _bccmails = Nothing

        'KPO 01/07/2002
        'if eHost send with CDO
        Dim berroryn As Boolean
        Dim iamount As Integer = _maxamount
        If iamount = 0 Then iamount = 50

        'use cdonts.
        If _mUseCDOYN Then
            Dim objMail As Object = Nothing
            Try
                objMail = CreateObject("CDONTS.NewMail")
            Catch ex As Exception
                mError = String.Concat("[EmailModule] Can't create CDONTS Object : ", ex.Message)
                berroryn = True
                WriteLog("mError1:" & mError)
            End Try
            If Not berroryn Then
                objMail.From = _mFrom
                objMail.Subject = _mSubject
                'kpo 02/07/2002
                objMail.BodyFormat = 0 'CdoBodyFormatHTML
                'kpo 04/07/2002 need for HTML format
                objMail.MailFormat = 0 'CdoMailFormatHTML 
                objMail.Body = _mBody
                If Not file_path Is Nothing Then
                    Try
                        For Each s_file In file_path
                            objMail.AttachFile(Trim(s_file), Dir(target_path, FileAttribute.Normal))
                        Next
                    Catch ex As Exception
                        mError = String.Concat("[EmailModule] Can't Attach the file (", s_file, ") : ", ex.Message)
                        berroryn = True
                        WriteLog("mError:2" & mError)
                    End Try
                End If
            End If

            If Not berroryn Then
                Dim icount As Integer
                Dim mst_mail As New System.Text.StringBuilder
                For Each _bccmail In _bccmailcollect
                    icount += 1
                    mst_mail.Append(_bccmail)
                    mst_mail.Append(";")

                    If icount = iamount Then
                        Try
                            objMail.bcc = mst_mail.ToString
                            objMail.Send()
                        Catch ex As Exception
                            mError = String.Concat("[EmailModule] Can't Send the email : ", ex.Message)
                            berroryn = True
                            WriteLog("mError3:" & mError)
                            Exit For
                        End Try
                        'reset loop
                        mst_mail = New System.Text.StringBuilder
                        icount = 0
                    End If
                Next
                'last loop
                If icount > 0 Then
                    Try
                        objMail.bcc = mst_mail.ToString
                        objMail.Send()
                    Catch ex As Exception
                        mError = String.Concat("[EmailModule] Can't Send the email : ", ex.Message)
                        berroryn = True
                        WriteLog("mError4:" & mError)
                    End Try
                End If
                mst_mail = Nothing
            End If
            objMail = Nothing

        Else 'use .net mail

            Dim mclient As New Mail.SmtpClient
            Dim mmessage As Mail.MailMessage = Nothing
            Dim breturn As Integer
            Dim m_to As String
            Try
                mmessage = New Mail.MailMessage
                mmessage.From = New Net.Mail.MailAddress(URemoveDuplicateEmail(_mFrom))
                m_to = URemoveDuplicateEmail(_mTo)

                If m_to <> "" Then UAddingEmailfromlist(mmessage.To, m_to)

                mmessage.Subject = _mSubject
                mmessage.IsBodyHtml = True
                'tko : 20/12/2004 <-- encoding emailbody
                mmessage.BodyEncoding = System.Text.Encoding.UTF8
                mmessage.Body = _mBody

                Dim mfiles() As String
                If Not file_path Is Nothing Then
                    mfiles = file_path
                Else
                    If IO.Directory.Exists(target_path) Then mfiles = IO.Directory.GetFiles(target_path)
                End If
                If Not mfiles Is Nothing Then
                    For Each s_file In mfiles
                        mmessage.Attachments.Add(New Mail.Attachment(s_file))
                    Next
                End If

            Catch ex As Exception
                breturn = -1
                mError = String.Concat("[EmailModule] Can't create the email : ", ex.Message)
                WriteLog("mError5:" & mError)
            End Try

            mclient.Host = _mSMTPServerName
            mclient.Port = mAuthenticalPort
            mclient.EnableSsl = mUseSSLYN
            If (mAuthenticalUser = "") Or (mAuthenticalPwd = "") Then
            Else
                mclient.Credentials = New Net.NetworkCredential(mAuthenticalUser, mAuthenticalPwd)
            End If

            If breturn = 0 Then

                'send bcc addreess.
                Dim bsendyn As Boolean
                Dim icount As Integer
                Dim mst_mail As New System.Text.StringBuilder
                '_bccmailcollect = not duplicate and good email
                For Each _bccmail In _bccmailcollect
                    If mst_mail.Length > 0 Then mst_mail.Append(";")
                    icount += 1
                    mst_mail.Append(_bccmail)
                    If icount = iamount Then
                        Try
                            mmessage.Bcc.Clear()
                            UAddingEmailfromlist(mmessage.Bcc, mst_mail.ToString)

                            mclient.Send(mmessage)
                            bsendyn = True
                        Catch ex As Exception
                            mError = String.Concat("[EmailModule] Can't Send the email : ", ex.Message)
                            berroryn = True
                            WriteLog("mError6:" & mError)
                            Exit For
                        End Try
                        'reset loop
                        mst_mail = New System.Text.StringBuilder
                        icount = 0
                    End If
                Next
                'last loop
                If icount > 0 Then
                    Try
                        mmessage.Bcc.Clear()
                        UAddingEmailfromlist(mmessage.Bcc, mst_mail.ToString)

                        mclient.Send(mmessage)
                        bsendyn = True
                    Catch ex As Exception
                        mError = String.Concat("[EmailModule] Can't Send the email : ", ex.Message)
                        berroryn = True
                        WriteLog("mError7:" & mError)
                    End Try
                End If
                mst_mail = Nothing

                'send without BCC
                If Not bsendyn Then
                    Try
                        mclient.Send(mmessage)
                        bsendyn = True
                    Catch ex As Exception
                        mError = String.Concat("[EmailModule] Can't Send the email : ", ex.Message)
                        berroryn = True
                        WriteLog("mError8:" & mError)
                    End Try
                End If
            End If

            mmessage = Nothing
            mclient = Nothing

        End If
    End Function

    'remove duplicate
    'validate that email is validate
    Private Function UISGoodEmailFormat(ByVal mCheckMail As String) As Boolean
        Dim mmail As Mail.MailAddress
        Try
            mmail = New Mail.MailAddress(mCheckMail)
            UISGoodEmailFormat = True
        Catch ex As Exception
            UISGoodEmailFormat = False
        End Try
        mmail = Nothing
    End Function

End Module