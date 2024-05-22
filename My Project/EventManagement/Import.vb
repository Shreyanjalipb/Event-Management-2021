Imports System
Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Collections.Generic
Imports System.Data.OleDb

Public Class Import
    Private sumNewRec As Integer
    Private sumDuplicateRec As Integer
    Private sumRecnotSave As Integer
    Private DuplicateRec As String = ""
    Private RecnotSave As String = ""
    Private NewRec As String = ""
    Private tempOrderNo As String = ""
    Private tempOrderDate As String = ""
    Private ColumnS As String
    Private Sub Btnbrown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnbrown.Click
        Dim _path As String
        Try
            clearData()
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                _path = OpenFileDialog1.FileName
                If _path <> "" Then
                    txtpath.Text = _path
                Else
                    txtpath.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox("Cannot Select Path for Excel File: " & ex.Message)
        End Try
    End Sub

    Private Sub Btnok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnok.Click
        If txtpath.Text = "" Then
            MsgBox("File Name ห้ามว่าง.")
            txtpath.Focus()
            Exit Sub

        Else
            ImportFile(txtpath.Text)
        End If

        '     If InStr(txtpath.Text, ".csv", ) Then

        'If InsertContact() Then

        'End If



        '  ImportFile(txtpath.Text)
        'InsertContact(1,"t1", "t2", _
        '               "t3", "t4", "t5", _
        '               "_Contact_Country", "_Home_Address1 ", "_Home_Address2 ", _
        '               "_Home_Suburb","_Home_State", "_Home_Postcode ", "_Home_Country", _
        '               " _Contact_Number", "_Position ","_Given_Name ","_Surname", _
        '               "_Phone_Bus","_Phone_Direct","_Phone_Pri","_Phone_Mobile", _
        '               "_Fax", "_EMail ","_Other","_Department","_Debtor_Number","_Charge_Code", _
        '               "_Web_Page","_Mailing_Address","_Contact_Colour","_Sequence ","_Created As String, _
        '               "_Special_Notes As String, ByVal _Title As String, ByVal _Inactive_Contact_Date As Date, ByVal _LastEdited As Date, _
        '                       ByVal _Prefix As String, ByVal _Created_by As String, ByVal _Last_Edited_by As String, ByVal _PrivateContact As String, _
        '                       ByVal _Mail_Preference_Date As Date, ByVal _Mail_Preference_Subscribed As String, ByVal _Contact_Address_3 As String, ByVal _Home_Address_3 As String, ByVal _Email2 As String) As Boolean

        '  Else

        '  End If

    End Sub

    Private Sub Import_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx"
        'this.openFileDialog1.FileName = "*.txt";

        ProgressBar1.Hide()
        lblstate.Hide()
        'clearData()

    End Sub
    Public Sub clearData()
        txtpath.Text = ""
        lblduplicate.Text = "0"
        lblNotImportRec.Text = "0"
        lblstate.Text = "0"
        lblSumnew.Text = "0"
        ProgressBar1.Value = 0
    End Sub

    Private Sub ImportFile(ByVal pathFile As String)
        Try
            DataNotImport = ""
            ProgressBar1.Show()
            lblstate.Show()

            Dim Tstate As String




            Dim objConn As New OleDbConnection
            Dim dtAdapter As OleDbDataAdapter
            Dim dtTbName As New DataTable
            Dim dtdata As New DataTable
            Dim Tbname As String


            Dim strConnString As String
            strConnString = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source=" & pathFile & ";Extended Properties=Excel 8.0;"
            objConn = New OleDbConnection(strConnString)
            objConn.Open()

            '------------------get table name--------------------------'
            dtTbName = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            If dtTbName.Rows.Count > 0 Then
                Tbname = dtTbName.Rows(0).Item("TABLE_NAME")
            End If


            Dim strSQL As String
            strSQL = "SELECT * FROM [" + Tbname + "]"
            dtAdapter = New OleDbDataAdapter(strSQL, objConn)
            dtAdapter.Fill(dtdata)
            ColumnS = "|"
            Dim i As Integer
            For i = 0 To dtdata.Columns.Count - 1
                ColumnS = ColumnS + dtdata.Columns(i).ColumnName + "|"
            Next i

            Dim statusContact As Integer
            Dim index As Integer
            For index = 0 To dtdata.Rows.Count - 1
                If dtdata.Rows(index).Item("Company").ToString() <> "" Then

                    'check company name
                    If MyCompanyClass.CheckCompany(dtdata.Rows(index).Item("Company").ToString()) Then
                        statusContact = InsertContact(dtdata.Rows(index))
                        If statusContact = 0 Then
                            sumNewRec = sumNewRec + 1
                            NewRec = NewRec + "," + dtdata.Rows(index).Item("Contact").ToString()
                        ElseIf statusContact = 1 Then
                            sumRecnotSave = sumRecnotSave + 1
                            RecnotSave = RecnotSave + "," + dtdata.Rows(index).Item("Contact").ToString()
                        ElseIf statusContact = 2 Then
                            sumDuplicateRec = sumDuplicateRec + 1
                            DuplicateRec = DuplicateRec + "," + dtdata.Rows(index).Item("Contact").ToString()
                        End If
                    Else
                        If InsertCompany(dtdata.Rows(index)) Then
                            statusContact = InsertContact(dtdata.Rows(index))
                            If statusContact = 0 Then
                                sumNewRec = sumNewRec + 1
                                NewRec = NewRec + "," + dtdata.Rows(index).Item("Contact").ToString()
                            ElseIf statusContact = 1 Then
                                sumRecnotSave = sumRecnotSave + 1
                                RecnotSave = RecnotSave + "," + dtdata.Rows(index).Item("Contact").ToString()
                            ElseIf statusContact = 2 Then
                                sumDuplicateRec = sumDuplicateRec + 1
                                DuplicateRec = DuplicateRec + "," + dtdata.Rows(index).Item("Contact").ToString()
                            End If
                        End If
                    End If


                Else
                    sumRecnotSave = sumRecnotSave + 1
                    RecnotSave = RecnotSave + "," + dtdata.Rows(index).Item("Contact").ToString()

                End If


                ProgressBar1.Value = (index + 1) / (dtdata.Rows.Count - 1) * 100
                Tstate = ProgressBar1.Value
                lblstate.Text = Tstate + "%"
                Application.DoEvents()

            Next index


            ProgressBar1.Value = 100
            Tstate = ProgressBar1.Value
            lblstate.Text = Tstate + "%"

            DataNotImport = RecnotSave
            lblNotImportRec.Text = sumRecnotSave
            lblduplicate.Text = sumDuplicateRec
            lblSumnew.Text = sumNewRec
            tempOrderNo = ""
            tempOrderDate = ""
            sumRecnotSave = 0
            sumDuplicateRec = 0
            sumNewRec = 0

            dtAdapter = Nothing
            dtdata = Nothing
            dtTbName = Nothing
            objConn.Close()
            objConn = Nothing

            MsgBox("Import data successfully.")

        Catch ex As Exception
            MsgBox("Import File Error : " + ex.ToString())
        End Try

    End Sub


    'ByVal _Company_RecNo As Long, ByVal _Contact_Address_1 As String, ByVal _Contact_Post_Code As String, _
    '                            ByVal _Contact_State As String, ByVal _Contact_Address_2 As String, ByVal _Contact_City As String, _
    '                            ByVal _Contact_Country As String, ByVal _Home_Address1 As String, ByVal _Home_Address2 As String, _
    '                            ByVal _Home_Suburb As String, ByVal _Home_State As String, ByVal _Home_Postcode As String, ByVal _Home_Country As String, _
    '                            ByVal _Contact_Number As String, ByVal _Position As String, ByVal _Given_Name As String, ByVal _Surname As String, _
    '                            ByVal _Phone_Bus As String, ByVal _Phone_Direct As String, ByVal _Phone_Pri As String, ByVal _Phone_Mobile As String, _
    '                            ByVal _Fax As String, ByVal _EMail As String, ByVal _Other As String, ByVal _Department As String, ByVal _Debtor_Number As String, ByVal _Charge_Code As String, _
    '                            ByVal _Web_Page As String, ByVal _Mailing_Address As Long, ByVal _Contact_Colour As Long, ByVal _Sequence As Long, ByVal _Created As String, _
    '                            ByVal _Special_Notes As String, ByVal _Title As String, ByVal _Inactive_Contact_Date As Date, ByVal _LastEdited As Date, _
    '                            ByVal _Prefix As String, ByVal _Created_by As String, ByVal _Last_Edited_by As String, ByVal _PrivateContact As String, _
    '                            ByVal _Mail_Preference_Date As Date, ByVal _Mail_Preference_Subscribed As String, ByVal _Contact_Address_3 As String, ByVal _Home_Address_3 As String, ByVal _Email2 As String


    Private Function InsertContact(ByVal dt As DataRow) As Integer

        Dim mreturn As Integer
        Try

            MyContactClass.Contact_RecNo = Convert.ToInt32(MyContactClass.GetTopData("Contact_RecNo")) + 1

            If ChkColumnName("|Company|") Then
                If Not IsDBNull(dt.Item("Company")) Then
                    MyContactClass.Company_RecNo = MyCompanyClass.GetCompanyRecNo(dt.Item("Company"))
                Else
                    MyCompanyClass.Company_RecNo = ""
                End If
            End If
            If ChkColumnName("|Address 1|") Then
                If Not IsDBNull(dt.Item("Address 1")) Then
                    MyContactClass.Contact_Address_1 = dt.Item("Address 1")
                    MyContactClass.Home_Address1 = dt.Item("Address 1")
                Else
                    MyContactClass.Contact_Address_1 = ""
                    MyContactClass.Home_Address1 = ""
                End If
            End If

            If ChkColumnName("|Address 2|") Then
                If Not IsDBNull(dt.Item("Address 2")) Then
                    MyContactClass.Contact_Address_2 = dt.Item("Address 2")
                    MyContactClass.Home_Address2 = dt.Item("Address 2")
                Else
                    MyContactClass.Contact_Address_2 = ""
                    MyContactClass.Home_Address2 = ""
                End If
            End If


            If ChkColumnName("|Address 3|") Then
                If Not IsDBNull(dt.Item("Address 3")) Then
                    MyContactClass.Contact_Address_3 = dt.Item("Address 3")
                    MyContactClass.Home_Address_3 = dt.Item("Address 3")
                Else
                    MyContactClass.Contact_Address_3 = ""
                    MyContactClass.Home_Address_3 = ""
                End If
            End If

            If ChkColumnName("|Zip|") Then
                If Not IsDBNull(dt.Item("Zip")) Then
                    MyContactClass.Contact_Post_Code = dt.Item("Zip")
                    MyContactClass.Home_Postcode = dt.Item("Zip")
                Else
                    MyContactClass.Contact_Post_Code = ""
                    MyContactClass.Home_Postcode = ""
                End If
            End If

            If ChkColumnName("|State|") Then
                If Not IsDBNull(dt.Item("State")) Then
                    MyContactClass.Contact_State = dt.Item("State")
                    MyContactClass.Home_State = dt.Item("State")
                Else
                    MyContactClass.Contact_State = ""
                    MyContactClass.Home_State = ""
                End If
            End If
            If ChkColumnName("|City|") Then
                If Not IsDBNull(dt.Item("City")) Then
                    MyContactClass.Contact_City = dt.Item("City")
                Else
                    MyContactClass.Contact_City = ""
                End If
            End If

            If ChkColumnName("|Country|") Then
                If Not IsDBNull(dt.Item("Country")) Then
                    MyContactClass.Contact_Country = dt.Item("Country")
                    MyContactClass.Home_Country = dt.Item("Country")
                Else
                    MyContactClass.Contact_Country = ""
                    MyContactClass.Home_Country = ""
                End If
            End If


            ' MyContactClass.Home_Suburb = "_Home_Suburb"
            If ChkColumnName("|Position|") Then
                If Not IsDBNull(dt.Item("Position")) Then
                    MyContactClass.Position = dt.Item("Position")
                Else
                    MyContactClass.Position = ""
                End If
            End If


            MyContactClass.Contact_Number = MyContactClass.Contact_RecNo.ToString()

            '      MyContactClass.Home_Suburb = "_Home_Suburb"
            If ChkColumnName("|First Name|") Then
                If Not IsDBNull(dt.Item("First Name")) Then
                    MyContactClass.Given_Name = dt.Item("First Name")
                Else
                    MyContactClass.Given_Name = ""
                End If
            End If
            If ChkColumnName("|Last Name|") Then
                If Not IsDBNull(dt.Item("Last Name")) Then
                    MyContactClass.Surname = dt.Item("Last Name")
                Else
                    MyContactClass.Surname = ""
                End If
            End If
            If ChkColumnName("|Phone|") Then
                If Not IsDBNull(dt.Item("Phone")) Then
                    MyContactClass.Phone_Bus = dt.Item("Phone")
                Else
                    MyContactClass.Phone_Bus = ""
                End If
            End If
            If ChkColumnName("|Mobile|") Then
                If Not IsDBNull(dt.Item("Mobile")) Then
                    MyContactClass.Phone_Mobile = dt.Item("Mobile")
                Else
                    MyContactClass.Phone_Mobile = ""
                End If
            End If


            If ChkColumnName("|E-mail|") Then
                If Not IsDBNull(dt.Item("E-mail")) Then
                    MyContactClass.EMail = dt.Item("E-mail")
                Else
                    MyContactClass.EMail = ""
                End If
            End If
            If ChkColumnName("|Web|") Then
                If Not IsDBNull(dt.Item("Web")) Then
                    MyContactClass.Web_Page = dt.Item("Web")
                Else
                    MyContactClass.Web_Page = ""
                End If
            End If
            If ChkColumnName("|Fax|") Then
                If Not IsDBNull(dt.Item("Fax")) Then
                    MyContactClass.Fax = dt.Item("Fax")
                Else
                    MyContactClass.Fax = ""
                End If
            End If
            If ChkColumnName("|Title|") Then
                If Not IsDBNull(dt.Item("Title")) Then
                    MyContactClass.Title = dt.Item("Title")
                Else
                    MyContactClass.Title = ""
                End If
            End If
            MyContactClass.Special_Notes = ""
            MyContactClass.Phone_Direct = ""
            MyContactClass.Phone_Pri = ""

            MyContactClass.Other = ""
            MyContactClass.Department = ""
            MyContactClass.Debtor_Number = ""
            MyContactClass.Charge_Code = ""

            MyContactClass.Mailing_Address = 0
            MyContactClass.Contact_Colour = 0
            MyContactClass.Sequence = 0
            MyContactClass.Created = Date.Now

            MyContactClass.Inactive_Contact_Date = Date.Now
            MyContactClass.LastEdited = Date.Now
            MyContactClass.Prefix = ""
            MyContactClass.Created_by = "bkk Office"
            MyContactClass.Last_Edited_by = "bkk Office"
            MyContactClass.PrivateContact = False
            MyContactClass.Mail_Preference_Date = Date.Now
            MyContactClass.Mail_Preference_Subscribed = True

            MyContactClass.Email2 = ""



            If MyContactClass.ChkContact(MyContactClass.Given_Name, MyContactClass.Surname) Then
                If MyContactClass.Save("") Then
                    mreturn = 0
                Else
                    mreturn = 1
                End If

            Else
                mreturn = 2

            End If

        Catch ex As Exception
            MessageBox.Show("InsertContact Error : " + ex.ToString())
        End Try


        Return mreturn
    End Function
    Private Function InsertCompany(ByVal dt As DataRow) As Boolean
        Dim mreturn As Boolean

        MyCompanyClass.Company_RecNo = Convert.ToInt32(MyCompanyClass.GetTopData("Company_RecNo")) + 1
        MyCompanyClass.Member_Number = MyCompanyClass.Company_RecNo

        If ChkColumnName("|Company|") Then
            If Not IsDBNull(dt.Item("Company")) Then
                MyCompanyClass.Company_Name = dt.Item("Company")
            Else
                MyCompanyClass.Company_Name = ""
            End If
        End If

        If ChkColumnName("|Address 1|") Then
            If Not IsDBNull(dt.Item("Address 1")) Then
                MyCompanyClass.Address_1 = dt.Item("Address 1")
            Else
                MyCompanyClass.Address_1 = ""
            End If
        End If
        If ChkColumnName("|Address 2|") Then
            If Not IsDBNull(dt.Item("Address 2")) Then
                MyCompanyClass.Address_2 = dt.Item("Address 2")
            Else
                MyCompanyClass.Address_2 = ""
            End If

        End If
        If ChkColumnName("|Address 3|") Then
            If Not IsDBNull(dt.Item("Address 3")) Then
                MyCompanyClass.Address_3 = dt.Item("Address 3")
            Else
                MyCompanyClass.Address_3 = ""
            End If
        End If

        If ChkColumnName("|City|") Then
            If Not IsDBNull(dt.Item("City")) Then
                MyCompanyClass.City = dt.Item("City")
            Else
                MyCompanyClass.City = ""
            End If
        End If

        If ChkColumnName("|State|") Then
            If Not IsDBNull(dt.Item("State")) Then
                MyCompanyClass.State = dt.Item("State")
            Else
                MyCompanyClass.State = ""
            End If

        End If
        If ChkColumnName("|Zip|") Then
            If Not IsDBNull(dt.Item("Zip")) Then
                MyCompanyClass.Post_Code = dt.Item("Zip")
            Else
                MyCompanyClass.Post_Code = ""
            End If
        End If
        If ChkColumnName("|Country|") Then
            If Not IsDBNull(dt.Item("Country")) Then
                MyCompanyClass.Country = dt.Item("Country")
            Else
                MyCompanyClass.Country = ""
            End If
        End If
        If ChkColumnName("|Other|") Then
            If Not IsDBNull(dt.Item("Other")) Then
                MyCompanyClass.Other = dt.Item("Other")
            Else
                MyCompanyClass.Other = ""
            End If
        End If
        If ChkColumnName("|Phone|") Then
            If Not IsDBNull(dt.Item("Phone")) Then
                MyCompanyClass.Phone = dt.Item("Phone")
            Else
                MyCompanyClass.Phone = ""
            End If
        End If
        If ChkColumnName("|Fax|") Then
            If Not IsDBNull(dt.Item("Fax")) Then
                MyCompanyClass.Fax = dt.Item("Fax")
            Else
                MyCompanyClass.Fax = ""
            End If
        End If
        If ChkColumnName("|Web|") Then
            If Not IsDBNull(dt.Item("Web")) Then
                MyCompanyClass.Web_Page = dt.Item("Web")
            Else
                MyCompanyClass.Web_Page = ""
            End If
        End If
        MyCompanyClass.Company_Colour = 0
        MyCompanyClass.Created = Date.Now
        MyCompanyClass.Inactive_Company_Date = Date.Now
        MyCompanyClass.LastEdited = Date.Now
        MyCompanyClass.CompanyUserDefinedDate = Date.Now

        MyCompanyClass.AnotherField = ""
        MyCompanyClass.Debtor_Number = ""
        MyCompanyClass.UserDefined1 = ""
        MyCompanyClass.UserDefined2 = ""
        MyCompanyClass.UserDefined3 = ""
        MyCompanyClass.UserDefined4 = ""
        MyCompanyClass.UserDefined5 = ""
        MyCompanyClass.UserDefined6 = ""



        If MyCompanyClass.Save("") Then
            mreturn = True
        Else

            mreturn = False
        End If

        Return mreturn
    End Function



    'ProgressBar1.Show()
    'Dim Pstatus As Byte
    'Dim ai As Integer
    'Dim Tstate As String
    '    Pstatus = 0
    '    For ai = 0 To 100000 'filenames.Length - 1
    '        If Pstatus = 1 Then Exit For
    '        ProgressBar1.Value = (ai + 1) / 100000 * 100 'filenames.Length * 100
    '        Tstate = ProgressBar1.Value
    '        Label9.Text = Tstate + "%"
    '        Application.DoEvents()
    '    Next ai

    'Private Sub Btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.Close()
    'End Sub
    Private Sub BtnDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDetail1.Click
        Dim frm As New ImportDetail
        frm.txtDuplicate.Text = ""

        If RecnotSave.Length > 0 Then
            frm.txtdontImport.Text = RecnotSave.Substring(1, RecnotSave.Length - 1)
        End If
        If DuplicateRec.Length > 0 Then
            frm.txtDuplicate.Text = DuplicateRec.Substring(1, DuplicateRec.Length - 1)
        End If
        frm.ShowDialog()
    End Sub
    Private Function ChkColumnName(ByVal name As String) As Boolean
        If ColumnS.IndexOf(name) > -1 Then
            Return True
        Else
            Return False
        End If

    End Function

End Class