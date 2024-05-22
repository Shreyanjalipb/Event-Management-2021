Public Class CompanyClass

    'Company_RecNo,Member_Number,Company_Name,Address_1,Address_2,City,State
    'Post_Code,Country,Other,Phone,Fax,Web_Page,Company_Colour
    'Created,Inactive_Company_Date,LastEdited,CompanyUserDefinedDate
    'AnotherField,Debtor_Number,UserDefined1,UserDefined2
    'UserDefined3,UserDefined4,UserDefined5,UserDefined6,Address_3 
    Private mCompany_RecNo As Long = 0
    Property Company_RecNo() As Long
        Get
            Return mCompany_RecNo
        End Get
        Set(ByVal value As Long)
            mCompany_RecNo = value
        End Set
    End Property
    Private mMember_Number As String = ""
    Property Member_Number() As String
        Get
            Return mMember_Number
        End Get
        Set(ByVal value As String)
            mMember_Number = value
        End Set
    End Property
    Private mCompany_Name As String = ""
    Property Company_Name() As String
        Get
            Return mCompany_Name
        End Get
        Set(ByVal value As String)
            mCompany_Name = value
        End Set
    End Property
    Private mAddress_1 As String = ""
    Property Address_1() As String
        Get
            Return mAddress_1
        End Get
        Set(ByVal value As String)
            mAddress_1 = value
        End Set
    End Property
    Private mAddress_2 As String = ""
    Property Address_2() As String
        Get
            Return mAddress_2
        End Get
        Set(ByVal value As String)
            mAddress_2 = value
        End Set
    End Property
    Private mCity As String = ""
    Property City() As String
        Get
            Return mCity
        End Get
        Set(ByVal value As String)
            mCity = value
        End Set
    End Property
    Private mState As String = ""
    Property State() As String
        Get
            Return mState
        End Get
        Set(ByVal value As String)
            mState = value
        End Set
    End Property
    Private mPost_Code As String = ""
    Property Post_Code() As String
        Get
            Return mPost_Code
        End Get
        Set(ByVal value As String)
            mPost_Code = value
        End Set
    End Property
    Private mCountry As String = ""
    Property Country() As String
        Get
            Return mCountry
        End Get
        Set(ByVal value As String)
            mCountry = value
        End Set
    End Property
    Private mOther As String = ""
    Property Other() As String
        Get
            Return mOther
        End Get
        Set(ByVal value As String)
            mOther = value
        End Set
    End Property
    Private mPhone As String = ""
    Property Phone() As String
        Get
            Return mPhone
        End Get
        Set(ByVal value As String)
            mPhone = value
        End Set
    End Property
    Private mFax As String = ""
    Property Fax() As String
        Get
            Return mFax
        End Get
        Set(ByVal value As String)
            mFax = value
        End Set
    End Property
    Private mWeb_Page As String = ""
    Property Web_Page() As String
        Get
            Return mWeb_Page
        End Get
        Set(ByVal value As String)
            mWeb_Page = value
        End Set
    End Property
    Private mCompany_Colour As Long
    Property Company_Colour() As Long
        Get
            Return mCompany_Colour
        End Get
        Set(ByVal value As Long)
            mCompany_Colour = value
        End Set
    End Property
    Private mCreated As Date
    Property Created() As Date
        Get
            Return mCreated
        End Get
        Set(ByVal value As Date)
            mCreated = value
        End Set
    End Property
    Private mInactive_Company_Date As Date
    Property Inactive_Company_Date() As Date
        Get
            Return mInactive_Company_Date
        End Get
        Set(ByVal value As Date)
            mInactive_Company_Date = value
        End Set
    End Property
    Private mLastEdited As Date
    Property LastEdited() As Date
        Get
            Return mLastEdited
        End Get
        Set(ByVal value As Date)
            mLastEdited = value
        End Set
    End Property
    Private mCompanyUserDefinedDate As Date
    Property CompanyUserDefinedDate() As Date
        Get
            Return mCompanyUserDefinedDate
        End Get
        Set(ByVal value As Date)
            mCompanyUserDefinedDate = value
        End Set
    End Property
    Private mAnotherField As String = ""
    Property AnotherField() As String
        Get
            Return mAnotherField
        End Get
        Set(ByVal value As String)
            mAnotherField = value
        End Set
    End Property
    Private mDebtor_Number As String = ""
    Property Debtor_Number() As String
        Get
            Return mDebtor_Number
        End Get
        Set(ByVal value As String)
            mDebtor_Number = value
        End Set
    End Property
    Private mUserDefined1 As String = ""
    Property UserDefined1() As String
        Get
            Return mUserDefined1
        End Get
        Set(ByVal value As String)
            mUserDefined1 = value
        End Set
    End Property
    Private mUserDefined2 As String = ""
    Property UserDefined2() As String
        Get
            Return mUserDefined2
        End Get
        Set(ByVal value As String)
            mUserDefined2 = value
        End Set
    End Property
    Private mUserDefined3 As String = ""
    Property UserDefined3() As String
        Get
            Return mUserDefined3
        End Get
        Set(ByVal value As String)
            mUserDefined3 = value
        End Set
    End Property
    Private mUserDefined4 As String = ""
    Property UserDefined4() As String
        Get
            Return mUserDefined4
        End Get
        Set(ByVal value As String)
            mUserDefined4 = value
        End Set
    End Property
    Private mUserDefined5 As String = ""
    Property UserDefined5() As String
        Get
            Return mUserDefined5
        End Get
        Set(ByVal value As String)
            mUserDefined5 = value
        End Set
    End Property
    Private mUserDefined6 As String = ""
    Property UserDefined6() As String
        Get
            Return mUserDefined6
        End Get
        Set(ByVal value As String)
            mUserDefined6 = value
        End Set
    End Property
    Private mAddress_3 As String = ""
    Property Address_3() As String
        Get
            Return mAddress_3
        End Get
        Set(ByVal value As String)
            mAddress_3 = value
        End Set
    End Property


    Function Save(Optional ByVal mGUID As String = "")
        Dim ReMsg As Boolean
        Try
            Dim param_Company_RecNo As New SqlClient.SqlParameter("@Company_RecNo", SqlDbType.BigInt, 4)
            param_Company_RecNo.Value = Company_RecNo
            Dim param_Member_Number As New SqlClient.SqlParameter("@Member_Number", SqlDbType.NVarChar, 10)
            param_Member_Number.Value = Member_Number
            Dim param_Company_Name As New SqlClient.SqlParameter("@Company_Name", SqlDbType.NVarChar, 255)
            param_Company_Name.Value = Company_Name
            Dim param_Address_1 As New SqlClient.SqlParameter("@Address_1", SqlDbType.NVarChar, 50)
            param_Address_1.Value = Address_1
            Dim param_Address_2 As New SqlClient.SqlParameter("@Address_2", SqlDbType.NVarChar, 50)
            param_Address_2.Value = Address_2
            Dim param_City As New SqlClient.SqlParameter("@City", SqlDbType.NVarChar, 50)
            param_City.Value = City
            Dim param_State As New SqlClient.SqlParameter("@State", SqlDbType.NVarChar, 20)
            param_State.Value = State
            Dim param_Post_Code As New SqlClient.SqlParameter("@Post_Code", SqlDbType.NVarChar, 20)
            param_Post_Code.Value = Post_Code
            Dim param_Country As New SqlClient.SqlParameter("@Country", SqlDbType.NVarChar, 30)
            param_Country.Value = Country
            Dim param_Other As New SqlClient.SqlParameter("@Other", SqlDbType.NVarChar, 50)
            param_Other.Value = Other
            Dim param_Phone As New SqlClient.SqlParameter("@Phone", SqlDbType.NVarChar, 20)
            param_Phone.Value = Phone
            Dim param_Fax As New SqlClient.SqlParameter("@Fax", SqlDbType.NVarChar, 20)
            param_Fax.Value = Fax
            Dim param_Web_Page As New SqlClient.SqlParameter("@Web_Page", SqlDbType.NVarChar, 255)
            param_Web_Page.Value = Web_Page
            Dim param_Company_Colour As New SqlClient.SqlParameter("@Company_Colour", SqlDbType.BigInt, 4)
            param_Company_Colour.Value = Company_Colour
            Dim param_Created As New SqlClient.SqlParameter("@Created", SqlDbType.DateTime)
            param_Created.Value = Created
            'Dim param_Inactive_Company_Date As New SqlClient.SqlParameter("@Inactive_Company_Date", SqlDbType.DateTime)
            'param_Inactive_Company_Date.Value = Inactive_Company_Date
            Dim param_LastEdited As New SqlClient.SqlParameter("@LastEdited", SqlDbType.DateTime)
            param_LastEdited.Value = LastEdited
            Dim param_CompanyUserDefinedDate As New SqlClient.SqlParameter("@CompanyUserDefinedDate", SqlDbType.DateTime)
            param_CompanyUserDefinedDate.Value = CompanyUserDefinedDate
            Dim param_AnotherField As New SqlClient.SqlParameter("@AnotherField", SqlDbType.NVarChar, 50)
            param_AnotherField.Value = AnotherField
            Dim param_Debtor_Number As New SqlClient.SqlParameter("@Debtor_Number", SqlDbType.NVarChar, 20)
            param_Debtor_Number.Value = Debtor_Number
            Dim param_UserDefined1 As New SqlClient.SqlParameter("@UserDefined1", SqlDbType.NVarChar, 50)
            param_UserDefined1.Value = UserDefined1
            Dim param_UserDefined2 As New SqlClient.SqlParameter("@UserDefined2", SqlDbType.NVarChar, 50)
            param_UserDefined2.Value = UserDefined2
            Dim param_UserDefined3 As New SqlClient.SqlParameter("@UserDefined3", SqlDbType.NVarChar, 50)
            param_UserDefined3.Value = UserDefined3
            Dim param_UserDefined4 As New SqlClient.SqlParameter("@UserDefined4", SqlDbType.NVarChar, 50)
            param_UserDefined4.Value = UserDefined4
            Dim param_UserDefined5 As New SqlClient.SqlParameter("@UserDefined5", SqlDbType.NVarChar, 50)
            param_UserDefined5.Value = UserDefined5
            Dim param_UserDefined6 As New SqlClient.SqlParameter("@UserDefined6", SqlDbType.NVarChar, 50)
            param_UserDefined6.Value = UserDefined6
            Dim param_Address_3 As New SqlClient.SqlParameter("@Address_3", SqlDbType.NVarChar, 50)
            param_Address_3.Value = Address_3

            Dim mSql As String
            mSql = String.Concat("select [Company_RecNo] from [COMPANY] where [Company_RecNo] = @Company_RecNo ", _
                    "if @@Rowcount = 0 ", _
                        "insert into [COMPANY] (Company_RecNo,", _
                      "Member_Number,Company_Name,Address_1,Address_2,City,State,", _
                      "Post_Code,Country,Other,Phone,Fax,Web_Page,Company_Colour,", _
                      "Created,LastEdited,CompanyUserDefinedDate,", _
                      "AnotherField,Debtor_Number,UserDefined1,UserDefined2,", _
                      "UserDefined3,UserDefined4,UserDefined5,UserDefined6,Address_3 ", _
                       ")", _
                        " values (@Company_RecNo,", _
                       "@Member_Number,@Company_Name,@Address_1,@Address_2,@City,@State,", _
                      "@Post_Code,@Country,@Other,@Phone,@Fax,@Web_Page,@Company_Colour,", _
                      "@Created,@LastEdited,@CompanyUserDefinedDate,", _
                      "@AnotherField,@Debtor_Number,@UserDefined1,@UserDefined2,", _
                      "@UserDefined3,@UserDefined4,@UserDefined5,@UserDefined6,@Address_3 ", _
                       ") ", _
                    " else ", _
                        "update [COMPANY] set ", _
                        "Company_RecNo=@Company_RecNo,", _
                        "Member_Number=@Member_Number,Company_Name=@Company_Name,Address_1=@Address_1,Address_2=@Address_2,City=@City,State=@State,", _
                      "Post_Code=@Post_Code,Country=@Country,Other=@Other,Phone=@Phone,Fax=@Fax,Web_Page=@Web_Page,Company_Colour=@Company_Colour,", _
                      "Created=@Created,LastEdited=@LastEdited,CompanyUserDefinedDate=@CompanyUserDefinedDate,", _
                      "AnotherField=@AnotherField,Debtor_Number=@Debtor_Number,UserDefined1=@UserDefined1,UserDefined2=@UserDefined2,", _
                      "UserDefined3=@UserDefined3,UserDefined4=@UserDefined4,UserDefined5=@UserDefined5,UserDefined6=@UserDefined6,Address_3=@Address_3", _
                        " where [Company_RecNo] = @Company_RecNo ")
            ReMsg = URunSql(mSql, param_Company_RecNo, param_Member_Number, param_Company_Name, param_Address_1, param_Address_2, param_City, param_State, _
            param_Post_Code, param_Country, param_Other, param_Phone, param_Fax, param_Web_Page, param_Company_Colour, _
            param_Created, param_LastEdited, param_CompanyUserDefinedDate, _
            param_AnotherField, param_Debtor_Number, param_UserDefined1, param_UserDefined2, _
            param_UserDefined3, param_UserDefined4, param_UserDefined5, param_UserDefined6, param_Address_3)


            param_Company_RecNo = Nothing
            param_Member_Number = Nothing
            param_Company_Name = Nothing
            param_Address_1 = Nothing
            param_Address_2 = Nothing
            param_City = Nothing
            param_State = Nothing
            param_Post_Code = Nothing
            param_Country = Nothing
            param_Other = Nothing
            param_Phone = Nothing
            param_Fax = Nothing
            param_Web_Page = Nothing
            param_Company_Colour = Nothing
            param_Created = Nothing
            ' param_Inactive_Company_Date = Nothing
            param_LastEdited = Nothing
            param_CompanyUserDefinedDate = Nothing
            param_AnotherField = Nothing
            param_Debtor_Number = Nothing
            param_UserDefined1 = Nothing
            param_UserDefined2 = Nothing
            param_UserDefined3 = Nothing
            param_UserDefined4 = Nothing
            param_UserDefined5 = Nothing
            param_UserDefined6 = Nothing
            param_Address_3 = Nothing
        Catch ex As Exception
            MsgBox("Save CompanyClass Error:" & ex.ToString)
        End Try
        Return ReMsg
    End Function

    Function GetTopData(ByVal FieldName As String) As String
        Dim mreturn As String = ""
        Dim sql As String = ""
        sql = String.Concat("SELECT TOP 1 ", FieldName, " FROM [COMPANY] ORDER BY ", FieldName, " DESC ")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item(FieldName)) Then
                mreturn = ""
            Else
                mreturn = dt.Rows(0).Item(FieldName)
            End If
        End If
        Return mreturn
    End Function

    Function GetCompanyRecNo(ByVal CompanyName As String) As Long
        Dim mreturn As Long
        Dim sql As String = ""
        sql = String.Concat("SELECT Company_RecNo FROM [COMPANY] WHERE Company_Name='" + CompanyName + "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("Company_RecNo")) Then
                mreturn = 0
            Else
                mreturn = Convert.ToInt32(dt.Rows(0).Item("Company_RecNo").ToString())
            End If
        End If
        Return mreturn
    End Function

    Function CheckCompany(ByVal CompanyName As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT Company_Name FROM [COMPANY] where Company_Name='" & CompanyName & "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count < 1 Then
            Mreturn = False
        Else
            Mreturn = True
        End If
        Return Mreturn
    End Function
End Class
