
Public Class ClassUser
    Private mGUID As String
    Private mFullName As String
    Private mUserName As String
    Private mPassword As String
    Private mConfirm As String
    Private mUserType As String
    Private mDeleteYN As Boolean
    Private mCheckUser As Boolean
    Private mUserDup As Boolean
    Private mStatus As String
    Property Status() As String
        Get
            Return mStatus
        End Get
        Set(ByVal value As String)
            mStatus = value
        End Set
    End Property
    Property UserDup() As Boolean
        Get
            Return mUserDup
        End Get
        Set(ByVal value As Boolean)
            mUserDup = value
        End Set
    End Property
    Public Property FullName() As String
        Get
            Return mFullName
        End Get
        Set(ByVal value As String)
            mFullName = value
        End Set
    End Property
    Property UserName() As String
        Get
            Return mUserName
        End Get
        Set(ByVal value As String)
            mUserName = value
        End Set
    End Property
    Property PassWord() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property
    Property ConFirmPass() As String
        Get
            Return mConfirm
        End Get
        Set(ByVal value As String)
            mConfirm = value
        End Set
    End Property
    Property UserType() As String
        Get
            Return mUserType
        End Get
        Set(ByVal value As String)
            mUserType = value
        End Set
    End Property
    Public Property CheckedUser() As Boolean
        Get
            Return mCheckUser
        End Get
        Set(ByVal value As Boolean)
            mCheckUser = value
        End Set
    End Property
    Public Property DeleteYN() As Boolean
        Get
            Return mDeleteYN
        End Get
        Set(ByVal value As Boolean)
            mDeleteYN = value
        End Set
    End Property
    ''' <summary>
    ''' Save Value From Form UserAdd
    ''' </summary>
    ''' <param name="mGUID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Save(Optional ByVal mGUID As String = "")
        Dim param_GUID As New SqlClient.SqlParameter("@GUID", SqlDbType.NVarChar, 50)
        param_GUID.Value = System.Guid.NewGuid.ToString

        Dim param_FullName As New SqlClient.SqlParameter("@FullName", SqlDbType.NVarChar, 50)
        param_FullName.Value = FullName
        Dim param_UserName As New SqlClient.SqlParameter("@UserName", SqlDbType.NVarChar, 20)
        param_UserName.Value = UserName
        Dim param_Password As New SqlClient.SqlParameter("@Password", SqlDbType.NVarChar, 20)
        param_Password.Value = PassWord
        Dim param_UserType As New SqlClient.SqlParameter("@UserType", SqlDbType.NVarChar, 10)
        param_UserType.Value = UserType

        Dim prLastEditBy As New SqlClient.SqlParameter("@LastEditBy", SqlDbType.NVarChar, 20)
        prLastEditBy.Value = LoginId()

        SetLangUS()
        Dim prLastEditDate As New SqlClient.SqlParameter("@LastEditDate", SqlDbType.DateTime)
        prLastEditDate.Value = Now.Date


        Dim mSql As String
        mSql = String.Concat("select [Id] from [User] where [UserName] = @UserName ", _
                "if @@Rowcount = 0 ", _
                    "insert into [User] (Id,", _
                    "FullName,UserName,Password,Type,LastEditBy,LastEditDate", _
                    ")", _
                    " values (@GUID,", _
                    "@FullName,@UserName,@Password,@UserType,@LastEditBy,@LastEditDate", _
                    ") ", _
                " else ", _
                    "update [User] set ", _
                    "FullName = @FullName,UserName = @UserName,Password = @Password,Type = @UserType,LastEditBy=@LastEditBy,LastEditDate=@LastEditDate", _
                    " where [UserName] = @UserName ")

        Save = UExecuteQuery(mSql, param_GUID, param_FullName, param_UserName, param_Password, param_UserType, prLastEditBy, prLastEditDate)
        param_GUID = Nothing

        param_FullName = Nothing
        param_Password = Nothing
        param_UserName = Nothing
        param_UserType = Nothing
    End Function
    Function UpdateYN(ByVal mSelect As String)
        Dim mSql As String
        SetLangUS()
        mSql = String.Concat("update [User] set ", _
                    "DeleteYN = '1',LastEditBy='", LoginId(), "',LastEditDate='", Now.Date, "'", _
                    " where [UserName] = '" & mSelect & "'")
        UpdateYN = UExecuteQuery(mSql)
    End Function
    Function RetriveUser(Optional ByVal mData As String = "", Optional ByVal mselect As String = "") As DataTable
        Dim SQL As String
        If Trim(mData) = "" And Trim(mselect) = "" Then
            SQL = "select * from [User] where DeleteYN = 'False'"
        Else
            If Trim(mselect) = "U" Then
                SQL = String.Concat("select * from [User] where UserName like '" & mData & "%' and DeleteYN = 'False'")
            Else
                SQL = String.Concat("select * from [User] where FullName like '" & mData & "%' and DeleteYN = 'False'")
            End If
        End If

        Return UReadSqlbase(ConnectionString, SQL, "")
    End Function
    Function RetriveAllUser(ByVal username As String) As DataTable
        Dim SQL As String
        SQL = String.Concat("select * from [User] where UserName = '" & username & "' and [User].DeleteYN='False'")
        Return UReadSqlbase(ConnectionString, SQL, "")
    End Function

    Function CheckAddUser(ByVal Fullname As String, ByVal UserName As String, ByVal Password As String, ByVal Type As String) As Boolean
        Dim mreturn As Boolean = False
        Dim SQL As String = String.Concat("select * from [User] where [User].Fullname ='", Fullname, "' and [User].UserName='", UserName, "' and [User].Password='", Password, "' and [User].Type='", Type, "' and [User].DeleteYN ='True'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
            mreturn = UpdateYN_N(UserName)
        End If

        'Id, Fullname, UserName, Password, Type, DeleteYN, LastEditBy, LastEditDate
        Return mreturn
    End Function
    Function UpdateYN_N(ByVal mSelect As String)
        Dim mSql As String
        SetLangUS()
        mSql = String.Concat("update [User] set ", _
                    "DeleteYN = '0',LastEditBy='", LoginId(), "',LastEditDate='", Now.Date, "'", _
                    " where [UserName] = '" & mSelect & "'")
        UpdateYN_N = UExecuteQuery(mSql)
    End Function

End Class
