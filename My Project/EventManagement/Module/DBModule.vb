
Public Module DBModule

    Private mvConnectionString As String
    Public Sub UInitialClass()
        'new param here
        'With GParam

        '    mvConnectionString = String.Concat("server=", .Sql_ServerName, ";uid=", .Sql_UserName, _
        '            ";pwd=", .Sql_Password, ";database=", .Sql_DatabaseName)
        'End With

        'age_kop   mvConnectionString = "Data Source=192.168.10.3;Initial Catalog=EventsPerfect;Integrated Security=True;User ID=sa;Password=;"
        mvConnectionString = "Data Source=" + MyConfigClass.SqlServerName + ";Initial Catalog=" + MyConfigClass.SqlDBName + ";Integrated Security=True;User ID=sa;Password=;"

        ' mvConnectionString = "Data Source=TOSHIBA-PC\SQLEXPRESS;Initial Catalog=Service;User Id=sa;Password=efa;"
        'MyConfigClass.LoadConFig()
        'mvConnectionString = "Data Source=" + MyConfigClass.DbSeverName + ";Initial Catalog=" + MyConfigClass.DbName + ";User Id=" + MyConfigClass.UserName + ";Password=" + MyConfigClass.Password + ";"
    End Sub

    ReadOnly Property ConnectionString() As String
        Get
            If mvConnectionString = "" Then UInitialClass()
            Return mvConnectionString
        End Get
    End Property

    '------------------------------------------------------
    'purpose : return the string or dbnull.value
    'params  : mString - string value
    'return  : return string or dbnull.value
    'modified: kpo 30/03/2006 created.
    '------------------------------------------------------
    Function USentStringtoDB(ByVal mString As String) As Object
        Try
            Return IIf(Trim(mString) <> "", mString, DBNull.Value)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Function USentDatetoDB(ByVal mDate As Date) As Object
        Try
            Return IIf(mDate > New Date(0), mDate, DBNull.Value)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    '------------------------------------------------------
    'purpose : return the database object (prevent crash because nothing)
    'params  : mDBobj - fieldvalue
    'return  : field value can return
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function UGetdbValue(ByVal mDBObj As Object) As Object
        Try
            Return mDBObj
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '------------------------------------------------------
    'purpose : connect the database and query the 'sql'
    'params  : msql - sql query
    '          mParams() - sqlparam array object
    'return  : datareader of the query
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function UReadSqlRead(ByVal mSql As String, ByVal ParamArray mParam() As Object) As SqlClient.SqlDataReader
        Return UReadSqlbaseRead(ConnectionString, mSql, mParam)
    End Function

    '------------------------------------------------------
    'purpose : connect the database and query the 'sql' specified the connection string
    'params  : mConnection - connection string
    '          msql - sql query
    '          mParams() - sqlparam array object
    'return  : datareader of the query
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function UReadSqlbaseRead(ByVal mConnection As String, ByVal mSql As String, ByVal ParamArray mParam() As Object) As SqlClient.SqlDataReader

        Dim mConn As New SqlClient.SqlConnection(mConnection)
        Try
            mConn.Open()
        Catch ex As Exception
            WriteLog(mSql)
            WriteLog(String.Format("Error while Get Sql Data from Query : {0}{1}{2}with Error : {3}{4}", vbCrLf, mSql, vbCrLf, vbCrLf, ex.Message))
            UReadSqlbaseRead = Nothing
            Exit Function
        End Try
        Dim mComm As New SqlClient.SqlCommand(mSql, mConn)
        Try
            Dim i_for As Integer
            For i_for = 0 To mParam.Length - 1
                mComm.Parameters.Add(mParam(i_for))
            Next
            Return mComm.ExecuteReader(CommandBehavior.CloseConnection)
        Catch ex As Exception
            WriteLog(mSql)
            WriteLog(String.Format("Error while Get Sql Data from Query : {0}{1}{2}with Error : {3}{4}", vbCrLf, mSql, vbCrLf, vbCrLf, ex.Message))
            UReadSqlbaseRead = Nothing
        End Try
    End Function

    '------------------------------------------------------
    'purpose : connect the database and query the 'sql' 
    'params  : msql - sql query
    '          mTable - tablename of datatable
    '          mParams() - sqlparam array object
    'return  : datatable of the query
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function UReadSql(ByVal mSql As String, ByVal mTable As String, ByVal ParamArray mParam() As Object) As DataTable
        Return UReadSqlbase(ConnectionString, mSql, mTable, mParam)
    End Function

    '------------------------------------------------------
    'purpose : connect the database and query the 'sql' specified the connection string
    'params  : mConnection - connection string
    '          msql - sql query
    '          mTable - tablename of datatable
    '          mParams() - sqlparam array object
    'return  : datatable of the query
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function UReadSqlbase(ByVal mConnection As String, ByVal mSql As String, ByVal mTable As String, ByVal ParamArray mParam() As Object) As DataTable

        Dim mConn As New SqlClient.SqlConnection(mConnection)
        Try
            mConn.Open()
        Catch ex As Exception
            WriteLog(mSql)
            WriteLog(String.Format("Error while Get Sql Data from Query : {0}{1}{2}with Error : {3}{4}", vbCrLf, mSql, vbCrLf, vbCrLf, ex.Message))
            MessageBox.Show(ex.Message)
            UReadSqlbase = Nothing
            Exit Function
        End Try
        Dim mComm As New SqlClient.SqlDataAdapter(mSql, mConn)
        Dim i_for As Integer
        For i_for = 0 To mParam.Length - 1
            mComm.SelectCommand.Parameters.Add(mParam(i_for))
        Next

        Dim mTable1 As New DataTable(mTable)
        Try
            mComm.Fill(mTable1)
        Catch ex As Exception
            WriteLog(mSql)
            WriteLog(String.Format("Error while Get Sql Data from Query : {0}{1}{2}with Error : {3}{4}", vbCrLf, mSql, vbCrLf, vbCrLf, ex.Message))
        End Try

        mComm = Nothing
        If (mConn.State = ConnectionState.Open) Then mConn.Close()
        Return mTable1

    End Function

    '------------------------------------------------------
    'purpose : connect the database and execute query 'sql' 
    'params  : msql - sql query
    '          mParams() - sqlparam array object
    'return  : success to execute query or not.
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function URunSql(ByVal mSql As String, ByVal ParamArray mParam() As Object) As Boolean
        Return URunSqlbase(ConnectionString, mSql, mParam)
    End Function

    '------------------------------------------------------
    'purpose : connect the database and  execute query 'sql' specified the connection string
    'params  : mConnection - connection string
    '          msql - sql query
    '          mParams() - sqlparam array object
    'return  : success to execute query or not.
    'modified: kpo 03/01/2005 created.
    '------------------------------------------------------
    Function URunSqlbase(ByVal mConnection As String, ByVal mSql As String, ByVal ParamArray mParam() As Object) As Boolean

        Dim mConn As New SqlClient.SqlConnection(mConnection)
        Try
            mConn.Open()
        Catch e As Exception
            WriteLog(mSql)
            WriteLog(String.Format("Error while execute Sql Data from Query : {0}{1}{2}with Error : {3}{4}", vbCrLf, mSql, vbCrLf, vbCrLf, e.Message))
            Exit Function
        End Try
        Dim mComm As New SqlClient.SqlCommand(mSql, mConn)
        Try
            Dim i_for As Integer
            For i_for = 0 To mParam.Length - 1
                mComm.Parameters.Add(mParam(i_for))
            Next
            mComm.ExecuteNonQuery()
            URunSqlbase = True
        Catch e As Exception
            MsgBox("Error in the URunSqlbase() :: " & vbCrLf & e.ToString)
            WriteLog(mSql)
            WriteLog(String.Format("Error while execute Sql Data from Query : {0}{1}{2}with Error : {3}{4}", vbCrLf, mSql, vbCrLf, vbCrLf, e.Message))
        End Try
        mComm.Parameters.Clear()
        mComm = Nothing
        If (mConn.State = ConnectionState.Open) Then mConn.Close()
    End Function
    Function UExecuteQuery(ByVal mSql As String, ByVal ParamArray mParam() As Object) As Boolean
        Return UExecuteQueryBase(ConnectionString, mSql, mParam)
    End Function
    Function UExecuteQueryBase(ByVal mConnection As String, ByVal mSql As String, ByVal ParamArray mParam() As Object) As Boolean

        Dim mConn As New SqlClient.SqlConnection(mConnection)
        Try
            mConn.Open()
        Catch e As Exception
            MsgBox("Error in the URunSqlbase() :: " & vbCrLf & e.ToString)
            Exit Function
        End Try
        Dim mComm As New SqlClient.SqlCommand(mSql, mConn)
        Try
            Dim i_for As Integer
            For i_for = 0 To mParam.Length - 1
                mComm.Parameters.Add(mParam(i_for))
            Next
            mComm.ExecuteNonQuery()
            UExecuteQueryBase = True
        Catch e As Exception
            MsgBox("Error in the URunSqlbase() :: " & vbCrLf & e.ToString)
        End Try
        mComm.Parameters.Clear()
        mComm = Nothing
        If (mConn.State = ConnectionState.Open) Then mConn.Close()
    End Function

    
End Module
