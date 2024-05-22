Public Class OrderClass
    Private mOrderNo As String
    Private mOrderDate As String
    Private mDate As Date
    Private mPartcode As String
    Private mDeleteYN As Boolean

    Property OrderNo() As String
        Get
            Return mOrderNo
        End Get
        Set(ByVal value As String)
            mOrderNo = value
        End Set
    End Property
    Property OrderDate() As String
        Get
            Return mOrderDate
        End Get
        Set(ByVal value As String)
            mOrderDate = value
        End Set
    End Property
    Property _Date() As Date
        Get
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
        End Set
    End Property
    Property Partcode() As String
        Get
            Return mPartcode
        End Get
        Set(ByVal value As String)
            mPartcode = value
        End Set
    End Property
    Property DeleteYN() As Boolean
        Get
            Return mDeleteYN
        End Get
        Set(ByVal value As Boolean)
            mDeleteYN = value
        End Set
    End Property

    Function Save(ByVal Id As String)
        Dim ReMsg As Boolean
        Try

            Dim Sql As String
            'LastEditBy, LastEditDate
            Dim prOrderNo As New SqlClient.SqlParameter("@OrderNo", SqlDbType.NChar, 8)
            prOrderNo.Value = Id

            Dim prOrderDate As New SqlClient.SqlParameter("@OrderDate", SqlDbType.NChar, 8)
            prOrderDate.Value = OrderDate

            Dim prDate As New SqlClient.SqlParameter("@Date", SqlDbType.DateTime)
            prDate.Value = _Date

            Dim prDeleteYN As New SqlClient.SqlParameter("@DeleteYN", SqlDbType.Bit)
            prDeleteYN.Value = DeleteYN

            Dim prLastEditBy As New SqlClient.SqlParameter("@LastEditBy", SqlDbType.NVarChar, 20)
            prLastEditBy.Value = LoginId()

            SetLangUS()
            Dim prLastEditDate As New SqlClient.SqlParameter("@LastEditDate", SqlDbType.DateTime)
            prLastEditDate.Value = Now.Date

            Sql = String.Concat(" select [Order].OrderNo from [Order] where [Order].OrderNo = @OrderNo ", _
                                    " if @@Rowcount = 0 ", _
                                    " INSERT INTO [Order] ([Order].OrderNo,[Order].OrderDate,[Order].Date,[Order].DeleteYN,[Order].LastEditBy,[Order].LastEditDate)", _
                                    " VALUES (@OrderNo,@OrderDate,@Date,@DeleteYN,@LastEditBy,@LastEditDate)", _
                                    " else", _
                                    " update [Order] set", _
                                    " [Order].OrderNo = @OrderNo,[Order].OrderDate = @OrderDate,[Order].Date = @Date,[Order].DeleteYN = @DeleteYN,[Order].LastEditBy=@LastEditBy,[Order].LastEditDate=@LastEditDate", _
                                    " where [Order].OrderNo = @OrderNo")

            ReMsg = URunSql(Sql, prOrderNo, prOrderDate, prDate, prDeleteYN, prLastEditBy, prLastEditDate)

            prOrderNo.Value = Nothing
            prOrderDate.Value = Nothing
            prDate.Value = Nothing
            prDeleteYN.Value = Nothing
        Catch ex As Exception

        End Try

        Return ReMsg
    End Function
    Function CheckOrder(ByVal OrderNo As String, ByVal OrderDate As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT OrderNo FROM [Order] where OrderNo='" & OrderNo & "' and OrderDate='", OrderDate, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count < 1 Then
            Mreturn = True
        End If
        Return Mreturn
    End Function

    Function RetriveOrderPartTag(ByVal OrderDate As String) As DataTable
        Dim SQL As String
        SQL = String.Concat("select * from [Order] where OrderDate like'" & OrderDate & "%' and DeleteYN='False'")
        Return UReadSqlbase(ConnectionString, SQL, "")
    End Function


    Function GetOrderdate(ByVal OrderNo As String) As String
        Dim Mreturn As String = ""
        Dim sql As String = ""
        sql = String.Concat("SELECT OrderDate FROM [Order] where OrderNo='" & OrderNo & "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("OrderDate")) Then
                Mreturn = UFormatDate2(Now)
            Else
                Mreturn = dt.Rows(0).Item("OrderDate")
            End If
        End If
        Return Mreturn
    End Function
    Function DeleteOrder(ByVal OrderNo As String) As Boolean
        Dim sql As String = ""
        SetLangUS()
        sql = String.Concat("update [Order] set [Order].DeleteYN =1,[Order].LastEditBy='", LoginId(), "',[Order].LastEditDate='", Now.Date, "' where [Order].OrderNo ='", OrderNo, "'")
        Return UExecuteQuery(sql)
    End Function
End Class
