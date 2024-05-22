Public Class OrderProcessClass
    Private mId As String
    Private mOrderDetailId As String
    Private mOrderNo As String
    Private mSerialNo As String
    Private mPackingDate As Date
    Private mReceiveDate As Date
    Private mDeliveryDate As Date

    Property Id() As String
        Get
            Return mId
        End Get
        Set(ByVal value As String)
            mId = value
        End Set
    End Property
    Property OrderDetailId() As String
        Get
            Return mOrderDetailId
        End Get
        Set(ByVal value As String)
            mOrderDetailId = value
        End Set
    End Property
    Property OrderNo() As String
        Get
            Return mOrderNo
        End Get
        Set(ByVal value As String)
            mOrderNo = value
        End Set
    End Property
    Property SerialNo() As String
        Get
            Return mSerialNo
        End Get
        Set(ByVal value As String)
            mSerialNo = value
        End Set
    End Property

    Property PackingDate() As Date
        Get
            Return mPackingDate
        End Get
        Set(ByVal value As Date)
            mPackingDate = value
        End Set
    End Property
    Property ReceiveDate() As Date
        Get
            Return mReceiveDate
        End Get
        Set(ByVal value As Date)
            mReceiveDate = value
        End Set
    End Property
    Property DeliveryDate() As Date
        Get
            Return mDeliveryDate
        End Get
        Set(ByVal value As Date)
            mDeliveryDate = value
        End Set
    End Property

    Function Save(ByVal Id As String)
        Dim ReMsg As Boolean

        Try
            'PId, OrderDetailId, OrderNo, SerialNo, PackingDate, ReceiveDate, DeliveryDate
            Dim Sql As String
            Dim prId As New SqlClient.SqlParameter("@PId", SqlDbType.NVarChar, 50)
            prId.Value = Id

            Dim prOrderDetailId As New SqlClient.SqlParameter("@OrderDetailId", SqlDbType.NVarChar, 50)
            prOrderDetailId.Value = OrderDetailId

            Dim prOrderNo As New SqlClient.SqlParameter("@OrderNo", SqlDbType.NVarChar, 8)
            prOrderNo.Value = OrderNo

            Dim prSerialNo As New SqlClient.SqlParameter("@SerialNo", SqlDbType.NVarChar, 50)
            prSerialNo.Value = SerialNo

            Dim prLastEditBy As New SqlClient.SqlParameter("@LastEditBy", SqlDbType.NVarChar, 20)
            prLastEditBy.Value = LoginId()

            SetLangUS()
            Dim prLastEditDate As New SqlClient.SqlParameter("@LastEditDate", SqlDbType.DateTime)
            prLastEditDate.Value = Now.Date

            'Dim prPackingDate As New SqlClient.SqlParameter("@PackingDate", SqlDbType.DateTime)
            'prPackingDate.Value = PackingDate

            'Dim prReceiveDate As New SqlClient.SqlParameter("@ReceiveDate", SqlDbType.DateTime)
            'prReceiveDate.Value = ReceiveDate

            'Dim prDeliveryDate As New SqlClient.SqlParameter("@DeliveryDate", SqlDbType.DateTime)
            'prDeliveryDate.Value = DeliveryDate

            'Sql = String.Concat(" select PId from OrderProcess where PId = @PId ", _
            '                        " if @@Rowcount = 0 ", _
            '                        " INSERT INTO OrderProcess(PId, OrderDetailId, OrderNo, SerialNo, PackingDate, ReceiveDate, DeliveryDate)", _
            '                        " VALUES (@PId, @OrderDetailId, @OrderNo, @SerialNo, @PackingDate, @ReceiveDate, @DeliveryDate)", _
            '                        " else", _
            '                        " update OrderProcess set", _
            '                        " PId=@PId, OrderDetailId=@OrderDetailId, OrderNo=@OrderNo, SerialNo=@SerialNo, PackingDate=@PackingDate, ReceiveDate=@ReceiveDate, DeliveryDate=@DeliveryDate", _
            '                        " where PId=@PId")


            Sql = String.Concat(" select PId from OrderProcess where PId = @PId ", _
                                    " if @@Rowcount = 0 ", _
                                    " INSERT INTO OrderProcess(PId, OrderDetailId, OrderNo, SerialNo,LastEditBy,LastEditDate)", _
                                    " VALUES (@PId, @OrderDetailId, @OrderNo, @SerialNo,@LastEditBy,@LastEditDate)", _
                                    " else", _
                                    " update OrderProcess set", _
                                    " PId=@PId, OrderDetailId=@OrderDetailId, OrderNo=@OrderNo, SerialNo=@SerialNo,LastEditBy=@LastEditBy,LastEditDate=@LastEditDate", _
                                    " where PId=@PId")


            'PId, OrderDetailId, OrderNo, SerialNo, PackingDate, ReceiveDate, DeliveryDate
            ' ReMsg = URunSql(Sql, prId, prOrderDetailId, prOrderNo, prSerialNo, prPackingDate, prReceiveDate, prDeliveryDate)
            ReMsg = URunSql(Sql, prId, prOrderDetailId, prOrderNo, prSerialNo, prLastEditBy, prLastEditDate)
            prId.Value = ""
            prOrderDetailId.Value = ""
            prOrderNo.Value = ""
            prSerialNo.Value = ""

        Catch ex As Exception
            MsgBox("Error:" & ex.ToString)
        End Try
        Return ReMsg
    End Function

    Function CheckCount(ByVal OrderDetailId As String) As Integer
        Dim mreturn As Integer

        Dim SQL As String = String.Concat("select OrderDetailId from [OrderProcess] where OrderDetailId ='", OrderDetailId, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        mreturn = dt.Rows.Count
        Return mreturn
    End Function

    Function CheckOrder(ByVal SerialNo As String) As String
        Dim mreturn As String = ""
        '  PackingDate, ReceiveDate, DeliveryDate
        Dim SQL As String = String.Concat("select * from [OrderProcess] where SerialNo ='", SerialNo, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("PackingDate")) And IsDBNull(dt.Rows(0).Item("ReceiveDate")) And IsDBNull(dt.Rows(0).Item("DeliveryDate")) Then
                mreturn = MyOrderDetailClass.GetOrderNo(Trim(dt.Rows(0).Item("OrderDetailId").ToString()))
            Else
                If dt.Rows(0).Item("PackingDate").ToString() <> "" And dt.Rows(0).Item("ReceiveDate").ToString() <> "" And dt.Rows(0).Item("DeliveryDate").ToString() <> "" Then
                    mreturn = ""
                Else
                    mreturn = MyOrderDetailClass.GetOrderNo(Trim(dt.Rows(0).Item("OrderDetailId").ToString()))

                End If

            End If
        End If

        Return mreturn
    End Function
    Function LoadOrderProcessInList(ByVal OrderDetailId As String) As DataTable
        Dim sql As String = ""
        sql = String.Concat("SELECT * from [OrderProcess] where OrderDetailId ='", OrderDetailId, "'")
        Return UReadSqlbase(ConnectionString, sql, "")
    End Function
    Function UpDateOrderProcess(ByVal FieldUpdate As String, ByVal Mvalue As Date, ByVal OrderProcessId As String) As Boolean
        Dim mreturn As Boolean = False
        Try
            Dim sql As String = ""
            sql = String.Concat(" update [OrderProcess] set [OrderProcess].", FieldUpdate, " = '", Mvalue, "',[OrderProcess].LastEditBy='", LoginId(), "',[OrderProcess].LastEditDate='", Mvalue, "' where PId='", OrderProcessId, "'")
            mreturn = UExecuteQuery(sql)
        Catch ex As Exception
            Return False
        End Try
        Return mreturn
    End Function
    Function CheckUpdateStatus(ByVal OrderProcessId As String) As Boolean
        Dim mreturn As Boolean
        Dim SQL As String = String.Concat("select * from [OrderProcess] where PId ='", OrderProcessId, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
       
           If IsDBNull(dt.Rows(0).Item("PackingDate")) And IsDBNull(dt.Rows(0).Item("ReceiveDate")) And IsDBNull(dt.Rows(0).Item("DeliveryDate")) Then

            Else
                If dt.Rows(0).Item("PackingDate").ToString() <> "" And dt.Rows(0).Item("ReceiveDate").ToString() <> "" And dt.Rows(0).Item("DeliveryDate").ToString() <> "" Then

                    If MySerailClass.UpdateSerialStatus(dt.Rows(0).Item("SerialNo").ToString(), "F") Then
                        mreturn = True
                    End If
                End If
               
            End If


        End If

        Return mreturn
    End Function


    

End Class
