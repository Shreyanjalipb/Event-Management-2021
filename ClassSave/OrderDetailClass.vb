Public Class OrderDetailClass
    Private mId As String
    Private mOrderNo As String
    Private mOrderDate As String
    Private mKPBNo As String
    Private mMaker As String
    Private mPartNo As String
    Private mPartName As String
    Private mQty As String
    Private mDelDate As String
    Property Id() As String
        Get
            Return mId
        End Get
        Set(ByVal value As String)
            mId = value
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
    Property OrderDate() As String
        Get
            Return mOrderDate
        End Get
        Set(ByVal value As String)
            mOrderDate = value
        End Set
    End Property
    Property KPBNo() As String
        Get
            Return mKPBNo
        End Get
        Set(ByVal value As String)
            mKPBNo = value
        End Set
    End Property
    Property Maker() As String
        Get
            Return mMaker
        End Get
        Set(ByVal value As String)
            mMaker = value
        End Set
    End Property
    Property PartNo() As String
        Get
            Return mPartNo
        End Get
        Set(ByVal value As String)
            mPartNo = value
        End Set
    End Property
    Property PartName() As String
        Get
            Return mPartName
        End Get
        Set(ByVal value As String)
            mPartName = value
        End Set
    End Property
    Property Qty() As String
        Get
            Return mQty
        End Get
        Set(ByVal value As String)
            mQty = value
        End Set
    End Property
    Property DelDate() As String
        Get
            Return mDelDate
        End Get
        Set(ByVal value As String)
            mDelDate = value
        End Set
    End Property
    Function Save(ByVal Id As String)
        Dim ReMsg As Boolean
        Try

        
            Dim Sql As String

            Dim prId As New SqlClient.SqlParameter("@OrderDetailId", SqlDbType.NVarChar, 50)
            prId.Value = Id

            Dim prOrderNo As New SqlClient.SqlParameter("@OrderNo", SqlDbType.NChar, 8)
            prOrderNo.Value = OrderNo

            Dim prOrderDate As New SqlClient.SqlParameter("@OrderDate", SqlDbType.NChar, 8)
            prOrderDate.Value = OrderDate

            Dim prKPBNo As New SqlClient.SqlParameter("@KPBNo", SqlDbType.NChar, 10)
            prKPBNo.Value = KPBNo

            Dim prMaker As New SqlClient.SqlParameter("@Maker", SqlDbType.NVarChar, 250)
            prMaker.Value = Maker

            Dim prPartNo As New SqlClient.SqlParameter("@PartNo", SqlDbType.NVarChar, 12)
            prPartNo.Value = PartNo

            Dim prPartName As New SqlClient.SqlParameter("@PartName", SqlDbType.NVarChar, 250)
            prPartName.Value = PartName

            Dim prQty As New SqlClient.SqlParameter("@Qty", SqlDbType.NChar, 10)
            prQty.Value = Qty

            Dim prDelDate As New SqlClient.SqlParameter("@DelDate", SqlDbType.NChar, 8)
            prDelDate.Value = DelDate

            Dim prLastEditBy As New SqlClient.SqlParameter("@LastEditBy", SqlDbType.NVarChar, 20)
            prLastEditBy.Value = LoginId()

            SetLangUS()
            Dim prLastEditDate As New SqlClient.SqlParameter("@LastEditDate", SqlDbType.DateTime)
            prLastEditDate.Value = Now.Date

            Sql = String.Concat(" select [OrderDetail].OrderDetailId from [OrderDetail] where [OrderDetail].OrderDetailId = @OrderDetailId ", _
                                    " if @@Rowcount = 0 ", _
                                    " INSERT INTO [OrderDetail] ([OrderDetail].OrderDetailId,[OrderDetail].OrderNo, [OrderDetail].OrderDate, [OrderDetail].KPBNo, [OrderDetail].Maker, [OrderDetail].PartNo, [OrderDetail].PartName, [OrderDetail].Qty, [OrderDetail].DelDate,[OrderDetail].LastEditBy,[OrderDetail].LastEditDate)", _
                                    " VALUES (@OrderDetailId, @OrderNo, @OrderDate, @KPBNo, @Maker, @PartNo, @PartName, @Qty, @DelDate,@LastEditBy,@LastEditDate)", _
                                    " else", _
                                    " update [OrderDetail] set", _
                                    " [OrderDetail].OrderDetailId=@OrderDetailId,[OrderDetail].OrderNo=@OrderNo, [OrderDetail].OrderDate=@OrderDate, [OrderDetail].KPBNo=@KPBNo, [OrderDetail].Maker=@Maker, [OrderDetail].PartNo=@PartNo, [OrderDetail].PartName=@PartName, [OrderDetail].Qty=@Qty, [OrderDetail].DelDate=@DelDate,[OrderDetail].LastEditBy=@LastEditBy,[OrderDetail].LastEditDate=@LastEditDate", _
                                    " where [OrderDetail].OrderDetailId = @OrderDetailId")

            ReMsg = URunSql(Sql, prId, prOrderNo, prOrderDate, prKPBNo, prMaker, prPartNo, prPartName, prQty, prDelDate, prLastEditBy, prLastEditDate)

            prId = Nothing
            prOrderNo = Nothing
            prOrderDate = Nothing
            prKPBNo = Nothing
            prMaker = Nothing
            prPartNo = Nothing
            prPartName = Nothing
            prQty = Nothing
            prDelDate = Nothing
        Catch ex As Exception

        End Try
        Return ReMsg
    End Function

    Function CheckDuplicate(ByVal OrderNo As String, ByVal Orderdate As String, ByVal PartNo As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT PartNo FROM OrderDetail where PartNo ='", PartNo, "' and OrderNo ='", OrderNo, "' and OrderDate='", Orderdate, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        'If dt.Rows.Count < 1 Then
        'SBU
        If dt.Rows.Count > 0 Then
            'End sbu 
            Mreturn = True
        End If
        Return Mreturn
    End Function
    Function RetriveDataOrder(Optional ByVal mOrder As String = "") As DataTable
        Dim SQL As String
        If mOrder = "" Then
            SQL = String.Concat("select *, CAST( 0 AS INT) AS Delivery from [OrderDetail] ")
        Else
            SQL = String.Concat("select *, CAST( 0 AS INT) AS Delivery from [OrderDetail] where OrderNo like '", mOrder, "%'")
        End If

        Return UReadSqlbase(ConnectionString, SQL, "")
    End Function

    Function RetriveDataOrder_Del(Optional ByVal mOrder As String = "") As DataTable
        Dim SQL As String
        If mOrder = "" Then
            SQL = String.Concat("select *, CAST( 0 AS INT) AS Delivery from [OrderDetail] ")
        Else
            SQL = String.Concat("select *, CAST( 0 AS INT) AS Delivery from [OrderDetail] where OrderNo like '", mOrder, "%'")
        End If

        Return UReadSqlbase(ConnectionString, SQL, "")
    End Function


    Function LoadDataOrderDetail(Optional ByVal mOrder As String = "") As DataTable
        Dim SQL As String
        Dt_Show = New DataTable
        If mOrder = "" Then
            SQL = String.Concat("select *, CAST( 0 AS INT) AS Delivery from [OrderDetail] ")
        Else
            SQL = String.Concat("select *, CAST( 0 AS INT) AS Delivery from [OrderDetail] where OrderNo like '", mOrder, "%'")
        End If
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
            'OrderDetailId, OrderNo, OrderDate, KPBNo, Maker, PartNo, PartName, Qty, DelDate

            Dt_Show.Columns.Add("OrderDetailId")
            Dt_Show.Columns.Add("OrderNo")
            Dt_Show.Columns.Add("OrderDate")
            Dt_Show.Columns.Add("KPBNo")
            Dt_Show.Columns.Add("Maker")
            Dt_Show.Columns.Add("PartNo")
            Dt_Show.Columns.Add("PartName")
            Dt_Show.Columns.Add("Qty")
            Dt_Show.Columns.Add("DelDate")
            Dt_Show.Columns.Add("Delivery")
            Dim dtRow As DataRow

            For i = 0 To dt.Rows.Count - 1
                dtRow = Dt_Show.NewRow()
                If IsDBNull(dt.Rows(i).Item("OrderDetailId")) Then
                    dtRow.Item("OrderDetailId") = ""
                Else
                    dtRow.Item("OrderDetailId") = dt.Rows(i).Item("OrderDetailId")
                End If
                If IsDBNull(dt.Rows(i).Item("OrderNo")) Then
                    dtRow.Item("OrderNo") = ""
                Else
                    dtRow.Item("OrderNo") = dt.Rows(i).Item("OrderNo")
                End If

                If IsDBNull(dt.Rows(i).Item("OrderDate")) Then
                    dtRow.Item("OrderDate") = ""
                Else
                    dtRow.Item("OrderDate") = dt.Rows(i).Item("OrderDate")
                End If

                If IsDBNull(dt.Rows(i).Item("KPBNo")) Then
                    dtRow.Item("KPBNo") = ""
                Else
                    dtRow.Item("KPBNo") = dt.Rows(i).Item("KPBNo")
                End If

                If IsDBNull(dt.Rows(i).Item("Maker")) Then
                    dtRow.Item("Maker") = ""
                Else
                    dtRow.Item("Maker") = dt.Rows(i).Item("Maker")
                End If

                If IsDBNull(dt.Rows(i).Item("PartNo")) Then
                    dtRow.Item("PartNo") = ""
                Else
                    dtRow.Item("PartNo") = dt.Rows(i).Item("PartNo")
                End If

                If IsDBNull(dt.Rows(i).Item("PartName")) Then
                    dtRow.Item("PartName") = ""
                Else
                    dtRow.Item("PartName") = dt.Rows(i).Item("PartName")
                End If

                If IsDBNull(dt.Rows(i).Item("Qty")) Then
                    dtRow.Item("Qty") = ""
                Else
                    dtRow.Item("Qty") = dt.Rows(i).Item("Qty")
                End If

                If IsDBNull(dt.Rows(i).Item("DelDate")) Then
                    dtRow.Item("DelDate") = ""
                Else
                    dtRow.Item("DelDate") = dt.Rows(i).Item("DelDate")
                End If

                dtRow.Item("Delivery") = MyOrderProcessClass.CheckCount(Trim(dt.Rows(i).Item("OrderDetailId").ToString()))
                Dt_Show.Rows.Add(dtRow)
            Next

        Else

            Dt_Show = dt
        End If

        Return Dt_Show
    End Function

    Function GetOrderDetailId(ByVal Partno As String, ByVal Orderno As String) As String
        Dim mreturn As String = ""
        Dim SQL As String = String.Concat("select OrderDetailId from [OrderDetail] where PartNo ='", Partno, "' and OrderNo ='", Orderno, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("OrderDetailId")) Then
                mreturn = ""
            Else
                mreturn = dt.Rows(0).Item("OrderDetailId").ToString()
            End If

        End If
        Return mreturn
    End Function
    Function Delete(ByVal OrderDetailId As String) As Boolean
        Dim mreturn As Boolean = False
        Try
            Dim sql As String = ""
            sql = String.Concat("DELETE FROM OrderDetail WHERE OrderDetailId='", OrderDetailId, "'")
            mreturn = UExecuteQuery(sql)
        Catch ex As Exception

        End Try

        Return mreturn
    End Function
    Function GetOrderQty(ByVal OrderDetailId As String) As Integer
        Dim mreturn As Integer = 0
        Dim SQL As String = String.Concat("select Qty from [OrderDetail] where OrderDetailId ='", OrderDetailId, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("Qty")) Then
                mreturn = ""
            Else
                mreturn = dt.Rows(0).Item("Qty")
            End If

        End If
        Return mreturn
    End Function

    Function CheckOrder1(ByVal OrderNo As String, ByVal OrderDate As String, ByVal PartNo As String) As String
        Dim Mreturn As String = ""
        Dim sql As String = ""
        sql = String.Concat("SELECT OrderDetailId FROM [OrderDetail] where OrderNo='" & OrderNo & "' and OrderDate='", OrderDate, "' and PartNo='", PartNo, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            Mreturn = dt.Rows(0).Item("OrderDetailId")
        Else
            Mreturn = ""
        End If
        Return Mreturn
    End Function

    Function GetOrderNo(ByVal OrderDetailId As String) As String
        Dim mreturn As String = ""
        Dim SQL As String = String.Concat("select OrderNo from [OrderDetail] where OrderDetailId ='", OrderDetailId, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("OrderNo")) Then
                mreturn = ""
            Else
                mreturn = dt.Rows(0).Item("OrderNo")
            End If

        End If
        Return mreturn
    End Function
    Function DeleteOrderDetail(ByVal OrderNo As String) As Boolean
        Dim mreturn As Boolean
        Try
            Dim SQL As String = String.Concat("select OrderDetailId from [OrderDetail] where OrderNo ='", OrderNo, "'")
            Dim dt As DataTable = UReadSqlbase(ConnectionString, SQL, "")
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    If MyOrderProcessClass.CheckCount(Trim(dt.Rows(i).Item("OrderDetailId").ToString())) = 0 Then
                        mreturn = Delete(Trim(dt.Rows(i).Item("OrderDetailId").ToString()))
                    End If

                Next

            End If

        Catch ex As Exception

        End Try
        
        Return mreturn
    End Function
End Class
