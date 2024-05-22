Public Class SerailClass
    Private mId As String
    Private mSerialNo As String
    Private mPartNo As String
    Private mDeleteYN As Boolean
    Private mSerialStatus As String
    Property Id() As String
        Get
            Return mId
        End Get
        Set(ByVal value As String)
            mId = value
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
    Property PartNo() As String
        Get
            Return mPartNo
        End Get
        Set(ByVal value As String)
            mPartNo = value
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
    Property SerialStatus() As String
        Get
            Return mSerialStatus
        End Get
        Set(ByVal value As String)
            mSerialStatus = value
        End Set
    End Property


    Function Save(ByVal Id As String)
        Dim ReMsg As Boolean

        Try

            Dim Sql As String
          
            Dim prId As New SqlClient.SqlParameter("@SerialNo", SqlDbType.NChar, 50) ' BuyDate,
            prId.Value = Id

            Dim prSerialNo As New SqlClient.SqlParameter("@Serial", SqlDbType.NVarChar, 10) ' BuyDate,
            prSerialNo.Value = SerialNo

            Dim prPartNo As New SqlClient.SqlParameter("@PartNo", SqlDbType.NVarChar, 11) ' BuyDate,
            prPartNo.Value = PartNo

            Dim prDeleteYN As New SqlClient.SqlParameter("@DeleteYN", SqlDbType.NVarChar, 50) ' BuyDate,
            prDeleteYN.Value = DeleteYN

            Dim prSerialStatus As New SqlClient.SqlParameter("@SerialStatus", SqlDbType.NChar, 10) ' BuyDate,
            prSerialStatus.Value = SerialStatus

            Dim prLastEditBy As New SqlClient.SqlParameter("@LastEditBy", SqlDbType.NVarChar, 20)
            prLastEditBy.Value = LoginId()

            SetLangUS()
            Dim prLastEditDate As New SqlClient.SqlParameter("@LastEditDate", SqlDbType.DateTime)
            prLastEditDate.Value = Now.Date

            Sql = String.Concat(" select SerialNo from Serial where SerialNo = @SerialNo ", _
                                    " if @@Rowcount = 0 ", _
                                    " INSERT INTO Serial(SerialNo,Serial,PartNo,DeleteYN,SerialStatus,LastEditBy,LastEditDate)", _
                                    " VALUES (@SerialNo,@Serial,@PartNo,@DeleteYN,@SerialStatus,@LastEditBy,@LastEditDate)", _
                                    " else", _
                                    " update Serial set", _
                                    " SerialNo = @SerialNo,Serial = @Serial,PartNo = @PartNo,DeleteYN = @DeleteYN,SerialStatus= @SerialStatus,LastEditBy=@LastEditBy,LastEditDate=@LastEditDate", _
                                    " where SerialNo = @SerialNo")

            
            ReMsg = URunSql(Sql, prId, prSerialNo, prPartNo, prDeleteYN, prSerialStatus, prLastEditBy, prLastEditDate)
            prId.Value = ""
            prSerialNo.Value = ""
            prPartNo.Value = ""
            prDeleteYN.Value = ""
        Catch ex As Exception
            MsgBox("Error:" & ex.ToString)
        End Try
        Return ReMsg
    End Function
    Function UpdateSerialStatus(ByVal _Serial As String, ByVal _SerialStatus As String)
        Dim ReMsg As Boolean
        Try
            Dim prSerialNo As New SqlClient.SqlParameter("@Serial", SqlDbType.NVarChar, 10)
            prSerialNo.Value = SerialNo
            Dim prSerialStatus As New SqlClient.SqlParameter("@SerialStatus", SqlDbType.NChar, 10)
            prSerialStatus.Value = SerialStatus

            SetLangUS()
            Dim Sql As String = String.Concat("update Serial set SerialStatus = '", _SerialStatus, "',LastEditBy='", LoginId(), "',LastEditDate='", Now.Date, "' where SerialNo ='", _Serial, "'")

            ReMsg = UExecuteQuery(Sql)
        Catch ex As Exception
            ReMsg = False
        End Try

        Return ReMsg
    End Function

    Function LoadDataInList(ByVal _PartNo As String) As DataTable
        Dim sql As String = ""
        sql = String.Concat("SELECT *FROM Serial where PartNo='" & _PartNo & "' and DeleteYN ='False'")
        Return UReadSqlbase(ConnectionString, sql, "")
    End Function
    Function LoadDataSerial(ByVal SerialNumber As String) As DataTable
        Dim sql As String = ""
        sql = String.Concat("SELECT SerialNo FROM Serial where SerialNo='" & SerialNumber & "'")
        Return UReadSqlbase(ConnectionString, sql, "")
    End Function
    Function CheckSerialNumber(ByVal SerialNumber As String, ByVal PartNo As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT Serial FROM Serial where Serial='", SerialNumber, "' and PartNo='", PartNo, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            Mreturn = False
        Else
            Mreturn = True
        End If
        Return Mreturn
    End Function


    Function Delete(ByVal SerialNumber As String) As Boolean
        Dim mreturn As Boolean = False
        Try
            SetLangUS()
            Dim sql As String = ""
            sql = String.Concat("update Serial set DeleteYN =1 ,LastEditBy='", LoginId(), "',LastEditDate='", Now.Date, "' where SerialNo ='", SerialNumber, "'")
            mreturn = UExecuteQuery(sql)
        Catch ex As Exception

        End Try
        
        Return mreturn
    End Function


    Function CheckSerial(ByVal Serial As String, ByVal SerialStatus As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT * FROM Serial where SerialNo='", Serial, "'and DeleteYN ='False' and SerialStatus ='", SerialStatus, "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("PartNo")) Then
                _PartNo = ""
            Else
                _PartNo = dt.Rows(0).Item("PartNo").ToString()
            End If

            Mreturn = True
        Else
            Mreturn = False
        End If
        Return Mreturn
    End Function

End Class
