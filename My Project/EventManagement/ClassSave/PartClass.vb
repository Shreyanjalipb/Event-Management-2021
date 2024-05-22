Public Class PartClass
    Private mNo As Integer
    Private mPartNo As String
    Private mPartname As String
    Private mPartcode As String
    Private mQty As String
    Private mRuibetsu As String
    Private mDeleteYN As Boolean
    Private mItem As Integer
    Private mKPBNo As String
    Property No() As Integer
        Get
            Return mNo
        End Get
        Set(ByVal value As Integer)
            mNo = value
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
    Property KPBNo() As String
        Get
            Return mKPBNo
        End Get
        Set(ByVal value As String)
            mKPBNo = value
        End Set
    End Property
    Property Partname() As String
        Get
            Return mPartname
        End Get
        Set(ByVal value As String)
            mPartname = value
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
    Property Qty() As String
        Get
            Return mQty
        End Get
        Set(ByVal value As String)
            mQty = value
        End Set
    End Property
    Property Ruibetsu() As String
        Get
            Return mRuibetsu
        End Get
        Set(ByVal value As String)
            mRuibetsu = value
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
    Property Item() As Integer
        Get
            Return mItem
        End Get
        Set(ByVal value As Integer)
            mItem = value
        End Set
    End Property

    Function Save(ByVal Id As String)
        Dim ReMsg As Boolean
        'If Id = "" Then
        '    'New
        'Else
        '    'Edit
        'End If
        Dim Sql As String
        Dim prPartNo As New SqlClient.SqlParameter("@PartNo", SqlDbType.NVarChar, 11) ' BuyDate,
        prPartNo.Value = Id

        Dim prKPBNo As New SqlClient.SqlParameter("@KPBNo", SqlDbType.NChar, 10) ' BuyDate,
        prKPBNo.Value = KPBNo

        Dim prPartName As New SqlClient.SqlParameter("@PartName", SqlDbType.NVarChar, 250) ' BuyDate,
        prPartName.Value = Partname

        Dim prPartcode As New SqlClient.SqlParameter("@Partcode", SqlDbType.NVarChar, 50) ' BuyDate,
        prPartcode.Value = Partcode

        Dim prQty As New SqlClient.SqlParameter("@Qty", SqlDbType.NChar, 10) ' BuyDate,
        prQty.Value = Qty

        Dim prRuibetsu As New SqlClient.SqlParameter("@Ruibetsu", SqlDbType.NChar, 2) ' BuyDate,
        prRuibetsu.Value = Ruibetsu

        Dim prDelleteYN As New SqlClient.SqlParameter("@DelleteYN", SqlDbType.Bit) ' BuyDate,
        prDelleteYN.Value = DeleteYN

        Dim prItem As New SqlClient.SqlParameter("@Item", SqlDbType.Int) ' BuyDate,
        prItem.Value = Item

        Dim prLastEditBy As New SqlClient.SqlParameter("@LastEditBy", SqlDbType.NVarChar, 20)
        prLastEditBy.Value = LoginId()

        SetLangUS()
        Dim prLastEditDate As New SqlClient.SqlParameter("@LastEditDate", SqlDbType.DateTime)
        prLastEditDate.Value = Now.Date

        Sql = String.Concat(" select PartNo from Part where PartNo = @PartNo ", _
                                " if @@Rowcount = 0 ", _
                                " INSERT INTO Part (PartNo,KPBNo,PartName,PartCode,Qty", _
                                " ,Ruibetsu,DelleteYN,Item,LastEditBy,LastEditDate)", _
                                " VALUES (@PartNo,@KPBNo,@PartName,@Partcode,@Qty,@Ruibetsu", _
                                " ,@DelleteYN,@Item,@LastEditBy,@LastEditDate)", _
                                " else", _
                                " update Part set", _
                                " PartNo = @PartNo, KPBNo = @KPBNo,PartName = @PartName,Partcode = @Partcode", _
                                " ,Qty = @Qty,Ruibetsu = @Ruibetsu,DelleteYN = @DelleteYN,Item=@Item,LastEditBy=@LastEditBy,LastEditDate=@LastEditDate", _
                                " where PartNo = @PartNo")

        ReMsg = URunSql(Sql, prPartNo, prKPBNo, prPartName, prPartcode, prQty, prRuibetsu, prDelleteYN, prItem, prLastEditBy, prLastEditDate)

        prKPBNo.Value = Nothing
        prPartNo.Value = Nothing
        prPartName.Value = Nothing
        prPartcode.Value = Nothing
        prQty.Value = 0
        prRuibetsu.Value = Nothing
        prDelleteYN.Value = Nothing

        Return ReMsg
    End Function
    Function Delete(ByVal Id As Boolean)

        Dim Sql As String
        Sql = String.Concat(" select KPBNo from Part where Kanban = @KPBNo ", _
                                " if @@Rowcount = 0 ", _
                                " INSERT INTO Part(KPBNo,PartNo,PartName,PartCode,Qty", _
                                " ,Ruibetsu,DelleteYN)", _
                                " VALUES (@KPBNo,@PartNo,@PartName,@Partcode,@Qty,@Ruibetsu", _
                                " ,@DelleteYN)", _
                                " else", _
                                " update Part set", _
                                " KPBNo = @KPBNo,PartNo = @PartNo,PartName = @PartName,Partcode = @Partcode", _
                                " ,Qty = @Qty,Ruibetsu = @Ruibetsu,DelleteYN = @DelleteYN)", _
                                " where KPBNo = @KPBNo")
        Dim ReMsg As Boolean
        ReMsg = URunSql(Sql, "")
        Return ReMsg
    End Function
    Function RetriveDataInList() As DataTable
        Dim sql As String = ""
        sql = String.Concat("SELECT *FROM Part where DelleteYN=0 Order by  Item")
        Return UReadSqlbase(ConnectionString, sql, "")
    End Function
    Function GetKanbanvData(ByVal Id As String) As DataTable
        Dim sql As String = ""
        sql = String.Concat("SELECT *FROM Part where PartNo ='" & Id & "' ")
        Return UReadSqlbase(ConnectionString, sql, "")
    End Function
    Function DeleteKanban(ByVal PartNo As String) As Boolean
        Dim sql As String = ""
        SetLangUS()
        sql = String.Concat("update Part set DelleteYN =1,LastEditBy='", LoginId(), "',LastEditDate='", Now.Date, "' where PartNo ='", PartNo, "'")
        Return UExecuteQuery(sql)
    End Function
    Function CheckKanbanNo(ByVal KanbanNo As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT KPBNo FROM Part where KPBNo ='" & KanbanNo & "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count < 1 Then
            Mreturn = True
        End If
        Return Mreturn
    End Function

    Function CheckPartNo(ByVal PartNo As String) As Boolean
        Dim Mreturn As Boolean = False
        Dim sql As String = ""
        sql = String.Concat("SELECT PartNo FROM Part where PartNo ='" & PartNo & "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count < 1 Then
            Mreturn = True
        End If
        Return Mreturn
    End Function
    Function GetPartName(ByVal PartNo As String) As String
        Dim Mreturn As String = ""
        Dim sql As String = ""
        sql = String.Concat("SELECT PartName FROM Part where PartNo ='" & PartNo & "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("PartName")) Then
                Mreturn = ""
            Else
                Mreturn = dt.Rows(0).Item("PartName").ToString()
            End If
        End If
        Return Mreturn
    End Function
End Class
