Imports System.Collections.Generic
Public Module GModule
    Public myUser As New ClassUser
    Private Ds As Data.DataTable
    Public Dt_Show As DataTable
    'Private mTranslate As FrontTrans.Translator
    Public ConfigDay As Integer
    Public ConfigDis As Integer
    'Dim CMain As New MainClass
    Public ROUseId As String
    Public RoUserRepair As String
    Public WO_Use_ROId As String
    Private mLoginId As String
    Private mLoginEmpId As String
    Private mLoginFullName As String
    Private mLoginApprove As Boolean
    Private mLoginType As Integer
    Private mLoginRO As Boolean 'can use RO
    Private mLoginWO As Boolean 'can use WO
    Private mLoginCO As Boolean 'can use CO
    Private mvtmpID As String
    Private mvtmpSearch As String
    Public KanbanEdit As Boolean
    Private mKanbanId As String
    Private mKanban_PartNoKanban As String
    Private mSerialNubberEdit As Boolean
    Private mSerialNubber As String
    Private _mPartNo As String
    Public Morder As String
    Public DataNotImport As String
    Public Const DEFAULT_ENCKEY As String = "123135566"
    'Private FRM As New FrmMain
    'Private mparam As DTCParam

    'Calendar for employee
    Private mCalendar As String
    Private mCalendarValue As String

    Private _LangFormat As String
    ''' <summary>
    ''' Start
    ''' </summary>
    ''' <remarks></remarks>
    Sub Main()
        ''login
        'If GParam.IsparamExistYN = False Then
        '    'Load Config
        '    Dim Frm_config As New FrmConfig
        '    Frm_config.ShowDialog()
        'Else
        '    GParam.ReadProperty()
        'End If
        ''start main
        'Dim frm As New Main
        'Dim Ds As New Data.DataTable
        'Ds = CMain.Config
        'Dim i As Integer
        'If Not Ds Is Nothing Then
        '    For i = 0 To Ds.Rows.Count - 1
        '        ConfigDay = Ds.Rows(i).Item("Day")
        '        ConfigDis = Ds.Rows(i).Item("Distance")
        '    Next
        'End If
        'Ds = Nothing
        'frm.ShowDialog()
        Dim frm As New Main
        frm.ShowDialog()
    End Sub
    Property PMDs() As Data.DataTable
        Get
            Return Ds
        End Get
        Set(ByVal value As Data.DataTable)
            Ds = value
        End Set
    End Property

    Property LangFormat() As String
        Get
            Return _LangFormat
        End Get
        Set(ByVal value As String)
            _LangFormat = value
        End Set
    End Property

    Property LoginId() As String
        Get
            Return mLoginId
        End Get
        Set(ByVal value As String)
            mLoginId = value
        End Set
    End Property
    Property LoginEmpId() As String
        Get
            Return mLoginEmpId
        End Get
        Set(ByVal value As String)
            mLoginEmpId = value
        End Set
    End Property
    Property LoginFullName() As String
        Get
            Return mLoginFullName
        End Get
        Set(ByVal value As String)
            mLoginFullName = value
        End Set
    End Property
    Property LoginApprove() As Boolean
        Get
            Return mLoginApprove
        End Get
        Set(ByVal value As Boolean)
            mLoginApprove = value
        End Set
    End Property
    Property LoginType() As Integer
        Get
            Return mLoginType
        End Get
        Set(ByVal value As Integer)
            mLoginType = value
        End Set
    End Property
    Property LoginRO() As Boolean
        Get
            Return mLoginRO
        End Get
        Set(ByVal value As Boolean)
            mLoginRO = value
        End Set
    End Property
    Property LoginWO() As Boolean
        Get
            Return mLoginWO
        End Get
        Set(ByVal value As Boolean)
            mLoginWO = value
        End Set
    End Property
    Property LoginCO() As Boolean
        Get
            Return mLoginCO
        End Get
        Set(ByVal value As Boolean)
            mLoginCO = value
        End Set
    End Property

    Property empCalendar() As String
        Get
            Return mCalendar
        End Get
        Set(ByVal value As String)
            mCalendar = value
        End Set
    End Property
    Property empCalendarValue() As Date
        Get
            Return mCalendarValue
        End Get
        Set(ByVal value As Date)
            mCalendarValue = value
        End Set
    End Property
    Property tmpID() As String
        Get
            Return mvtmpID
        End Get
        Set(ByVal value As String)
            mvtmpID = value
        End Set
    End Property
    Property tmpSearch() As String
        Get
            Return mvtmpSearch
        End Get
        Set(ByVal value As String)
            mvtmpSearch = value
        End Set
    End Property

    'Public ReadOnly Property GParam() As DTCParam
    '    Get
    '        If mparam Is Nothing Then mparam = New DTCParam
    '        Return mparam
    '    End Get
    'End Property
    Public Function Translate(ByVal Tagid As String, ByVal ParamArray param() As String) As String
        Dim mreturn As String
        'If mTranslate Is Nothing Then
        '    mTranslate = New FrontTrans.Translator(True, "frontwaretranslator.fw")
        '    mTranslate.Application.ApplicationID = Val(GParam.Trans_AppID)
        '    'mTranslate.DebugMode = GParam.Trans_DebugTransYN
        '    mTranslate.Language.ISO = GParam.Trans_LangID
        'End If
        'mreturn = mTranslate.Translate(Tagid, param)
        Return mreturn
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub WriteLog(ByVal ParamArray mlog() As String)
        'GParam.WriteLog(mlog)
    End Sub
    Private _MyPartClass As PartClass
    Public ReadOnly Property MyPartClass() As PartClass
        Get
            If _MyPartClass Is Nothing Then
                _MyPartClass = New PartClass
            End If
            Return _MyPartClass
        End Get
    End Property
    Private _MySerailClass As SerailClass
    Public ReadOnly Property MySerailClass() As SerailClass
        Get
            If _MySerailClass Is Nothing Then
                _MySerailClass = New SerailClass
            End If
            Return _MySerailClass
        End Get
    End Property

    Private _MyOrderClass As OrderClass
    Public ReadOnly Property MyOrderClass() As OrderClass
        Get
            If _MyOrderClass Is Nothing Then
                _MyOrderClass = New OrderClass
            End If
            Return _MyOrderClass
        End Get
    End Property
    Private _MyOrderDetailClass As OrderDetailClass
    Public ReadOnly Property MyOrderDetailClass() As OrderDetailClass
        Get
            If _MyOrderDetailClass Is Nothing Then
                _MyOrderDetailClass = New OrderDetailClass
            End If
            Return _MyOrderDetailClass
        End Get
    End Property
    Private _MyOrderProcessClass As OrderProcessClass
    Public ReadOnly Property MyOrderProcessClass() As OrderProcessClass
        Get
            If _MyOrderProcessClass Is Nothing Then
                _MyOrderProcessClass = New OrderProcessClass
            End If
            Return _MyOrderProcessClass
        End Get
    End Property
    Private _MyConfigClass As ConfigClass
    Public ReadOnly Property MyConfigClass() As ConfigClass
        Get
            If _MyConfigClass Is Nothing Then
                _MyConfigClass = New ConfigClass
            End If
            Return _MyConfigClass
        End Get
    End Property



    Private _MyCompanyClass As CompanyClass
    Public ReadOnly Property MyCompanyClass() As CompanyClass
        Get
            If _MyCompanyClass Is Nothing Then
                _MyCompanyClass = New CompanyClass
            End If
            Return _MyCompanyClass
        End Get
    End Property
    Private _MyContactClass As ContactClass
    Public ReadOnly Property MyContactClass() As ContactClass
        Get
            If _MyContactClass Is Nothing Then
                _MyContactClass = New ContactClass
            End If
            Return _MyContactClass
        End Get
    End Property

    Private _ContactList As List(Of ContactClass)
    Property MyContactList() As List(Of ContactClass)
        Get
            If _ContactList Is Nothing Then
                _ContactList = New List(Of ContactClass)
            End If
            Return _ContactList
        End Get
        Set(ByVal value As List(Of ContactClass))
            _ContactList = value
        End Set
    End Property


    Property KanbanId() As String
        Get
            Return mKanbanId
        End Get
        Set(ByVal value As String)
            mKanbanId = value
        End Set
    End Property
    Property Kanban_PartNoKanban() As String
        Get
            Return mKanban_PartNoKanban
        End Get
        Set(ByVal value As String)
            mKanban_PartNoKanban = value
        End Set
    End Property
    Property SerialNubber() As String
        Get
            Return mSerialNubber
        End Get
        Set(ByVal value As String)
            mSerialNubber = value
        End Set
    End Property
    Property SerialNubberEdit() As Boolean
        Get
            Return mSerialNubberEdit
        End Get
        Set(ByVal value As Boolean)
            mSerialNubberEdit = value
        End Set
    End Property
    Property _PartNo() As String
        Get
            Return _mPartNo
        End Get
        Set(ByVal value As String)
            _mPartNo = value
        End Set
    End Property

    Public Function Get_RandomNumber(ByVal Low As Integer, ByVal High As Integer) As Integer
        Return Math.Floor((High - Low + 1) * Rnd() + Low)
    End Function




End Module
