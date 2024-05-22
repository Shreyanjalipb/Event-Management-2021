
Public Class ConfigClass
    'Log file Parameters
    Private Const HKEY_LOCAL_MACHINE As Integer = REG_MODULE.HKEY_REGISTER.HKEY_LOCAL_MACHINE
    Private Const HKEY_CLASSES_ROOT As Integer = REG_MODULE.HKEY_REGISTER.HKEY_CLASSES_ROOT
    Private Const HKEY_CURRENT_USER As Integer = REG_MODULE.HKEY_REGISTER.HKEY_CURRENT_USER

    Private mvarLoadParams As Boolean
    Private Const DEFAULT_ENTRY_REGISTRY As String = "SOFTWARE\RoyalCliff\EventManagement"


    Private Const DEFAULT_DATABASE As String = "EventsPerfect_test"
    Private Const DEFAULT_SQLSERVERNAME As String = "192.168.10.247"
    Private Const DEFAULT_SQLUSERNAME As String = "sa"
    Private Const DEFAULT_SQLPASSWORD As String = ""

    Private mvarSqlDBName As String
    Private mvarSqlServerName As String
    Private mvarSqlUsername As String
    Private mvarSqlPassword As String


    Private Const REGPATH_SQLSERVER As String = "\SqlServer"

    Friend Property SqlServerName() As String
        Get
            SqlServerName = IIf(mvarLoadParams, mvarSqlServerName, GetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlServerName", DEFAULT_SQLSERVERNAME))
        End Get
        Set(ByVal Value As String)
            SetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlServerName", Value)
        End Set
    End Property

    Friend Property SqlUsername() As String
        Get
            SqlUsername = IIf(mvarLoadParams, mvarSqlUsername, GetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlUsername", DEFAULT_SQLUSERNAME))
        End Get
        Set(ByVal Value As String)
            SetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlUsername", Value)
        End Set
    End Property

    Friend Property SqlPassword() As String
        Get
            SqlPassword = IIf(mvarLoadParams, mvarSqlPassword, UStrDecode(GetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlPassword", UStrEncode(DEFAULT_SQLPASSWORD, DEFAULT_ENCKEY)), DEFAULT_ENCKEY))
        End Get
        Set(ByVal Value As String)
            SetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlPassword", UStrEncode(Value, DEFAULT_ENCKEY))
        End Set
    End Property

    Friend Property SqlDBName() As String
        Get
            SqlDBName = IIf(mvarLoadParams, mvarSqlDBName, GetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlDBName", DEFAULT_DATABASE))
        End Get
        Set(ByVal Value As String)
            SetRegValue(HKEY_LOCAL_MACHINE, String.Concat(DEFAULT_ENTRY_REGISTRY, REGPATH_SQLSERVER), "SqlDBName", Value)
        End Set
    End Property

End Class
