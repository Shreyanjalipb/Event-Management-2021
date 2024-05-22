Imports Microsoft.Win32

'-------------------------------------------------------------
'  
' register module for .NET 
' KPO 22/03/2002
' source from : http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dndotnet/html/replaceapicalls.asp
'------------------------------------------------------------

Module REG_MODULE
    Private myReg As RegistryKey

    Public _ErrorRegModule_ As String

    Public Enum HKEY_REGISTER
        HKEY_CLASSES_ROOT = 1
        HKEY_LOCAL_MACHINE = 2
        HKEY_CURRENT_USER = 3
    End Enum
    Private Structure RegData
        Public Data As Object   ' Exact data 
        Public Value As String

        Public Overrides _
         Function ToString() As String
            Return Me.Value
        End Function
    End Structure

    Public Function GetRegValue(ByVal hKey As HKEY_REGISTER, ByVal lpszSubKey As String, ByVal szKey As String, ByVal szDefault As String) As Object
        Try

        Select Case hKey
            Case HKEY_REGISTER.HKEY_CLASSES_ROOT
                myReg = Registry.ClassesRoot
            Case HKEY_REGISTER.HKEY_CURRENT_USER
                myReg = Registry.CurrentUser
            Case HKEY_REGISTER.HKEY_LOCAL_MACHINE
                myReg = Registry.LocalMachine
        End Select
        'kPO 31/05/2002 so remove create key when 
        'strange problem occure when use in web         
        GetRegValue = szDefault
        If Not myReg Is Nothing Then

            myReg = myReg.OpenSubKey(lpszSubKey, False)
            If Not myReg Is Nothing Then
                GetRegValue = myReg.GetValue(szKey, szDefault)
            End If
        Else
        End If

        Catch ex As Exception
        _ErrorRegModule_ = ex.Message
        Return ""
        End Try
    End Function

    'change value input to object
    Public Function SetRegValue(ByVal hKey As HKEY_REGISTER, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As Object) As Boolean
        Try

        Select Case hKey
            Case HKEY_REGISTER.HKEY_CLASSES_ROOT
                myReg = Registry.ClassesRoot
                myReg = Registry.ClassesRoot.CreateSubKey(lpszSubKey)
            Case HKEY_REGISTER.HKEY_CURRENT_USER
                myReg = Registry.CurrentUser
                myReg = Registry.CurrentUser.CreateSubKey(lpszSubKey)
            Case HKEY_REGISTER.HKEY_LOCAL_MACHINE
                myReg = Registry.LocalMachine
                myReg = Registry.LocalMachine.CreateSubKey(lpszSubKey)
        End Select

        'myReg = Registry.LocalMachine.CreateSubKey(lpszSubKey)
        'myReg = Registry.CurrentUser
        'myReg.CreateSubKey(lpszSubKey)

        If Not myReg Is Nothing Then myReg.SetValue(sSetValue, sValue)
        SetRegValue = True

        Catch ex As Exception
        _ErrorRegModule_ = ex.Message
        End Try
    End Function


    Public Function DeleteRegKey(ByVal hKey As HKEY_REGISTER, ByVal SubKeyToDelete As String) As Long
        Try
        Select Case hKey
            Case HKEY_REGISTER.HKEY_CLASSES_ROOT
                myReg = Registry.ClassesRoot
            Case HKEY_REGISTER.HKEY_CURRENT_USER
                myReg = Registry.CurrentUser
            Case HKEY_REGISTER.HKEY_LOCAL_MACHINE
                myReg = Registry.LocalMachine
        End Select
        myReg.DeleteSubKeyTree(SubKeyToDelete)
        DeleteRegKey = True

        Catch ex As Exception
        _ErrorRegModule_ = ex.Message
        End Try
    End Function

    Public Function GetAllKeyValue(ByVal hKey As HKEY_REGISTER, ByVal lpszSubKey As String) As Object()
        Try

        ' Register function for .NET
        ' NGR : 26/05/2003
        ' Source from : http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dndotnet/html/replaceapicalls.asp
        Select Case hKey
            Case HKEY_REGISTER.HKEY_CLASSES_ROOT
                myReg = Registry.ClassesRoot
            Case HKEY_REGISTER.HKEY_CURRENT_USER
                myReg = Registry.CurrentUser
            Case HKEY_REGISTER.HKEY_LOCAL_MACHINE
                myReg = Registry.LocalMachine
        End Select
        Dim strKey As String = Trim(lpszSubKey)
        Dim strArrValue() As String
        Dim strValue As String = ""
        Dim strArrData() As Object
        Dim recData As RegData
        Dim intIdx As Integer = 0
        myReg = myReg.OpenSubKey(strKey) ' open key 
        ReDim strArrData(myReg.GetValueNames.Length - 1)
        strArrValue = myReg.GetValueNames()  ' return key to array of string
        For Each strValue In strArrValue
            recData.Value = strValue.ToString
            recData.Data = myReg.GetValue(strValue)
            strArrData(intIdx) = recData.Data
            intIdx += 1
        Next
        GetAllKeyValue = strArrData ' return array of string
        Erase strArrValue
        Erase strArrData
        strArrData = Nothing
        strArrValue = Nothing

        Catch ex As Exception
            _ErrorRegModule_ = ex.Message
            Return Nothing
        End Try
    End Function

End Module

