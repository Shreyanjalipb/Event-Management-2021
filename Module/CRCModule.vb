Imports System.Security.Cryptography

Module CRCModule
    '------------------------------------------------------
    'purpose  : calculate CRC
    'params   : mData - data to calculate crc
    '           mKey - data to mask
    'return   : crc string of data
    'kpo 11/08/2006 created
    '------------------------------------------------------
    '
    Public Function CalcCRCNET(ByVal mData As String, ByVal mKey As String) As String
       Dim _shacrypt As New SHA1CryptoServiceProvider
       _shacrypt.ComputeHash(System.Text.Encoding.UTF8.GetBytes(String.Concat(mKey, mData)))

        CalcCRCNET = Convert.ToBase64String(_shacrypt.Hash)

        _shacrypt.Clear()
        _shacrypt = Nothing
    End Function

    'for password never back..
    Function StrEncrypt(ByVal mData As String) As String
        Dim md5 As MD5CryptoServiceProvider
        Dim bytValue() As Byte
        Dim bytHash() As Byte

        ' Create New Crypto Service Provider Object
        md5 = New MD5CryptoServiceProvider

        ' Convert the original string to array of Bytes
        bytValue = System.Text.Encoding. _
            UTF8.GetBytes(mData)

        ' Compute the Hash, returns an array of Bytes
        bytHash = md5.ComputeHash(bytValue)

        ' Return a base 64 encoded string of the Hash value
        StrEncrypt = Convert.ToBase64String(bytHash)

        md5.Clear()
        Erase bytHash
        Erase bytValue

    End Function
    'check there is crypt ?
    Function ISThereISCrypt(ByVal mData As String) As Boolean
        Try
            Dim mtemp As String
            mtemp = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(mData))
            Return True
        Catch ex As Exception

        End Try
    End Function
    '------------------------------------------------------
    'purpose  : strencode - encode the string
    'params   : mData - data to encode
    '           mKey - key to mask
    'return   : string encode
    'kpo 11/08/2006 created
    '------------------------------------------------------
    '.net Encodeing.
    Function UStrEncode(ByVal mData As String, ByVal mKey As String) As String
        Dim md5 As RC2CryptoServiceProvider
        Dim bytValue() As Byte
        Dim bytHash() As Byte

        ' Create New Crypto Service Provider Object
        md5 = New RC2CryptoServiceProvider
        md5.Mode = CipherMode.CBC

        bytValue = System.Text.UTF7Encoding.UTF7.GetBytes(mData)
        bytHash = GetKeyByteArray(mKey)


        md5.IV = bytHash
        md5.Key = bytHash

        Dim mem1 As New IO.MemoryStream
        Dim encStream As New CryptoStream(mem1, _
            md5.CreateEncryptor, CryptoStreamMode.Write)

        encStream.Write(bytValue, 0, bytValue.Length)
        encStream.Close()

        bytValue = mem1.ToArray
        UStrEncode = Convert.ToBase64String(bytValue)

        md5.Clear()
        Erase bytHash
        Erase bytValue
        encStream = Nothing
        mem1.Close()
        mem1 = Nothing
    End Function

    '------------------------------------------------------
    'purpose  : strdecode - decode the string
    'params   : mData - data to decode
    '           mKey - key to mask
    'return   : string decode
    'kpo 11/08/2006 created
    '------------------------------------------------------
    Function UStrDecode(ByVal mData As String, ByVal mKey As String) As String
        If mData = "" Then Exit Function

        Dim md5 As RC2CryptoServiceProvider
        Dim bytValue() As Byte
        Dim bytHash() As Byte

        '' Create New Crypto Service Provider Object
        md5 = New RC2CryptoServiceProvider
        '' Convert the original string to array of Bytes
        bytValue = Convert.FromBase64String(mData)
        Dim mem1 As New IO.MemoryStream(bytValue, 0, bytValue.Length)

        bytHash = GetKeyByteArray(mKey)

        md5.Mode = CipherMode.CBC
        md5.Key = bytHash
        md5.IV = bytHash


        Dim encStream As New CryptoStream(mem1, _
            md5.CreateDecryptor, CryptoStreamMode.Read)
        Dim encstreamread As New IO.StreamReader(encStream)
        UStrDecode = encstreamread.ReadToEnd()

        encstreamread.Close()
        encstreamread = Nothing
        encStream.Close()

        md5.Clear()
        Erase bytHash
        Erase bytValue
        encStream = Nothing
        mem1.Close()
        mem1 = Nothing
    End Function


    '------------------------------------------------------
    'purpose  : get array of 8 chars use by strencode func.
    'params   : sPassword - string
    'return   : byte array 
    'kpo 11/08/2006 created
    '------------------------------------------------------
    Private Function GetKeyByteArray(ByVal sPassword As String) As Byte()
        Dim byteTemp(7) As Byte
        sPassword = sPassword.PadRight(8)   ' make sure we have 8 chars

        Dim iCharIndex As Integer
        For iCharIndex = 0 To 7
            byteTemp(iCharIndex) = Asc(Mid$(sPassword, iCharIndex + 1, 1))
        Next

        Return byteTemp
    End Function
End Module
