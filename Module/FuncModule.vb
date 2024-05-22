Imports System.Security.Cryptography

'******************************************************************
'  ** important 
'  besure that function you put here is ok, and be sure that you get
'  latest everytime you use 
'
'  this module objective : to put same function to use later
'******************************************************************

Public Module FuncModule
    Function UTrim(ByVal mString As String) As String
        Return Trim(mString.Replace(Chr(0), ""))
    End Function

    'convert format hh:mm minutes
    Function _UConvert2Time(ByVal mString As String) As Integer
        If Len(mString) >= 5 Then
            Dim mtime_hour As Integer = Val(Left(mString, 2))
            Dim mtime_min As Integer = Val(Mid(mString, 4, 2))
            Return (mtime_hour * 60) + mtime_min
        Else
            Return 0
        End If
    End Function
    'convert minutes to format hh:mm
    'sno : change from mminutes/60 to mminutes\60
    Function _UConvertFromTime(ByVal mminutes As Integer) As String
        Dim mhours As Integer = mminutes \ 60
        Dim mmins As Integer = mminutes Mod 60
        Return String.Format("{0:00}:{1:00}", mhours, mmins)
    End Function

    'convert Hours to Minute 8 to 480
    Function _UConvertHoursToMinute(ByVal mHours As Integer) As Integer
        Return mHours * 60
    End Function
    'convert Hours to Minute 480 to 8
    Function _UConvertMinuteToHours(ByVal mHours As Integer) As Integer
        Return mHours / 60
    End Function




    Function _UCombineStringSpace(ByVal mSeperator As String, ByVal ParamArray mData() As String) As String
        Dim mst1 As New System.Text.StringBuilder
        Dim ifor As Integer
        Dim b1stYN As Boolean = False
        For ifor = 0 To mData.Length - 1
                If Trim(mData(ifor)) <> "" Then
                   If b1stYN Then mst1.Append(mSeperator)
                   mst1.Append(mData(ifor))
                   b1stYN = True
                End If
        Next
        Return mst1.ToString
    End Function

    Function UGetDbString(ByVal mDbobj As Object) As String
        If IsDBNull(mDbobj) Then
            Return ""
        Else
            Return Trim(mDbobj)
        End If
    End Function

    Function _UCountingFreeTextLine(ByVal mtxtarea As String, ByVal mMaxCharsinLine As Integer) As Integer
        If trim(mtxtarea) = "" Then Exit Function
        Dim mtxtareas() As String
        mtxtareas = mtxtarea.Split(vbCrLf)
        Dim ifor As Integer
        Dim iline As Double
        For ifor = 0 To mtxtareas.Length - 1
            If mMaxCharsinLine = 0 Then
                iline += 1
            Else
                iline += Math.Ceiling(Len(mtxtareas(ifor)) / mMaxCharsinLine)
            End If
        Next
        If Trim(mtxtarea) = "" Then iline = 0
        Return iline
    End Function

    Function UFormatcurrency(ByVal mValue As Double, Optional ByVal mValidateYN As Boolean = True) As String
       If mValidateYN Then
          If (mValue >= 9999999.98) Or (mValue >= 9999999.99) Then Return "0,00"
       End If
       Return Replace(Replace(Replace(String.Format("{0:n2}", mValue), ",", "x"), ".", ","), "x", ".")
    End Function

    'format = dd/mm
    Function UformatECHEtoMF(ByVal mEche As String) As Integer
        Try
        Return Val(String.Concat(Left(mEche, 2), Mid(mEche, 4, 2)))
        Catch ex As Exception

        End Try
    End Function

    Function UFormatECHEFromMF(ByVal mECHE As String) As String
        Dim mtemp As String = String.Format("{0:0000}", Val(mECHE))
        Return String.Concat(Left(mtemp, 2), "/", Right(mtemp, 2))
    End Function

    '------------------------------------------------------
    'purpose  : write format date yyyymmddhhmmss
    'params   : mdate - date value
    'return   : date string
    'modified : sno 04/01/2006
    '           kpo 10/01/2006 if date is date 0 or error return ""
    '------------------------------------------------------
    Function UFormatDate2(ByVal mDate As Date, Optional ByVal mShowTimeYN As Boolean = False) As String
        UFormatDate2 = ""
        Try
            Dim mtimeformat As String = "{3:00}{4:00}{5:00}"
            If Not mShowTimeYN Then mtimeformat = ""
            If mDate > New Date(0) Then
                UFormatDate2 = String.Format(String.Concat("{2:0000}{1:00}{0:00}", mtimeformat), mDate.Day, mDate.Month, mDate.Year)
            End If
        Catch ex As Exception
        End Try
    End Function
    

    'purpose  : write format time from date
    'params   : mtime - time value
    'return   : time  string
    'modified : sno 15/03/2006
    'modified : sno 13/05/2008
    '------------------------------------------------------
    Function UFormatTime(ByVal mDateTime As Date) As String
        Dim mreturn As String = "00:00"
        If mDateTime > New Date(0) Then
            mreturn = String.Format("{0:00}:{1:00}", mDateTime.Hour, mDateTime.Minute)
        End If
        Return mreturn
    End Function
    '------------------------------------------------------
    'purpose : 'replace . as "", and replace, as .
    'param   : mst - string to cut
    '          mInd - start index
    '          mlength - length of string
    '          [mRemoveCRYN] remove enter ?
    'return  : string with specify length
    '------------------------------------------------------
    'get block of string
    'return  : prase string
    'modified: kpo 05/07/2006 
    '          kpo 11/10/2006 remove cr/lf
    '------------------------------------------------------
Function mParseCOMM(ByVal mst As String, ByVal mInd As Integer, ByVal mlength As Integer, _
    Optional ByVal mRemoveCRYN As Boolean = False) As String

    Dim iindex As Integer
    iindex = IIf(mInd = 1, 1, ((mInd - 1) * mlength) + 1)
    Dim mreturn As String = Mid(mst, iindex, mlength)
    If mRemoveCRYN Then mreturn = Replace(Replace(mreturn, vbCr, ""), vbLf, "")

    Return mreturn
End Function

    '------------------------------------------------------
    'purpose : 'replace . as "", and replace, as .
    'param   : mvalue - french currency
    'return  : real double
    'modified: kpo 30/05/2006 moved.
    '-----------------------------------------------------
    Public Function UConvertVal(ByVal mValue As String) As Double
        Return Val(Trim(mValue).Replace(".", "").Replace(",", "."))
    End Function
    Public Function UConvertCurrenctVal(ByVal mValue As String) As Double
        Return Val(Trim(mValue).Replace(",", "."))
    End Function
    Public Function UConvertCurrenctToCommaVal(ByVal mValue As String) As String
        Return Val(Trim(mValue).Replace(".", ","))
    End Function
'------------------------------------------------------
'purpose : format bank account number from xxx-xxxxxxx-xx to xxxxxxxxxxxx
'params  : mBankAcc - bankaccount format
'return  : bank account format for save in DB
'modified: kpo 23/03/2006 created.
'------------------------------------------------------
Function UFormatBANKtoDB(ByVal mBankAcc As String) As String
    UFormatBANKtoDB = ""
    If mBankAcc = "" Then Exit Function
    Return Replace(mBankAcc, "-", "")
End Function
'------------------------------------------------------
'purpose : format bank account number from xxxxxxxxxxxx to xxx-xxxxxxx-xx
'params  : mBankAccDB - bankaccount from db
'return  : bank account format for display From DB
'modified: kpo 23/03/2006 created.
'        : kpo 08/11/2007 optional
'------------------------------------------------------
Function UFormatBANKFromDB(ByVal mBankAccDB As String, Optional ByVal mEmptyZeroYN As Boolean = False) As String
    Dim mbanka As String
    mbanka = Strings.Right(String.Concat(Strings.StrDup(12, "0"), Trim(mBankAccDB)), 12)
    If mEmptyZeroYN Then
        If Val(mbanka) = 0 Then Return ""
    End If

    Return String.Concat(Mid(mbanka, 1, 3), "-", Mid(mbanka, 4, 7), "-", Mid(mbanka, 11, 2))
End Function
'------------------------------------------------------
'purpose : format NNAT number from xxxxxxxxxxxx to xxxxxx-xxx-xx
'params  : mNNATDB - NNAT from db
'return  : NNAT format for display From DB
'modified: kpo 11/12/2007 created.
'------------------------------------------------------
Function UFormatNNATFromDB(ByVal mNNATDB As String, Optional ByVal mEmptyZeroYN As Boolean = False) As String
    Dim mbanka As String
    mbanka = Strings.Right(String.Concat(Strings.StrDup(11, "0"), Trim(mNNATDB)), 11)
    If mEmptyZeroYN Then
        If Val(mbanka) = 0 Then Return ""
    End If

    Return String.Concat(Mid(mbanka, 1, 6), "-", Mid(mbanka, 7, 3), "-", Mid(mbanka, 10, 2))
End Function
'------------------------------------------------------
'purpose : format number from MF to good format
'params  : mstring = Number string from MF
'          mlength = lenght of field (include decimal)
'          mDecimal = number of decimal 
'          [original] = don't care 9999 or 99998
'return  : good number string format
'modified: kpo 11/01/2006 created
'        : kpo 01/06/2007 if all 9999 or 999998 then zero
'------------------------------------------------------
Public Function UFormatNUMFromMF(ByVal mString As String, ByVal mlength As Integer, ByVal mDecimal As Integer, Optional ByVal mOriginalValueYN As Boolean = False) As String
    Dim mreturn As String
    If Not mOriginalValueYN Then
        If mString = Strings.StrDup(mlength, "9") Then mString = "0"
        If mString = String.Concat(Strings.StrDup(mlength - 1, "9"), "8") Then mString = "0"
    End If

    mreturn = Strings.Right(String.Concat(Strings.StrDup(mlength, "0"), Trim(mString)), mlength)
    'if there is decimal.
    If mDecimal > 0 Then
       If mlength > mDecimal Then mreturn = String.Concat(Strings.Left(mString, mlength - mDecimal), ".", Strings.Right(mString, mDecimal))
    End If
    Return Val(mreturn)
End Function
'------------------------------------------------------
'purpose  : write format time from transaction to class
'params   : mtime - time value
'return   : time  date
'modified : sno 15/03/2006
'------------------------------------------------------
Function UFormatTimeFromMF(ByVal mtime As String) As Date
    If mtime <> "" Then
        Return New Date(1970, 1, 1, Strings.Left(mtime, 2), Strings.Right(mtime, 2), "00")
    Else
        Return New Date(0)
    End If
    End Function
    '------------------------------------------------------
    'purpose  : write format time from MF
    'params   : mtime - time value
    'return   : time (hh:mm)
    'modified : sno 11/05/2007
    '           kpo fixed problem.
    '------------------------------------------------------

    Function UFormatToTIMEFromMF(ByVal mTime As String) As String
        Try
            Dim ilentimest As String
            If Len(Trim(mTime)) > 4 Then
                ilentimest = "{0:000000}"
            Else
                ilentimest = "{0:0000}"
            End If

            Dim mtemp As String = String.Format(ilentimest, Val(mTime))

            If mtemp.Length >= 4 Then
                UFormatToTIMEFromMF = String.Concat(Mid(mtemp, 1, 2), ":", Mid(mtemp, 3, 2))
            Else
                UFormatToTIMEFromMF = "00:00"
            End If
        Catch ex As Exception
            UFormatToTIMEFromMF = "00:00"
        End Try
    End Function

'------------------------------------------------------
'purpose  : write format time from class to transaction
'params   : mtime - time value
'return   : time  string
'modified : sno 15/03/2006
'------------------------------------------------------
   Function UFormatTimetoMF(ByVal mDateTime As Date) As String
        Dim mreturn As String = "0000"
        If mDateTime > New Date(0) Then
            mreturn = String.Format("{0:00}{1:00}", mDateTime.Hour, mDateTime.Minute)
        End If
        Return mreturn
    End Function
'------------------------------------------------------
'purpose : format number to MF from double
'params  : mValue = Number 
'          mlength = lenght of field (include decimal)
'          mDecimal = number of decimal 
'return  : good number string format
'modified: kpo 17/04/2006 created
'          kpo 31/05/2006 fix problem about currency
'------------------------------------------------------
Public Function UFormatNUMToMF(ByVal mValue As Double, ByVal mlength As Integer, ByVal mDecimal As Integer) As String
    Dim mvaspli As String = "."
    'detect french
    If Trim(12.2).IndexOf(",") > 0 Then mvaspli = ","

    'detect sign
    Dim msignup As Byte = IIf(mValue < 0, 1, 0)
    Dim msignups As String = IIf(mValue < 0, "-", "")
    Dim mvalst As String = Trim(Math.Abs(mValue))
    Dim mvaliss() As String
    mvaliss = mvalst.Split(mvaspli)
    'get 1st frac part
    Dim mreturn As String
    mreturn = Strings.Right(String.Concat(Strings.StrDup(mlength - mDecimal - msignup, "0"), mvaliss(0)), mlength - mDecimal - msignup)
    'get 2st decimal part
    If mvaliss.Length > 1 Then
        mreturn = String.Concat(msignups, mreturn, _
            Strings.Left(String.Concat(mvaliss(1), Strings.StrDup(mDecimal, "0")), mDecimal))
    Else
        mreturn = String.Concat(msignups, mreturn, Strings.StrDup(mDecimal, "0"))
    End If
    Return mreturn
End Function
    '------------------------------------------------------
    'purpose  : write format date dd/MM/yyyy hh:mm:ss
    'params   : mdate - date value
    'return   : date string
    'modified : sno 04/01/2006
    '           kpo 10/01/2006 if date is date 0 or error return ""
    '------------------------------------------------------
    Function UFormatDate(ByVal mDate As Date, Optional ByVal mShowTimeYN As Boolean = False) As String
        UFormatDate = ""
        Try
            Dim mtimeformat As String = " {3:00}:{4:00}:{5:00}"
            If Not mShowTimeYN Then mtimeformat = ""
            If mDate > New Date(0) Then
                UFormatDate = String.Format(String.Concat("{0:00}/{1:00}/{2:0000}", mtimeformat), mDate.Day, mDate.Month, mDate.Year, mDate.Hour, mDate.Minute, mDate.Second)
            End If
        Catch ex As Exception
        End Try
    End Function

'------------------------------------------------------
'purpose  : convert date from MF (YYYYMMDD) to date
'return   : date
'modified : KPO 10/01/2006
'           KPO 01/03/2006
'------------------------------------------------------
Function UFormatDateFromMF(ByVal mDate As String) As Date
    Try
        UFormatDateFromMF = New Date(0)
        Dim mtemp As String = mDate
        If mtemp.Length = 8 Then UFormatDateFromMF = New Date(Mid(mtemp, 1, 4), Mid(mtemp, 5, 2), Mid(mtemp, 7, 2))
        Catch ex As Exception
        UFormatDateFromMF = New Date("00000000")
    End Try
End Function

'------------------------------------------------------
'purpose  : write format date for mainframe (ddmmyyyy) from format dd/mm/yyyy
'return   : ddmmyyyy string.
'modified : KPO 06/01/2006
'------------------------------------------------------
Function UFormatDatetoMF(ByVal mDate As String) As String
    Dim mreturn As String = "00000000"
        If Len(mDate) = 10 Then mreturn = String.Format("{2:0000}{1:00}{0:00}", Mid(mDate, 1, 2), Mid(mDate, 4, 2), Mid(mDate, 7, 4))
    Return mreturn
End Function

Function UFormatDatetoMF(ByVal mDate As Date) As String
    Return UFormatDatetoMF(UFormatDate(mDate))
End Function


    '------------------------------------------------------
    'purpose  : convert string to  format date
    'params   : mDate - Date string with format dd/MM/yyyy
    '           mtime - time string format HH:mm or HH:mm:ss
    'return   : date value
    'modified : kpo 11/01/2006 create
    '------------------------------------------------------
    Function UFormat2Date(ByVal mDate As String, Optional ByVal mTime As String = "") As Date
        Dim mtemp As String = mDate
        Dim mreturndate As Date
        Dim mhour, mmin, msec As Integer

        Try
            If mTime.Length >= 5 Then
                mhour = Mid(mTime, 1, 2)
                mmin = Mid(mTime, 4, 2)
            End If
            If mTime.Length = 8 Then msec = Mid(mTime, 7, 2)
            If mtemp.Length = 10 Then mreturndate = New Date(Mid(mtemp, 7, 4), Mid(mtemp, 4, 2), _
                    Mid(mtemp, 1, 2), mhour, mmin, msec)

        Catch ex As Exception
            mreturndate = New Date(0)
        End Try
        Return mreturndate
    End Function

    '--------------------------------------------------------
    'remove TAG <gen:ang></gen:ang> from input string
    'paramin : _InputString - input string.
    'return  : string without TAG <gen:ang></gen:ang>
    'TKO            created.
    'kpo 11/08/2006 opimized and verify
    '--------------------------------------------------------
    Function Func_RemoveTagGENANG(ByVal _mInputString As String) As String
        Dim stemp As String
        Dim sleft As String
        Dim sright As String
        stemp = _mInputString

        If Not stemp.Replace("&lt;gen:ang&gt;", "").Length = stemp.Length Then
            sleft = Mid(stemp, 1, stemp.IndexOf("&lt;gen:ang&gt;"))
            sright = Mid(stemp, stemp.IndexOf("&lt;/gen:ang&gt;") + ("&lt;/gen:ang&gt;".Length + 1), stemp.Length)
            Return String.Concat(sleft, sright)
        ElseIf Not stemp.Replace("<gen:ang>", "").Length = stemp.Length Then
            sleft = Mid(stemp, 1, stemp.IndexOf("<gen:ang>"))
            sright = Mid(stemp, stemp.IndexOf("</gen:ang>") + "</gen:ang>".Length + 1, stemp.Length)
            Return String.Concat(sleft, sright)
        Else
            Return stemp
        End If

    End Function

   'Get file by using FileOpen
    Public Function UGetFile(ByVal File_Path As String) As String

        Dim s_build As New System.Text.StringBuilder
        Dim mfile As IO.TextReader
        Dim l_len As Long
        l_len = FileLen(File_Path)

        Dim mChar(1023) As Char
        Dim i_read As Integer
        mfile = IO.File.OpenText(File_Path)

        i_read = mfile.Read(mChar, 0, 1024)
        Do While (i_read > 0)
            If (s_build.Length + 1024) > l_len Then
                s_build.Append(mChar, 0, CType(l_len - s_build.Length, Integer))
            Else
                s_build.Append(mChar, 0, i_read)
            End If
            i_read = mfile.Read(mChar, 0, 1024)
        Loop

        mfile.Close()
        Erase mChar

        UGetFile = s_build.ToString

        mChar = Nothing
        mfile = Nothing
        s_build = Nothing
    End Function

    '------------------------------------------------------------
    'TKO
    'Use this function to replace singlecode(')  by doublecode('').
    '------------------------------------------------------------
    'Function ReplaceString(ByVal s_val As String) As String
    '    Return Trim(Replace(s_val, "'", "''"))
    'End Function

    '==================================================================
    'Purpose    : checking the previous instance of Ansgate
    'Parameter  : none
    'Return     : Ansgate already run ?
    'Author     : KPO 20/11/2002 
    'Modified   : KPO 20/05/2003 only check previos when ISPREV =true
    '==================================================================
    Function PrevInstance() As Boolean
        Dim mApp As String
        mApp = IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly.Location.ToString)
        Dim mTest As Threading.Mutex = New Threading.Mutex(True, mApp)
        If Not mTest.WaitOne(1, False) Then
            'Application is already running
            PrevInstance = True
        Else
            PrevInstance = False
        End If
        If Not mTest Is Nothing Then mTest.Close()
        mTest = Nothing
    End Function

    '------------------------------------------------------
    'purpose  : validate the currency settins.
    'params   : none
    'return   : true - if not use . as decimal separator. 
    'kpo 11/08/2006 created
    '------------------------------------------------------
    Function UISFrenchCur() As Boolean
        Return (Trim(CDbl("12.2")) <> "12.2")
    End Function

     '------------------------------------------------------
    'purpose  : get version number
    'params   : none
    'return   : Version of current application
    'kpo 11/08/2006 created
    '------------------------------------------------------
    '
    Public Function UGetVersionInfo() As String
        Try
            Return Trim(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart) + "." + Trim(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart) + "." + Trim(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileVersion)
        Catch ex As Exception
            Return "No Version"
        End Try
    End Function

    '------------------------------------------------------
    'purpose : remove enter change to \n
    'params  : mText - text to show in java
    'return  : text properly display in java
    'modified: kpo 27/04/2006 create
    '          kpo 23/05/2006 trim
    '          kpo 22/06/2007 check when only cr or lf is detect not put 2 times /n
    '------------------------------------------------------
    Public Function UDoStringJava(ByVal mText As String) As String
        Dim mtemp As String = mText
        mtemp = Replace(mtemp, "\", "\\")
        mtemp = Replace(mtemp, vbCrLf, "\n")

        Dim balreadycrYN As Boolean
        Dim mtemp2 As String = mtemp
        mtemp = Replace(mtemp, vbCr, "\n")
        If mtemp2 <> mtemp Then balreadycrYN = True

        mtemp = Replace(mtemp, vbLf, IIf(balreadycrYN, "", "\n"))
        mtemp = Replace(mtemp, "'", "\'")
        'mtemp = Replace(mtemp, "&", "&amp;")
        Return Trim(mtemp)
    End Function
    'this function for show text in textbox
    Function UMakeStringForJava(ByVal mString As String) As String
        Dim ret As String
        ret = Trim(mString)
        ret = Replace(ret, "\", "\\")
        ret = Replace(ret, "'", "\'")
        ret = Replace(ret, """", "\""")

        Dim balreadycrYN As Boolean
        Dim mtemp2 As String = ret
        ret = Replace(ret, vbCr, "<br>")
        If mtemp2 <> ret Then balreadycrYN = True

        ret = Replace(ret, vbLf, IIf(balreadycrYN, "", "<br>"))
        Return ret
    End Function

    Public Function WriteFile(ByVal mFilename As String, ByVal mData As Object, Optional ByRef mErrorstring As String = "") As Boolean
            Dim io_file As IO.FileStream
            Try
                io_file = New IO.FileStream(mFilename, IO.FileMode.OpenOrCreate)
                io_file.Write(mData, 0, mData.length)
                io_file.Close()
                WriteFile = True
            Catch ex As Exception
                mErrorstring = String.Concat("Error Writing file : ", ex.Message)
            End Try
            io_file = Nothing
    End Function

    Public Function UArrayToString(ByVal mArray As Object, Optional ByRef mErrorstring As String = "") As String
        Dim ArrayLength As Long
        ArrayLength = (4 / 3) * mArray.Length
        If ArrayLength Mod 4 <> 0 Then
            ArrayLength = ArrayLength + 4 - ArrayLength Mod 4
        End If

        Dim base64CharArray(ArrayLength - 1) As Char
        Try
            System.Convert.ToBase64CharArray(mArray, _
                                           0, _
                                           mArray.Length, _
                                           base64CharArray, 0)
        Catch ex As Exception
            mErrorstring = String.Concat("Error in UArrayToString : ", ex.Message)
        End Try

        UArrayToString = New String(base64CharArray)
        Erase base64CharArray
    End Function

    Public Function UStringToArray(ByVal mString As String, Optional ByRef mErrorstring As String = "") As Object
        Try
            UStringToArray = System.Convert.FromBase64String(mString)
        Catch ex As Exception
            mErrorstring = String.Concat("Error in UStringToArray : ", ex.Message)
            UStringToArray = Nothing
        End Try
    End Function
 'convert from IP string to IP value
    Public Function CoverttoIPValue(ByVal mIP As String) As Long
        Try
            Dim mvIP As String = Trim(mIP)
            If mvIP = "" Then Exit Function
            If mvIP.IndexOf(".") <= 0 Then Exit Function

            Dim bnotok As Boolean
            Dim digit() As String
            digit = Split(mvIP, ".")
            If UBound(digit) <> 3 Then bnotok = True

            Dim i As Integer
            If Not bnotok Then
                For i = 0 To 3
                    If Val(digit(i)) > 255 Then
                        bnotok = True
                        Exit For
                    End If
                Next i
            End If
            If Not bnotok Then
                If Val(digit(0)) = 0 And Val(digit(1)) = 0 And Val(digit(2)) = 0 And Val(digit(3)) = 0 Then bnotok = True
            End If

            If Not bnotok Then CoverttoIPValue = (Val(digit(0)) * 256 ^ 3) + (Val(digit(1)) * 256 ^ 2) + (Val(digit(2)) * 256) + Val(digit(3))

            Erase digit
            digit = Nothing
        Catch ex As Exception
            Return 0
        End Try
        
    End Function
    'convert IP value to IP String
    Public Function ConverttoIPST(ByVal rxValue As Long) As String
        Dim rxTemp As Long
        Dim rxIP(3) As String

        rxIP(3) = rxValue Mod 256

        rxTemp = ((rxValue - Val(rxIP(3))) Mod (256 ^ 2))
        rxIP(2) = rxTemp / 256

        rxTemp = (rxValue - Val(rxIP(2)) * 256 - Val(rxIP(3))) Mod (256 ^ 3)
        rxIP(1) = rxTemp / 256 ^ 2

        rxTemp = rxValue - Val(rxIP(1)) * 256 ^ 2 - Val(rxIP(2)) * 256 - Val(rxIP(3))
        rxIP(0) = rxTemp / 256 ^ 3

        'Txt_IPAddress.Text = rxIP(0) & "." & rxIP(1) & "." & rxIP(2) & "." & rxIP(3)
        ConverttoIPST = String.Concat(rxIP(0), ".", rxIP(1), ".", rxIP(2), ".", rxIP(3))

        Erase rxIP
        rxIP = Nothing

    End Function

    '------------------------------------------------------
    'display 99/999.99.99 if first 2 digits are 02 or 03, else display 999/99.99.99
    'purpose : convert TeLSUB data from MF to display format
    'params  : mTELSUB - TELSUB from MF
    'return  : TELSUB in display format
    'modified: kpo 06/01/2006 created.
    '          if all zero then show blank
    '          kpo 23/05/2006 fixed for 2nd display 
    '          kpo 01/06/2006 fixed
    'Only GSM numbers (= 10 digits) must be displayed as 9999/99.99.99

    'Normal phone numbers (= 9 digits) have to be displayed as 99/999.99.99 (02/... -> 09/...) or 999/99.99.99 (010/.. -> ...)
    '------------------------------------------------------
    Function UFormatTELSUB(ByVal mTELSUB As String, Optional ByVal mGSMYN As Boolean = False) As String
        Dim mtemp As String = Replace(Replace(mTELSUB, "/", ""), ".", "")
        mtemp = String.Format("{0:0000000000}", Val(mtemp))
        If mGSMYN Then
        Else
            mtemp = Strings.Right(mtemp, 9)
        End If
        Dim mtemp2 As String = UFormat2TELSUB(mtemp)
        If mtemp2 = "" Then
            UFormatTELSUB = ""
        ElseIf UHideWhenAllZero(mtemp2) = "" Then
            UFormatTELSUB = " "
        Else
            If mGSMYN Then
                UFormatTELSUB = String.Concat(Strings.Left(mtemp, 4), "/", Mid(mtemp, 5, 2), ".", Mid(mtemp, 7, 2), ".", Mid(mtemp, 9, 2))
            Else
                If (Strings.Left(Trim(mtemp), 2) = "02") Or (Strings.Left(Trim(mtemp), 2) = "03") Or (Strings.Left(Trim(mtemp), 2) = "04") Or (Strings.Left(Trim(mtemp), 2) = "09") Then
                    UFormatTELSUB = String.Concat(Strings.Left(Trim(mtemp), 2), "/", Mid(mtemp2, 3, 3), ".", Mid(mtemp2, 6, 2), ".", Mid(mtemp2, 8, 2))
                Else
                    UFormatTELSUB = String.Concat(Strings.Left(mtemp, 3), "/", Mid(mtemp, 4, 2), ".", Mid(mtemp, 6, 2), ".", Mid(mtemp, 8, 2))
                End If
            End If
        End If
    End Function

    '------------------------------------------------------
    'display 99/999.99.99 if first 2 digits are 02 or 03, else display 999/99.99.99
    'purpose : convert TeLSUB data to MF from display format
    'params  : mTELSUB - TELSUB from display
    'return  : TELSUB to MF format
    'modified: kpo 05/04/2006 created.
    '------------------------------------------------------
    Function UFormat2TELSUB(ByVal mTELSUB As String) As String
        Dim mtempsub As String
        mtempsub = Replace(Replace(Trim(mTELSUB), "/", ""), ".", "")
        If mtempsub = "" Then mtempsub = Strings.StrDup(9, "0")
        Return mtempsub
    End Function

    '------------------------------------------------------
    'purpose : if all data = 0 then show space
    'params  : mData - mData wantto display
    'return  : good format display
    'modified: kpo 03/05/2006 created.
    '------------------------------------------------------
    Function UHideWhenAllZero(ByVal mData As String) As String
        If Trim(Replace(mData, "0", "")) = "" Then
            Return ""
        Else
            Return mData
        End If
    End Function
'------------------------------------------------------
'purpose : encode string to 64basestring
'params  : mstring = string to convert
'return  : 64base string
'modified: kpo 10/01/2006 created
'------------------------------------------------------
Function UEncodeString(ByVal mString As String) As String
    UEncodeString = ""
    If mString = "" Then Exit Function
    Return System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(mString))
End Function

'------------------------------------------------------
'purpose : decode string from 64basestring
'params  : mstring = 64basestring want to convert back
'return  : normal string
'modified: kpo 10/01/2006 created
'------------------------------------------------------
Function UDecodeString(ByVal mString As String) As String
    Try
        Return System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(mString))
    Catch ex As Exception
        Return ""
    End Try
    End Function


    Function UConvertDecToHex(ByVal Dec As Double) As String
        Return Hex(Dec)
    End Function
    
    Function ReplaceString(ByVal text As String, ByVal OldChar As String, ByVal NewChar As String) As String
        Dim Strreturn As String = text
        Try
            Dim intIndexOfMatch As Integer = Strreturn.IndexOf(OldChar)
            While intIndexOfMatch > 0
                Strreturn = Strreturn.Replace(OldChar, NewChar)
                intIndexOfMatch = Strreturn.IndexOf(OldChar)

            End While
        Catch generatedExceptionName As Exception
        End Try
        Return Strreturn
    End Function
    ''' <summary>
    ''' DateString format yyyymmdd
    ''' return dd/mm/yyyy
    ''' </summary>
    ''' <param name="DateString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function UConvertString2Date(ByVal DateString As String) As Date
        Dim mdate As Date
        Try
            Dim format As System.IFormatProvider = New System.Globalization.CultureInfo("en-US")

            If DateString.Length > 0 Then
                Dim Newdate As String = String.Concat(DateString.Substring(6, 2), "/", DateString.Substring(4, 2), "/", DateString.Substring(0, 4))

                mdate = DateTime.Parse(Newdate, format)
            End If
        Catch ex As Exception
            Return Now
        End Try
        Return mdate
    End Function
    Function UConvertString2Date1(ByVal DateString As String) As Date
        Dim mdate As Date
        Try
            Dim format As System.IFormatProvider = New System.Globalization.CultureInfo("th-TH")

            If DateString.Length > 0 Then
                Dim m1 As Integer = CInt(DateString.Substring(0, 4)) + 543

                Dim Newdate As String = String.Concat(DateString.Substring(6, 2), "/", DateString.Substring(4, 2), "/", m1)

                mdate = DateTime.Parse(Newdate, format)
            End If
        Catch ex As Exception
            Return Now
        End Try
        Return mdate
    End Function



    Function UConvertString2Date2(ByVal DateString As String) As String
        Dim _mdate As String = ""
        Try
            '   Dim format As System.IFormatProvider = New System.Globalization.CultureInfo("th-TH")

            If DateString.Length > 0 Then
                Dim m1() As String
                m1 = DateString.Split("/")
                _mdate = String.Concat(m1(1), "/", m1(0), "/", m1(2))
                ' Dim Newdate As String = DateString

                ' mdate = DateTime.Parse(Newdate, format)
            End If
        Catch ex As Exception
            Return DateString
        End Try
        Return _mdate
    End Function

    Public Sub SetLangUS()
        If System.Threading.Thread.CurrentThread.CurrentCulture.ToString() <> "en-US" Then
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        End If

    End Sub
End Module
