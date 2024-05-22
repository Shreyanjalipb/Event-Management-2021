Public Class ContactClass
    'Contact_RecNo
    'Company_RecNo,Contact_Address_1,Contact_Post_Code
    'Contact_State,Contact_Address_2,Contact_City,Contact_Country
    'Home_Address1,Home_Address2,Home_Suburb,Home_State
    'Home_Postcode,Home_Country
    'Contact_Number,Position,Given_Name,Surname,
    'Phone_Bus,Phone_Direct,Phone_Pri,Phone_Mobile
    'Fax,EMail,Other,Department,Debtor_Number,Charge_Code
    'Web_Page,Mailing_Address,Contact_Colour,Sequence,Created
    'Special_Notes,Title,Inactive_Contact_Date,LastEdited
    'Prefix,Created_by,Last_Edited_by,PrivateContact
    'Mail_Preference_Date,Mail_Preference_Subscribed
    'Contact_Address_3,Home_Address_3,Email2
    Private mContact_RecNo As Long
    Property Contact_RecNo() As Long
        Get
            Return mContact_RecNo
        End Get
        Set(ByVal value As Long)
            mContact_RecNo = value
        End Set
    End Property
    Private mCompany_RecNo As Long
    Property Company_RecNo() As Long
        Get
            Return mCompany_RecNo
        End Get
        Set(ByVal value As Long)
            mCompany_RecNo = value
        End Set
    End Property
    Private mContact_Address_1 As String = ""
    Property Contact_Address_1() As String
        Get
            Return mContact_Address_1
        End Get
        Set(ByVal value As String)
            mContact_Address_1 = value
        End Set
    End Property

    Private mContact_Post_Code As String = ""
    Property Contact_Post_Code() As String
        Get
            Return mContact_Post_Code
        End Get
        Set(ByVal value As String)
            mContact_Post_Code = value
        End Set
    End Property

    Private mContact_State As String = ""
    Property Contact_State() As String
        Get
            Return mContact_State
        End Get
        Set(ByVal value As String)
            mContact_State = value
        End Set
    End Property
    Private mContact_Address_2 As String = ""
    Property Contact_Address_2() As String
        Get
            Return mContact_Address_2
        End Get
        Set(ByVal value As String)
            mContact_Address_2 = value
        End Set
    End Property
    Private mContact_City As String = ""
    Property Contact_City() As String
        Get
            Return mContact_City
        End Get
        Set(ByVal value As String)
            mContact_City = value
        End Set
    End Property
    Private mContact_Country As String = ""
    Property Contact_Country() As String
        Get
            Return mContact_Country
        End Get
        Set(ByVal value As String)
            mContact_Country = value
        End Set
    End Property
    Private mHome_Address1 As String = ""
    Property Home_Address1() As String
        Get
            Return mHome_Address1
        End Get
        Set(ByVal value As String)
            mHome_Address1 = value
        End Set
    End Property
    Private mHome_Address2 As String = ""
    Property Home_Address2() As String
        Get
            Return mHome_Address2
        End Get
        Set(ByVal value As String)
            mHome_Address2 = value
        End Set
    End Property
    Private mHome_Suburb As String = ""
    Property Home_Suburb() As String
        Get
            Return mHome_Suburb
        End Get
        Set(ByVal value As String)
            mHome_Suburb = value
        End Set
    End Property
    Private mHome_State As String = ""
    Property Home_State() As String
        Get
            Return mHome_State
        End Get
        Set(ByVal value As String)
            mHome_State = value
        End Set
    End Property
    Private mHome_Postcode As String = ""
    Property Home_Postcode() As String
        Get
            Return mHome_Postcode
        End Get
        Set(ByVal value As String)
            mHome_Postcode = value
        End Set
    End Property
    Private mHome_Country As String = ""
    Property Home_Country() As String
        Get
            Return mHome_Country
        End Get
        Set(ByVal value As String)
            mHome_Country = value
        End Set
    End Property
    Private mContact_Number As String = ""
    Property Contact_Number() As String
        Get
            Return mContact_Number
        End Get
        Set(ByVal value As String)
            mContact_Number = value
        End Set
    End Property
    Private mPosition As String = ""
    Property Position() As String
        Get
            Return mPosition
        End Get
        Set(ByVal value As String)
            mPosition = value
        End Set
    End Property
    Private mGiven_Name As String = ""
    Property Given_Name() As String
        Get
            Return mGiven_Name
        End Get
        Set(ByVal value As String)
            mGiven_Name = value
        End Set
    End Property
    Private mSurname As String = ""
    Property Surname() As String
        Get
            Return mSurname
        End Get
        Set(ByVal value As String)
            mSurname = value
        End Set
    End Property
    Private mPhone_Bus As String = ""
    Property Phone_Bus() As String
        Get
            Return mPhone_Bus
        End Get
        Set(ByVal value As String)
            mPhone_Bus = value
        End Set
    End Property
    Private mPhone_Direct As String = ""
    Property Phone_Direct() As String
        Get
            Return mPhone_Direct
        End Get
        Set(ByVal value As String)
            mPhone_Direct = value
        End Set
    End Property
    Private mPhone_Pri As String = ""
    Property Phone_Pri() As String
        Get
            Return mPhone_Pri
        End Get
        Set(ByVal value As String)
            mPhone_Pri = value
        End Set
    End Property
    Private mPhone_Mobile As String = ""
    Property Phone_Mobile() As String
        Get
            Return mPhone_Mobile
        End Get
        Set(ByVal value As String)
            mPhone_Mobile = value
        End Set
    End Property
    Private mFax As String = ""
    Property Fax() As String
        Get
            Return mFax
        End Get
        Set(ByVal value As String)
            mFax = value
        End Set
    End Property
    Private mEMail As String = ""
    Property EMail() As String
        Get
            Return mEMail
        End Get
        Set(ByVal value As String)
            mEMail = value
        End Set
    End Property
    Private mOther As String = ""
    Property Other() As String
        Get
            Return mOther
        End Get
        Set(ByVal value As String)
            mOther = value
        End Set
    End Property
    Private mDepartment As String = ""
    Property Department() As String
        Get
            Return mDepartment
        End Get
        Set(ByVal value As String)
            mDepartment = value
        End Set
    End Property
    Private mDebtor_Number As String = ""
    Property Debtor_Number() As String
        Get
            Return mDebtor_Number
        End Get
        Set(ByVal value As String)
            mDebtor_Number = value
        End Set
    End Property
    Private mCharge_Code As String = ""
    Property Charge_Code() As String
        Get
            Return mCharge_Code
        End Get
        Set(ByVal value As String)
            mCharge_Code = value
        End Set
    End Property
    Private mWeb_Page As String = ""
    Property Web_Page() As String
        Get
            Return mWeb_Page
        End Get
        Set(ByVal value As String)
            mWeb_Page = value
        End Set
    End Property
    Private mMailing_Address As Long
    Property Mailing_Address() As Long
        Get
            Return mMailing_Address
        End Get
        Set(ByVal value As Long)
            mMailing_Address = value
        End Set
    End Property
    Private mContact_Colour As Long
    Property Contact_Colour() As Long
        Get
            Return mContact_Colour
        End Get
        Set(ByVal value As Long)
            mContact_Colour = value
        End Set
    End Property
    Private mSequence As Long
    Property Sequence() As Long
        Get
            Return mSequence
        End Get
        Set(ByVal value As Long)
            mSequence = value
        End Set
    End Property
    Private mCreated As Date
    Property Created() As Date
        Get
            Return mCreated
        End Get
        Set(ByVal value As Date)
            mCreated = value
        End Set
    End Property
    Private mSpecial_Notes As String = ""
    Property Special_Notes() As String
        Get
            Return mSpecial_Notes
        End Get
        Set(ByVal value As String)
            mSpecial_Notes = value
        End Set
    End Property
    Private mTitle As String = ""
    Property Title() As String
        Get
            Return mTitle
        End Get
        Set(ByVal value As String)
            mTitle = value
        End Set
    End Property
    Private mInactive_Contact_Date As Date
    Property Inactive_Contact_Date() As Date
        Get
            Return mInactive_Contact_Date
        End Get
        Set(ByVal value As Date)
            mInactive_Contact_Date = value
        End Set
    End Property
    Private mLastEdited As Date
    Property LastEdited() As Date
        Get
            Return mLastEdited
        End Get
        Set(ByVal value As Date)
            mLastEdited = value
        End Set
    End Property

    Private mPrefix As String = ""
    Property Prefix() As String
        Get
            Return mPrefix
        End Get
        Set(ByVal value As String)
            mPrefix = value
        End Set
    End Property
    Private mCreated_by As String = ""
    Property Created_by() As String
        Get
            Return mCreated_by
        End Get
        Set(ByVal value As String)
            mCreated_by = value
        End Set
    End Property
    Private mLast_Edited_by As String = ""
    Property Last_Edited_by() As String
        Get
            Return mLast_Edited_by
        End Get
        Set(ByVal value As String)
            mLast_Edited_by = value
        End Set
    End Property
    Private mPrivateContact As Boolean
    Property PrivateContact() As Boolean
        Get
            Return mPrivateContact
        End Get
        Set(ByVal value As Boolean)
            mPrivateContact = value
        End Set
    End Property
    Private mMail_Preference_Date As Date
    Property Mail_Preference_Date() As Date
        Get
            Return mMail_Preference_Date
        End Get
        Set(ByVal value As Date)
            mMail_Preference_Date = value
        End Set
    End Property
    Private mMail_Preference_Subscribed As Boolean
    Property Mail_Preference_Subscribed() As Boolean
        Get
            Return mMail_Preference_Subscribed
        End Get
        Set(ByVal value As Boolean)
            mMail_Preference_Subscribed = value
        End Set
    End Property
    Private mContact_Address_3 As String = ""
    Property Contact_Address_3() As String
        Get
            Return mContact_Address_3
        End Get
        Set(ByVal value As String)
            mContact_Address_3 = value
        End Set
    End Property

    Private mHome_Address_3 As String = ""
    Property Home_Address_3() As String
        Get
            Return mHome_Address_3
        End Get
        Set(ByVal value As String)
            mHome_Address_3 = value
        End Set
    End Property
    Private mEmail2 As String = ""
    Property Email2() As String
        Get
            Return mEmail2
        End Get
        Set(ByVal value As String)
            mEmail2 = value
        End Set
    End Property

    Function Save(Optional ByVal mGUID As String = "")
        Dim ReMsg As Boolean
        Try
            Dim param_Contact_RecNo As New SqlClient.SqlParameter("@Contact_RecNo", SqlDbType.BigInt, 4)
            param_Contact_RecNo.Value = Contact_RecNo
            Dim param_Company_RecNo As New SqlClient.SqlParameter("@Company_RecNo", SqlDbType.BigInt, 4)
            param_Company_RecNo.Value = Company_RecNo
            Dim param_Contact_Address_1 As New SqlClient.SqlParameter("@Contact_Address_1", SqlDbType.NVarChar, 50)
            param_Contact_Address_1.Value = Contact_Address_1
            Dim param_Contact_Post_Code As New SqlClient.SqlParameter("@Contact_Post_Code", SqlDbType.NVarChar, 20)
            param_Contact_Post_Code.Value = Contact_Post_Code
            Dim param_Contact_State As New SqlClient.SqlParameter("@Contact_State", SqlDbType.NVarChar, 20)
            param_Contact_State.Value = Contact_State
            Dim param_Contact_Address_2 As New SqlClient.SqlParameter("@Contact_Address_2", SqlDbType.NVarChar, 50)
            param_Contact_Address_2.Value = Contact_City
            Dim param_Contact_City As New SqlClient.SqlParameter("@Contact_City", SqlDbType.NVarChar, 50)
            param_Contact_City.Value = Contact_City

            Dim param_Contact_Country As New SqlClient.SqlParameter("@Contact_Country", SqlDbType.NVarChar, 30)
            param_Contact_Country.Value = Contact_Country

            Dim param_Home_Address1 As New SqlClient.SqlParameter("@Home_Address1", SqlDbType.NVarChar, 50)
            param_Home_Address1.Value = Home_Address1

            Dim param_Home_Address2 As New SqlClient.SqlParameter("@Home_Address2", SqlDbType.NVarChar, 50)
            param_Home_Address2.Value = Home_Address2

            Dim param_Home_Suburb As New SqlClient.SqlParameter("@Home_Suburb", SqlDbType.NVarChar, 50)
            param_Home_Suburb.Value = Home_Suburb

            Dim param_Home_State As New SqlClient.SqlParameter("@Home_State", SqlDbType.NVarChar, 20)
            param_Home_State.Value = Home_State
            
            Dim param_Home_Postcode As New SqlClient.SqlParameter("@Home_Postcode", SqlDbType.NVarChar, 20)
            param_Home_Postcode.Value = Home_Postcode

            Dim param_Home_Country As New SqlClient.SqlParameter("@Home_Country", SqlDbType.NVarChar, 30)
            param_Home_Country.Value = Home_Country

            Dim param_Contact_Number As New SqlClient.SqlParameter("@Contact_Number", SqlDbType.NVarChar, 10)
            param_Contact_Number.Value = Contact_Number

            Dim param_Position As New SqlClient.SqlParameter("@Position", SqlDbType.NVarChar, 50)
            param_Position.Value = Position

            Dim param_Given_Name As New SqlClient.SqlParameter("@Given_Name", SqlDbType.NVarChar, 50)
            param_Given_Name.Value = Given_Name

            Dim param_Surname As New SqlClient.SqlParameter("@Surname", SqlDbType.NVarChar, 50)
            param_Surname.Value = Surname

            Dim param_Phone_Bus As New SqlClient.SqlParameter("@Phone_Bus", SqlDbType.NVarChar, 20)
            param_Phone_Bus.Value = Phone_Bus

            Dim param_Phone_Direct As New SqlClient.SqlParameter("@Phone_Direct", SqlDbType.NVarChar, 20)
            param_Phone_Direct.Value = Phone_Direct

            Dim param_Phone_Pri As New SqlClient.SqlParameter("@Phone_Pri", SqlDbType.NVarChar, 20)
            param_Phone_Pri.Value = Phone_Pri

            Dim param_Phone_Mobile As New SqlClient.SqlParameter("@Phone_Mobile", SqlDbType.NVarChar, 20)
            param_Phone_Mobile.Value = Phone_Mobile

            Dim param_Fax As New SqlClient.SqlParameter("@Fax", SqlDbType.NVarChar, 20)
            param_Fax.Value = Fax

            Dim param_EMail As New SqlClient.SqlParameter("@EMail", SqlDbType.NVarChar, 100)
            param_EMail.Value = EMail

            Dim param_Other As New SqlClient.SqlParameter("@Other", SqlDbType.NVarChar, 100)
            param_Other.Value = Other

            Dim param_Department As New SqlClient.SqlParameter("@Department", SqlDbType.NVarChar, 50)
            param_Department.Value = Department

            Dim param_Debtor_Number As New SqlClient.SqlParameter("@Debtor_Number", SqlDbType.NVarChar, 20)
            param_Debtor_Number.Value = Debtor_Number

            Dim param_Charge_Code As New SqlClient.SqlParameter("@Charge_Code", SqlDbType.NVarChar, 20)
            param_Charge_Code.Value = Charge_Code

            Dim param_Web_Page As New SqlClient.SqlParameter("@Web_Page", SqlDbType.NVarChar, 255)
            param_Web_Page.Value = Web_Page

            Dim param_Mailing_Address As New SqlClient.SqlParameter("@Mailing_Address", SqlDbType.BigInt, 4)
            param_Mailing_Address.Value = Mailing_Address

            Dim param_Contact_Colour As New SqlClient.SqlParameter("@Contact_Colour", SqlDbType.BigInt, 4)
            param_Contact_Colour.Value = Contact_Colour

            Dim param_Sequence As New SqlClient.SqlParameter("@Sequence", SqlDbType.BigInt, 4)
            param_Sequence.Value = Sequence

            Dim param_Created As New SqlClient.SqlParameter("@Created", SqlDbType.DateTime)
            param_Created.Value = Created

            Dim param_Special_Notes As New SqlClient.SqlParameter("@Special_Notes", SqlDbType.NVarChar, 16)
            param_Special_Notes.Value = Special_Notes

            Dim param_Title As New SqlClient.SqlParameter("@Title", SqlDbType.NVarChar, 65)
            param_Title.Value = Title
            '''''

            '    Dim param_Inactive_Contact_Date As New SqlClient.SqlParameter("@Inactive_Contact_Date", SqlDbType.DateTime)
            '   param_Inactive_Contact_Date.Value = Inactive_Contact_Date

            Dim param_LastEdited As New SqlClient.SqlParameter("@LastEdited", SqlDbType.DateTime)
            param_LastEdited.Value = LastEdited

            Dim param_Prefix As New SqlClient.SqlParameter("@Prefix", SqlDbType.NVarChar, 5)
            param_Prefix.Value = Prefix

            Dim param_Created_by As New SqlClient.SqlParameter("@Created_by", SqlDbType.NVarChar, 255)
            param_Created_by.Value = Created_by

            Dim param_Last_Edited_by As New SqlClient.SqlParameter("@Last_Edited_by", SqlDbType.NVarChar, 255)
            param_Last_Edited_by.Value = Last_Edited_by

            Dim param_PrivateContact As New SqlClient.SqlParameter("@PrivateContact", SqlDbType.Bit, 1)
            param_PrivateContact.Value = PrivateContact

            Dim param_Mail_Preference_Date As New SqlClient.SqlParameter("@Mail_Preference_Date", SqlDbType.DateTime)
            param_Mail_Preference_Date.Value = Mail_Preference_Date

            Dim param_Mail_Preference_Subscribed As New SqlClient.SqlParameter("@Mail_Preference_Subscribed", SqlDbType.Bit, 1)
            param_Mail_Preference_Subscribed.Value = Mail_Preference_Subscribed

            Dim param_Contact_Address_3 As New SqlClient.SqlParameter("@Contact_Address_3", SqlDbType.NVarChar, 50)
            param_Contact_Address_3.Value = Contact_Address_3

            Dim param_Home_Address_3 As New SqlClient.SqlParameter("@Home_Address_3", SqlDbType.NVarChar, 50)
            param_Home_Address_3.Value = Home_Address_3

            Dim param_Email2 As New SqlClient.SqlParameter("@Email2", SqlDbType.NVarChar, 100)
            param_Email2.Value = Email2
           

            Dim mSql As String
            mSql = String.Concat("select [Contact_RecNo] from [CONTACT] where [Contact_RecNo] = @Contact_RecNo ", _
                    "if @@Rowcount = 0 ", _
                        "insert into [CONTACT] (Contact_RecNo,", _
                       "Company_RecNo,Contact_Address_1,Contact_Post_Code,Contact_State,Contact_Address_2,Contact_City,", _
                       "Contact_Country, Home_Address1, Home_Address2, Home_Suburb, Home_State, Home_Postcode, ", _
                       "Home_Country, Contact_Number, Position, Given_Name, Surname, ", _
                       "Phone_Bus, Phone_Direct, Phone_Pri, Phone_Mobile, Fax, EMail, Other, Department, Debtor_Number, ", _
                       "Charge_Code, Web_Page, Mailing_Address, Contact_Colour,", _
                       " Sequence, Created, Special_Notes, Title, ", _
                        " LastEdited, Prefix, Created_by, Last_Edited_by, ", _
                        "PrivateContact, Mail_Preference_Date, Mail_Preference_Subscribed, ", _
                         "Contact_Address_3, Home_Address_3, Email2", _
                       ")", _
                        " values (@Contact_RecNo,", _
                         "@Company_RecNo,@Contact_Address_1,@Contact_Post_Code,@Contact_State,@Contact_Address_2,@Contact_City,", _
                         "@Contact_Country, @Home_Address1, @Home_Address2, @Home_Suburb, @Home_State, @Home_Postcode,", _
                         "@Home_Country, @Contact_Number, @Position, @Given_Name, @Surname, ", _
                         "@Phone_Bus, @Phone_Direct, @Phone_Pri, @Phone_Mobile, @Fax, @EMail, @Other, @Department, @Debtor_Number,", _
                         " @Charge_Code, @Web_Page, @Mailing_Address, @Contact_Colour,", _
                         "@Sequence, @Created, @Special_Notes, @Title, ", _
                        " @LastEdited, @Prefix, @Created_by, @Last_Edited_by, ", _
                        "@PrivateContact, @Mail_Preference_Date, @Mail_Preference_Subscribed, ", _
                         "@Contact_Address_3, @Home_Address_3, @Email2", _
                       ") ", _
                    " else ", _
                        "update [CONTACT] set ", _
                        "Contact_RecNo=@Contact_RecNo,", _
                        "Company_RecNo=@Company_RecNo,Contact_Address_1=@Contact_Address_1,Contact_Post_Code=@Contact_Post_Code,", _
                        "Contact_State=@Contact_State,Contact_Address_2=@Contact_Address_2,Contact_City=@Contact_City,", _
                        "Contact_Country=@Contact_Country, Home_Address1=@Home_Address1, Home_Address2=@Home_Address2, Home_Suburb=@Home_Suburb, Home_State=@Home_State, Home_Postcode=@Home_Postcode,", _
                        "Home_Country=@Home_Country, Contact_Number=@Contact_Number, Position=@Position, Given_Name=@Given_Name, Surname=@Surname, ", _
                        "Phone_Bus=@Phone_Bus,Phone_Direct=@Phone_Direct, Phone_Pri=@Phone_Pri, Phone_Mobile=@Phone_Mobile, Fax=@Fax, EMail=@EMail, Other=@Other, Department=@Department,Debtor_Number=@Debtor_Number,", _
                        " Charge_Code=@Charge_Code, Web_Page=@Web_Page, Mailing_Address=@Mailing_Address,Contact_Colour=@Contact_Colour,", _
                        " Sequence=@Sequence, Created=@Created, Special_Notes=@Special_Notes,Title= @Title, ", _
                         " LastEdited=@LastEdited, Prefix=@Prefix, Created_by=@Created_by, Last_Edited_by=@Last_Edited_by, ", _
                        "PrivateContact=@PrivateContact, Mail_Preference_Date=@Mail_Preference_Date, Mail_Preference_Subscribed=@Mail_Preference_Subscribed, ", _
                         "Contact_Address_3=@Contact_Address_3, Home_Address_3=@Home_Address_3, Email2=@Email2", _
                        " where [Contact_RecNo] = @Contact_RecNo ")
            ReMsg = URunSql(mSql, param_Contact_RecNo, param_Company_RecNo, param_Contact_Address_1, _
                            param_Contact_Post_Code, param_Contact_State, param_Contact_Address_2, param_Contact_City, _
                            param_Contact_Country, param_Home_Address1, param_Home_Address2, param_Home_Suburb, param_Home_State, _
                            param_Home_Postcode, param_Home_Country, param_Contact_Number, param_Position, param_Given_Name, param_Surname, _
                            param_Phone_Bus, param_Phone_Direct, param_Phone_Pri, param_Phone_Mobile, param_Fax, param_EMail, param_Other, param_Department, param_Debtor_Number, _
                            param_Charge_Code, param_Web_Page, param_Mailing_Address, param_Contact_Colour, _
                            param_Sequence, param_Created, param_Special_Notes, param_Title, _
                             param_LastEdited, param_Prefix, param_Created_by, param_Last_Edited_by, _
                            param_PrivateContact, param_Mail_Preference_Date, param_Mail_Preference_Subscribed, _
                            param_Contact_Address_3, param_Home_Address_3, param_Email2)


            param_Contact_RecNo = Nothing
            param_Company_RecNo = Nothing
            param_Contact_Address_1 = Nothing
            param_Contact_Post_Code = Nothing
            param_Contact_State = Nothing
            param_Contact_Address_2 = Nothing
            param_Contact_City = Nothing
            param_Contact_Country = Nothing
            param_Home_Address1 = Nothing
            param_Home_Address2 = Nothing
            param_Home_Suburb = Nothing
            param_Home_State = Nothing
            param_Home_Postcode = Nothing
            param_Home_Country = Nothing
            param_Contact_Number = Nothing
            param_Position = Nothing
            param_Given_Name = Nothing
            param_Surname = Nothing
            param_Phone_Bus = Nothing
            param_Phone_Direct = Nothing
            param_Phone_Pri = Nothing
            param_Phone_Mobile = Nothing
            param_Fax = Nothing
            param_EMail = Nothing
            param_Other = Nothing
            param_Department = Nothing
            param_Debtor_Number = Nothing
            param_Charge_Code = Nothing
            param_Web_Page = Nothing
            param_Mailing_Address = Nothing
            param_Contact_Colour = Nothing
            param_Sequence = Nothing
            param_Created = Nothing
            param_Special_Notes = Nothing
            param_Title = Nothing
            '  param_Inactive_Contact_Date = Nothing
            param_LastEdited = Nothing
            param_Prefix = Nothing
            param_Created_by = Nothing
            param_Last_Edited_by = Nothing
            param_PrivateContact = Nothing
            param_Mail_Preference_Date = Nothing
            param_Mail_Preference_Subscribed = Nothing
            param_Contact_Address_3 = Nothing
            param_Home_Address_3 = Nothing
            param_Email2 = Nothing
        Catch ex As Exception
            MsgBox("Save ContactClass Error:" & ex.ToString)
        End Try
        Return ReMsg
    End Function
    '    SELECT     TOP 1 Contact_RecNo
    'FROM         CONTACT
    'ORDER BY Contact_RecNo DESC
    Function GetTopData(ByVal FieldName As String) As String
        Dim mreturn As String = ""
        Dim sql As String = ""
        sql = String.Concat("SELECT TOP 1 ", FieldName, " FROM [CONTACT] ORDER BY ", FieldName, " DESC ")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item(FieldName)) Then
                mreturn = ""
            Else
                mreturn = dt.Rows(0).Item(FieldName)
            End If
        End If
        Return mreturn
    End Function

    Function ChkContact(ByVal Name As String, ByVal LastName As String) As Boolean
        Dim mreturn As Boolean
        Dim sql As String = ""
        sql = String.Concat("SELECT Contact_RecNo FROM [CONTACT] WHERE Given_Name='" + Name + "' and Surname='" + LastName + "'")
        Dim dt As DataTable = UReadSqlbase(ConnectionString, sql, "")
        If dt.Rows.Count > 0 Then
            mreturn = False

        Else
            mreturn = True
        End If
        Return mreturn
    End Function
End Class
