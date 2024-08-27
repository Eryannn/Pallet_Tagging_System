Public Class CommonFunctionsCls

    Public Function EncryptPassword(ByVal Password As String) As String

        'Encrypt the Password
        Dim sEncryptedPassword As String = ""
        Dim sEncryptKey As String = "P@SSW@RD@09" 'Should be minimum 8 characters

        Try
            sEncryptedPassword = EncryptDecryptClass.EncryptPasswordMD5(Password, sEncryptKey)

        Catch ex As Exception
            Return sEncryptedPassword
        End Try

        Return sEncryptedPassword
    End Function


    Public Function DecryptPassword(ByVal Password As String) As String
        'Encrypt the Password
        Dim sDecryptedPassword As String = ""
        Dim sEncryptKey As String = "P@SSW@RD@09" 'Should be minimum 8 characters

        Try
            sDecryptedPassword = EncryptDecryptClass.DecryptPasswordMD5(Password, sEncryptKey)

        Catch ex As Exception
            Return sDecryptedPassword
        End Try

        Return sDecryptedPassword
    End Function

End Class
