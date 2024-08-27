Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Configuration
Imports System.Security.Cryptography
Public Class Form3
    Dim con As SqlConnection
    Dim cmd As SqlCommand
    Dim FuncCls As New CommonFunctionsCls()

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            'IF TRUE, IT SHOWS THE TEXT
            txtpassword.UseSystemPasswordChar = False
            txtconfirmpass.UseSystemPasswordChar = False
        Else
            'IF FALSE, IT WILL HIDE THE TEXT AND IT WILL TURN INTO BULLETS.
            txtpassword.UseSystemPasswordChar = True
            txtconfirmpass.UseSystemPasswordChar = True
        End If
    End Sub

    'Private Function Encrypt(clearText As String) As String
    '    Dim EncryptionKey As String = "MAKV2SPBNI99212"
    '    Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
    '    Using encryptor As Aes = Aes.Create()
    '        Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
    '     &H65, &H64, &H76, &H65, &H64, &H65,
    '     &H76})
    '        encryptor.Key = pdb.GetBytes(32)
    '        encryptor.IV = pdb.GetBytes(16)
    '        Using ms As New MemoryStream()
    '            Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
    '                cs.Write(clearBytes, 0, clearBytes.Length)
    '                cs.Close()
    '            End Using
    '            clearText = Convert.ToBase64String(ms.ToArray())
    '        End Using
    '    End Using
    '    Return clearText
    'End Function

    'Private Function Decrypt(cipherText As String) As String
    '    Dim EncryptionKey As String = "MAKV2SPBNI99212"
    '    Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
    '    Using encryptor As Aes = Aes.Create()
    '        Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
    '     &H65, &H64, &H76, &H65, &H64, &H65,
    '     &H76})
    '        encryptor.Key = pdb.GetBytes(32)
    '        encryptor.IV = pdb.GetBytes(16)
    '        Using ms As New MemoryStream()
    '            Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
    '                cs.Write(cipherBytes, 0, cipherBytes.Length)
    '                cs.Close()
    '            End Using
    '            cipherText = Encoding.Unicode.GetString(ms.ToArray())
    '        End Using
    '    End Using
    '    Return cipherText
    'End Function

    Private Sub load_table()

        Dim constr As String = "Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011"
        Using con As New SqlConnection(constr)
            Using cmd As New SqlCommand("SELECT Site, Emp_num, Name, Dept, Position, Password FROM Employee WHERE IsActive = 1", con)
                Using ada As New SqlDataAdapter()
                    Dim dt As New DataTable()
                    cmd.CommandType = CommandType.Text
                    cmd.Connection = con
                    ada.SelectCommand = cmd
                    ada.Fill(dt)
                    Form4.DataGridView1.DataSource = dt
                    dt.Columns(5).ColumnName = "Password"
                    dt.Columns.Add("decrypt_password")
                    For Each row As DataRow In dt.Rows
                        row("decrypt_password") = Decrypt(row("Password").ToString)
                    Next
                End Using
            End Using
        End Using
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        opendb()
        registeruser(cmbsite.Text, txtempnum.Text, txtname.Text, txtdept.Text, txtsection.Text, txtposition.Text, Encrypt(txtpassword.Text.Trim()), cmb_access.Text, cmb_rights.Text)
        loadtable()
        closedb()
    End Sub


    Private Sub Form3_Load(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbsite.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnexit.Click
        Form4.Show()
        Me.Close()
    End Sub

    Private Sub Form3_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class