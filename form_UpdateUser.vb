Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Configuration
Imports System.Security.Cryptography
Public Class form_UpdateUser

    Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnupdate.Click

        Dim cmdupdate As SqlCommand = New SqlCommand("
        UPDATE Employee
        SET
        Name=@name,
        Dept=@dept,
        Section=@section,
        Position=@position,
        Password=@password,
        Access=@access,
        Rights=@rights
        WHERE
        Emp_num=@empnum AND
        Password=@prevpass AND
        Site=@site
        ", con)

        cmdupdate.Parameters.AddWithValue("@name", txtname.Text)
        cmdupdate.Parameters.AddWithValue("@dept", txtdept.Text)
        cmdupdate.Parameters.AddWithValue("@section", txtsection.Text)
        cmdupdate.Parameters.AddWithValue("@position", txtposition.Text)
        cmdupdate.Parameters.AddWithValue("@password", Encrypt(txtpassword.Text))
        cmdupdate.Parameters.AddWithValue("@site", cmbsite.Text)
        cmdupdate.Parameters.AddWithValue("@prevpass", txtprevpass.Text)
        cmdupdate.Parameters.AddWithValue("@empnum", txtempnum.Text)
        cmdupdate.Parameters.AddWithValue("@access", cmb_access.Text)
        cmdupdate.Parameters.AddWithValue("@rights", cmb_rights.Text)

        Try
            con.Open()

            If txtpassword.Text = txtconfirmpass.Text Then
                MsgBox("Updated Successfully")

                cmdupdate.ExecuteNonQuery()
                Form4.Show()
                Me.Close()
            Else
                MsgBox("Password doesn't match")

            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub btnexit_Click(sender As Object, e As EventArgs) Handles btnexit.Click
        Form4.Show()
        Me.Close()
    End Sub
End Class