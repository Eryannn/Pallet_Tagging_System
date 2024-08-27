Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Security.Cryptography

Public Class Form4
    Dim con As SqlConnection
    Dim cmd As SqlCommand
    Dim FuncCls As New CommonFunctionsCls()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.Show()
        Dispose()
    End Sub

    Private Sub load_table()

        Dim constr As String = "Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011"
        Using con As New SqlConnection(constr)
            Using cmd As New SqlCommand("SELECT Site, Emp_num, Name, Dept, Section, Position, Password, Access, Rights FROM Employee WHERE IsActive = 1", con)
                Using ada As New SqlDataAdapter()
                    Dim dt As New DataTable()
                    cmd.CommandType = CommandType.Text
                    cmd.Connection = con
                    ada.SelectCommand = cmd
                    ada.Fill(dt)
                    DataGridView1.DataSource = dt
                    dt.Columns(6).ColumnName = "Password"
                    dt.Columns.Add("decrypt_password")
                    For Each row As DataRow In dt.Rows
                        row("decrypt_password") = Decrypt(row("Password").ToString)
                    Next
                End Using
            End Using
        End Using
    End Sub
    Private Sub load_data()
        'Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        'Dim cmd As New SqlCommand("SELECT Emp_num FROM Employee", con)

        'Try
        '    con.Open()
        '    Dim reader As SqlDataReader = cmd.ExecuteReader()

        '    While reader.Read()
        '        ComboBox1.Items.Add(reader("Emp_num").ToString())
        '    End While

        '    reader.Close()
        'Catch ex As Exception

        'Finally
        '    con.Close()
        'End Try


    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'opendb()
        'loadusertable()
        'closedb()
        load_table()
        'load_data()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick

    End Sub

    Private Sub btnupdate_Click(sender As Object, e As EventArgs) Handles btnupdate.Click
        Dim site As String = DataGridView1.SelectedCells(0).Value.ToString
        Dim empnum As String = DataGridView1.SelectedCells(1).Value.ToString
        Dim Name As String = DataGridView1.SelectedCells(2).Value.ToString
        Dim dept As String = DataGridView1.SelectedCells(3).Value.ToString
        Dim section As String = DataGridView1.SelectedCells(4).Value.ToString
        Dim position As String = DataGridView1.SelectedCells(5).Value.ToString
        Dim prevpass As String = DataGridView1.SelectedCells(6).Value.ToString
        Dim access As String = DataGridView1.SelectedCells(7).Value.ToString
        Dim rights As String = DataGridView1.SelectedCells(8).Value.ToString
        form_UpdateUser.Show()
        form_UpdateUser.cmbsite.Text = site
        form_UpdateUser.txtempnum.Text = empnum
        form_UpdateUser.txtname.Text = Name
        form_UpdateUser.txtdept.Text = dept
        form_UpdateUser.txtsection.Text = section
        form_UpdateUser.txtposition.Text = position
        form_UpdateUser.txtprevpass.Text = prevpass
        form_UpdateUser.cmb_access.Text = access
        form_UpdateUser.cmb_rights.Text = rights
        Me.Close()

    End Sub
End Class