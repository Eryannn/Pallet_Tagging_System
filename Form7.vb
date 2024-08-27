Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Configuration
Imports System.Security.Cryptography
Public Class Form7
    Dim dtp As DateTimePicker = New DateTimePicker()
    Dim rect As Rectangle





    'Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
    '    dtp.Format = DateTimePickerFormat.Time
    '    dtp.ShowUpDown = True
    '    Select Case DataGridView1.Columns(e.ColumnIndex).Name
    '        Case "StartTime"
    '            rect = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
    '            dtp.Size = New Size(rect.Width, rect.Height)
    '            dtp.Location = New Point(rect.X, rect.Y)
    '            dtp.Visible = True
    '        Case "EndTime"
    '            rect = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
    '            dtp.Size = New Size(rect.Width, rect.Height)
    '            dtp.Location = New Point(rect.X, rect.Y)
    '            dtp.Visible = True
    '    End Select
    'End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        dtp.Format = DateTimePickerFormat.Custom
        dtp.CustomFormat = "yyyy-MM-dd HH:mm" ' You can customize the format based on your preference
        dtp.ShowUpDown = True

        ' Assuming startTimeColumnIndex is the index of the "StartTime" column
        Dim startTimeColumnIndex As Integer = 1

        If e.RowIndex >= 0 AndAlso e.ColumnIndex = startTimeColumnIndex Then
            Dim rect As Rectangle = DataGridView1.GetCellDisplayRectangle(startTimeColumnIndex, e.RowIndex, True)
            dtp.Size = New Size(rect.Width, rect.Height)
            dtp.Location = New Point(rect.X, rect.Y)
            dtp.Visible = True
            AddHandler dtp.ValueChanged, AddressOf dtp_ValueChanged

            If DataGridView1.CurrentCell.Value IsNot DBNull.Value Then
                dtp.Value = DateTime.Parse(DataGridView1.CurrentCell.Value.ToString())
            End If
        Else
            dtp.Visible = False
        End If



        'dtp.Format = DateTimePickerFormat.Custom
        'dtp.CustomFormat = "yyyy-MM-dd HH:mm" ' You can customize the format based on your preference
        'dtp.ShowUpDown = True

        'If e.RowIndex >= 0 AndAlso e.ColumnIndex = 2 Then
        '    ' Assuming yourTimeColumnIndex is the index of the column for which you want to show the date-time picker
        '    Dim rect As Rectangle = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
        '    dtp.Size = New Size(rect.Width, rect.Height)
        '    dtp.Location = New Point(rect.X, rect.Y)
        '    dtp.Visible = True
        '    AddHandler dtp.ValueChanged, AddressOf dtp_ValueChanged

        '    If DataGridView1.CurrentCell.Value IsNot DBNull.Value Then
        '        dtp.Value = DateTime.Parse(DataGridView1.CurrentCell.Value.ToString())
        '    End If
        'Else
        '    dtp.Visible = False
        'End If


        'dtp.Format = DateTimePickerFormat.Custom
        'dtp.CustomFormat = "HH:mm"
        'dtp.ShowUpDown = True
        'If e.RowIndex >= 0 AndAlso e.ColumnIndex = 2 Then
        '    ' Assuming yourTimeColumnIndex is the index of the column for which you want to show the time picker
        '    Dim rect As Rectangle = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
        '    dtp.Size = New Size(rect.Width, rect.Height)
        '    dtp.Location = New Point(rect.X, rect.Y)
        '    dtp.Visible = True
        '    AddHandler dtp.ValueChanged, AddressOf dtp_ValueChanged

        '    If DataGridView1.CurrentCell.Value IsNot DBNull.Value Then
        '        dtp.Value = DateTime.Parse(DataGridView1.CurrentCell.Value.ToString())
        '    End If
        'Else
        '    dtp.Visible = False
        'End If
    End Sub
    Private Sub dtp_ValueChanged(sender As Object, e As EventArgs)
        ' Update the cell value with the selected time
        DataGridView1.CurrentCell.Value = dtp.Value.ToString("HH:mm")
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        If e.Context = DataGridViewDataErrorContexts.Commit Then
            MsgBox("Invalid value in the DataGridViewComboBoxCell. Please select a valid item from the list.")
        End If

        ' You can also choose to suppress the exception:
        e.ThrowException = False
    End Sub

    Private Sub dtp_textchange(ByVal sender As Object, ByVal e As EventArgs)
        DataGridView1.CurrentCell.Value = dtp.Text.ToString
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim con1 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd As New SqlCommand("SELECT Seq, StartTime, EndTime, DTHrs, Category, CauseCode, NoteExistsFlag, Notes  from IT_KPI_DTDetail", con1)
        Dim cmd1 As New SqlCommand("SELECT Category ,Section from IT_KPI_SecDTCause where Section ='Die Cutting'", con1)

        Dim a As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        a.Fill(dt)
        DataGridView1.DataSource = dt

        DataGridView1.Controls.Add(dtp)
        dtp.Format = DateTimePickerFormat.Custom
        dtp.Visible = False
        AddHandler dtp.TextChanged, AddressOf dtp_textchange

        'Try
        '    con1.Open()

        '    Dim reader As SqlDataReader = cmd1.ExecuteReader
        '    Dim comboboxsite As New List(Of String)

        '    While reader.Read()
        '        comboboxsite.Add(reader(0))
        '    End While
        '    For Each row As DataGridViewRow In DataGridView1.Rows
        '        If row.Cells("Category").Value IsNot Nothing Then
        '            Dim test As New DataGridViewComboBoxCell
        '            test.DataSource = comboboxsite
        '            row.Cells("Category") = test
        '        End If


        '    Next
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally
        '    con1.Close()
        'End Try

    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        Dim con1 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd As New SqlCommand("SELECT Seq, StartTime, EndTime, DTHrs, Category, CauseCode, NoteExistsFlag, Notes  from IT_KPI_DTDetail", con1)
        Dim cmd1 As New SqlCommand("SELECT Category ,Section from IT_KPI_SecDTCause where Section ='Die Cutting'", con1)

        'Dim a As New SqlDataAdapter(cmd)
        'Dim dt As New DataTable
        'a.Fill(dt)
        'DataGridView1.DataSource = dt

        'DataGridView1.Controls.Add(dtp)
        'dtp.Format = DateTimePickerFormat.Custom
        'dtp.Visible = False
        'AddHandler dtp.TextChanged, AddressOf dtp_textchange

        Try
            con1.Open()

            Dim reader As SqlDataReader = cmd1.ExecuteReader
            Dim comboboxsite As New List(Of String)

            While reader.Read()
                comboboxsite.Add(reader(0))
            End While

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("Category").Value IsNot Nothing Then
                    Dim test As New DataGridViewComboBoxCell
                    test.DataSource = comboboxsite
                    row.Cells("Category") = test
                End If
            Next
            'For I As Integer = e.RowIndex To e.RowIndex + e.RowIndex - 1
            '    Dim row As DataGridViewRow = DataGridView1.Rows(I)

            '    If row.Cells("Category").Value IsNot Nothing Then
            '        Dim test As New DataGridViewComboBoxCell
            '        test.DataSource = comboboxsite
            '        row.Cells("Category") = test
            '    End If
            'Next
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            con1.Close()
        End Try




        'Try
        '    con1.Open()

        '    Dim reader As SqlDataReader = cmd1.ExecuteReader
        '    Dim comboboxsite As New List(Of String)

        '    While reader.Read()
        '        comboboxsite.Add(reader(0))
        '    End While
        '    For Each row As DataGridViewRow In DataGridView1.Rows
        '        If row.Cells("Category").Value IsNot Nothing Then
        '            Dim test As New DataGridViewComboBoxCell
        '            test.DataSource = comboboxsite
        '            row.Cells("Category") = test
        '        End If


        '    Next
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally
        '    con1.Close()
        'End Try
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) 
    End Sub
End Class