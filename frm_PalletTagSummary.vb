Imports System.Data.SqlClient
Public Class Frm_PalletTagSummary
    Dim con1 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            con1.Open()
            Dim cmd_generatetable As SqlCommand = New SqlCommand("IT_View_Pallet_Tag_Summary_By_Job", con1)

            cmd_generatetable.CommandType = CommandType.StoredProcedure

            cmd_generatetable.Parameters.AddWithValue("hdrjob", SqlDbType.NVarChar).Value = txt_job.Text

            Dim a As New SqlDataAdapter(cmd_generatetable)
            Dim dt As New DataTable
            a.Fill(dt)
            dgv_reports.DataSource = dt

            AutofitColumns(dgv_reports)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con1.Close()
        End Try

    End Sub

    Private Sub AutofitColumns(dataGridView As DataGridView)
        ' Auto-resize columns to fit their content
        For Each column As DataGridViewColumn In dataGridView.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        ' Perform the actual auto-resizing based on content
        dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
    End Sub
End Class