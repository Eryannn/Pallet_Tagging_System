Imports System.Data.SqlClient
Public Class Frm_Transaction_Summary_Report
    Public con_pallet As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
    Public con_pisp As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_app;User ID=sa;Password=pi_dc_2011")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If cmb_report_type.Text = "TIME EVALUATION REPORT" Then
            frm_preview_TimeEvaluationReport.Show()

        Else
            frm_preview_outputvolumereport.Show()
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'Dim report7 As New RPT_Transaction_Summary
        'Dim user As String = "sa"
        'Dim pwd As String = "pi_dc_2011"

        'report7.SetDataSource(dg_preview)

        'report7.SetDatabaseLogon(user, pwd)
        'Form5.CrystalReportViewer1.ReportSource = report7
        'Form5.CrystalReportViewer1.Refresh()
        'Form5.CrystalReportViewer1.Zoom(50)
        'Form5.Show()
    End Sub

    Private Sub cmb_machine_DropDown(sender As Object, e As EventArgs) Handles cmb_machine.DropDown

        cmb_machine.Items.Clear()
        populaterscmachine()

    End Sub

    Private Sub cmb_machine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_machine.SelectedIndexChanged

    End Sub

    Private Sub Frm_Transaction_Summary_Report_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub populaterscmachine()
        Dim cmd_view_machine As SqlCommand = New SqlCommand("IT_Select_Machine", con_pallet)
        cmd_view_machine.CommandType = CommandType.StoredProcedure

        cmd_view_machine.Parameters.Add("@section", SqlDbType.NVarChar).Value = cmb_section.Text

        Try
            con_pallet.Open()
            Dim reader As SqlDataReader = cmd_view_machine.ExecuteReader
            If reader.HasRows Then
                While reader.Read()
                    cmb_machine.Items.Add(reader("RESID").ToString())
                End While
                reader.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con_pallet.Close()
        End Try
    End Sub

    Private Sub cmb_report_type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_report_type.SelectedIndexChanged
        If cmb_report_type.Text = "TIME EVALUATION REPORT" Then
            cmb_machine.Enabled = True
        Else
            cmb_machine.Enabled = False
            cmb_machine.Text = ""
        End If
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Show()
        Form2.txtoper_num.Text = Form1.cmbuserid.Text
        Form2.txtoper_numt3.Text = Form1.cmbuserid.Text
        Form2.TextBox3.Text = Form1.ComboBox2.Text
        Form2.txtshift.Text = Form1.cmbshift.Text
        Form2.lblshiftt2.Text = Form1.cmbshift.Text
        Form2.lblshift_reports.Text = Form1.cmbshift.Text
        Form2.lblusert2.Text = Form1.cmbuserid.Text.ToUpper
        Form2.txtshiftt3.Text = Form1.cmbshift.Text
        Form2.txtsite_subcon.Text = Form1.ComboBox2.Text
        Form2.txtshift_subcon.Text = Form1.cmbshift.Text
        Me.Close()
    End Sub
End Class