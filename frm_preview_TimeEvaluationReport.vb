Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms
Public Class frm_preview_TimeEvaluationReport
    Public con_pallet As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
    Public con_pisp As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_app;User ID=sa;Password=pi_dc_2011")
    Private Sub frm_preview_TimeEvaluationReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim reportPath As String = System.IO.Path.Combine(Application.StartupPath, "\\192.168.10.27\Shopfloor_Folder\HYDRA PROJECT\RDL\IT_RPT_TimeEvaluationReport.rdl")
        ReportViewer1.ProcessingMode = ProcessingMode.Local
        ReportViewer1.LocalReport.ReportPath = reportPath
        Dim dataSource1 As New ReportDataSource("DataSet1", GetData())

        '' Clear existing data sources (optional, depending on your needs)
        ReportViewer1.LocalReport.DataSources.Clear()



        '' Add the new data source
        ReportViewer1.LocalReport.DataSources.Add(dataSource1)

        Dim reportParameters As New List(Of ReportParameter)()

        '' Add parameters to the list. Replace "ParameterName" with the actual name of the parameter
        '' defined in your report, and provide the value you want to pass.
        reportParameters.Add(New ReportParameter("machine", Frm_Transaction_Summary_Report.cmb_machine.Text))
        reportParameters.Add(New ReportParameter("startdate", Frm_Transaction_Summary_Report.dtp_startdate.Value.ToString))
        reportParameters.Add(New ReportParameter("enddate", Frm_Transaction_Summary_Report.dtp_enddate.Value.ToString))
        ''reportParameters.Add(New ReportParameter("empnum", EmpValue))

        '' Set the parameters to the report.
        ReportViewer1.LocalReport.SetParameters(reportParameters)

        Me.ReportViewer1.RefreshReport()


    End Sub

    Private Function GetData() As DataTable
        Dim datatable1 As New DataTable()
        Try
            con_pallet.Open()

            Dim cmd_timeeval_rpt As New SqlCommand("IT_RPT_TimeEvaluationReport", con_pallet)
            cmd_timeeval_rpt.CommandType = CommandType.StoredProcedure

            cmd_timeeval_rpt.Parameters.AddWithValue("machine", SqlDbType.NVarChar).Value = Frm_Transaction_Summary_Report.cmb_machine.Text
            cmd_timeeval_rpt.Parameters.AddWithValue("startdate", SqlDbType.Date).Value = Frm_Transaction_Summary_Report.dtp_startdate.Value.ToString
            cmd_timeeval_rpt.Parameters.AddWithValue("enddate", SqlDbType.Date).Value = Frm_Transaction_Summary_Report.dtp_enddate.Value.ToString

            Dim adapter As New SqlDataAdapter(cmd_timeeval_rpt)

            adapter.Fill(datatable1)

        Catch ex As Exception
        Finally
            con_pallet.Close()
        End Try
        Return datatable1
    End Function
End Class