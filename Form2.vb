Imports System.Data.SqlClient
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports CrystalDecisions.Shared
Imports ZXing
Imports System.Windows.Forms

'Public Class WcItem
'    Public Property Wc As String
'    Public Property Description As String

'    Public Overrides Function ToString() As String
'        Return Wc
'    End Function
'End Class
Public Class Form2
    Dim con As SqlConnection
    Dim cmd As SqlCommand
    Dim da As SqlDataAdapter
    Dim dt As New DataSet
    Dim dtp As DateTimePicker = New DateTimePicker()
    Dim rect As Rectangle


    Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Declare Function Wow64EnableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Private osk As String = "C:\Windows\System32\osk.exe"

    Public Property Wc As String
    Public Property Description As String

    Public Overrides Function ToString() As String
        Return Wc
    End Function


    Private Sub bindgrid()

        Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
            Using cmd As New SqlCommand("SELECT * FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
                cmd.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                cmd.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
                cmd.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)
                cmd.CommandType = CommandType.Text
                Using sda As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        sda.Fill(dt)
                        DataGridView1.DataSource = dt
                    End Using
                End Using
            End Using
        End Using

        'Add a CheckBox Column to the DataGridView at the first position.
        Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
        checkBoxColumn.HeaderText = ""
        checkBoxColumn.Width = 30
        checkBoxColumn.Name = "checkBoxColumn"
        DataGridView1.Columns.Insert(0, checkBoxColumn)
    End Sub

    Private Sub getuserdetails()
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd1 As SqlCommand = New SqlCommand("SELECT Site, Emp_num, Name, section from Employee WHERE Emp_num=@empnum AND Site=@site", con1)
        'cmd1.Parameters.AddWithValue("@empnum", txtoper_num.Text)
        'cmd1.Parameters.AddWithValue("@site", TextBox3.Text)
        cmd1.Parameters.AddWithValue("@empnum", Form1.cmbuserid.Text)
        cmd1.Parameters.AddWithValue("@site", Form1.ComboBox2.Text)
        Dim currenttime As DateTime = DateTime.Now

        Try
            con1.Open()
            Dim sqlreader As SqlDataReader = cmd1.ExecuteReader()
            While sqlreader.Read()
                'MsgBox(sqlreader("Name").ToString)
                txtoper_name.Text = sqlreader("Name").ToString()
                txtoper_namet3.Text = sqlreader("Name").ToString()
                lblnamet2.Text = sqlreader("Name").ToString()

                lblsitet2.Text = sqlreader("Site").ToString()
                txtsitet3.Text = sqlreader("Site").ToString
                TextBox2.Text = sqlreader("Emp_num").ToString()
                txtoper_numt3.Text = sqlreader("Emp_num").ToString
                txtempnumt3.Text = sqlreader("Emp_num").ToString
                txtempnum_subcon.Text = sqlreader("Emp_num").ToString
                txtempname_subcon.Text = sqlreader("Name").ToString
                txtempnum_subcon_hdr.Text = sqlreader("Emp_num").ToString
                cmbsection_schedule.Text = sqlreader("section").ToString
                lblsite_reports.Text = sqlreader("Site").ToString
                lblempnum_reports.Text = sqlreader("Emp_num").ToString
                lblempname_report.Text = sqlreader("Name").ToString

                'lblnamet4.Text = sqlreader("Name").ToString
                'lblsitet4.Text = sqlreader("Site").ToString()


            End While
            sqlreader.Close()
        Catch ex As Exception
        Finally
            con1.Close()
        End Try
    End Sub
    Private Sub populaterscmach()
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        'Dim cmd1 As SqlCommand = New SqlCommand("SELECT RS.RESID,RS.DESCR FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID where wc.wc=@wc", con1)
        Dim cmd1 As SqlCommand = New SqlCommand("SELECT DISTINCT RS.RESID,RS.DESCR, RS.Uf_Resrc_Section FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid RIGHT JOIN RESRC000 RS ON RS.RESID=RG.RESID WHERE RS.Uf_Resrc_Section = @section", con1)
        cmd1.Parameters.AddWithValue("@section", cmbsection_schedule.Text)
        Try
            con1.Open()
            Dim reader As SqlDataReader = cmd1.ExecuteReader
            If reader.HasRows Then
                While reader.Read()
                    ComboBox2.Items.Add(reader("RESID").ToString())
                End While
                reader.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub populatewcdebug()
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd2 As SqlCommand = New SqlCommand("SELECT WC.wc, wc.description, wc.wc + ' - ' + wc.description as result FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID", con1)

        Try
            con1.Open()
            Dim sqlreader1 As SqlDataReader = cmd2.ExecuteReader()
            While sqlreader1.Read()
                Dim wc As String = sqlreader1("wc").ToString
                'ComboBox1.Items.Add(wc)
                'ComboBox1.ValueMember = sqlreader1("wc").ToString
            End While
            sqlreader1.Close()

        Catch ex As Exception

        End Try
    End Sub


    Private Sub populatewc()

        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd2 As SqlCommand = New SqlCommand("SELECT DISTINCT WC.wc, wc.description, wc.wc + ' - ' + wc.description as result FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID ORDER BY WC.WC ASC ", con1)

        Try
            con1.Open()
            Dim sqlreader1 As SqlDataReader = cmd2.ExecuteReader()
            While sqlreader1.Read()
                Dim wc As String = sqlreader1("wc").ToString + " - " + sqlreader1("description").ToString
                'ComboBox1.Items.Add(wc)
                'ComboBox1.ValueMember = sqlreader1("wc").ToString
            End While
            sqlreader1.Close()

        Catch ex As Exception

        End Try
    End Sub

    'Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs)
    '    Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '    Dim cmd As SqlCommand = New SqlCommand("SELECT WC.wc,WC.description FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID where wc.wc =@wc", con1)
    '    'cmd.Parameters.AddWithValue("@wc", ComboBox1.Text)
    '    Try
    '        con.Open()
    '        Dim reader As SqlDataReader = cmd.ExecuteReader
    '        If reader.HasRows Then
    '            MsgBox("has rows")
    '            While reader.Read()
    '                txtdescp2.Text = reader("description").ToString
    '            End While
    '        Else
    '            MsgBox("no rows")
    '        End If
    '    Catch ex As Exception
    '        'MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub load_table()
        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            'Dim cmd As SqlCommand = New SqlCommand("SELECT Tag_num as TAG#, Tag_date as DATE, Tag_qty as QTY/PALLET, Emp_num as EMP NO, Shift,Comment FROM Ptag_Line", con)
            Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM Ptag_Line", con)
            Dim a As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            a.Fill(dt)
            DataGridView1.DataSource = dt

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        getuserdetails()
        DataGridView1.DefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView1.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView1.RowHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView2.DefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView2.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView2.RowHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView3.DefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView3.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView3.RowHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView_subcon.DefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView_subcon.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        DataGridView_subcon.RowHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        'DataGridView4.DefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        'DataGridView4.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}
        'DataGridView4.RowHeadersDefaultCellStyle = New DataGridViewCellStyle With {.Font = New Font("Arial Narrow", 12), .BackColor = Color.White}

        Dim currenttime As DateTime = DateTime.Now
        'If currenttime.Hour >= 7 AndAlso currenttime.Hour < 19 Then
        '    cmbshiftt4.SelectedIndex = cmbshiftt4.Items.IndexOf("DS")
        'Else
        '    cmbshiftt4.SelectedIndex = cmbshiftt4.Items.IndexOf("NS")
        'End If

        'populatewc()

        'txtdate.Text = Today + " " + TimeOfDay
        'txtdatet3.Text = Today + " " + TimeOfDay
        txtdate.Text = DateTime.Now.ToString("MMMM dd, yyyy")
        txtdatet3.Text = DateTime.Now.ToString("MMMM dd, yyyy")
        txtdate_subcon.Text = DateTime.Now.ToString("MMMM dd, yyyy")
        TabPage1.Size = New Size(1000, 968)
    End Sub

    Private Sub AutofitColumns(dataGridView As DataGridView)
        ' Auto-resize columns to fit their content
        For Each column As DataGridViewColumn In dataGridView.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        ' Perform the actual auto-resizing based on content
        dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
    End Sub
    Private Sub txtjoboperation_TextChanged_1(sender As Object, e As EventArgs) Handles txtjoboperation.TextChanged
        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM Ptag_Hdr WHERE Job=@jonumber And Suffix=@josuffix And Oper_num=@operationnum", con)

        '      Dim cmd As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
        'from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
        'inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job =@jonumber and jrtrg.suffix=@josuffix and jrtrg.oper_num=@operationnum", con)

        cmd.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
        cmd.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
        cmd.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty, running_qty,  Emp_num, Name, Shift, Comment FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable
        a.Fill(dt)
        DataGridView1.DataSource = dt
        AutofitColumns(DataGridView1)
        disablecheckbox()
        Try
            con.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim latestrunningqty As Integer = getlatestrunningqty()



            If txtjoboperation.Text.Length >= 1 Then

                If reader.HasRows Then
                    btnsave.Enabled = False
                    btngenerate.Enabled = True


                    While reader.Read()
                        txtwc.Text = reader("Wc").ToString
                        txtnextop.Text = reader("Next_wc").ToString
                        rtbmachine.Text = reader("Resource").ToString
                        txtjobname.Text = reader("Item").ToString
                        txtjobdesc.Text = reader("Item_desc").ToString
                        txtjobqty.Text = reader("Job_qty").ToString
                        Dim tagqty As Integer = CInt(txtjobqty.Text)
                        Dim runningqty As Integer = tagqty - latestrunningqty
                        If runningqty < 0 Then
                            runningqty = 0
                        End If
                        txtremainingqty.Text = runningqty
                        txtnumtag.Text = reader("Total_tags").ToString
                        Label13.Text = reader("U_m").ToString
                        Label66.Text = reader("U_m").ToString
                        lblpoum.Text = reader("PO_Um").ToString
                        txtpoqty.Text = reader("PO_qty").ToString
                        txtdesc.Text = reader("Job_name").ToString
                        txtlot.Text = reader("Lot_no").ToString
                        txtstockcode.Text = reader("stock_desc").ToString
                        txtbigsize.Text = reader("Sheet_size").ToString
                        txtcutsize.Text = reader("Cut_size").ToString
                        txtnumout.Text = reader("num_out").ToString
                        If txtnumout.Text = "" Then
                            txtnumout.Text = 0
                        Else
                            txtnumout.Text = txtnumout.Text
                        End If
                    End While

                Else
                    reader.Close()
                    btnsave.Enabled = True
                    btngenerate.Enabled = False
                    'Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
                    'Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                    '    from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                    '    inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and REPLACE('',jrtrg.suffix,null)=@josuffix and jrtrg.oper_num=@operationnum", con1)
                    Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
                    Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                        from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                        inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and jrtrg.suffix = @josuffix and jrtrg.oper_num=@operationnum", con1)
                    cmd1.Parameters.AddWithValue("@jonumber", txtjoborder.Text)

                    cmd1.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
                    cmd1.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

                    Dim cmd2 As SqlCommand = New SqlCommand("select j.job,j.suffix, j.item,j.description ,itm.Uf_itemdesc_ext, j.qty_released as job_qty, itm.u_m
                        from job j inner join item itm on itm.item=j.item
                        where j.job=@jonumber and j.suffix=@josuffix", con1)
                    cmd2.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                    cmd2.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)

                    Dim cmd3 As SqlCommand = New SqlCommand("select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description
                         from jobroute jrt
                         inner join wc on wc.wc=jrt.wc 
                         where jrt.job =@jonumber
                         and jrt.suffix=@josuffix
                         and jrt.oper_num=@operationnum
                         order by  jrt.oper_num ASC", con1)
                    cmd3.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                    cmd3.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
                    cmd3.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

                    Dim cmd4 As SqlCommand = New SqlCommand("select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description
                         from jobroute jrt
                         inner join wc on wc.wc=jrt.wc 
                         where jrt.job =@jonumber
                         and jrt.suffix=@josuffix
                         and jrt.oper_num>@operationnum
                         order by  jrt.oper_num ASC", con1)
                    cmd4.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                    cmd4.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
                    cmd4.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

                    Dim cmd5 As SqlCommand = New SqlCommand("SELECT JOB.item,qty_released,ITEM.U_M FROM JOB INNER JOIN item ON JOB.item=item.item WHERE TYPE='J' AND JOB.suffix=0 AND JOB.JOB=@jonumber", con1)
                    cmd5.Parameters.AddWithValue("@jonumber", txtjoborder.Text)

                    Dim cmd6 As SqlCommand = New SqlCommand("select ref_job as job ,ref_suf as suffix,ref_oper as oper_num ,wc.wc,wc.description
                         from job jb
                         left join jobroute jr on jb.ref_job=jr.job
                         and jb.ref_suf=jr.suffix 
                         and jb.ref_oper=jr.oper_num
                         inner join wc on wc.wc=jr.wc
                          where jb.job=@jonumber and jb.suffix=@josuffix", con1)
                    cmd6.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                    cmd6.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)

                    Dim cmd7 As SqlCommand = New SqlCommand("SELECT 
		                    job.job,
		                    job.item, 
		                    job.qty_released, 
		                    item.description, 
		                    item.u_m 
                    INTO #job2 from job
                    INNER JOIN item on
		                    item.item = job.item
                    where job.type = 'J' AND job.suffix = 0 

                    SELECT 
		                    jroute.Uf_Operation_SC_LotNo,
		                    jroute.Uf_Operation_SC_StockCode,
		                    jroute.Uf_Operation_SC_BigSheet,
		                    jroute.Uf_Operation_SC_StockCode,
		                    jroute.Uf_Operation_SC_NumOuts,
		                    CONVERT(varchar, CAST(ROUND(item.Uf_Item_Width, 0) AS int)) + ' x ' + CONVERT(varchar, CAST(ROUND(item.Uf_Item_Length, 0) AS int)) + ' MM' AS [Cut_size],
                            job2.description
	                    FROM job
                    INNER JOIN jobroute jroute on 
		                    job.job = jroute.job AND
		                    jroute.suffix = job.suffix
                    INNER JOIN job_sch jsch on
		                    job.job = jsch.job AND
		                    job.suffix = jsch.suffix
                    INNER JOIN item on
		                    job.item = item.item
                    INNER JOIN jrtresourcegroup jrsrcgrp on
		                    jroute.job = jrsrcgrp.job AND
		                    jroute.suffix = jrsrcgrp.suffix AND
		                    jroute.oper_num = jrsrcgrp.oper_num
                    INNER JOIN RGRPMBR000 R000 on
		                    jrsrcgrp.rgid = r000.RGID
                    INNER JOIN RESRC000 C000 on
		                    r000.RESID = c000.RESID
                    INNER JOIN wc on 
		                    jroute.wc = wc.wc
                    INNER JOIN #job2 job2 on
		                    job.job = job2.job

                    WHERE 
		                    job.job = @job AND
		                    JOB.suffix = @suffix AND
		                    job.type = 'J' AND
		                    job.stat = 'R'", con1)

                    cmd7.Parameters.AddWithValue("@job", txtjoborder.Text)
                    cmd7.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)



                    Try
                        con1.Open()
                        Dim reader1 As SqlDataReader = cmd1.ExecuteReader
                        If reader1.HasRows Then
                            While reader1.Read()
                                rtbmachine.Text = reader1("RESID").ToString + " " + reader1("DESCR").ToString
                                'TextBox3.Text = reader1("DESCR").ToString
                            End While
                            reader1.Close()
                            Dim reader2 As SqlDataReader = cmd2.ExecuteReader
                            If reader2.HasRows Then
                                While reader2.Read()
                                    txtjobname.Text = reader2("item").ToString
                                    txtjobdesc.Text = reader2("description").ToString
                                    txtjobqty.Text = reader2("job_qty").ToString
                                    Label13.Text = reader2("U_m").ToString
                                End While
                                reader2.Close()
                                Dim reader3 As SqlDataReader = cmd3.ExecuteReader
                                If reader3.HasRows Then
                                    While reader3.Read()
                                        txtwc.Text = reader3("wc").ToString + " " + reader3("description").ToString
                                    End While
                                    reader3.Close()
                                    Dim reader4 As SqlDataReader = cmd4.ExecuteReader
                                    If reader4.HasRows Then
                                        While reader4.Read()
                                            txtnextop.Text = reader4("description").ToString
                                        End While
                                        reader4.Close()
                                        Dim reader5 As SqlDataReader = cmd5.ExecuteReader
                                        If reader5.HasRows Then
                                            While reader5.Read()
                                                txtpoqty.Text = reader5("qty_released").ToString
                                                lblpoum.Text = reader5("U_M").ToString
                                            End While
                                            reader5.Close()
                                            Dim reader7 As SqlDataReader = cmd7.ExecuteReader
                                            If reader7.HasRows Then
                                                While reader7.Read()
                                                    txtlot.Text = reader7("Uf_Operation_SC_LotNo").ToString
                                                    txtstockcode.Text = reader7("Uf_Operation_SC_StockCode").ToString
                                                    txtbigsize.Text = reader7("Uf_Operation_SC_BigSheet").ToString
                                                    txtdesc.Text = reader7("description").ToString
                                                    'txtnumout.Text = CInt(reader7("Uf_Operation_SC_NumOuts").ToString)
                                                    txtnumout.Text = reader7("Uf_Operation_SC_NumOuts").ToString
                                                    If txtnumout.Text = "" Then
                                                        txtnumout.Text = 0
                                                    Else
                                                        txtnumout.Text = txtnumout.Text
                                                    End If
                                                    txtcutsize.Text = reader7("Cut_size").ToString
                                                End While
                                                reader7.Close()
                                            End If
                                        End If
                                    Else
                                        reader4.Close()
                                        Dim reader6 As SqlDataReader = cmd6.ExecuteReader
                                        If reader6.HasRows Then
                                            While reader6.Read()
                                                txtnextop.Text = reader6("description").ToString
                                            End While
                                            reader6.Close()
                                            Dim reader5 As SqlDataReader = cmd5.ExecuteReader
                                            If reader5.HasRows Then
                                                While reader5.Read()
                                                    txtpoqty.Text = reader5("qty_released").ToString
                                                    lblpoum.Text = reader5("U_M").ToString
                                                End While
                                                reader5.Close()
                                                Dim reader7 As SqlDataReader = cmd7.ExecuteReader
                                                If reader7.HasRows Then
                                                    While reader7.Read()
                                                        txtlot.Text = reader7("Uf_Operation_SC_LotNo").ToString
                                                        txtstockcode.Text = reader7("Uf_Operation_SC_StockCode").ToString
                                                        txtbigsize.Text = reader7("Uf_Operation_SC_BigSheet").ToString
                                                        txtdesc.Text = reader7("description").ToString
                                                        'txtnumout.Text = CInt(reader7("Uf_Operation_SC_NumOuts").ToString)
                                                        txtnumout.Text = reader7("Uf_Operation_SC_NumOuts").ToString
                                                        If txtnumout.Text = "" Then
                                                            txtnumout.Text = 0
                                                        Else
                                                            txtnumout.Text = txtnumout.Text
                                                        End If
                                                        txtcutsize.Text = reader7("Cut_size").ToString
                                                    End While
                                                    reader7.Close()
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            txtqtyperpallet.Clear()
                            txtpoqty.Clear()
                            txtwc.Clear()
                            txtnextop.Clear()
                            rtbmachine.Clear()
                            txtjobname.Clear()
                            txtjobqty.Clear()
                            txtjobdesc.Clear()
                            txtnumtag.Clear()
                            txtremainingqty.Clear()
                            Label66.Text = If(cleartext() = 0, "", cleartext().ToString())
                            Label13.Text = If(cleartext() = 0, "", cleartext().ToString())
                            lblpoum.Text = If(cleartext() = 0, "", cleartext().ToString())
                            txtdesc.Clear()
                            txtbigsize.Clear()
                            txtnumout.Clear()
                            txtsheetqty.Clear()
                            txtcutsize.Clear()
                            txtlot.Clear()
                            txt_nextopernum.Clear()
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Error: " & ex.Message)
                    Finally
                        con1.Close()
                    End Try
                End If
            ElseIf txtjoboperation.Text.Length = 0
                txtqtyperpallet.Clear()
                txtpoqty.Clear()
                txtwc.Clear()
                txtnextop.Clear()
                rtbmachine.Clear()
                txtjobname.Clear()
                txtjobqty.Clear()
                txtjobdesc.Clear()
                txtnumtag.Clear()
                txtremainingqty.Clear()
                Label66.Text = If(cleartext() = 0, "", cleartext().ToString())
                Label13.Text = If(cleartext() = 0, "", cleartext().ToString())
                lblpoum.Text = If(cleartext() = 0, "", cleartext().ToString())
                txtdesc.Clear()
                txtbigsize.Clear()
                txtnumout.Clear()
                txtsheetqty.Clear()
                txtcutsize.Clear()
                txtlot.Clear()
                txt_nextopernum.Clear()
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            con.Close()
        End Try

        Try
            If txtnextop.Text.Length >= 1 Then
                Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")

                con1.Open()

                Dim cmd_nextop As SqlCommand = New SqlCommand("SELECT job, suffix, oper_num, wc from jobroute WHERE wc = 'P610' AND job = @job AND Suffix = @suffix", con1)

                cmd_nextop.Parameters.AddWithValue("job", txtjoborder.Text)
                cmd_nextop.Parameters.AddWithValue("suffix", txtjobsuffix.Text)

                Dim read_cmdnextop As SqlDataReader = cmd_nextop.ExecuteReader
                If read_cmdnextop.HasRows Then
                    While read_cmdnextop.Read
                        txt_nextopernum.Text = read_cmdnextop("oper_num").ToString
                    End While
                    read_cmdnextop.Close()
                End If
                con1.Close()
            Else
                txt_nextopernum.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM Ptag_Hdr WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)

        '      Dim cmd As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
        'from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
        'inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job =@jonumber and jrtrg.suffix=@josuffix and jrtrg.oper_num=@operationnum", con)

        cmd.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
        cmd.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
        cmd.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

        Dim a As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        a.Fill(dt)
        DataGridView1.DataSource = dt
        Try
            con.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read()
                    rtbmachine.Text = reader("Item").ToString
                End While
            Else
                reader.Close()
                Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
                Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and jrtrg.suffix=@josuffix and jrtrg.oper_num=@operationnum", con1)
                cmd1.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                cmd1.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
                cmd1.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

                Dim cmd2 As SqlCommand = New SqlCommand(" select j.job,j.suffix, j.item,j.description ,itm.Uf_itemdesc_ext, j.qty_released as job_qty, itm.u_m
                from job j inner join item itm on itm.item=j.item
                where j.job=@jonumber and j.suffix=@josuffix", con1)
                cmd2.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                cmd2.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)


                Try
                    con1.Open()
                    Dim reader1 As SqlDataReader = cmd1.ExecuteReader
                    If reader1.HasRows Then
                        While reader1.Read()
                            rtbmachine.Text = reader1("RESID").ToString
                            txtjobname.Text = reader1("DESCR").ToString
                        End While
                        reader1.Close()
                        Dim reader2 As SqlDataReader = cmd2.ExecuteReader
                        If reader2.HasRows Then
                            While reader2.Read()
                                txtjobqty.Text = reader2("job_qty").ToString
                                Label13.Text = reader2("u_m").ToString
                            End While

                        End If
                        'reader1.Close()

                        'Dim cmd2 As SqlCommand = New SqlCommand(" select j.job,j.suffix, j.item,j.description ,itm.Uf_itemdesc_ext, j.qty_released as job_qty, itm.u_m
                        'from job j inner join item itm on itm.item=j.item
                        'where j.job=@jonumber and j.suffix=@josuffix", con1)
                        'cmd2.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
                        'cmd2.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)

                        'Try
                        '    Dim reader2 As SqlDataReader = cmd2.ExecuteReader
                        '    While reader2.Read()
                        '        TextBox6.Text = reader2("job_qty").ToString
                        '    End While
                        '    reader2.Close()
                        'Catch ex As Exception

                        'End Try
                    Else
                        MessageBox.Show("No Rows detected")
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message)
                Finally
                    con1.Close()
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            con.Close()
        End Try

    End Sub


    Private Function getlatestsequence() As Integer
        Dim latestseq As Integer = 0

        Try
            Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
                con.Open()

                Using cmd As New SqlCommand("select max(seq) from IT_KPI_DTDetail where RefRowPointer = @refrowpointer", con)
                    'cmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)

                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot DBNull.Value Then
                        latestseq = Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch ex As Exception

        End Try

        Return latestseq
    End Function

    Private Function GetLatestTagNumberFromDatabaset3() As Integer
        Dim latestTagNumber As Integer = 0

        Try
            ' Open a connection to your SQL Server database
            Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
                con.Open()

                ' Create a SQL command to retrieve the latest tag number
                Using cmd As New SqlCommand("Select MAX(Tag_num) FROM NCFPtag_Line WHERE Site=@site AND Job=@job AND Suffix=@suffix AND Oper_num=@opernum", con)
                    cmd.Parameters.AddWithValue("@site", txtsitet3.Text)
                    cmd.Parameters.AddWithValue("@job", txtjobordert3.Text)
                    cmd.Parameters.AddWithValue("@suffix", txtjobsuffixt3.Text)
                    cmd.Parameters.AddWithValue("@opernum", txtjoboperationt3.Text)

                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot DBNull.Value Then
                        latestTagNumber = Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle any exceptions here
            MsgBox(ex.Message)
        End Try

        Return latestTagNumber
    End Function
    Private Function GetLatestTagNumberFromDatabase() As Integer
        Dim latestTagNumber As Integer = 0

        Try
            ' Open a connection to your SQL Server database
            Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
                con.Open()

                ' Create a SQL command to retrieve the latest tag number
                Using cmd As New SqlCommand("SELECT MAX(Tag_num) FROM Ptag_Line WHERE Site=@site AND Job=@job AND Suffix=@suffix AND Oper_num=@opernum", con)
                    cmd.Parameters.AddWithValue("@site", TextBox3.Text)
                    cmd.Parameters.AddWithValue("@job", txtjoborder.Text)
                    cmd.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
                    cmd.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot DBNull.Value Then
                        latestTagNumber = Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle any exceptions here
            MsgBox(ex.Message)
        End Try

        Return latestTagNumber
    End Function


    Private Function getlatestrunningqty() As Integer
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        'Using cmd As New SqlCommand("SELECT TOP 1 running_qty FROM Ptag_Line where job = @job and Suffix = @suffix and Oper_num = @opernum ORDER BY Tag_date DESC", con)
        Using cmd As New SqlCommand("SELECT TOP 1 running_qty FROM Ptag_Line where job = @job and Suffix = @suffix and Oper_num = @opernum ORDER BY Tag_num DESC", con)

            cmd.Parameters.AddWithValue("@job", txtjoborder.Text)
            cmd.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

            Try
                con.Open()
                Dim result As Object = cmd.ExecuteScalar
                con.Close()

                If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                    Return CInt(result)
                    MsgBox(result)
                Else
                    Return 0 ' Default value if no rows are found
                End If
            Catch ex As Exception
            End Try
        End Using
        Return 0
    End Function


    Private Sub btngenerate_Click(sender As Object, e As EventArgs) Handles btngenerate.Click
        btngenerate.Enabled = False

        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #], Tag_date, Tag_qty, running_qty, Emp_num, Name, Shift, Comment FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable

        Dim tagqty As Integer = CInt(txtqtyperpallet.Text)
        Dim runningqty As Integer = getlatestrunningqty() + tagqty

        Try
            con.Open()

            Dim cmd As SqlCommand




            cmd = New SqlCommand("INSERT INTO Ptag_Line ([Select], Site, Job, Suffix, Oper_num, Tag_num, Tag_date, Tag_qty, running_qty, Emp_num, Name, Shift, Comment, sheet_qty) VALUES (@select, @site, @job, @suffix, @opernum, @tagnum, @tagdate, @tagqty, @runningqty, @empnum, @name, @shift, @comment, @sheetqty)", con)

            cmd.Parameters.AddWithValue("@select", 0)
            cmd.Parameters.AddWithValue("@site", TextBox3.Text)
            cmd.Parameters.AddWithValue("@job", txtjoborder.Text)
            cmd.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

            ' Retrieve the latest tag number from the database
            Dim latestTagNumber As Integer = GetLatestTagNumberFromDatabase()
            ' Increment the latest tag number
            latestTagNumber += 1
            cmd.Parameters.AddWithValue("@tagnum", latestTagNumber)

            cmd.Parameters.AddWithValue("@tagdate", DateTime.Now)
            cmd.Parameters.AddWithValue("@tagqty", txtqtyperpallet.Text)
            cmd.Parameters.AddWithValue("@runningqty", runningqty)
            cmd.Parameters.AddWithValue("@empnum", txtoper_num.Text)
            cmd.Parameters.AddWithValue("@name", txtoper_name.Text)
            cmd.Parameters.AddWithValue("@shift", txtshift.Text)
            cmd.Parameters.AddWithValue("@comment", cmbcomment.Text)

            cmd.Parameters.AddWithValue("@sheetqty", CInt(txtsheetqty.Text))
            cmd.ExecuteNonQuery()
            btngenerate.Enabled = True

            Dim remainingtagqty As Integer = CInt(txtjobqty.Text) - runningqty
            If remainingtagqty < 0 Then
                remainingtagqty = 0
            End If
            txtremainingqty.Text = remainingtagqty

            cleartext()
            a.Fill(dt)
            DataGridView1.DataSource = dt
            MsgBox("Generate Tag Successfully")
            disablecheckbox()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(getlatestrunningqty)
            MsgBox(txtsheetqty.Text)
        Finally
            con.Close()
        End Try


        'Try
        '    con.Open()

        '    Dim cmd As SqlCommand
        '    If Integer.Parse(txtgentags.Text) <= Integer.Parse(txtnumtag.Text) Then
        '        If Integer.TryParse(txtgentags.Text, Nothing) AndAlso CInt(txtgentags.Text) >= 1 Then
        '            Dim tagCount As Integer = CInt(txtgentags.Text)

        '            ' Retrieve the latest tag number from the database
        '            Dim latestTagNumber As Integer = GetLatestTagNumberFromDatabase()
        '            Dim latestrunningqty As Integer = getlatestrunningqty()

        '            'MsgBox(latestrunningqty)

        '            For i As Integer = 1 To tagCount

        '                Dim tagqty As Integer = CInt(txtqtyperpallet.Text)
        '                Dim runningqty As Integer = latestrunningqty + tagqty

        '                cmd = New SqlCommand("INSERT INTO Ptag_Line ([Select], Site, Job, Suffix, Oper_num, Tag_num, Tag_date, Tag_qty, running_qty, Emp_num, Name, Shift, Comment) VALUES (@select, @site, @job, @suffix, @opernum, @tagnum, @tagdate, @tagqty, @runningqty, @empnum, @name, @shift, @comment)", con)

        '                cmd.Parameters.AddWithValue("@select", 0)
        '                cmd.Parameters.AddWithValue("@site", TextBox3.Text)
        '                cmd.Parameters.AddWithValue("@job", txtjoborder.Text)
        '                cmd.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
        '                cmd.Parameters.AddWithValue("@opernum", txtjoboperation.Text)
        '                ' Increment the latest tag number for each row
        '                latestTagNumber += 1
        '                cmd.Parameters.AddWithValue("@tagnum", latestTagNumber)
        '                cmd.Parameters.AddWithValue("@tagdate", DateTime.Now)
        '                cmd.Parameters.AddWithValue("@tagqty", txtqtyperpallet.Text)
        '                cmd.Parameters.AddWithValue("@runningqty", runningqty)
        '                cmd.Parameters.AddWithValue("@empnum", txtoper_num.Text)
        '                cmd.Parameters.AddWithValue("@name", txtoper_name.Text)
        '                cmd.Parameters.AddWithValue("@shift", txtshift.Text)
        '                cmd.Parameters.AddWithValue("@comment", cmbcomment.Text)

        '                If Integer.TryParse(txtgentags.Text, Nothing) AndAlso CInt(txtgentags.Text) >= 0 Then
        '                    cmd.ExecuteNonQuery()
        '                    btngenerate.Enabled = True
        '                Else
        '                    MsgBox("No. of tags is not available")
        '                    Exit For ' Exit the loop if there's an issue
        '                End If
        '                latestrunningqty = runningqty
        '                Dim remainingtagqty As Integer = CInt(txtjobqty.Text)
        '                Dim result As Integer = remainingtagqty - latestrunningqty

        '                If result < 0 Then
        '                    result = 0
        '                End If
        '                txtremainingqty.Text = result
        '            Next
        '            cleartext()
        '            a.Fill(dt)
        '            DataGridView1.DataSource = dt
        '            MsgBox("Generate Tags Successfully")
        '            disablecheckbox()
        '        End If
        '    Else
        '        MsgBox("Invalid number of tags")
        '        btngenerate.Enabled = True
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally
        '    con.Close()
        'End Try

        Try
            con.Open()
            Dim cmd_updatehdr As New SqlCommand("UPDATE Ptag_Hdr Set total_tags=@totaltags, Pallet_size=@palletqty WHERE job=@job AND suffix=@suffix AND Oper_num=@opernum", con)

            Dim latestTagNumberhdr As Integer = GetLatestTagNumberFromDatabase()

            cmd_updatehdr.Parameters.AddWithValue("@totaltags", latestTagNumberhdr)
            cmd_updatehdr.Parameters.AddWithValue("@palletqty", txtremainingqty.Text)
            cmd_updatehdr.Parameters.AddWithValue("@job", txtjoborder.Text)
            cmd_updatehdr.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
            cmd_updatehdr.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

            cmd_updatehdr.ExecuteNonQuery()
            ' MsgBox(latestTagNumberhdr)

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("debugging in updating hdr")
        Finally
            con.Close()
        End Try

    End Sub
    Private Sub btnsave_Click_1(sender As Object, e As EventArgs) Handles btnsave.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            If txtjoborder.Text.Length = 10 And txtjobsuffix.Text.Length >= 1 And txtjoboperation.Text.Length >= 1 Then
                Dim cmd As SqlCommand = New SqlCommand("INSERT INTO Ptag_Hdr (
                        Site,
                        Job,
                        Suffix,
                        Oper_num,
                        Wc,
                        Wc_desc,
                        Next_wc, 
                        Machine, 
                        Resource,
                        Item,
                        Item_desc,
                        U_m,
                        Job_qty,
                        PO_qty,
                        PO_Um,
                        Pallet_size,
                        Total_tags,
                        Createdate,
                        Job_name,
                        Stock_desc,
                        Sheet_size,
                        Cut_size,
                        Lot_no,
                        num_out
                        ) 
                VALUES (
                        @site,
                        @job,
                        @suffix,
                        @opernum,
                        @wc,
                        @wcdesc,
                        @nextwc,
                        @machine,
                        @resource,
                        @item,
                        @itemdesc,
                        @um,
                        @jobqty,
                        @poqty,
                        @poum,
                        @palletsize,
                        @totalsize,
                        @createdate,
                        @jobname,
                        @stockdesc,
                        @sheetsize,
                        @cutsize,
                        @lotno,
                        @numout)", con)
                Dim machine As String() = rtbmachine.Text.Split(" "c)
                Dim wcdesc As String() = txtwc.Text.Split(" "c)

                cmd.Parameters.AddWithValue("@site", TextBox3.Text)
                cmd.Parameters.AddWithValue("@job", txtjoborder.Text)
                cmd.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
                cmd.Parameters.AddWithValue("@opernum", txtjoboperation.Text)
                cmd.Parameters.AddWithValue("@wc", wcdesc(0))
                If wcdesc.Length >= 2 Then
                    ' Join all elements except the first one (the code)
                    Dim description As String = String.Join(" ", wcdesc.Skip(1))
                    cmd.Parameters.AddWithValue("@wcdesc", description)
                End If
                'cmd.Parameters.AddWithValue("@wcdesc", wcdesc(1)) 'ERIAN add this line for wcdesc
                cmd.Parameters.AddWithValue("@nextwc", txtnextop.Text)

                'cmd.Parameters.AddWithValue("@resource", rtbmachine.Text)
                cmd.Parameters.AddWithValue("@machine", machine(0))
                cmd.Parameters.AddWithValue("@resource", rtbmachine.Text)
                cmd.Parameters.AddWithValue("@item", txtjobname.Text)
                cmd.Parameters.AddWithValue("@itemdesc", txtjobdesc.Text)
                cmd.Parameters.AddWithValue("@um", Label13.Text)
                cmd.Parameters.AddWithValue("@jobqty", txtjobqty.Text)
                If txtpoqty IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtpoqty.Text) Then
                    cmd.Parameters.AddWithValue("@poqty", txtpoqty.Text)
                Else
                    cmd.Parameters.AddWithValue("@poqty", DBNull.Value)
                End If
                'cmd.Parameters.AddWithValue("@poqty", If(Integer.TryParse(txtpoqty.Text, Nothing), txtpoqty.Text, DBNull.Value))
                'cmd.Parameters.AddWithValue("@poqty", txtpoqty.Text)
                cmd.Parameters.AddWithValue("@poum", lblpoum.Text)
                cmd.Parameters.AddWithValue("@palletsize", txtqtyperpallet.Text)
                cmd.Parameters.AddWithValue("@totalsize", txtnumtag.Text)
                cmd.Parameters.AddWithValue("@createdate", DateTime.Now)

                cmd.Parameters.AddWithValue("@jobname", txtdesc.Text)
                cmd.Parameters.AddWithValue("@stockdesc", txtstockcode.Text)
                cmd.Parameters.AddWithValue("@sheetsize", txtbigsize.Text)
                cmd.Parameters.AddWithValue("@cutsize", txtcutsize.Text)
                cmd.Parameters.AddWithValue("@lotno", txtlot.Text)
                cmd.Parameters.AddWithValue("@numout", CInt(txtnumout.Text))

                cmd.ExecuteNonQuery()
                MessageBox.Show("Saved Successfully")
                btnsave.Enabled = False
                btngenerate.Enabled = True
                cleartext()
            Else
                MsgBox("Invalid Job")
                btnsave.Enabled = True
                btngenerate.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub txtqtyperpallet_TextChanged_1(sender As Object, e As EventArgs) Handles txtqtyperpallet.TextChanged

        If Not String.IsNullOrEmpty(txtjobqty.Text) AndAlso Not String.IsNullOrEmpty(txtqtyperpallet.Text) AndAlso IsNumeric(txtjobqty.Text) AndAlso IsNumeric(txtqtyperpallet.Text) AndAlso Convert.ToDecimal(txtqtyperpallet.Text) <> 0 Then
            txtnumtag.Text = Convert.ToInt32(Decimal.Parse(txtjobqty.Text) / Decimal.Parse(txtqtyperpallet.Text)).ToString()
            If txtnumout.Text = "" Then
                txtnumout.Text = 0
            End If
            If txtnumout.Text = 0 Then
                txtsheetqty.Text = 0
            Else
                txtsheetqty.Text = CInt(txtqtyperpallet.Text) / CInt(txtnumout.Text)
            End If

        Else
            txtnumtag.Text = "" ' or any other value to indicate division by zero or invalid input
        End If

        If txtnumout.Text = "" Then
            txtsheetqty.Text = 0
        End If

    End Sub
    Private Sub btnlogout_Click_1(sender As Object, e As EventArgs) Handles btnlogout.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            'Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where name=@name AND emp_num = @empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            ' Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where Name=@name and Emp_num=@empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)
            Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)

            cmd.Parameters.AddWithValue("name", txtoper_name.Text)
            cmd.Parameters.AddWithValue("empnum", txtoper_num.Text)
            cmd.Parameters.AddWithValue("job", txtjoborder.Text)
            cmd.Parameters.AddWithValue("suffix", txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("opernum", txtjoboperation.Text)

            cmd1.Parameters.AddWithValue("name", txtoper_namet3.Text)
            cmd1.Parameters.AddWithValue("empnum", txtoper_numt3.Text)
            cmd1.Parameters.AddWithValue("job", txtjobordert3.Text)
            cmd1.Parameters.AddWithValue("suffix", txtjobsuffixt3.Text)
            cmd1.Parameters.AddWithValue("opernum", txtjoboperationt3.Text)

            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
            Form1.txtpassword.Clear()
            Form1.cmbuserid.SelectedValue = 0
            Form1.Show()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub
    Private Sub btnkeyboard_Click_1(sender As Object, e As EventArgs) Handles btnkeyboard.Click
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Try
            con1.Open()
            Dim cmd7 As SqlCommand = New SqlCommand("SELECT 
		                    job.job,
		                    job.item, 
		                    job.qty_released, 
		                    item.description, 
		                    item.u_m 
                    INTO #job2 from job
                    INNER JOIN item on
		                    item.item = job.item
                    where job.type = 'J' AND job.suffix = 0 

                    SELECT 
		                    jroute.Uf_Operation_SC_LotNo,
		                    jroute.Uf_Operation_SC_StockCode,
		                    jroute.Uf_Operation_SC_BigSheet,
		                    jroute.Uf_Operation_SC_StockCode,
		                    jroute.Uf_Operation_SC_NumOuts,
		                    CONVERT(varchar, CAST(ROUND(item.Uf_Item_Width, 0) AS int)) + ' x ' + CONVERT(varchar, CAST(ROUND(item.Uf_Item_Length, 0) AS int)) + ' MM' AS [Cut Size],
		                    jroute.Uf_Operation_SC_LotNo
	                    FROM job
                    INNER JOIN jobroute jroute on 
		                    job.job = jroute.job AND
		                    jroute.suffix = job.suffix
                    INNER JOIN job_sch jsch on
		                    job.job = jsch.job AND
		                    job.suffix = jsch.suffix
                    INNER JOIN item on
		                    job.item = item.item
                    INNER JOIN jrtresourcegroup jrsrcgrp on
		                    jroute.job = jrsrcgrp.job AND
		                    jroute.suffix = jrsrcgrp.suffix AND
		                    jroute.oper_num = jrsrcgrp.oper_num
                    INNER JOIN RGRPMBR000 R000 on
		                    jrsrcgrp.rgid = r000.RGID
                    INNER JOIN RESRC000 C000 on
		                    r000.RESID = c000.RESID
                    INNER JOIN wc on 
		                    jroute.wc = wc.wc
                    INNER JOIN #job2 job2 on
		                    job.job = job2.job

                    WHERE 
		                    job.job = @job AND
		                    job.suffix = @suffix AND
		                    job.type = 'J' AND
		                    job.stat = 'R'", con1)

            cmd7.Parameters.AddWithValue("@job", txtjoborder.Text)
            cmd7.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)

            Dim reader7 As SqlDataReader = cmd7.ExecuteReader
            If reader7.HasRows Then
                While reader7.Read()
                    txtlot.Text = reader7("Uf_Operation_SC_LotNo").ToString
                    txtstockcode.Text = reader7("Uf_Operation_SC_StockCode").ToString
                    txtbigsize.Text = reader7("Uf_Operation_SC_BigSheet").ToString
                End While
                MsgBox(reader7("Uf_Operation_SC_LotNo").ToString)
                MsgBox(reader7("Uf_Operation_SC_BigSheet").ToString)
                reader7.Close()
            Else
                MsgBox("no rows")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con1.Close()
        End Try


        'Dim old As Long
        'If Environment.Is64BitOperatingSystem Then
        '    If Wow64DisableWow64FsRedirection(old) Then
        '        Process.Start(osk)
        '        Wow64EnableWow64FsRedirection(old)
        '    End If
        'Else
        '    Process.Start(osk)
        'End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim checkboxcell As DataGridViewCheckBoxCell = TryCast(row.Cells(0), DataGridViewCheckBoxCell)
                If checkboxcell IsNot Nothing Then
                    checkboxcell.Value = True
                End If
            End If
        Next
    End Sub
    Private Sub txtjoborder_TextChanged(sender As Object, e As EventArgs) Handles txtjoborder.TextChanged
    End Sub

    Private Sub txtjobsuffix_TextChanged_1(sender As Object, e As EventArgs)
        If txtjobsuffix.Text.Length >= 1 Then
            txtjoboperation.Focus()
        End If
    End Sub
    Private Sub Button5_Click_2(sender As Object, e As EventArgs) Handles Button5.Click

        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim selectedrows As New List(Of Integer)

        Try
            con.Open()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value IsNot Nothing AndAlso CBool(DataGridView1.Rows(i).Cells(0).Value) = True Then
                    selectedrows.Add(i)
                End If
            Next

            Dim report4 As New CrystalReport4
            Dim user As String = "sa"
            Dim pwd As String = "pi_dc_2011"

            For Each rowIndex As Integer In selectedrows
                ' Create a new report instance for each selected row
                'Dim selectedvalue As Boolean

                report4.SetParameterValue("job", txtjoborder.Text)
                report4.SetParameterValue("suffix", txtjobsuffix.Text)
                report4.SetParameterValue("oper_num", txtjoboperation.Text)
                'report4.SetParameterValue("job_operation", txtjoboperation.Text)
                'report4.SetParameterValue("Select", selectedvalue)
                'report4.SetParameterValue("total_tagnum", DataGridView1.Rows.Count)
                report4.SetParameterValue("operator", txtoper_num.Text)

            Next

            report4.SetDatabaseLogon(user, pwd)
            Form5.CrystalReportViewer1.ReportSource = report4
            Form5.CrystalReportViewer1.Refresh()
            Form5.CrystalReportViewer1.Zoom(50)
            Form5.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try


    End Sub

    ' ============================ THIS SECTION IS FOR THE 2ND TAB ============================
    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown
        ComboBox2.Items.Clear()
        populaterscmach()

        'Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        'Dim cmd As SqlCommand = New SqlCommand("SELECT RS.RESID,RS.DESCR FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID where wc.wc=@wc", con1)
        'cmd.Parameters.AddWithValue("@wc", ComboBox1.Text)
        'Try
        '    ComboBox2.Items.Clear()
        '    con1.Open()
        '    Dim reader As SqlDataReader = cmd.ExecuteReader
        '    While reader.Read()
        '        Dim rscmch As String = reader("RESID").ToString()
        '        ComboBox2.Items.Add(rscmch)r
        '    End While
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '    Dim cmd As SqlCommand = New SqlCommand("SELECT WC.wc,WC.description, wc.wc + ' - ' + wc.description as Result FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID where (wc.wc + ' - ' + wc.description)=@wc", con1)
    '    cmd.Parameters.AddWithValue("@wc", ComboBox1.Text)

    '    'MsgBox(ComboBox1.ValueMember)

    '    Try
    '        con1.Open()
    '        Dim reader As SqlDataReader = cmd.ExecuteReader
    '        If reader.HasRows Then
    '            While reader.Read()
    '                txtdescp2.Text = reader("description").ToString
    '                'ComboBox1.SelectedItem = -1
    '                'ComboBox1.SelectedText = reader("wc").ToString
    '                populatewcdebug()
    '                ComboBox1.Text = reader("wc").ToString
    '            End While

    '        End If
    '        reader.Close()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    Finally
    '        con1.Close()
    '    End Try
    'End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'rs.RESID AS Machine,
        '        jr.job ,
        '        jr.suffix,
        '        jr.oper_num As Oper,
        '        js.start_date AS [Date],
        '        jo.item As Item ,
        '        itm.description,
        '        CAST(jo.qty_released As INT) As QTY,
        '        itm.u_m AS UM
        'From job  jo  
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd As SqlCommand = New SqlCommand("
        BEGIN TRAN
        
       

        SELECT 
	            jr.job ,
		        jr.suffix,
		        jr.oper_num as Oper,
		        js.start_date AS [Date],
		        jo.item AS Item ,
		        itm.description,
		        CAST(jo.qty_released AS INT) AS QTY,
		        itm.u_m AS UM
        FROM job  jo  
        INNER JOIN jobroute jr 
		           ON jo.job=jr.job AND jo.suffix=jr.suffix
        INNER JOIN job_sch js
		           ON jo.job=js.job AND jo.suffix=js.suffix
        INNER JOIN item itm 
		           ON itm.item=jo.item
        INNER JOIN wc 
		           ON wc.wc =jr.wc
        INNER JOIN jrtresourcegroup jrs 
		           ON jr.job=jrs.job AND jr.suffix=jrs.suffix 
		                             AND jr.oper_num=jrs.oper_num
        INNER JOIN RGRPMBR000 rg 
		           ON jrs.rgid=rg.RGID 
        INNER JOIN RESRC000 rs 
		           ON rg.RESID=rs.RESID

        WHERE jo.type='J' 
          AND jo.stat='R'
          AND js.start_date >=@startdate
          AND js.end_date <=@enddate
          AND rs.RESID =@resid
          
        ORDER BY DATE ASC
        ROLLBACK TRAN", con1)
        'remove this line from query AND jo1.qty_complete < jo1.qty_released select 
        'job,
        '        suffix,
        '        qty_released,
        '        qty_complete
        'INTO #JOB1
        'From job
        'Where
        ''        suffix = 0
        'INNER Join #JOB1 jo1
        '           On jo1.job = jo.job
        'cmd.Parameters.AddWithValue("@wc", ComboBox1.Text)AND JR.wc=@wc
        cmd.CommandTimeout = 120
        cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = DateTimePicker1.Value.Date
        'cmd.Parameters.Add("@enddate", SqlDbType.DateTime).Value = DateTimePicker2.Value.Date
        cmd.Parameters.Add("@enddate", SqlDbType.DateTime).Value = DateTimePicker2.Value.AddDays(1)
        cmd.Parameters.AddWithValue("@resid", ComboBox2.Text)

        Dim a As New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        a.Fill(dt)
        DataGridView2.DataSource = dt
        Dim qtyColumn As DataGridViewColumn = DataGridView2.Columns(6)
        Dim datecolumn As DataGridViewColumn = DataGridView2.Columns(3)
        ' Set the format for the quantity column to display commas
        qtyColumn.DefaultCellStyle.Format = "N0"
        qtyColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        datecolumn.DefaultCellStyle.Format = "MM/dd/yyyy"
        AutofitColumns(DataGridView2)
    End Sub

    Private Sub txtjobsuffix_TextChanged(sender As Object, e As EventArgs) Handles txtjobsuffix.TextChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where name=@name AND emp_num = @empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where Name=@name and Emp_num=@empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)

            cmd.Parameters.AddWithValue("name", txtoper_name.Text)
            cmd.Parameters.AddWithValue("empnum", txtoper_num.Text)
            cmd.Parameters.AddWithValue("job", txtjoborder.Text)
            cmd.Parameters.AddWithValue("suffix", txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("opernum", txtjoboperation.Text)

            cmd1.Parameters.AddWithValue("name", txtoper_namet3.Text)
            cmd1.Parameters.AddWithValue("empnum", txtoper_numt3.Text)
            cmd1.Parameters.AddWithValue("job", txtjobordert3.Text)
            cmd1.Parameters.AddWithValue("suffix", txtjobsuffixt3.Text)
            cmd1.Parameters.AddWithValue("opernum", txtjoboperationt3.Text)

            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
            Form1.txtpassword.Clear()
            Form1.cmbuserid.SelectedValue = 0
            Form1.Show()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub
    Private Sub ComboBox1_Click(sender As Object, e As EventArgs)
        ComboBox2.SelectedIndex = -1
    End Sub
    '================================================ THIS SECTION IS FOR 3rd TAB =======================================================
    Private Sub txtjoboperationt3_TextChanged(sender As Object, e As EventArgs) Handles txtjoboperationt3.TextChanged
        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd As SqlCommand = New SqlCommand("SELECT * FROM NCFPtag_Hdr WHERE Job=@jonumber And Suffix=@josuffix And Oper_num=@operationnum", con)

        '      Dim cmd As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
        'from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
        'inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job =@jonumber and jrtrg.suffix=@josuffix and jrtrg.oper_num=@operationnum", con)

        cmd.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
        cmd.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
        cmd.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty as 'PALLET QTY', Qty_Affected AS 'QTY AFFECTED',  Emp_num as Empnum, Name, Shift, Comment FROM NCFPtag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable
        a.Fill(dt)
        DataGridView3.DataSource = dt
        AutofitColumns(DataGridView3)
        disablecheckbox1()



        Try
            con.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader

            If txtjoboperationt3.Text.Length >= 1 Then

                If reader.HasRows Then
                    btnsavet3.Enabled = False
                    btngeneratet3.Enabled = True
                    While reader.Read()
                        txtwct3.Text = reader("Wc").ToString
                        txtnextop_t3.Text = reader("Next_wc").ToString
                        rtbmachinet3.Text = reader("Resource").ToString
                        txtjobnamet3.Text = reader("Item").ToString
                        txtjobdesct3.Text = reader("Item_desc").ToString
                        txtjobqtyt3.Text = reader("Job_qty").ToString
                        Label28.Text = reader("U_m").ToString
                    End While


                Else
                    reader.Close()
                    btnsavet3.Enabled = True
                    btngeneratet3.Enabled = False
                    'Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
                    'Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                    '    from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                    '    inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and REPLACE('',jrtrg.suffix,null)=@josuffix and jrtrg.oper_num=@operationnum", con1)
                    Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
                    Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                        from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                        inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and jrtrg.suffix = @josuffix and jrtrg.oper_num=@operationnum", con1)
                    cmd1.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
                    cmd1.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
                    cmd1.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

                    Dim cmd2 As SqlCommand = New SqlCommand(" select j.job,j.suffix, j.item,j.description ,itm.Uf_itemdesc_ext, j.qty_released as job_qty, itm.u_m
                        from job j inner join item itm on itm.item=j.item
                        where j.job=@jonumber and j.suffix=@josuffix", con1)
                    cmd2.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
                    cmd2.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)

                    Dim cmd3 As SqlCommand = New SqlCommand("select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description
                         from jobroute jrt
                         inner join wc on wc.wc=jrt.wc 
                         where jrt.job =@jonumber
                         and jrt.suffix=@josuffix
                         and jrt.oper_num=@operationnum
                         order by  jrt.oper_num ASC", con1)
                    cmd3.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
                    cmd3.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
                    cmd3.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

                    Dim cmd4 As SqlCommand = New SqlCommand("select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description
                         from jobroute jrt
                         inner join wc on wc.wc=jrt.wc 
                         where jrt.job =@jonumber
                         and jrt.suffix=@josuffix
                         and jrt.oper_num>@operationnum
                         order by  jrt.oper_num ASC", con1)
                    cmd4.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
                    cmd4.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
                    cmd4.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

                    Try
                        con1.Open()
                        Dim reader1 As SqlDataReader = cmd1.ExecuteReader
                        If reader1.HasRows Then
                            While reader1.Read()
                                rtbmachinet3.Text = reader1("RESID").ToString + " " + reader1("DESCR").ToString
                                'TextBox3.Text = reader1("DESCR").ToString
                            End While
                            reader1.Close()
                            Dim reader2 As SqlDataReader = cmd2.ExecuteReader
                            If reader2.HasRows Then
                                While reader2.Read()
                                    txtjobnamet3.Text = reader2("item").ToString
                                    txtjobdesct3.Text = reader2("description").ToString
                                    txtjobqtyt3.Text = reader2("job_qty").ToString
                                    Label28.Text = reader2("U_m").ToString
                                End While
                                reader2.Close()
                                Dim reader3 As SqlDataReader = cmd3.ExecuteReader
                                If reader3.HasRows Then
                                    While reader3.Read()
                                        txtwct3.Text = reader3("wc").ToString + " " + reader3("description").ToString
                                    End While
                                    reader3.Close()
                                    Dim reader4 As SqlDataReader = cmd4.ExecuteReader
                                    If reader4.HasRows Then
                                        While reader4.Read()
                                            txtnextop_t3.Text = reader4("description").ToString
                                        End While
                                        reader4.Close()
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Error: " & ex.Message)
                    Finally
                        con1.Close()
                    End Try
                End If
            ElseIf txtjoboperationt3.Text.Length <= 1
                txtwct3.Clear()
                txtgeneratetagt3.Clear()
                txtnextop_t3.Clear()
                rtbmachinet3.Clear()
                txtjobnamet3.Clear()
                txtjobqtyt3.Clear()
                txtjobdesct3.Clear()
                Label28.Text = If(cleartext() = 0, "", cleartext().ToString())
                txtjobdesct3.Clear()
                txtjobnamet3.Clear()

            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            con.Close()
        End Try
    End Sub


    Private Sub load_data()
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            da = New SqlDataAdapter("SELECT [Select], Tag_num, RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty as 'PALLET QTY', Qty_Affected AS 'QTY AFFECTED', Emp_num, Name, Shift, Comment FROM NCFPtag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)

            da.SelectCommand.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
            da.SelectCommand.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
            da.SelectCommand.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

            dt.Clear()
            da.Fill(dt)
            DataGridView1.DataSource = dt.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            con.Close()
        End Try




        'da = New SqlDataAdapter("SELECT [Select], Tag_num , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty as 'PALLET QTY', Qty_Affected AS 'QTY AFFECTED',  Emp_num, Name, Shift, Comment FROM NCFPtag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)

        'da.SelectCommand.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
        'da.SelectCommand.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
        'da.SelectCommand.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

        'dt.Clear()
        'da.Fill(dt)
        'DataGridView1.DataSource = dt.Tables(0)
        'con.Close()
    End Sub
    Private Sub btnupdate_Click(sender As Object, e As EventArgs) Handles btnupdate.Click

        Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
            con.Open()
            Dim updateCommandText As String = "UPDATE NCFPtag_Line SET Tag_qty = @tagqty, Qty_Affected = @qtyaffected, Comment=@comment WHERE Tag_num=@tagnum AND Site = @site AND Job=@job AND Suffix=@suffix AND Oper_num=@opernum AND Name=@name"
            Dim cmd As New SqlCommand(updateCommandText, con)

            ' Add parameters to the SqlCommand
            cmd.Parameters.Add(New SqlParameter("@tagqty", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@qtyaffected", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@tagnum", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@site", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@job", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@suffix", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@opernum", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@name", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@comment", SqlDbType.NVarChar))

            Try
                ' Iterate through the DataGridView rows
                For Each row As DataGridViewRow In DataGridView3.Rows
                    ' Check if the row is not the header row
                    If Not row.IsNewRow Then
                        ' Set parameter values based on the data in the current row
                        cmd.Parameters("@tagqty").Value = row.Cells("PALLET QTY").Value
                        cmd.Parameters("@qtyaffected").Value = row.Cells("QTY AFFECTED").Value
                        cmd.Parameters("@tagnum").Value = row.Cells("Tag #").Value
                        cmd.Parameters("@site").Value = txtsitet3.Text
                        cmd.Parameters("@job").Value = txtjobordert3.Text
                        cmd.Parameters("@suffix").Value = txtjobsuffixt3.Text
                        cmd.Parameters("@opernum").Value = txtjoboperationt3.Text
                        cmd.Parameters("@name").Value = txtoper_namet3.Text
                        cmd.Parameters("@comment").Value = row.Cells("Comment").Value
                        ' Execute the update command for the current row
                        cmd.ExecuteNonQuery()
                    End If
                Next
                MsgBox("Update Successfully")
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        End Using


    End Sub

    Private Sub btnsavet3_Click(sender As Object, e As EventArgs) Handles btnsavet3.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            If txtjobordert3.Text.Length = 10 And txtjobsuffixt3.Text.Length >= 1 And txtjoboperationt3.Text.Length >= 1 Then
                Dim cmd As SqlCommand = New SqlCommand("INSERT INTO NCFPtag_Hdr (Site, Job, Suffix, Oper_num, Wc, Wc_desc, Next_wc, Machine, Resource, Item, Item_desc, U_m, Job_qty, Createdate) VALUES (@site,@job,@suffix,@opernum,@wc,@wcdesc,@nextwc, @machine,@resource,@item,@itemdesc,@um,@jobqty,@createdate)", con)

                Dim machinet3 As String() = rtbmachinet3.Text.Split(" "c)
                Dim wcdesct3 As String() = txtwct3.Text.Split(" "c)

                cmd.Parameters.AddWithValue("@site", txtsitet3.Text)
                cmd.Parameters.AddWithValue("@job", txtjobordert3.Text)
                cmd.Parameters.AddWithValue("@suffix", txtjobsuffixt3.Text)
                cmd.Parameters.AddWithValue("@opernum", txtjoboperationt3.Text)
                'cmd.Parameters.AddWithValue("@wc", txtwct3.Text)

                cmd.Parameters.AddWithValue("@nextwc", txtnextop_t3.Text)
                cmd.Parameters.AddWithValue("@machine", machinet3(0))
                cmd.Parameters.AddWithValue("@resource", rtbmachinet3.Text)
                cmd.Parameters.AddWithValue("@item", txtjobnamet3.Text)
                cmd.Parameters.AddWithValue("@itemdesc", txtjobdesct3.Text)
                cmd.Parameters.AddWithValue("@um", Label28.Text)
                cmd.Parameters.AddWithValue("@jobqty", txtjobqtyt3.Text)
                cmd.Parameters.AddWithValue("@createdate", DateTime.Now)


                If wcdesct3.Length >= 2 Then
                    ' Join all elements except the first one (the code)
                    Dim description As String = String.Join(" ", wcdesct3.Skip(1))
                    cmd.Parameters.AddWithValue("@wcdesc", description)
                End If

                cmd.Parameters.AddWithValue("@wc", wcdesct3(0))
                cmd.ExecuteNonQuery()

                MessageBox.Show("Saved Successfully")
                btnsavet3.Enabled = False
                btngeneratet3.Enabled = True
                cleartext()
            Else
                MsgBox("Invalid Job")
                btnsavet3.Enabled = True
                btngeneratet3.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub btngeneratet3_Click(sender As Object, e As EventArgs) Handles btngeneratet3.Click
        btngeneratet3.Enabled = False

        Dim con3 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")


        Dim tablecmd3 As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty as 'PALLET QTY', Qty_Affected AS 'QTY AFFECTED',  Emp_num, Name, Shift, Comment FROM NCFPtag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con3)
        tablecmd3.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
        tablecmd3.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
        tablecmd3.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

        Dim a3 As New SqlDataAdapter(tablecmd3)
        Dim dt3 As New DataTable

        Try
            con3.Open()

            Dim cmd3 As SqlCommand


            ' Retrieve the latest tag number from the database
            Dim latestTagNumber3 As Integer = GetLatestTagNumberFromDatabaset3()

            cmd3 = New SqlCommand("INSERT INTO NCFPtag_Line ([Select], Site, Job, Suffix, Oper_num, Tag_num, Tag_date, Emp_num, Name, Shift) VALUES (@Select, @site, @job, @suffix, @opernum, @tagnum, @tagdate, @empnum, @name, @shift)", con3)

            cmd3.Parameters.AddWithValue("@Select", 0)
            cmd3.Parameters.AddWithValue("@site", txtsitet3.Text)
            cmd3.Parameters.AddWithValue("@job", txtjobordert3.Text)
            cmd3.Parameters.AddWithValue("@suffix", txtjobsuffixt3.Text)
            cmd3.Parameters.AddWithValue("@opernum", txtjoboperationt3.Text)
            ' Increment the latest tag number for each row
            latestTagNumber3 += 1
            cmd3.Parameters.AddWithValue("@tagnum", latestTagNumber3)
            cmd3.Parameters.AddWithValue("@tagdate", DateTime.Now)
            cmd3.Parameters.AddWithValue("@empnum", txtoper_numt3.Text)
            cmd3.Parameters.AddWithValue("@name", txtoper_namet3.Text)
            cmd3.Parameters.AddWithValue("@shift", txtshiftt3.Text)

            cmd3.ExecuteNonQuery()
            btngeneratet3.Enabled = True

            cleartext()
            a3.Fill(dt3)
            DataGridView3.DataSource = dt3
            MsgBox("Generate Tags Successfully")
            disablecheckbox1()

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("debugging")
        Finally
            con3.Close()
        End Try

    End Sub

    Private Sub DataGridView1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseUp
        ' Check if the clicked cell is in the "Select" column


        If e.ColumnIndex = DataGridView1.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
            ' This code will make sure the checkbox value is toggled immediately
            DataGridView1.EndEdit()
        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Try
            con.Open()

            For Each row As DataGridViewRow In DataGridView1.Rows
                Dim ischecked As Boolean = Convert.ToBoolean(row.Cells(0).Value)

                If ischecked Then
                    Dim cmdchecked As New SqlCommand("UPDATE Ptag_Line Set [Select] = 1 WHERE Job = @job And Suffix = @suffix And Oper_num = @opernum And Tag_num = @tagnum", con)

                    cmdchecked.Parameters.AddWithValue("job", txtjoborder.Text)
                    cmdchecked.Parameters.AddWithValue("suffix", txtjobsuffix.Text)
                    cmdchecked.Parameters.AddWithValue("opernum", txtjoboperation.Text)
                    cmdchecked.Parameters.AddWithValue("tagnum", row.Cells(1).Value)

                    cmdchecked.ExecuteNonQuery()
                Else
                    Dim cmdunchecked As New SqlCommand("UPDATE Ptag_Line Set [Select] = 0 WHERE Job = @job And Suffix = @suffix And Oper_num = @opernum And Tag_num = @tagnum", con)

                    cmdunchecked.Parameters.AddWithValue("job", txtjoborder.Text)
                    cmdunchecked.Parameters.AddWithValue("suffix", txtjobsuffix.Text)
                    cmdunchecked.Parameters.AddWithValue("opernum", txtjoboperation.Text)
                    cmdunchecked.Parameters.AddWithValue("tagnum", row.Cells(1).Value)

                    cmdunchecked.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub UpdateDatabase(selectedValue As Boolean, rowIndex As Integer)

        Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
            con.Open()
            Using cmd As New SqlCommand("Update Ptag_line Set [Select] = @selectedvalue where [Select] = @Select", con)
                cmd.Parameters.AddWithValue("@selectedvalue", If(selectedValue, 1, 0))
                cmd.Parameters.AddWithValue("@Select", rowIndex)

                cmd.ExecuteNonQuery()

            End Using
        End Using

    End Sub


    Private Sub txtjobordert3_TextChanged(sender As Object, e As EventArgs) Handles txtjobordert3.TextChanged

    End Sub

    Private Sub txtjobsuffixt3_TextChanged(sender As Object, e As EventArgs) Handles txtjobsuffixt3.TextChanged

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim selectedrows As New List(Of Integer)

        For i As Integer = 0 To DataGridView3.Rows.Count - 1
            If DataGridView3.Rows(i).Cells(0).Value IsNot Nothing AndAlso CBool(DataGridView3.Rows(i).Cells(0).Value) = True Then
                selectedrows.Add(i)
            End If
        Next

        Dim report4 As New CrystalReport5
        Dim user As String = "sa"
        Dim pwd As String = "pi_dc_2011"

        For Each rowIndex As Integer In selectedrows
            ' Create a new report instance for each selected row
            'Dim selectedvalue As Boolean

            report4.SetParameterValue("job", txtjobordert3.Text)
            report4.SetParameterValue("suffix", txtjobsuffixt3.Text)
            report4.SetParameterValue("oper_num", txtjoboperationt3.Text)
            'report4.SetParameterValue("job_operation", txtjoboperation.Text)
            'report4.SetParameterValue("Select", selectedvalue)
            report4.SetParameterValue("Total_tagnum", DataGridView3.Rows.Count)
            report4.SetParameterValue("operator", txtoper_namet3.Text)

        Next
        report4.SetDatabaseLogon(user, pwd)
        Form5.CrystalReportViewer1.ReportSource = report4

        Form5.CrystalReportViewer1.Refresh()
        Form5.CrystalReportViewer1.Zoom(50)


        Form5.Show()
    End Sub

    Private Sub DataGridView3_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellValueChanged
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Try
            con.Open()

            For Each row As DataGridViewRow In DataGridView3.Rows
                Dim ischecked As Boolean = Convert.ToBoolean(row.Cells(0).Value)
                If ischecked Then
                    Dim cmdchecked As New SqlCommand("UPDATE NCFPtag_Line Set [Select] = 1 WHERE Job = @job And Suffix = @suffix And Oper_num = @opernum And Tag_num = @tagnum", con)

                    cmdchecked.Parameters.AddWithValue("job", txtjobordert3.Text)
                    cmdchecked.Parameters.AddWithValue("suffix", txtjobsuffixt3.Text)
                    cmdchecked.Parameters.AddWithValue("opernum", txtjoboperationt3.Text)
                    cmdchecked.Parameters.AddWithValue("tagnum", row.Cells(1).Value)

                    cmdchecked.ExecuteNonQuery()
                Else
                    Dim cmdunchecked As New SqlCommand("UPDATE NCFPtag_Line Set [Select] = 0 WHERE Job = @job And Suffix = @suffix And Oper_num = @opernum And Tag_num = @tagnum", con)

                    cmdunchecked.Parameters.AddWithValue("job", txtjobordert3.Text)
                    cmdunchecked.Parameters.AddWithValue("suffix", txtjobsuffixt3.Text)
                    cmdunchecked.Parameters.AddWithValue("opernum", txtjoboperationt3.Text)
                    cmdunchecked.Parameters.AddWithValue("tagnum", row.Cells(1).Value)

                    cmdunchecked.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView3_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.CellMouseUp
        If e.ColumnIndex = DataGridView3.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
            ' This code will make sure the checkbox value is toggled immediately
            DataGridView3.EndEdit()
        End If

    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles btnsellectall.Click

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim createdByUsername As String = row.Cells("Emp_num").Value.ToString
                Dim username As String = loggedinusername()

                If loggedinusername() = createdByUsername Then
                    Dim checkBoxCell As DataGridViewCheckBoxCell = TryCast(row.Cells("Select"), DataGridViewCheckBoxCell)
                    If checkBoxCell IsNot Nothing Then
                        checkBoxCell.Value = Not CBool(checkBoxCell.Value)
                    End If
                End If
            End If
        Next

    End Sub

    Private Function loggedinusername()

        Dim createdbyuser As String = String.Empty

        Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
            con.Open()

            Using cmd As New SqlCommand("Select Name, Emp_num from Ptag_line WHERE Emp_num=@empnum", con)
                cmd.Parameters.AddWithValue("@empnum", txtoper_num.Text)
                Dim reader As SqlDataReader = cmd.ExecuteReader()

                If reader.Read() Then
                    createdbyuser = reader("Emp_num").ToString()
                End If
            End Using
        End Using
        Return createdbyuser

    End Function
    Private Function disablecolumn()
        'For Each row As DataGridViewRow In DataGridView4.Rows
        '    row.Cells("DTHrs").ReadOnly = True
        '    row.Cells("RootCause").ReadOnly = True
        'Next
        'Return 0
    End Function
    Private Function disablecheckbox()
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim createdbyusername As String = row.Cells("Emp_num").Value.ToString

            If loggedinusername() = createdbyusername Then

                row.ReadOnly = True
                row.Cells("Comment").ReadOnly = False
                row.Cells("Select").ReadOnly = False

            Else
                'row.Cells("Select").ReadOnly = True
                row.ReadOnly = True
                row.Cells("Comment").ReadOnly = True
                row.Cells("Select").ReadOnly = True

            End If
        Next
        Return 0
    End Function

    Private Function disablecheckbox1()
        For Each row As DataGridViewRow In DataGridView3.Rows
            Dim createdbyusername As String = row.Cells("Empnum").Value.ToString

            If loggedinusername() = createdbyusername Then
                row.ReadOnly = False

                'row.Cells("Select").ReadOnly = False
                'row.Cells("PALLET QTY").ReadOnly = False
            Else
                row.ReadOnly = True
                'row.Cells("Select").ReadOnly = True
                'row.Cells("PALLET QTY").ReadOnly = True
            End If
        Next
        Return 0
    End Function

    Private Function disablecheckbox_subcon()
        For Each row As DataGridViewRow In DataGridView_subcon.Rows
            Dim createdbyusername As String = row.Cells("Emp_num").Value.ToString

            If loggedinusername() = createdbyusername Then

                row.ReadOnly = True
                row.Cells("Comment").ReadOnly = False
                row.Cells("Select").ReadOnly = False

            Else
                'row.Cells("Select").ReadOnly = True
                row.ReadOnly = True
                row.Cells("Comment").ReadOnly = True
                row.Cells("Select").ReadOnly = True

            End If
        Next
        Return 0
    End Function
    Private Sub Button9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'For Each row As DataGridViewRow In DataGridView3.Rows
        '    If Not row.IsNewRow Then
        '        Dim createdByUsername As String = row.Cells("Name").Value.ToString
        '        Dim username As String = loggedinusername()

        '        If loggedinusername() = createdByUsername Then
        '            Dim checkBoxCell As DataGridViewCheckBoxCell = TryCast(row.Cells("Select"), DataGridViewCheckBoxCell)
        '            If checkBoxCell IsNot Nothing Then
        '                checkBoxCell.Value = Not CBool(checkBoxCell.Value)
        '            End If
        '        End If
        '    End If
        'Next

        For Each row As DataGridViewRow In DataGridView3.Rows
            If Not row.IsNewRow Then
                Dim createdByUsername As String = row.Cells("Emp_Num").Value.ToString
                Dim username As String = loggedinusername()

                If loggedinusername() = createdByUsername Then
                    Dim checkBoxCell As DataGridViewCheckBoxCell = TryCast(row.Cells("Select"), DataGridViewCheckBoxCell)
                    If checkBoxCell IsNot Nothing Then
                        checkBoxCell.Value = Not CBool(checkBoxCell.Value)
                    End If
                End If

            End If

        Next
    End Sub

    'Private Sub ComboBox1_DropDown(sender As Object, e As EventArgs)
    '    ComboBox1.Items.Clear()
    '    populatewc()
    'End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim writer As New BarcodeWriter
        Dim writer1 As New BarcodeWriter

        writer1.Format = BarcodeFormat.QR_CODE
        writer.Format = BarcodeFormat.CODE_93

        Dim data1 As String = DataGridView2.SelectedCells(0).Value.ToString
        Form6.PictureBox1.Image = writer1.Write(DataGridView2.SelectedCells(0).Value.ToString)
        Form6.PictureBox2.Image = writer1.Write(DataGridView2.SelectedCells(1).Value.ToString)
        Form6.PictureBox3.Image = writer1.Write(DataGridView2.SelectedCells(2).Value.ToString)
        Form6.txtjob.Text = DataGridView2.SelectedCells(0).Value.ToString
        Form6.txtsuffix.Text = DataGridView2.SelectedCells(1).Value.ToString
        Form6.txtoperation.Text = DataGridView2.SelectedCells(2).Value.ToString
        Form6.rtbjobname.Text = DataGridView2.SelectedCells(5).Value.ToString
        Form6.txtjobqty.Text = DataGridView2.SelectedCells(6).Value.ToString
        Form6.PictureBox4.Image = writer1.Write(txtoper_num.Text)
        Form6.txtemployee.Text = lblnamet2.Text
        Form6.PictureBox5.Image = writer1.Write(lblshiftt2.Text)
        Form6.txtshift.Text = lblshiftt2.Text
        Form6.txtempnum.Text = txtoper_num.Text
        'Dim setuptxt As String = "SETUP"
        'Dim entertxt As String = "RUN"
        'Dim txtyes As String = "Y"
        'Dim txtno As String = "N"
        'Form6.PictureBox6.Image = writer1.Write("SETUP").ToString.ToUpper()
        'Form6.PictureBox6.Image = writer1.Write(setuptxt.ToUpper)
        'Form6.PictureBox7.Image = writer1.Write(entertxt.ToUpper)
        'Form6.PictureBox8.Image = writer1.Write(txtyes.ToUpper)
        'Form6.PictureBox9.Image = writer1.Write(txtno.ToUpper)

        ' Dim data As String = DataGridView1.Rows(1).Cells(0).Value.ToString()
        'Dim data1 As String = DataGridView2.SelectedCells(1).Value.ToString
        'Form6.PictureBox1.Image = writer.Write(DataGridView2.SelectedCells(1).Value.ToString)
        'Form6.PictureBox2.Image = writer.Write(DataGridView2.SelectedCells(2).Value.ToString)
        'Form6.PictureBox3.Image = writer.Write(DataGridView2.SelectedCells(3).Value.ToString)
        'Form6.txtjob.Text = DataGridView2.SelectedCells(1).Value.ToString
        'Form6.txtsuffix.Text = DataGridView2.SelectedCells(2).Value.ToString
        'Form6.txtoperation.Text = DataGridView2.SelectedCells(3).Value.ToString
        'Form6.rtbjobname.Text = DataGridView2.SelectedCells(6).Value.ToString
        'Form6.txtjobqty.Text = DataGridView2.SelectedCells(7).Value.ToString
        'Form6.PictureBox4.Image = writer1.Write(lblnamet2.Text)
        'Form6.txtemployee.Text = lblnamet2.Text
        'Form6.PictureBox5.Image = writer.Write(lblshiftt2.Text)
        'Form6.txtshift.Text = lblshiftt2.Text

        Form6.Refresh()
        Form6.Show()
        'PictureBox5.Image = writer.Write(data1)
        'PictureBox6.Image = writer.Write(DataGridView2.SelectedCells(2).Value.ToString)
        'PictureBox7.Image = writer.Write(DataGridView2.SelectedCells(3).Value.ToString)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        txtjoboperation.Clear()
        txtjoboperationt3.Clear()

        txtjoborder.Text = DataGridView2.SelectedCells(0).Value.ToString
        txtjobsuffix.Text = DataGridView2.SelectedCells(1).Value.ToString
        txtjoboperation.Text = DataGridView2.SelectedCells(2).Value.ToString
        txtjobordert3.Text = DataGridView2.SelectedCells(0).Value.ToString
        txtjobsuffixt3.Text = DataGridView2.SelectedCells(1).Value.ToString
        txtjoboperationt3.Text = DataGridView2.SelectedCells(2).Value.ToString
        'txtjobordert4.Text = DataGridView2.SelectedCells(0).Value.ToString
        'txtsuffixt4.Text = DataGridView2.SelectedCells(1).Value.ToString
        'txtoperationt4.Text = DataGridView2.SelectedCells(2).Value.ToString
        TabControl1.SelectedTab = TabPage1
        disablecheckbox()
        disablecheckbox1()

    End Sub

    'Private Sub txtoperationt4_TextChanged(sender As Object, e As EventArgs)
    '    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")

    '    Dim cmdviewhdr As New SqlCommand("Select * from IT_KPI_DTHeader where job = @job And Suffix = @suffix And shift = @shift And CONVERT(Date, DTDate, 120) = @date")
    '    cmdviewhdr.Parameters.AddWithValue("@job", txtjobordert4.Text)
    '    cmdviewhdr.Parameters.AddWithValue("@suffix", txtsuffixt4.Text)
    '    cmdviewhdr.Parameters.AddWithValue("@job", lblshiftt4.Text)
    '    cmdviewhdr.Parameters.AddWithValue("@date", DateTimePicker3.Value.Date)

    '    Dim cmd As New SqlCommand("Select jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
    '     from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
    '     inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job =@joborder and jrtrg.suffix=@jobsuffix and jrtrg.oper_num=@joboperation", con)

    '    cmd.Parameters.AddWithValue("@joborder", txtjobordert4.Text)
    '    cmd.Parameters.AddWithValue("@jobsuffix", txtsuffixt4.Text)
    '    cmd.Parameters.AddWithValue("@joboperation", txtoperationt4.Text)

    '    Dim cmd1 As New SqlCommand("select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description, Uf_Resrc_Section as section
    '    from jobroute jrt
    '    inner join wc on wc.wc=jrt.wc 
    '    inner join jrtresourcegroup jg on jrt.job=jg.job and jrt. suffix=jg.suffix and jrt.oper_num=jg.oper_num
    '    inner join RGRPMBR000 rg on rg.RGID=jg.rgid
    '    inner join RESRC000 re on re.RESID=rg.RESID
    '    where jrt.job =@joborder
    '    and jrt.suffix=@jobsuffix
    '    and jrt.oper_num=@joboperation", con)

    '    'BACKUP
    '    'Dim cmd1 As New SqlCommand(" select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description
    '    'from jobroute jrt
    '    'inner join wc on wc.wc=jrt.wc 
    '    'where jrt.job =@joborder
    '    'and jrt.suffix=@jobsuffix
    '    'and jrt.oper_num=@joboperation ", con)

    '    cmd1.Parameters.AddWithValue("@joborder", txtjobordert4.Text)
    '    cmd1.Parameters.AddWithValue("@jobsuffix", txtsuffixt4.Text)
    '    cmd1.Parameters.AddWithValue("@joboperation", txtoperationt4.Text)

    '    'Dim tablecmd As SqlCommand = New SqlCommand("SELECT Seq, StartTime, EndTime, DTHrs, Category, CauseCode, NoteExistsFlag, Notes  from IT_KPI_DTDetail WHERE RefRowPointer = @refrowpointer", con)
    '    Dim tablecmd As SqlCommand = New SqlCommand("SELECT dtdetail.Seq, dtdetail.StartTime, dtdetail.EndTime,CAST(DATEDIFF(MINUTE, dtdetail.StartTime, dtdetail.EndTime) / 60.0 AS DECIMAL(13, 9)) AS DTHrs, dtdetail.Category, dtdetail.CauseCode, dtcause.RootCause , Notes  from IT_KPI_DTDetail dtdetail
    '        LEFT JOIN IT_KPI_DTCausemaint dtcause on DTDETAIL.CauseCode = dtcause.CauseCode
    '        WHERE dtdetail.RefRowPointer = @refrowpointer", con)
    '    'tablecmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)



    '    Dim cmdrowpointer As New SqlCommand("select * from IT_KPI_DTHeader where job = @job and Suffix = @suffix AND OperNum=@opernum AND shift = @shift AND CONVERT(date, DTDate, 120) = @downtime", con)
    '    cmdrowpointer.Parameters.AddWithValue("@job", txtjobordert4.Text)
    '    cmdrowpointer.Parameters.AddWithValue("@suffix", txtsuffixt4.Text)
    '    cmdrowpointer.Parameters.AddWithValue("@shift", cmbshiftt4.Text)
    '    cmdrowpointer.Parameters.AddWithValue("@opernum", txtoperationt4.Text)
    '    cmdrowpointer.Parameters.AddWithValue("@downtime", DateTimePicker3.Value.Date)


    '    DataGridView4.Controls.Add(dtp)
    '    dtp.Format = DateTimePickerFormat.Custom
    '    dtp.Visible = False
    '    AddHandler dtp.TextChanged, AddressOf dtp_textchange


    '    Try
    '        con.Open()
    '        Dim reader1 As SqlDataReader = cmd1.ExecuteReader

    '        If reader1.HasRows Then
    '            While reader1.Read
    '                txtwct4.Text = reader1("wc").ToString
    '                txtwcdesct4.Text = reader1("description").ToString
    '                txtsectiont4.Text = reader1("section").ToString
    '            End While
    '            reader1.Close()
    '            Dim reader As SqlDataReader = cmd.ExecuteReader
    '            If reader.HasRows Then
    '                While reader.Read
    '                    txtresourcet4.Text = reader("RESID").ToString
    '                    'txtresourcet4.Text = reader("DESCR").ToString
    '                End While
    '                reader.Close()
    '                Dim readcmdrowpointer As SqlDataReader = cmdrowpointer.ExecuteReader
    '                If readcmdrowpointer.HasRows Then
    '                    Button2.Enabled = False
    '                    btnaddseq.Enabled = True
    '                    While readcmdrowpointer.Read
    '                        txtrowpointer.Text = readcmdrowpointer("RowPointer").ToString
    '                    End While
    '                    tablecmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)

    '                    Dim a As New SqlDataAdapter(tablecmd)
    '                    Dim dt As New DataTable
    '                    readcmdrowpointer.Close()
    '                    a.Fill(dt)
    '                    DataGridView4.DataSource = dt
    '                    AutofitColumns(DataGridView4)
    '                    disablecolumn()
    '                    'DataGridView1.Columns("DTHrs").ReadOnly = True
    '                Else
    '                    DataGridView4.DataSource = dt
    '                    Button2.Enabled = True
    '                    btnaddseq.Enabled = False
    '                End If
    '            End If
    '        Else
    '            DataGridView4.DataSource = dt
    '            txtcategory.Clear()
    '            txtrowpointer.Clear()
    '            txtwct4.Clear()
    '            txtwcdesct4.Clear()
    '            txtsectiont4.Clear()
    '            txtresourcet4.Clear()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    Finally
    '        con.Close()
    '    End Try
    'End Sub

    Private Sub Button3_Click_2(sender As Object, e As EventArgs) Handles Button3.Click

        Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
            con.Open()
            Dim updateCommandText As String = "UPDATE Ptag_Line SET Comment = @comment WHERE Tag_num=@tagnum AND Site = @site AND Job=@job AND Suffix=@suffix AND Oper_num=@opernum AND Name=@name"
            Dim cmd As New SqlCommand(updateCommandText, con)

            ' Add parameters to the SqlCommand
            cmd.Parameters.Add(New SqlParameter("@comment", SqlDbType.NVarChar))

            cmd.Parameters.Add(New SqlParameter("@tagnum", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@site", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@job", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@suffix", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@opernum", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@name", SqlDbType.NVarChar))

            Try
                ' Iterate through the DataGridView rows
                For Each row As DataGridViewRow In DataGridView1.Rows
                    ' Check if the row is not the header row
                    If Not row.IsNewRow Then
                        ' Set parameter values based on the data in the current row
                        cmd.Parameters("@comment").Value = row.Cells("Comment").Value

                        cmd.Parameters("@tagnum").Value = row.Cells("Tag #").Value
                        cmd.Parameters("@site").Value = TextBox3.Text
                        cmd.Parameters("@job").Value = txtjoborder.Text
                        cmd.Parameters("@suffix").Value = txtjobsuffix.Text
                        cmd.Parameters("@opernum").Value = txtjoboperation.Text
                        cmd.Parameters("@name").Value = txtoper_name.Text
                        ' Execute the update command for the current row
                        cmd.ExecuteNonQuery()
                    End If
                Next
                MsgBox("Update Successfully")
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try

        End Using

    End Sub

    Private Sub txtjoboperation_VisibleChanged(sender As Object, e As EventArgs) Handles txtjoboperation.VisibleChanged

    End Sub

    'Private Sub btnupdatet4_Click(sender As Object, e As EventArgs)
    '    Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '        Try
    '            con.Open()
    '            'Dim updateCommandText As String = "UPDATE IT_KPI_DTDetail SET StartTime=@starttime, EndTime=@endtime, Category=@category, CauseCode=@causecode, RecordDate=@recorddate, Notes=@notes WHERE RefRowPointer=@refrowpointer AND Seq=@seq"
    '            Dim updateCommandText As String = "UPDATE IT_KPI_DTDetail SET StartTime=@starttime, EndTime=@endtime, DTHrs=@dthrs, Category=@category, CauseCode=@causecode, RecordDate=@recorddate, CreatedBy=@createdby, Notes=@notes WHERE RefRowPointer=@refrowpointer AND Seq=@seq"
    '            Dim tablecmd As SqlCommand = New SqlCommand("SELECT dtdetail.Seq, dtdetail.StartTime, dtdetail.EndTime,CAST(DATEDIFF(MINUTE, dtdetail.StartTime, dtdetail.EndTime) / 60.0 AS DECIMAL(13, 9)) AS DTHrs, dtdetail.Category, dtdetail.CauseCode, dtcause.RootCause , Notes  from IT_KPI_DTDetail dtdetail
    '            LEFT JOIN IT_KPI_DTCausemaint dtcause on DTDETAIL.CauseCode = dtcause.CauseCode
    '            WHERE dtdetail.RefRowPointer = @refrowpointer", con)
    '            tablecmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)
    '            Dim a As New SqlDataAdapter(tablecmd)
    '            Dim dt As New DataTable

    '            Dim cmd As New SqlCommand(updateCommandText, con)

    '            cmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)
    '            cmd.Parameters.AddWithValue("@seq", txtseq.Text)
    '            cmd.Parameters.AddWithValue("@starttime", dtpstart.Value)
    '            cmd.Parameters.AddWithValue("@endtime", dtpend.Value)
    '            cmd.Parameters.AddWithValue("@dthrs", txtdthrs.Text)
    '            cmd.Parameters.AddWithValue("@category", cmbcategory.Text)
    '            cmd.Parameters.AddWithValue("@causecode", cmbcause.Text)
    '            cmd.Parameters.AddWithValue("@recorddate", Today())
    '            cmd.Parameters.AddWithValue("@createdby", lblnamet4.Text)
    '            cmd.Parameters.AddWithValue("@notes", txtnotes.Text)

    '            cmd.ExecuteNonQuery()

    '            a.Fill(dt)
    '            DataGridView4.DataSource = dt
    '            MsgBox("Update Successfully")
    '            txtseq.Clear()
    '            txtdthrs.Clear()
    '            txtnotes.Clear()
    '            txtrootcause.Clear()
    '            cmbcategory.Text = ""
    '            cmbcause.Text = ""
    '            dtpstart.Value = DateTime.Now
    '            dtpend.Value = DateTime.Now

    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try

    '    End Using
    'End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs)


    '    Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '        con.Open()

    '        Dim tablecmd As SqlCommand = New SqlCommand("SELECT dtdetail.Seq, dtdetail.StartTime, dtdetail.EndTime,CAST(DATEDIFF(MINUTE, dtdetail.StartTime, dtdetail.EndTime) / 60.0 AS DECIMAL(13, 9)) AS DTHrs, dtdetail.Category, dtdetail.CauseCode, dtcause.RootCause , Notes  from IT_KPI_DTDetail dtdetail
    '        LEFT JOIN IT_KPI_DTCausemaint dtcause on DTDETAIL.CauseCode = dtcause.CauseCode
    '        WHERE dtdetail.RefRowPointer = @refrowpointer", con)

    '        Dim cmdrowpointer As New SqlCommand("select * from IT_KPI_DTHeader where job = @job and Suffix = @suffix AND OperNum=@opernum AND shift = @shift AND CONVERT(date, DTDate, 120) = @downtime", con)
    '        cmdrowpointer.Parameters.AddWithValue("@job", txtjobordert4.Text)
    '        cmdrowpointer.Parameters.AddWithValue("@suffix", txtsuffixt4.Text)
    '        cmdrowpointer.Parameters.AddWithValue("@shift", cmbshiftt4.Text)
    '        cmdrowpointer.Parameters.AddWithValue("@opernum", txtoperationt4.Text)
    '        cmdrowpointer.Parameters.AddWithValue("@downtime", DateTimePicker3.Value.Date)

    '        Dim currentDateTime As DateTime = DateTime.Now
    '        Dim month As String = currentDateTime.ToString("MM") ' Month in two digits
    '        Dim day As String = currentDateTime.ToString("dd") ' Day in two digits
    '        Dim year As String = currentDateTime.ToString("yyyy") ' Year in four digits
    '        Dim time As String = currentDateTime.ToString("HHmm") ' Time in 24-hour format
    '        Dim amPm As String = If(currentDateTime.Hour < 12, "a", "b") ' AM (a) or PM (b)
    '        ' You need to maintain an incremental value and increment it as needed.
    '        ' Dim incrementalValue As String = getlatestrowpointnumber() ' Replace with your actual incremental value

    '        'Dim rowpointer As String = $"{month}{day}{year}-{year}-{time}-{amPm}000-{incrementalValue}"

    '        Dim updatehdr As String = "INSERT INTO IT_KPI_DTHeader (
    '        Section, 
    '        MachResc, 
    '        DTDate, 
    '        Shift,
    '        Job,
    '        Suffix,
    '        OperNum)
    '             values(
    '                 @section,
    '                 @machresc,
    '                 @dtdate,
    '                 @shift,
    '                 @job,
    '                 @suffix,
    '                 @opernum)"

    '        '@job,
    '        '@suffix,
    '        '@noteexistsflag,
    '        '@recorddate,
    '        '@rowpointer,
    '        '@createdby,
    '        '@updatedby,
    '        '@createdate,
    '        '@inworkflow,
    '        '@opernum)"

    '        ', 
    '        'Job,
    '        'Suffix,
    '        'NoteExistsFlag,
    '        'RecordDate,
    '        'rowpointer,
    '        'CreatedBy,
    '        'UpdatedBy,
    '        'CreateDate,
    '        'InWorkflow,
    '        'OperNum

    '        Dim cmd As New SqlCommand(updatehdr, con)

    '        Try
    '            cmd.Parameters.AddWithValue("@section", txtsectiont4.Text)
    '            cmd.Parameters.AddWithValue("@machresc", txtwcdesct4.Text)
    '            cmd.Parameters.AddWithValue("@dtdate", DateTimePicker3.Value)
    '            cmd.Parameters.AddWithValue("@shift", lblshiftt4.Text)
    '            cmd.Parameters.AddWithValue("@job", txtjobordert4.Text)
    '            cmd.Parameters.AddWithValue("@suffix", txtsuffixt4.Text)
    '            'cmd.Parameters.AddWithValue("@noteexistsflag", 0)
    '            'cmd.Parameters.AddWithValue("@recorddate", Today())
    '            'cmd.Parameters.AddWithValue("@rowpointer", rowpointer)
    '            'cmd.Parameters.AddWithValue("@createdby", lblnamet4.Text)
    '            'cmd.Parameters.AddWithValue("@updatedby", lblnamet4.Text)
    '            'cmd.Parameters.AddWithValue("@createdate", Today())
    '            'cmd.Parameters.AddWithValue("@inworkflow", 0)
    '            cmd.Parameters.AddWithValue("@opernum", txtoperationt4.Text)

    '            cmd.ExecuteNonQuery()
    '            MessageBox.Show("Saved Successfully")
    '            btnaddseq.Enabled = True
    '            Dim readcmdrowpointer As SqlDataReader = cmdrowpointer.ExecuteReader
    '            If readcmdrowpointer.HasRows Then
    '                Button2.Enabled = False
    '                While readcmdrowpointer.Read
    '                    txtrowpointer.Text = readcmdrowpointer("RowPointer").ToString
    '                End While
    '                tablecmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)

    '                Dim a As New SqlDataAdapter(tablecmd)
    '                Dim dt As New DataTable
    '                readcmdrowpointer.Close()
    '                a.Fill(dt)
    '                DataGridView4.DataSource = dt
    '                AutofitColumns(DataGridView4)
    '                disablecolumn()
    '                'DataGridView1.Columns("DTHrs").ReadOnly = True
    '            Else
    '                DataGridView4.DataSource = dt
    '                Button2.Enabled = True
    '            End If

    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        Finally
    '            con.Close()
    '        End Try

    '    End Using

    'End Sub

    'Private Sub Button9_Click_1(sender As Object, e As EventArgs)
    '    Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")

    '        Try
    '            con.Open()

    '            Dim tablecmd As SqlCommand = New SqlCommand("SELECT dtdetail.Seq, dtdetail.StartTime, dtdetail.EndTime,CAST(DATEDIFF(MINUTE, dtdetail.StartTime, dtdetail.EndTime) / 60.0 AS DECIMAL(13, 9)) AS DTHrs, dtdetail.Category, dtdetail.CauseCode, dtcause.RootCause , Notes  from IT_KPI_DTDetail dtdetail
    '            LEFT JOIN IT_KPI_DTCausemaint dtcause on DTDETAIL.CauseCode = dtcause.CauseCode
    '            WHERE dtdetail.RefRowPointer = @refrowpointer", con)
    '            tablecmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)
    '            Dim a As New SqlDataAdapter(tablecmd)
    '            Dim dt As New DataTable

    '            Dim insertcmd As New SqlCommand("INSERT INTO IT_KPI_DTDetail (RefRowPointer,Seq) values (@refrowpointer,@seq)", con)
    '            Dim latestseq As Integer = getlatestsequence()

    '            insertcmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)
    '            latestseq += 1
    '            insertcmd.Parameters.AddWithValue("@seq", latestseq)

    '            insertcmd.ExecuteNonQuery()
    '            a.Fill(dt)
    '            DataGridView4.DataSource = dt
    '            MsgBox("Sequence add successfully")
    '            txtseq.Text = latestseq
    '            disablecolumn()
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        Finally
    '            con.Close()
    '        End Try
    '    End Using
    'End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim old As Long
        If Environment.Is64BitOperatingSystem Then
            If Wow64DisableWow64FsRedirection(old) Then
                Process.Start(osk)
                Wow64EnableWow64FsRedirection(old)
            End If
        Else
            Process.Start(osk)
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim old As Long
        If Environment.Is64BitOperatingSystem Then
            If Wow64DisableWow64FsRedirection(old) Then
                Process.Start(osk)
                Wow64EnableWow64FsRedirection(old)
            End If
        Else
            Process.Start(osk)
        End If
    End Sub


    'Private Sub dtp_textchange(ByVal sender As Object, ByVal e As EventArgs)
    '    DataGridView4.CurrentCell.Value = dtp.Text.ToString
    'End Sub
    'Private Sub dtp_ValueChanged(sender As Object, e As EventArgs)
    '    If DataGridView4.CurrentCell IsNot Nothing AndAlso DataGridView4.CurrentCell.Value IsNot DBNull.Value Then
    '        ' Assuming the column type is DateTime
    '        If TypeOf DataGridView4.CurrentCell.Value Is DateTime Then
    '            DataGridView4.CurrentCell.Value = dtp.Value
    '        ElseIf TypeOf DataGridView4.CurrentCell.Value Is String Then
    '            ' If the column is stored as a string, parse and update
    '            DataGridView4.CurrentCell.Value = dtp.Value.ToString("yyyy-MM-dd HH:mm")
    '        End If
    '    End If
    'End Sub

    'Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs)

    '    txtseq.Text = DataGridView4.SelectedCells(0).Value.ToString

    '    'If DataGridView4.SelectedCells(1).Value Is Nothing Then
    '    '    dtpstart.Value = DateTime.Now
    '    'Else
    '    '    dtpstart.Value = DataGridView4.SelectedCells(1).Value.ToString
    '    'End If
    '    'If DataGridView4.SelectedCells(2).Value Is Nothing Then
    '    '    dtpend.Value = DateTime.Now
    '    'Else
    '    '    dtpend.Value = DataGridView4.SelectedCells(2).Value.ToString
    '    'End If
    '    '##############################################################################
    '    If DataGridView4.SelectedCells(1).Value Is Nothing OrElse String.IsNullOrEmpty(DataGridView4.SelectedCells(1).Value.ToString()) Then
    '        ' If the value is Nothing or an empty string, set a default value (e.g., DateTime.Now)
    '        dtpstart.Value = DateTime.Now
    '    Else
    '        ' If the value is not Nothing or an empty string, try to parse it to a Date
    '        If DateTime.TryParse(DataGridView4.SelectedCells(1).Value.ToString(), dtpstart.Value) Then
    '            dtpstart.Value = DataGridView4.SelectedCells(1).Value.ToString
    '        Else
    '            ' Handle the case where the date string is not valid
    '        End If
    '    End If

    '    If DataGridView4.SelectedCells(2).Value Is Nothing OrElse String.IsNullOrEmpty(DataGridView4.SelectedCells(2).Value.ToString()) Then
    '        ' If the value is Nothing or an empty string, set a default value (e.g., DateTime.Now)
    '        dtpend.Value = DateTime.Now
    '    Else
    '        ' If the value is not Nothing or an empty string, try to parse it to a Date
    '        If DateTime.TryParse(DataGridView4.SelectedCells(2).Value.ToString(), dtpend.Value) Then
    '            dtpend.Value = DataGridView4.SelectedCells(2).Value.ToString
    '        Else
    '            ' Handle the case where the date string is not valid
    '        End If
    '    End If

    '    txtdthrs.Text = DataGridView4.SelectedCells(3).Value.ToString
    '    cmbcategory.Text = DataGridView4.SelectedCells(4).Value.ToString
    '    cmbcause.Text = DataGridView4.SelectedCells(5).Value.ToString
    '    txtrootcause.Text = DataGridView4.SelectedCells(6).Value.ToString
    '    txtnotes.Text = DataGridView4.SelectedCells(7).Value.ToString

    'End Sub

    'Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs)
    '    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")

    '    Dim tablecmd As SqlCommand = New SqlCommand("SELECT dtdetail.Seq, dtdetail.StartTime, dtdetail.EndTime,CAST(DATEDIFF(MINUTE, dtdetail.StartTime, dtdetail.EndTime) / 60.0 AS DECIMAL(13, 9)) AS DTHrs, dtdetail.Category, dtdetail.CauseCode, dtcause.RootCause , Notes  from IT_KPI_DTDetail dtdetail
    '        LEFT JOIN IT_KPI_DTCausemaint dtcause on DTDETAIL.CauseCode = dtcause.CauseCode
    '        WHERE dtdetail.RefRowPointer = @refrowpointer", con)
    '    Try
    '        con.Open()
    '        Dim cmd As New SqlCommand("select * from IT_KPI_DTHeader where job = @job and Suffix = @suffix AND shift = @shift AND CONVERT(date, DTDate, 120) = @downtime", con)
    '        cmd.Parameters.AddWithValue("@job", txtjobordert4.Text)
    '        cmd.Parameters.AddWithValue("@suffix", txtsuffixt4.Text)
    '        cmd.Parameters.AddWithValue("@shift", cmbshiftt4.Text)
    '        cmd.Parameters.AddWithValue("@downtime", DateTimePicker3.Value.Date)

    '        Dim reader1 As SqlDataReader = cmd.ExecuteReader

    '        If reader1.HasRows Then
    '            Button2.Enabled = False
    '            btnaddseq.Enabled = True
    '            While reader1.Read
    '                txtrowpointer.Text = reader1("RowPointer").ToString
    '            End While
    '            tablecmd.Parameters.AddWithValue("@refrowpointer", txtrowpointer.Text)
    '            Dim a As New SqlDataAdapter(tablecmd)
    '            Dim dt As New DataTable
    '            reader1.Close()
    '            a.Fill(dt)
    '            DataGridView4.DataSource = dt
    '            AutofitColumns(DataGridView4)
    '            disablecolumn()
    '        Else
    '            Button2.Enabled = True
    '            btnaddseq.Enabled = False
    '        End If
    '        'MsgBox(DateTimePicker3.Value)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    Finally
    '        con.Close()
    '    End Try

    'End Sub

    'Private Sub DataGridView4_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
    '    If e.Context = DataGridViewDataErrorContexts.Commit Then
    '        MsgBox("Invalid value in the DataGridViewComboBoxCell. Please select a valid item from the list.")
    '    End If

    '    ' You can also choose to suppress the exception:
    '    e.ThrowException = False
    'End Sub

    Public Class causeitem

    End Class
    Private Sub DataGridView4_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        'Dim con1 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")

        'If DataGridView4.SelectedRows.Count > 0 Then

        '    Dim selectedRow As DataGridViewRow = DataGridView4.SelectedRows(0)
        '    Dim categoryValue As Object = selectedRow.Cells("Category").Value

        '    Dim causecodelist As New SqlCommand("select Category, CauseCode, RootCause from IT_KPI_DTCausemaint where Category=@category", con1)
        '    causecodelist.Parameters.AddWithValue("@category", categoryValue)

        '    Dim cmdcategorylist As New SqlCommand("SELECT Category ,Section from IT_KPI_SecDTCause where Section =@section", con1)
        '    cmdcategorylist.Parameters.AddWithValue("@section", txtwcdesct4.Text)

        '    If categoryValue IsNot Nothing Then
        '        ' txtcategory.Text = categoryValue
        '        Try
        '            con1.Open()
        '            Dim readcategorylist As SqlDataReader = cmdcategorylist.ExecuteReader
        '            Dim cmbcategorylist As New List(Of String)
        '            While readcategorylist.Read
        '                cmbcategorylist.Add(readcategorylist(0))
        '            End While
        '            readcategorylist.Close()
        '            Dim readcausecodelist As SqlDataReader = causecodelist.ExecuteReader
        '            Dim cmbcausecodelist As New List(Of String)
        '            Dim cmbcausecodeonly As New List(Of String)
        '            While readcausecodelist.Read()
        '                cmbcausecodelist.Add(readcausecodelist(1))
        '            End While

        '            For Each row As DataGridViewRow In DataGridView4.SelectedRows

        '                If row.Cells("Seq").Value IsNot Nothing Then
        '                    Dim cmbcategorycell As New DataGridViewComboBoxCell
        '                    cmbcategorycell.DataSource = cmbcategorylist
        '                    row.Cells("Category") = cmbcategorycell
        '                End If

        '                If row.Cells("CauseCode").Value IsNot Nothing Then
        '                    Dim cmbcausecodecell As New DataGridViewComboBoxCell
        '                    cmbcausecodecell.DataSource = cmbcausecodelist
        '                    row.Cells("CauseCode") = cmbcausecodecell
        '                End If
        '            Next
        '        Catch ex As Exception

        '        End Try
        '    Else
        '        'MsgBox("Category value is null.")
        '    End If

        'Else
        '    'MsgBox("No rows are selected.")
        'End If


    End Sub


    Private Sub DataGridView4_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)

    End Sub

    Private Sub DataGridView4_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
        'Dim con1 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        'Dim cmdcategorylist As New SqlCommand("SELECT Category ,Section from IT_KPI_SecDTCause where Section =@section", con1)
        'cmdcategorylist.Parameters.AddWithValue("@section", txtsectiont4.Text)

        'Try
        '    con1.Open()
        '    Dim readcausecodelist As SqlDataReader = cmdcategorylist.ExecuteReader
        '    Dim categorylist As New List(Of String)
        '    While readcausecodelist.Read()
        '        categorylist.Add(readcausecodelist(0))
        '    End While

        '    For Each row As DataGridViewRow In DataGridView4.SelectedRows
        '        If row.Cells("Seq").Value IsNot Nothing Then
        '            Dim cmbcategorycell As New DataGridViewComboBoxCell

        '            cmbcategorycell.DataSource = categorylist
        '            row.Cells("Category") = cmbcategorycell
        '        End If
        '    Next
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub Button9_Click_2(sender As Object, e As EventArgs)
        Dim old As Long
        If Environment.Is64BitOperatingSystem Then
            If Wow64DisableWow64FsRedirection(old) Then
                Process.Start(osk)
                Wow64EnableWow64FsRedirection(old)
            End If
        Else
            Process.Start(osk)
        End If
    End Sub

    'Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
    '    txtjobordert4.Text = DataGridView2.SelectedCells(0).Value.ToString
    '    txtsuffixt4.Text = DataGridView2.SelectedCells(1).Value.ToString
    '    txtoperationt4.Text = DataGridView2.SelectedCells(2).Value.ToString
    '    TabControl1.SelectedTab = TabPage4
    'End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Dim cmd As New SqlCommand("delete from Ptag_Line where job = @joborder and 
            suffix = @jobsuffix and Oper_num = @jobopernum and 
            (Tag_num = @tagnum OR Tag_num > @tagnum) and 
            Emp_num = @opernum and name = @name", con)
        cmd.Parameters.AddWithValue("@joborder", txtjoborder.Text)
        cmd.Parameters.AddWithValue("@jobsuffix", txtjobsuffix.Text)
        cmd.Parameters.AddWithValue("@jobopernum", txtjoboperation.Text)
        cmd.Parameters.AddWithValue("@opernum", txtoper_num.Text)
        cmd.Parameters.AddWithValue("@name", txtoper_name.Text)


        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty, running_qty,  Emp_num, Name, Shift, Comment FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable

        Dim promptmsg As DialogResult = MessageBox.Show("Are you sure you want to delete ? All tag number below will also be deleted.", "Delete Tag?", MessageBoxButtons.YesNo)

        If promptmsg = DialogResult.Yes Then
            Try
                con.Open()


                If DataGridView1.SelectedRows.Count > 0 Then
                    Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                    cmd.Parameters.AddWithValue("@tagnum", selectedRow.Cells(1).Value)
                    cmd.ExecuteNonQuery()
                    MsgBox("Tag Deleted")
                    a.Fill(dt)
                    DataGridView1.DataSource = dt
                    AutofitColumns(DataGridView1)
                    disablecheckbox()
                    Dim latestrunningqty As Integer = getlatestrunningqty()
                    Dim tagqty As Integer = CInt(txtjobqty.Text)
                    Dim result As Integer = tagqty - latestrunningqty
                    txtremainingqty.Text = result

                Else
                    MsgBox("Please Select a tag # to delete")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        Else
            MsgBox("Cancelled")
        End If


        Try
            con.Open()
            Dim cmd_updatehdr As New SqlCommand("UPDATE Ptag_Hdr Set total_tags=@totaltags, Pallet_size=@palletqty WHERE job=@job AND suffix=@suffix AND Oper_num=@opernum", con)

            Dim latestTagNumberhdr As Integer = GetLatestTagNumberFromDatabase()

            cmd_updatehdr.Parameters.AddWithValue("@totaltags", latestTagNumberhdr)
            cmd_updatehdr.Parameters.AddWithValue("@palletqty", txtremainingqty.Text)
            cmd_updatehdr.Parameters.AddWithValue("@job", txtjoborder.Text)
            cmd_updatehdr.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
            cmd_updatehdr.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

            cmd_updatehdr.ExecuteNonQuery()
            ' MsgBox(latestTagNumberhdr)

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("debugging in updating hdr")
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Dim cmd As New SqlCommand("delete from NCFPtag_Line where job = @joborder and 
            suffix = @jobsuffix and Oper_num = @jobopernum and 
            (Tag_num = @tagnum OR Tag_num > @tagnum) and 
            Emp_num = @opernum and name = @name", con)
        cmd.Parameters.AddWithValue("@joborder", txtjobordert3.Text)
        cmd.Parameters.AddWithValue("@jobsuffix", txtjobsuffixt3.Text)
        cmd.Parameters.AddWithValue("@jobopernum", txtjoboperationt3.Text)
        cmd.Parameters.AddWithValue("@opernum", txtoper_numt3.Text)
        cmd.Parameters.AddWithValue("@name", txtoper_namet3.Text)


        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty as 'PALLET QTY', Qty_Affected AS 'QTY AFFECTED',  Emp_num as Empnum, Name, Shift, Comment FROM NCFPtag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable

        Dim promptmsg As DialogResult = MessageBox.Show("Are you sure you want to delete ? All tag number below will also be deleted.", "Delete Tag?", MessageBoxButtons.YesNo)

        If promptmsg = DialogResult.Yes Then
            Try
                con.Open()

                If DataGridView3.SelectedRows.Count > 0 Then
                    Dim selectedRow As DataGridViewRow = DataGridView3.SelectedRows(0)
                    cmd.Parameters.AddWithValue("@tagnum", selectedRow.Cells(1).Value)
                    cmd.ExecuteNonQuery()
                    MsgBox("Tags Deleted!")
                    a.Fill(dt)
                    DataGridView3.DataSource = dt
                    AutofitColumns(DataGridView3)
                    disablecheckbox()
                Else
                    MsgBox("Please Select a tag # to delete")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        Else
            MsgBox("Cancelled")
        End If

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

    End Sub

    Private Sub cmbcomment_DropDown(sender As Object, e As EventArgs) Handles cmbcomment.DropDown
        cmbcomment.Items.Clear()
        populatecomment()
    End Sub
    Private Sub cmbcomment_subcon_DropDown(sender As Object, e As EventArgs) Handles cmbcomment_subcon.DropDown
        cmbcomment.Items.Clear()
        populatecomment()
    End Sub

    Private Sub populatecomment()

        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd As SqlCommand = New SqlCommand("SELECT * from Ptag_Comments", con)

        Try
            con.Open()
            Dim sqlreader1 As SqlDataReader = cmd.ExecuteReader()
            While sqlreader1.Read()
                Dim cmt As String = sqlreader1("comments").ToString
                cmbcomment.Items.Add(cmt)
                cmbcomment_subcon.Items.Add(cmt)
            End While
            sqlreader1.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        FilterData()
    End Sub
    Private Sub FilterData()
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Try
            Dim cmd As SqlCommand = New SqlCommand("
       SELECT 
            jr.job,
            jr.suffix,
            jr.oper_num as Oper,
            js.start_date AS [Date],
            jo.item AS Item,
            itm.description,
            CAST(jo.qty_released AS INT) AS QTY,
            itm.u_m AS UM
        FROM job  jo  
        INNER JOIN jobroute jr ON jo.job = jr.job AND jo.suffix = jr.suffix
        INNER JOIN job_sch js ON jo.job = js.job AND jo.suffix = js.suffix
        INNER JOIN item itm ON itm.item = jo.item
        INNER JOIN wc ON wc.wc = jr.wc
        INNER JOIN jrtresourcegroup jrs ON jr.job = jrs.job AND jr.suffix = jrs.suffix 
                                       AND jr.oper_num = jrs.oper_num
        INNER JOIN RGRPMBR000 rg ON jrs.rgid = rg.RGID 
        INNER JOIN RESRC000 rs ON rg.RESID = rs.RESID
        WHERE jo.type = 'J' 
          AND jo.stat = 'R'
          
          AND rs.RESID = @resid
		  AND js.start_date >=  @startdate
          AND js.end_date <=  @enddate
		ORDER BY 
    CASE 
        WHEN jr.job LIKE @description THEN 1
        ELSE 2
    END,
    itm.description ASC;", con1)

            'cmd.Parameters.AddWithValue("@wc", ComboBox1.Text)AND JR.wc = @wc
            cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = DateTimePicker1.Value.Date
            cmd.Parameters.Add("@enddate", SqlDbType.DateTime).Value = DateTimePicker2.Value.AddDays(1)
            cmd.Parameters.AddWithValue("@resid", ComboBox2.Text)
            cmd.Parameters.AddWithValue("@description", "%" & TextBox1.Text & "%")
            'cmd.Parameters.AddWithValue("@job", "%" & TextBox1.Text & "%")

            Dim a As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            a.Fill(dt)

            If DataGridView2.DataSource IsNot Nothing AndAlso TypeOf DataGridView2.DataSource Is DataView Then
                DirectCast(DataGridView2.DataSource, DataView).RowFilter = ""
            End If

            DataGridView2.DataSource = dt

            ' Formatting columns
            Dim qtyColumn As DataGridViewColumn = DataGridView2.Columns(6)
            Dim dateColumn As DataGridViewColumn = DataGridView2.Columns(3)

            ' Set the format for the quantity column to display commas
            qtyColumn.DefaultCellStyle.Format = "N0"
            qtyColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dateColumn.DefaultCellStyle.Format = "MM/dd/yyyy"

            AutofitColumns(DataGridView2)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con1.Close()
        End Try

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                        from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                        inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and jrtrg.suffix = @josuffix and jrtrg.oper_num=@operationnum", con1)
        cmd1.Parameters.AddWithValue("@jonumber", txtjoborder.Text)
        cmd1.Parameters.AddWithValue("@josuffix", txtjobsuffix.Text)
        cmd1.Parameters.AddWithValue("@operationnum", txtjoboperation.Text)

        Dim cmdupdatemachine As SqlCommand = New SqlCommand("Update Ptag_hdr set Machine=@machine, Resource=@resource where job=@Job AND Suffix=@suffix AND Oper_num=@opernum", con)
        cmdupdatemachine.Parameters.AddWithValue("@job", txtjoborder.Text)
        cmdupdatemachine.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
        cmdupdatemachine.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

        Try
            con1.Open()
            Dim reader1 As SqlDataReader = cmd1.ExecuteReader
            If reader1.HasRows Then
                While reader1.Read()
                    rtbmachine.Text = reader1("RESID").ToString + " " + reader1("DESCR").ToString
                End While
                reader1.Close()
                Try
                    con.Open()
                    Dim machine As String() = rtbmachine.Text.Split(" "c)
                    cmdupdatemachine.Parameters.AddWithValue("@machine", machine(0))
                    cmdupdatemachine.Parameters.AddWithValue("@resource", rtbmachine.Text)


                    cmdupdatemachine.ExecuteNonQuery()

                    MsgBox("Updated Machine Successfully")

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    con.Close()
                End Try
            End If
        Catch ex As Exception
        Finally
            con1.Close()
        End Try



    End Sub

    'Private Sub dtpend_ValueChanged(sender As Object, e As EventArgs)

    '    Dim startTime As DateTime = dtpstart.Value
    '    Dim endTime As DateTime = dtpend.Value

    '    ' Set seconds component to 0 for both DateTime objects
    '    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, startTime.Hour, startTime.Minute, 0)
    '    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, endTime.Hour, endTime.Minute, 0)

    '    ' Calculate the difference in minutes
    '    Dim minutesDifference As Double = DateDiff(DateInterval.Minute, startTime, endTime)

    '    ' Calculate the difference in hours without rounding
    '    Dim hoursDifference As Decimal = CDec(minutesDifference / 60.0)

    '    ' Display the result with at least 6 decimal places
    '    txtdthrs.Text = hoursDifference.ToString("0.######")


    'End Sub

    'Private Sub dtpstart_ValueChanged(sender As Object, e As EventArgs)

    '    Dim startTime As DateTime = dtpstart.Value
    '    Dim endTime As DateTime = dtpend.Value

    '    ' Set seconds component to 0 for both DateTime objects
    '    startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, startTime.Hour, startTime.Minute, 0)
    '    endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, endTime.Hour, endTime.Minute, 0)

    '    ' Calculate the difference in minutes
    '    Dim minutesDifference As Double = DateDiff(DateInterval.Minute, startTime, endTime)

    '    ' Calculate the difference in hours without rounding
    '    Dim hoursDifference As Decimal = CDec(minutesDifference / 60.0)

    '    ' Display the result with at least 6 decimal places
    '    txtdthrs.Text = hoursDifference.ToString("0.######")


    'End Sub

    'Private Sub cmbcategory_DropDown(sender As Object, e As EventArgs)
    '    cmbcategory.Items.Clear()
    '    cmbcause.SelectedIndex = -1
    '    categoryitems()
    'End Sub
    'Private Sub categoryitems()
    '    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '    Dim cmd As New SqlCommand("Select Category , Section from IT_KPI_SecDTCause where Section =@section", con)
    '    cmd.Parameters.AddWithValue("@section", txtwcdesct4.Text)

    '    Try
    '        con.Open()
    '        Dim readcmd As SqlDataReader = cmd.ExecuteReader()
    '        While readcmd.Read()
    '            cmbcategory.Items.Add(readcmd("Category"))
    '        End While
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub cmbcause_DropDown(sender As Object, e As EventArgs)
    '    cmbcause.Items.Clear()
    '    causeitems()
    'End Sub

    'Private Sub causeitems()
    '    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '    Dim cmd As New SqlCommand("Select Category, CauseCode, RootCause from IT_KPI_DTCausemaint where Category=@category", con)
    '    cmd.Parameters.AddWithValue("@category", cmbcategory.Text)

    '    Try
    '        con.Open()
    '        Dim readcmd As SqlDataReader = cmd.ExecuteReader()
    '        While readcmd.Read()
    '            Dim codeandcause As String = readcmd("CauseCode") + " - " + readcmd("RootCause")
    '            cmbcause.Items.Add(codeandcause)
    '        End While
    '        readcmd.Close()
    '    Catch ex As Exception
    '    Finally
    '        con.Close()
    '    End Try
    'End Sub
    'Private Sub cmbcause_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '    Dim cmd As New SqlCommand("select category, causecode, RootCause, CauseCode + ' - ' + RootCause as codecause from IT_KPI_DTCausemaint where (CauseCode + ' - ' + RootCause) = @causecode", con)
    '    cmd.Parameters.AddWithValue("@causecode", cmbcause.Text)

    '    Try
    '        con.Open()
    '        Dim readcmd As SqlDataReader = cmd.ExecuteReader
    '        If readcmd.HasRows Then
    '            While readcmd.Read()
    '                txtrootcause.Text = readcmd("RootCause").ToString
    '                populatecausecode()
    '                cmbcause.Text = readcmd("causecode").ToString
    '            End While
    '        End If
    '        readcmd.Close()

    '    Catch ex As Exception
    '    Finally
    '        con.Close()
    '    End Try

    'End Sub

    'Private Sub populatecausecode()
    '    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '    Dim cmd As New SqlCommand("select Category, CauseCode, RootCause from IT_KPI_DTCausemaint", con)

    '    Try
    '        con.Open()
    '        Dim readcmd As SqlDataReader = cmd.ExecuteReader
    '        While readcmd.Read
    '            Dim cc As String = readcmd("CauseCode").ToString
    '            cmbcause.Items.Add(cc)
    '        End While
    '        readcmd.Close()
    '    Catch ex As Exception

    '    End Try

    'End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd1 As SqlCommand = New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                        from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                        inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@jonumber and jrtrg.suffix = @josuffix and jrtrg.oper_num=@operationnum", con1)
        cmd1.Parameters.AddWithValue("@jonumber", txtjobordert3.Text)
        cmd1.Parameters.AddWithValue("@josuffix", txtjobsuffixt3.Text)
        cmd1.Parameters.AddWithValue("@operationnum", txtjoboperationt3.Text)

        Dim cmdupdatemachine As SqlCommand = New SqlCommand("Update NCFPtag_hdr set Resource=@resource where job=@Job AND Suffix=@suffix AND Oper_num=@opernum", con)
        cmdupdatemachine.Parameters.AddWithValue("@job", txtjobordert3.Text)
        cmdupdatemachine.Parameters.AddWithValue("@suffix", txtjobsuffixt3.Text)
        cmdupdatemachine.Parameters.AddWithValue("@opernum", txtjoboperationt3.Text)

        Try
            con1.Open()
            Dim reader1 As SqlDataReader = cmd1.ExecuteReader
            If reader1.HasRows Then
                While reader1.Read()
                    rtbmachinet3.Text = reader1("RESID").ToString + " " + reader1("DESCR").ToString
                End While
                reader1.Close()
                Try
                    con.Open()
                    'Dim machine As String() = rtbmachine.Text.Split(" "c)
                    'cmdupdatemachine.Parameters.AddWithValue("@machine", machine(0))
                    cmdupdatemachine.Parameters.AddWithValue("@resource", rtbmachinet3.Text)


                    cmdupdatemachine.ExecuteNonQuery()

                    MsgBox("Updated Machine Successfully")

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    con.Close()
                End Try
            End If
        Catch ex As Exception
        Finally
            con1.Close()
        End Try

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            'Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where name=@name AND emp_num = @empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            ' Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where Name=@name and Emp_num=@empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)
            Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)

            cmd.Parameters.AddWithValue("name", txtoper_name.Text)
            cmd.Parameters.AddWithValue("empnum", txtoper_num.Text)
            cmd.Parameters.AddWithValue("job", txtjoborder.Text)
            cmd.Parameters.AddWithValue("suffix", txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("opernum", txtjoboperation.Text)

            cmd1.Parameters.AddWithValue("name", txtoper_namet3.Text)
            cmd1.Parameters.AddWithValue("empnum", txtoper_numt3.Text)
            cmd1.Parameters.AddWithValue("job", txtjobordert3.Text)
            cmd1.Parameters.AddWithValue("suffix", txtjobsuffixt3.Text)
            cmd1.Parameters.AddWithValue("opernum", txtjoboperationt3.Text)

            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
            Form1.txtpassword.Clear()
            Form1.cmbuserid.SelectedValue = 0
            Form1.Show()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub btnprintcuttags_Click(sender As Object, e As EventArgs) Handles btnprintcuttags.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim selectedrows As New List(Of Integer)

        Try
            con.Open()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value IsNot Nothing AndAlso CBool(DataGridView1.Rows(i).Cells(0).Value) = True Then
                    selectedrows.Add(i)
                End If
            Next

            Dim report6 As New CrystalReport6
            Dim user As String = "sa"
            Dim pwd As String = "pi_dc_2011"

            For Each rowIndex As Integer In selectedrows
                ' Create a new report instance for each selected row
                'Dim selectedvalue As Boolean

                report6.SetParameterValue("job", txtjoborder.Text)
                report6.SetParameterValue("suffix", txtjobsuffix.Text)
                report6.SetParameterValue("oper_num", txtjoboperation.Text)
                'report4.SetParameterValue("job_operation", txtjoboperation.Text)
                'report4.SetParameterValue("Select", selectedvalue)
                'report4.SetParameterValue("total_tagnum", DataGridView1.Rows.Count)
                report6.SetParameterValue("operator", txtoper_num.Text)
                'If txtheight.Text IsNothing Then
                '    report6.SetParameterValue("Height", txtheight.Text)
                'Else

                '    report6.SetParameterValue("Height", txtheight.Text + """")
                'End If

                If txtheight.Text IsNot Nothing AndAlso txtheight.Text <> "" Then
                    report6.SetParameterValue("Height", txtheight.Text & """")
                Else
                    report6.SetParameterValue("Height", txtheight.Text)
                End If


                report6.SetParameterValue("metalic", cmbmetal.Text.ToUpper)
                report6.SetParameterValue("Lot", txtlot.Text.ToUpper)
                report6.SetParameterValue("bsheetsize", txtbigsize.Text)
                report6.SetParameterValue("cut_size", txtcutsize.Text)
                report6.SetParameterValue("stock_code", txtstockcode.Text)

            Next

            report6.SetDatabaseLogon(user, pwd)
            Form5.CrystalReportViewer1.ReportSource = report6
            Form5.CrystalReportViewer1.Refresh()
            Form5.CrystalReportViewer1.Zoom(50)
            Form5.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles txtopernum_subcon.TextChanged

        Dim con As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd_chkhdr As SqlCommand = New SqlCommand("SELECT * FROM Ptag_Hdr WHERE Job=@jonumber And Suffix=@josuffix And Oper_num=@operationnum", con)

        cmd_chkhdr.Parameters.AddWithValue("@jonumber", txtjob_subcon.Text)
        cmd_chkhdr.Parameters.AddWithValue("@josuffix", txtsuffix_subcon.Text)
        cmd_chkhdr.Parameters.AddWithValue("@operationnum", txtopernum_subcon.text)

        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty, running_qty,  Emp_num, Name, Shift, Comment FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjob_subcon.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtsuffix_subcon.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtopernum_subcon.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable
        a.Fill(dt)
        DataGridView_subcon.DataSource = dt
        AutofitColumns(DataGridView_subcon)
        disablecheckbox_subcon()

        If txttosite_subcon.Text = "FP-SP" Then
            Dim con_fpsp As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=FP-SP_App;User ID=sa;Password=pi_dc_2011")
            Dim con_pisp As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
            Dim cmd As New SqlCommand("select 
		                    job.type,
		                    job.job,
		                    job.suffix,
		                    job.job_date,
		                    job.cust_num,
		                    job.item,
		                    job.qty_released,
		                    job.ref_job,
		                    job.ref_suf,
		                    job.ref_oper,
		                    job.ref_seq,
		                    job.description
                    INTO #job3 from job
                    where type = 'J'

                    Select
		                    job.type,
		                    job.job,
		                    job.suffix,
		                    job.item,
		                    job.description
                    INTO #job4 from job
                    where type = 'J' and suffix = 0

                    SELECT 
		                    job4.job,
		                    job4.item,
		                    job4.description,
		                    'FPC' + ' - ' + wc.description as nextoperation
		                    FROM trnitem trn

                    INNER JOIN #job3 job3 ON 
		                    job3.job = trn.to_ref_num AND
		                    job3.suffix = trn.to_ref_line_suf

                    INNER JOIN jobroute jroute on
		                    jroute.job = job3.ref_job AND
		                    jroute.suffix = job3.ref_suf AND
		                    jroute.oper_num = job3.ref_oper

                    INNER JOIN wc on
		                    wc.wc = jroute.wc

                    INNER JOIN #job4 job4 on
		                    job4.job = job3.job

                    WHERE trn.trn_num = @trnnum AND trn.trn_line = @trnline AND jroute.oper_num=@opernum", con_fpsp)

            cmd.Parameters.AddWithValue("@trnnum", txttrnnum_subcon.Text)
            cmd.Parameters.AddWithValue("@trnline", txttrnline_subcon.Text)
            cmd.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

            Dim cmd_machine As New SqlCommand("SELECT jrtrg.job, jrtrg.suffix,jrtrg.oper_num,res.RESID,res.DESCR
                        from jrtresourcegroup jrtrg inner join RGRPMBR000 rg on jrtrg.rgid=rg.rgid
                        inner join RESRC000 res on  rg.RESID=res.RESID where jrtrg.job=@job and jrtrg.suffix = @suffix and jrtrg.oper_num=@opernum", con_pisp)

            cmd_machine.Parameters.AddWithValue("@job", txtjob_subcon.Text)
            cmd_machine.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
            cmd_machine.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

            Dim cmd_wc As New SqlCommand("select top 1 jrt.job,	jrt.suffix	,jrt.oper_num,	jrt.wc ,wc.description
                         from jobroute jrt
                         inner join wc on wc.wc=jrt.wc 
                         where jrt.job =@job
                         and jrt.suffix=@suffix
                         and jrt.oper_num=@opernum
                         order by  jrt.oper_num ASC", con_pisp)

            cmd_wc.Parameters.AddWithValue("@job", txtjob_subcon.Text)
            cmd_wc.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
            cmd_wc.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

            Dim cmd_jobqty As New SqlCommand("select j.job,j.suffix, j.item,j.description ,itm.Uf_itemdesc_ext, j.qty_released as job_qty, itm.u_m
                        from job j inner join item itm on itm.item=j.item
                        where j.job=@job and j.suffix=@suffix", con_pisp)

            cmd_jobqty.Parameters.AddWithValue("@job", txtjob_subcon.Text)
            cmd_jobqty.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)

            Try
                con.Open()


                Dim read_cmd_chkhdr As SqlDataReader = cmd_chkhdr.ExecuteReader
                If read_cmd_chkhdr.HasRows Then
                    btnsave_subcon.Enabled = False
                    btngenerate_subcon.Enabled = True
                    While read_cmd_chkhdr.Read
                        txtjob_subcon_source.Text = read_cmd_chkhdr("src_joborder").ToString
                        txtjobcode_subcon.Text = read_cmd_chkhdr("src_jobcode").ToString
                        txtjobname_subcon.Text = read_cmd_chkhdr("src_jobname").ToString
                        txtwc_subcon.Text = read_cmd_chkhdr("wc").ToString + " " + read_cmd_chkhdr("wc_desc").ToString
                        txtnextop_subcon.Text = read_cmd_chkhdr("next_wc").ToString
                        rtbmachine_subcon.Text = read_cmd_chkhdr("Resource").ToString
                        txtjoname_subcon.Text = read_cmd_chkhdr("item").ToString
                        txtjodesc_subcon.Text = read_cmd_chkhdr("item_desc").ToString
                        txtjobqty_subcon.Text = read_cmd_chkhdr("job_qty").ToString
                        txtremainingqty_subcon.Text = read_cmd_chkhdr("Pallet_size").ToString
                        txttotaltags_subcon.Text = read_cmd_chkhdr("Total_tags").ToString
                    End While
                    'read_cmd_chkhdr.Close()
                Else
                    read_cmd_chkhdr.Close()
                    btnsave_subcon.Enabled = True
                    btngenerate_subcon.Enabled = False
                    Try
                        con_fpsp.Open()
                        con_pisp.Open()
                        Dim read_cmd As SqlDataReader = cmd.ExecuteReader
                        If read_cmd.HasRows Then
                            While read_cmd.Read
                                txtjob_subcon_source.Text = read_cmd("job").ToString
                                txtjobcode_subcon.Text = read_cmd("item").ToString
                                txtjobname_subcon.Text = read_cmd("description").ToString
                                txtnextop_subcon.Text = read_cmd("nextoperation").ToString
                            End While
                            read_cmd.Close()
                            Dim read_cmd_machine As SqlDataReader = cmd_machine.ExecuteReader
                            If read_cmd_machine.HasRows Then
                                While read_cmd_machine.Read

                                    rtbmachine_subcon.Text = read_cmd_machine("RESID").ToString + " " + read_cmd_machine("DESCR").ToString
                                End While
                                read_cmd_machine.Close()
                                Dim read_cmd_wc As SqlDataReader = cmd_wc.ExecuteReader
                                If read_cmd_wc.HasRows Then
                                    While read_cmd_wc.Read
                                        txtwc_subcon.Text = read_cmd_wc("wc").ToString + " - " + read_cmd_wc("description").ToString
                                    End While
                                    read_cmd_wc.Close()
                                End If
                                Dim read_cmd_jobqty As SqlDataReader = cmd_jobqty.ExecuteReader
                                If read_cmd_jobqty.HasRows Then
                                    While read_cmd_jobqty.Read
                                        txtjobqty_subcon.Text = read_cmd_jobqty("job_qty").ToString
                                        lblum2_subcon.Text = read_cmd_jobqty("u_m").ToString
                                    End While
                                    read_cmd_jobqty.Close()
                                End If
                            End If
                        End If
                    Catch ex As Exception
                    Finally
                        con_fpsp.Close()
                        con_pisp.Close()
                    End Try
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        End If

        If txtopernum_subcon.Text.Length = 0 Then
            txtjob_subcon_source.Clear()
            txtjobcode_subcon.Clear()
            txtjobname_subcon.Clear()
            txtwc_subcon.Clear()
            txtnextop_subcon.Clear()
            rtbmachine_subcon.Clear()
            txtjobqty_subcon.Clear()

        End If

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles txtsuffix_subcon.TextChanged
        Dim con_pisp As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        Dim cmd As New SqlCommand("SELECT 
                        JOB,suffix, 
                        ord_num, 
                        ord_line, 
                        transfer.to_site,				
                        job.item,
				        job.description

                FROM JOB 
                        INNER JOIN TRANSFER ON JOB.ord_num = TRANSFER.trn_num
                        WHERE JOB.JOB = @job AND
                        job.suffix = @suffix", con_pisp)

        cmd.Parameters.AddWithValue("@job", txtjob_subcon.Text)
        cmd.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)

        Try
            con_pisp.Open()

            Dim read_cmd As SqlDataReader = cmd.ExecuteReader
            While read_cmd.Read
                txttosite_subcon.Text = read_cmd("to_site").ToString
                txttrnnum_subcon.Text = read_cmd("ord_num").ToString
                txttrnline_subcon.Text = read_cmd("ord_line").ToString
                txtjoname_subcon.Text = read_cmd("item").ToString
                txtjodesc_subcon.Text = read_cmd("description").ToString
            End While
            read_cmd.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con_pisp.Close()
        End Try


        If txtsuffix_subcon.Text.Length <= 0 Then
            txttosite_subcon.Clear()
            txttrnline_subcon.Clear()
            txttrnnum_subcon.Clear()
            txtjoname_subcon.Clear()
            txtjodesc_subcon.Clear()
            txtjobcode_subcon.Clear()
            txtjobname_subcon.Clear()
            txtjob_subcon_source.Clear()
            lblum2_subcon.Text = If(cleartext() = 0, "", cleartext().ToString())
        End If

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles btnsave_subcon.Click



        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            If txtjob_subcon.Text.Length = 10 And txtsuffix_subcon.Text.Length >= 1 And txtopernum_subcon.Text.Length >= 1 Then
                Dim cmd As SqlCommand = New SqlCommand("INSERT INTO Ptag_Hdr (
                        Site,
                        Job,
                        Suffix,
                        Oper_num,
                        Wc,
                        Wc_desc,
                        Next_wc, 
                        Machine, 
                        Resource,
                        Item,
                        Item_desc,
                        U_m,
                        Job_qty,
                        Pallet_size,
                        Createdate,
                        src_joborder,
                        src_jobcode,
                        src_jobname
                        ) 
                VALUES (
                        @site,
                        @job,
                        @suffix,
                        @opernum,
                        @wc,
                        @wcdesc,
                        @nextwc,
                        @machine,
                        @resource,
                        @item,
                        @itemdesc,
                        @um,
                        @jobqty,
                        @palletsize,
                        @createdate,
                        @srcjoborder,
                        @srcjobcode,
                        @srcjobname)", con)

                Dim machine As String() = rtbmachine_subcon.Text.Split(" "c)
                Dim wc As String() = txtwc_subcon.Text.Split("-"c)
                'Dim palletqty As Integer = CInt(txtremainingqty_subcon.Text)

                cmd.Parameters.AddWithValue("@site", txtsite_subcon.Text)
                cmd.Parameters.AddWithValue("@job", txtjob_subcon.Text)
                cmd.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
                cmd.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)
                cmd.Parameters.AddWithValue("@wc", wc(0))
                cmd.Parameters.AddWithValue("@wcdesc", wc(1))
                cmd.Parameters.AddWithValue("@nextwc", txtnextop_subcon.Text)
                cmd.Parameters.AddWithValue("@machine", machine(0))
                cmd.Parameters.AddWithValue("@resource", rtbmachine_subcon.Text)
                cmd.Parameters.AddWithValue("@item", txtjoname_subcon.Text)
                cmd.Parameters.AddWithValue("@itemdesc", txtjodesc_subcon.Text)
                cmd.Parameters.AddWithValue("@um", lblum2_subcon.Text)
                cmd.Parameters.AddWithValue("@jobqty", CInt(txtjobqty_subcon.Text))
                If txtremainingqty_subcon.Text = "" Then
                    cmd.Parameters.AddWithValue("@palletsize", CInt(txtjobqty_subcon.Text))
                End If
                cmd.Parameters.AddWithValue("@createdate", DateTime.Now)
                cmd.Parameters.AddWithValue("@srcjoborder", txtjob_subcon_source.Text)
                cmd.Parameters.AddWithValue("@srcjobcode", txtjobcode_subcon.Text)
                cmd.Parameters.AddWithValue("@srcjobname", txtjobname_subcon.Text)

                cmd.ExecuteNonQuery()

                txtremainingqty_subcon.Text = txtjobqty_subcon.Text
                MessageBox.Show("Saved Successfully")
                btnsave_subcon.Enabled = False
                btngenerate_subcon.Enabled = True
                cleartext()
            Else
                MsgBox("Invalid Job")
                btnsave_subcon.Enabled = True
                btngenerate_subcon.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub btngenerate_subcon_Click(sender As Object, e As EventArgs) Handles btngenerate_subcon.Click
        btngenerate.Enabled = False

        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #], Tag_date, Tag_qty, running_qty, Emp_num, Name, Shift, Comment FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjob_subcon.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtsuffix_subcon.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtopernum_subcon.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable

        Dim tagqty As Integer = CInt(txtqtypallet_subcon.Text)
        Dim runningqty As Integer = getlatestrunningqty_subcon() + tagqty

        Try
            con.Open()

            Dim cmd As SqlCommand

            cmd = New SqlCommand("INSERT INTO Ptag_Line ([Select],
                        Site, 
                        Job, 
                        Suffix, 
                        Oper_num, 
                        Tag_num, 
                        Tag_date, 
                        Tag_qty, 
                        running_qty, 
                        Emp_num, 
                        Name, 
                        Shift, 
                        Comment) 
                VALUES (@select, 
                        @site, 
                        @job, 
                        @suffix, 
                        @opernum,
                        @tagnum, 
                        @tagdate, 
                        @tagqty, 
                        @runningqty, 
                        @empnum, 
                        @name, 
                        @shift, 
                        @comment)", con)

            cmd.Parameters.AddWithValue("@select", 0)
            cmd.Parameters.AddWithValue("@site", txtsite_subcon.Text)
            cmd.Parameters.AddWithValue("@job", txtjob_subcon.Text)
            cmd.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
            cmd.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

            ' Retrieve the latest tag number from the database
            Dim latestTagNumber As Integer = GetLatestTagNumberFromDatabase_subcon()
            ' Increment the latest tag number
            latestTagNumber += 1
            cmd.Parameters.AddWithValue("@tagnum", latestTagNumber)

            cmd.Parameters.AddWithValue("@tagdate", DateTime.Now)
            cmd.Parameters.AddWithValue("@tagqty", txtqtypallet_subcon.Text)
            cmd.Parameters.AddWithValue("@runningqty", runningqty)
            cmd.Parameters.AddWithValue("@empnum", txtempnum_subcon.Text)
            cmd.Parameters.AddWithValue("@name", txtempname_subcon.Text)
            cmd.Parameters.AddWithValue("@shift", txtshift_subcon.Text)
            cmd.Parameters.AddWithValue("@comment", cmbcomment_subcon.Text)
            'cmd.Parameters.AddWithValue("@sheetqty", CInt(txtsheetqty.Text))
            cmd.ExecuteNonQuery()
            btngenerate.Enabled = True

            Dim remainingtagqty As Integer = CInt(txtjobqty_subcon.Text) - runningqty
            If remainingtagqty < 0 Then
                remainingtagqty = 0
            End If
            txtremainingqty_subcon.Text = remainingtagqty

            cleartext()
            a.Fill(dt)
            DataGridView_subcon.DataSource = dt
            MsgBox("Generate Tag Successfully")
            disablecheckbox()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("debugging in generating ptagline")
            MsgBox(runningqty)
        Finally
            con.Close()
        End Try

        Try
            con.Open()
            Dim cmd_updatehdr As New SqlCommand("UPDATE Ptag_Hdr Set total_tags=@totaltags, Pallet_size=@palletqty WHERE job=@job AND suffix=@suffix AND Oper_num=@opernum", con)

            Dim latestTagNumberhdr As Integer = GetLatestTagNumberFromDatabase_subcon()

            cmd_updatehdr.Parameters.AddWithValue("@totaltags", latestTagNumberhdr)
            cmd_updatehdr.Parameters.AddWithValue("@palletqty", txtremainingqty_subcon.Text)
            cmd_updatehdr.Parameters.AddWithValue("@job", txtjob_subcon.Text)
            cmd_updatehdr.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
            cmd_updatehdr.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

            cmd_updatehdr.ExecuteNonQuery()
            ' MsgBox(latestTagNumberhdr)

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("debugging in updating hdr")
        Finally
            con.Close()
        End Try

    End Sub

    Private Function GetLatestTagNumberFromDatabase_subcon() As Integer
        Dim latestTagNumber As Integer = 0

        Try
            ' Open a connection to your SQL Server database
            Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
                con.Open()

                ' Create a SQL command to retrieve the latest tag number
                Using cmd As New SqlCommand("SELECT MAX(Tag_num) FROM Ptag_Line WHERE Site=@site AND Job=@job AND Suffix=@suffix AND Oper_num=@opernum", con)
                    cmd.Parameters.AddWithValue("@site", txtsite_subcon.Text)
                    cmd.Parameters.AddWithValue("@job", txtjob_subcon.Text)
                    cmd.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
                    cmd.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot DBNull.Value Then
                        latestTagNumber = Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle any exceptions here
            MsgBox(ex.Message)
        End Try

        Return latestTagNumber
    End Function

    Private Function getlatestrunningqty_subcon() As Integer
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        'Using cmd As New SqlCommand("SELECT TOP 1 running_qty FROM Ptag_Line where job = @job and Suffix = @suffix and Oper_num = @opernum ORDER BY Tag_date DESC", con)
        Using cmd As New SqlCommand("SELECT TOP 1 running_qty FROM Ptag_Line where job = @job and Suffix = @suffix and Oper_num = @opernum ORDER BY Tag_num DESC", con)

            cmd.Parameters.AddWithValue("@job", txtjob_subcon.Text)
            cmd.Parameters.AddWithValue("@suffix", txtsuffix_subcon.Text)
            cmd.Parameters.AddWithValue("@opernum", txtopernum_subcon.Text)

            Try
                con.Open()
                Dim result As Object = cmd.ExecuteScalar
                con.Close()

                If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                    Return CInt(result)
                Else
                    Return 0 ' Default value if no rows are found
                End If
            Catch ex As Exception
            End Try
        End Using
        Return 0
    End Function

    Private Sub DataGridView_subcon_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_subcon.CellValueChanged
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Try
            con.Open()

            For Each row As DataGridViewRow In DataGridView_subcon.Rows
                Dim ischecked As Boolean = Convert.ToBoolean(row.Cells(0).Value)

                If ischecked Then
                    Dim cmdchecked As New SqlCommand("UPDATE Ptag_Line Set [Select] = 1 WHERE Job = @job And Suffix = @suffix And Oper_num = @opernum And Tag_num = @tagnum", con)

                    cmdchecked.Parameters.AddWithValue("job", txtjob_subcon.Text)
                    cmdchecked.Parameters.AddWithValue("suffix", txtsuffix_subcon.Text)
                    cmdchecked.Parameters.AddWithValue("opernum", txtopernum_subcon.Text)
                    cmdchecked.Parameters.AddWithValue("tagnum", row.Cells(1).Value)

                    cmdchecked.ExecuteNonQuery()
                Else
                    Dim cmdunchecked As New SqlCommand("UPDATE Ptag_Line Set [Select] = 0 WHERE Job = @job And Suffix = @suffix And Oper_num = @opernum And Tag_num = @tagnum", con)

                    cmdunchecked.Parameters.AddWithValue("job", txtjob_subcon.Text)
                    cmdunchecked.Parameters.AddWithValue("suffix", txtsuffix_subcon.Text)
                    cmdunchecked.Parameters.AddWithValue("opernum", txtopernum_subcon.Text)
                    cmdunchecked.Parameters.AddWithValue("tagnum", row.Cells(1).Value)

                    cmdunchecked.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView_subcon_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView_subcon.CellMouseUp
        ' Check if the clicked cell is in the "Select" column


        If e.ColumnIndex = DataGridView_subcon.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
            ' This code will make sure the checkbox value is toggled immediately
            DataGridView_subcon.EndEdit()
        End If
    End Sub

    Private Sub btnselectall_subcon_Click(sender As Object, e As EventArgs) Handles btnselectall_subcon.Click
        For Each row As DataGridViewRow In DataGridView_subcon.Rows
            If Not row.IsNewRow Then
                Dim createdByUsername As String = row.Cells("Emp_num").Value.ToString
                Dim username As String = loggedinusername()

                If loggedinusername() = createdByUsername Then
                    Dim checkBoxCell As DataGridViewCheckBoxCell = TryCast(row.Cells("Select"), DataGridViewCheckBoxCell)
                    If checkBoxCell IsNot Nothing Then
                        checkBoxCell.Value = Not CBool(checkBoxCell.Value)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            'Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where name=@name AND emp_num = @empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            ' Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where Name=@name and Emp_num=@empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)
            Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)
            Dim cmd1 As New SqlCommand("Update NCFPtag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum and [Select] = 1", con)
            Dim cmd2 As New SqlCommand("Update Ptag_line set [Select] = 0 where job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)

            cmd.Parameters.AddWithValue("empnum", txtoper_num.Text)
            cmd.Parameters.AddWithValue("job", txtjoborder.Text)
            cmd.Parameters.AddWithValue("suffix", txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("opernum", txtjoboperation.Text)


            cmd1.Parameters.AddWithValue("empnum", txtoper_numt3.Text)
            cmd1.Parameters.AddWithValue("job", txtjobordert3.Text)
            cmd1.Parameters.AddWithValue("suffix", txtjobsuffixt3.Text)
            cmd1.Parameters.AddWithValue("opernum", txtjoboperationt3.Text)

            cmd2.Parameters.AddWithValue("empnum", txtempnum_subcon.Text)
            cmd2.Parameters.AddWithValue("job", txtjob_subcon.Text)
            cmd2.Parameters.AddWithValue("suffix", txtsuffix_subcon.Text)
            cmd2.Parameters.AddWithValue("opernum", txtopernum_subcon.Text)

            cmd.ExecuteNonQuery()
            cmd1.ExecuteNonQuery()
            cmd2.ExecuteNonQuery()
            Form1.txtpassword.Clear()
            Form1.cmbuserid.SelectedValue = 0
            Form1.Show()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")

        Dim cmd As New SqlCommand("delete from Ptag_Line where job = @joborder and 
            suffix = @jobsuffix and Oper_num = @jobopernum and 
            (Tag_num = @tagnum OR Tag_num > @tagnum) and 
            Emp_num = @opernum", con)
        cmd.Parameters.AddWithValue("@joborder", txtjob_subcon.Text)
        cmd.Parameters.AddWithValue("@jobsuffix", txtsuffix_subcon.Text)
        cmd.Parameters.AddWithValue("@jobopernum", txtopernum_subcon.Text)
        cmd.Parameters.AddWithValue("@opernum", txtoper_num.Text)



        Dim tablecmd As SqlCommand = New SqlCommand("SELECT [Select], Tag_num as [Tag #] , RIGHT('0' + CAST(DAY(Tag_date) AS NVARCHAR(2)), 2) + ' ' + DATENAME(MONTH, Tag_date) + ' ' + CAST(YEAR(Tag_date) AS NVARCHAR(4)) AS 'Tag Date', Tag_qty, running_qty,  Emp_num, Name, Shift, Comment FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum", con)
        tablecmd.Parameters.AddWithValue("@jonumber", txtjob_subcon.Text)
        tablecmd.Parameters.AddWithValue("@josuffix", txtsuffix_subcon.Text)
        tablecmd.Parameters.AddWithValue("@operationnum", txtopernum_subcon.Text)

        Dim a As New SqlDataAdapter(tablecmd)
        Dim dt As New DataTable

        Dim promptmsg As DialogResult = MessageBox.Show("Are you sure you want to delete ? All tag number below will also be deleted.", "Delete Tag?", MessageBoxButtons.YesNo)

        If promptmsg = DialogResult.Yes Then
            Try
                con.Open()


                If DataGridView_subcon.SelectedRows.Count > 0 Then
                    Dim selectedRow As DataGridViewRow = DataGridView_subcon.SelectedRows(0)
                    cmd.Parameters.AddWithValue("@tagnum", selectedRow.Cells(1).Value)
                    cmd.ExecuteNonQuery()
                    MsgBox("Tag Deleted")
                    a.Fill(dt)
                    DataGridView_subcon.DataSource = dt
                    AutofitColumns(DataGridView_subcon)
                    disablecheckbox()
                    Dim latestrunningqty As Integer = getlatestrunningqty_subcon()
                    Dim tagqty As Integer = CInt(txtjobqty_subcon.Text)
                    Dim result As Integer = tagqty - latestrunningqty
                    txtremainingqty.Text = result

                Else
                    MsgBox("Please Select a tag # to delete")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try
        Else
            MsgBox("Cancelled")
        End If


        Try
            con.Open()
            Dim cmd_updatehdr As New SqlCommand("UPDATE Ptag_Hdr Set total_tags=@totaltags, Pallet_size=@palletqty WHERE job=@job AND suffix=@suffix AND Oper_num=@opernum", con)

            Dim latestTagNumberhdr As Integer = GetLatestTagNumberFromDatabase()

            cmd_updatehdr.Parameters.AddWithValue("@totaltags", latestTagNumberhdr)
            cmd_updatehdr.Parameters.AddWithValue("@palletqty", txtremainingqty.Text)
            cmd_updatehdr.Parameters.AddWithValue("@job", txtjoborder.Text)
            cmd_updatehdr.Parameters.AddWithValue("@suffix", txtjobsuffix.Text)
            cmd_updatehdr.Parameters.AddWithValue("@opernum", txtjoboperation.Text)

            cmd_updatehdr.ExecuteNonQuery()
            ' MsgBox(latestTagNumberhdr)

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("debugging in updating hdr")
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Using con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
            con.Open()
            Dim updateCommandText As String = "UPDATE Ptag_Line SET Comment = @comment WHERE Tag_num=@tagnum AND Site = @site AND Job=@job AND Suffix=@suffix AND Oper_num=@opernum"
            Dim cmd As New SqlCommand(updateCommandText, con)

            ' Add parameters to the SqlCommand
            cmd.Parameters.Add(New SqlParameter("@comment", SqlDbType.NVarChar))

            cmd.Parameters.Add(New SqlParameter("@tagnum", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@site", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@job", SqlDbType.NVarChar))
            cmd.Parameters.Add(New SqlParameter("@suffix", SqlDbType.Int))
            cmd.Parameters.Add(New SqlParameter("@opernum", SqlDbType.Int))

            Try
                ' Iterate through the DataGridView rows
                For Each row As DataGridViewRow In DataGridView_subcon.Rows
                    ' Check if the row is not the header row
                    If Not row.IsNewRow Then
                        ' Set parameter values based on the data in the current row
                        cmd.Parameters("@comment").Value = row.Cells("Comment").Value

                        cmd.Parameters("@tagnum").Value = row.Cells("Tag #").Value
                        cmd.Parameters("@site").Value = txtsite_subcon.Text
                        cmd.Parameters("@job").Value = txtjob_subcon.Text
                        cmd.Parameters("@suffix").Value = txtsuffix_subcon.Text
                        cmd.Parameters("@opernum").Value = txtopernum_subcon.Text
                        ' Execute the update command for the current row
                        cmd.ExecuteNonQuery()
                    End If
                Next
                MsgBox("Update Successfully")
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                con.Close()
            End Try

        End Using
    End Sub

    Private Sub Button18_Click_1(sender As Object, e As EventArgs) Handles btnprint_subcon.Click

        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim selectedrows As New List(Of Integer)

        Try
            con.Open()
            For i As Integer = 0 To DataGridView_subcon.Rows.Count - 1
                If DataGridView_subcon.Rows(i).Cells(0).Value IsNot Nothing AndAlso CBool(DataGridView_subcon.Rows(i).Cells(0).Value) = True Then
                    selectedrows.Add(i)
                End If
            Next

            Dim report7 As New CrystalReport7
            Dim user As String = "sa"
            Dim pwd As String = "pi_dc_2011"

            For Each rowIndex As Integer In selectedrows
                ' Create a new report instance for each selected row
                'Dim selectedvalue As Boolean

                report7.SetParameterValue("job", txtjob_subcon.Text)
                report7.SetParameterValue("suffix", txtsuffix_subcon.Text)
                report7.SetParameterValue("oper_num", txtopernum_subcon.Text)
                'report4.SetParameterValue("job_operation", txtjoboperation.Text)
                'report4.SetParameterValue("Select", selectedvalue)
                'report4.SetParameterValue("total_tagnum", DataGridView1.Rows.Count)
                report7.SetParameterValue("operator", txtempnum_subcon.Text)

            Next

            report7.SetDatabaseLogon(user, pwd)
            Form5.CrystalReportViewer1.ReportSource = report7
            Form5.CrystalReportViewer1.Refresh()
            Form5.CrystalReportViewer1.Zoom(50)
            Form5.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Sub

    Private Sub cmbsection_schedule_DropDown(sender As Object, e As EventArgs) Handles cmbsection_schedule.DropDown
        'cmbsection_schedule.Items.Clear()
        'populatesection()
    End Sub

    Private Sub populatesection()
        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim cmd_section As SqlCommand = New SqlCommand("select distinct section from Employee", con1)
        'cmd_section.Parameters.AddWithValue("@empnum", Form1.cmbuserid.Text)
        Try
            con1.Open()
            Dim read_cmd_section As SqlDataReader = cmd_section.ExecuteReader()
            While read_cmd_section.Read
                Dim section As String = read_cmd_section("section").ToString
                cmbsection_schedule.Items.Add(section)
            End While
            read_cmd_section.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con1.Close()
        End Try

    End Sub

    Private Sub Button18_Click_2(sender As Object, e As EventArgs) Handles Button18.Click
        Frm_PalletTagSummary.Show()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        'BACKUP
        'Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        'Dim selectedrows As New List(Of Integer)

        'Try
        '    con.Open()
        '    For i As Integer = 0 To DataGridView1.Rows.Count - 1
        '        If DataGridView1.Rows(i).Cells(0).Value IsNot Nothing AndAlso CBool(DataGridView1.Rows(i).Cells(0).Value) = True Then
        '            selectedrows.Add(i)
        '        End If
        '    Next

        '    Dim report4 As New CrystalReport_QRCODE
        '    Dim user As String = "sa"
        '    Dim pwd As String = "pi_dc_2011"

        '    For Each rowIndex As Integer In selectedrows
        '        ' Create a new report instance for each selected row
        '        'Dim selectedvalue As Boolean

        '        report4.SetParameterValue("job", txtjoborder.Text)
        '        report4.SetParameterValue("suffix", txtjobsuffix.Text)
        '        report4.SetParameterValue("oper_num", txtjoboperation.Text)
        '        'report4.SetParameterValue("job_operation", txtjoboperation.Text)
        '        'report4.SetParameterValue("Select", selectedvalue)
        '        'report4.SetParameterValue("total_tagnum", DataGridView1.Rows.Count)
        '        report4.SetParameterValue("operator", txtoper_num.Text)

        '    Next

        '    report4.SetDatabaseLogon(user, pwd)
        '    Form5.CrystalReportViewer1.ReportSource = report4
        '    Form5.CrystalReportViewer1.Refresh()
        '    Form5.CrystalReportViewer1.Zoom(50)
        '    Form5.Show()

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally
        '    con.Close()
        'End Try

        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Dim selectedrows As New List(Of Integer)

        Try
            con.Open()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(0).Value IsNot Nothing AndAlso CBool(DataGridView1.Rows(i).Cells(0).Value) = True Then
                    selectedrows.Add(i)
                End If
            Next

            Dim report4 As New CrystalReport_QRCODE
            Dim user As String = "sa"
            Dim pwd As String = "pi_dc_2011"

            Dim i2 As Integer
            i2 = DataGridView1.CurrentRow.Index

            'Label5.Text = DataGridView1.Item(6, i).Value

            Dim qri As Integer
            qri = DataGridView1.CurrentRow.Index

            For Each rowIndex As Integer In selectedrows
                ' Create a new report instance for each selected row
                'Dim selectedvalue As Boolean

                report4.SetParameterValue("job", txtjoborder.Text)
                report4.SetParameterValue("suffix", txtjobsuffix.Text)
                report4.SetParameterValue("nextopernum", txt_nextopernum.Text)
                report4.SetParameterValue("oper_num", txtjoboperation.Text)
                'report4.SetParameterValue("job_operation", txtjoboperation.Text)
                'report4.SetParameterValue("Select", selectedvalue)
                'report4.SetParameterValue("total_tagnum", DataGridView1.Rows.Count)
                'report4.SetParameterValue("tagnum", DataGridView1.Item(1, i2).Value)
                ' report4.SetParameterValue("tagqty", DataGridView1.SelectedRows(3).ToString)
                'report4.SetParameterValue("tagqty", DataGridView1.Item(3, i2).Value)
                report4.SetParameterValue("tagqty", DataGridView1.Item(3, qri).Value)
                report4.SetParameterValue("tag_num", DataGridView1.Item(1, qri).Value)
                report4.SetParameterValue("operator", txtoper_num.Text)

            Next

            report4.SetDatabaseLogon(user, pwd)
            Form5.CrystalReportViewer1.ReportSource = report4
            Form5.CrystalReportViewer1.Refresh()
            Form5.CrystalReportViewer1.Zoom(50)
            Form5.Show()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Form1.Show()
        Form1.txtpassword.Clear()
        Me.Close()
    End Sub

    Private Sub txtnumout_TextChanged(sender As Object, e As EventArgs) Handles txtnumout.TextChanged
        If Not String.IsNullOrEmpty(txtjobqty.Text) AndAlso Not String.IsNullOrEmpty(txtqtyperpallet.Text) AndAlso IsNumeric(txtjobqty.Text) AndAlso IsNumeric(txtqtyperpallet.Text) AndAlso Convert.ToDecimal(txtqtyperpallet.Text) <> 0 Then
            txtnumtag.Text = Convert.ToInt32(Decimal.Parse(txtjobqty.Text) / Decimal.Parse(txtqtyperpallet.Text)).ToString()
            If txtnumout.Text = "" Then
                txtnumout.Text = 0
            End If
            If txtnumout.Text = 0 Then
                txtsheetqty.Text = 0
            Else
                txtsheetqty.Text = CInt(txtqtyperpallet.Text) / CInt(txtnumout.Text)
            End If

        Else
            txtnumtag.Text = "" ' or any other value to indicate division by zero or invalid input
        End If

        If txtnumout.Text = "" Then
            txtsheetqty.Text = 0
        End If
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Frm_Transaction_Summary_Report.Show()
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        frm_MKTPT_Summary.Show()
    End Sub

End Class