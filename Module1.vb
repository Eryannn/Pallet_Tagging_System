Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Configuration
Imports System.Security.Cryptography
Module Module1
    Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
    Dim con1 As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")

    Function opendb()
        Try
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function

    Function closedb()
        Try
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function

    Function opendb1()
        Try
            con1.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function

    Function closedb1()
        Try
            con1.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function

    Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
         &H65, &H64, &H76, &H65, &H64, &H65,
         &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function

    Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
         &H65, &H64, &H76, &H65, &H64, &H65,
         &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function

    Function cleartext()
        Form2.txtjoborder.Clear()
        Form2.txtjobsuffix.Clear()
        Form2.txtjoboperation.Clear()
        Form2.txtwc.Clear()
        Form2.txtnextop.Clear()
        Form2.rtbmachine.Clear()
        Form2.txtjobname.Clear()
        Form2.txtjobdesc.Clear()
        Form2.txtjobqty.Clear()
        Form2.txtqtyperpallet.Clear()
        Form2.txtnumtag.Clear()
        Form2.Label13.Text = ""
        Form2.txtjoboperationt3.Clear()
        Form2.txtjobordert3.Clear()
        Form2.txtjobsuffixt3.Clear()

        Return 0
    End Function

    'Function populatewc()
    '    Try
    '        Dim combobox As ComboBox = Form2.TabControl1.TabPages("TabPage1").Controls("ComboBox1")
    '        Form6.ComboBox1.DataSource = Nothing
    '        Form6.ComboBox1.Items.Clear()
    '        Dim query As String = "SELECT WC.wc,WC.description FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID"
    '        Dim cmdquery As New SqlCommand(query, con1)

    '        Dim execquery As SqlDataReader = cmdquery.ExecuteReader
    '        While execquery.Read()
    '            Try
    '                Dim wc As String = execquery("wc").ToString() '+ "     " + execquery("descri tion").ToString
    '                'Form6.tabpage2.cmbwcp2.Items.Add(wc)
    '                combobox.Items.Add(wc)
    '                Form6.ComboBox1.Items.Add(wc)

    '            Catch ex As Exception
    '                MsgBox(ex.Message)
    '            End Try

    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    '    Return 0
    'End Function
    Function populateuser()
        Try
            Form1.cmbuserid.Items.Clear()
            Dim query As String = "SELECT Emp_num FROM Employee WHERE Site=@site AND IsActive = 1 GROUP BY Emp_num"
            Dim cmdquery As New SqlCommand(query, con)
            cmdquery.Parameters.AddWithValue("@site", Form1.ComboBox2.Text)
            Dim execquery As SqlDataReader = cmdquery.ExecuteReader
            While execquery.Read()
                Form1.cmbuserid.Items.Add(execquery(0))
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function

    Function loadtable()
        Dim query As String = "SELECT * FROM Ptag_Line WHERE Job=@jonumber AND Suffix=@josuffix AND Oper_num=@operationnum"
        Dim cmdquery As New SqlCommand(query, con)
        cmdquery.Parameters.AddWithValue("@jonumber", Form2.txtjoborder.Text)
        cmdquery.Parameters.AddWithValue("@josuffix", Form2.txtjobsuffix.Text)
        cmdquery.Parameters.AddWithValue("@operationnum", Form2.txtjoboperation.Text)
        Dim a As New SqlDataAdapter(cmdquery)
        Dim dt As New DataTable
        a.Fill(dt)
        Form2.DataGridView1.DataSource = dt
        Return 0
    End Function

    Function login(ByVal userid As String, ByVal password As String, ByVal site As String)
        Try
            Dim query As String = "SELECT * FROM Employee WHERE Emp_num=@userid AND Password=@password AND Site=@site"
            Dim cmdquery As New SqlCommand(query, con)

            cmdquery.Parameters.Add("@userid", SqlDbType.NVarChar).Value = userid
            cmdquery.Parameters.Add("@password", SqlDbType.NVarChar).Value = password
            cmdquery.Parameters.Add("@site", SqlDbType.NVarChar).Value = site

            Dim execquery As SqlDataReader = cmdquery.ExecuteReader
            Dim hasrow As Boolean = execquery.Read()

            If hasrow Then
                Dim IsAdmin As Boolean = Convert.ToBoolean(execquery("IsAdmin"))
                Dim IsActive As Boolean = Convert.ToBoolean(execquery("IsActive"))
                Dim useraccess As String = execquery("Access")
                Dim userrights As String = execquery("Rights")

                If IsAdmin AndAlso IsActive Then
                    Form4.Show()
                    Form1.Hide()
                ElseIf IsActive Then
                    If useraccess = "REPORT" Then
                        If userrights = "VIEW" Then
                            Form2.TabControl1.TabPages("TabPage1").Enabled = False
                            Form2.TabControl1.TabPages("TabPage2").Enabled = False
                            Form2.TabControl1.TabPages("TabPage3").Enabled = False
                            Form2.TabControl1.TabPages("TabPage4").Enabled = False
                            Form2.Button18.Enabled = False
                            'Frm_PalletTagSummary.Show()
                            Form2.lblempnum_reports.Text = Form1.cmbuserid.Text
                            Form2.lblshift_reports.Text = Form1.cmbshift.Text
                            Form2.lblsite_reports.Text = Form1.ComboBox2.Text
                            Form2.TabControl1.SelectedTab = Form2.TabPage5
                            Form2.lblempname_report.Text = execquery("Name").ToString
                            Form2.Show()
                            Form1.Hide()
                        End If
                    ElseIf useraccess = "TAGGING"
                        Dim form2 As New Form2()
                        form2.txtoper_num.Text = Form1.cmbuserid.Text
                        form2.txtoper_numt3.Text = Form1.cmbuserid.Text
                        form2.TextBox3.Text = Form1.ComboBox2.Text
                        form2.txtshift.Text = Form1.cmbshift.Text
                        form2.lblshiftt2.Text = Form1.cmbshift.Text
                        form2.lblshift_reports.Text = Form1.cmbshift.Text
                        form2.lblusert2.Text = Form1.cmbuserid.Text.ToUpper
                        form2.txtshiftt3.Text = Form1.cmbshift.Text
                        form2.txtsite_subcon.Text = Form1.ComboBox2.Text
                        form2.txtshift_subcon.Text = Form1.cmbshift.Text

                        'form2.cmbshiftt4.Text = Form1.cmbshift.Text
                        ' form2.lblshiftt4.Text = Form1.cmbshift.Text
                        form2.Show()
                        Form1.Hide()
                    End If
                Else
                    MessageBox.Show("Account is not Active")
                End If
            Else
                MessageBox.Show("Invalid Credentials!")
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            MessageBox.Show(password)
        End Try
        Return 0
    End Function

    Function getuserinfo(ByVal empnum As String, ByVal site As String)

        Return 0
    End Function

    Function registeruser(ByVal site As String, ByVal empnum As String, ByVal name As String, ByVal department As String, ByVal section As String, ByVal position As String, ByVal password As String, ByVal access As String, ByVal rights As String)
        Try
            Dim query As String = "EXEC insertpallettagginguser @site,@empnum, @name, @department, @section, @position, @encryptpassword, @access, @rights"
            Dim cmdquery As New SqlCommand(query, con)

            cmdquery.Parameters.Add("@site", SqlDbType.NVarChar).Value = site
            cmdquery.Parameters.Add("@empnum", SqlDbType.NVarChar).Value = empnum
            cmdquery.Parameters.Add("@name", SqlDbType.NVarChar).Value = name
            cmdquery.Parameters.Add("@department", SqlDbType.NVarChar).Value = department
            cmdquery.Parameters.Add("@section", SqlDbType.NVarChar).Value = section
            cmdquery.Parameters.Add("@position", SqlDbType.NVarChar).Value = position
            cmdquery.Parameters.Add("@encryptpassword", SqlDbType.NVarChar).Value = password
            cmdquery.Parameters.Add("@access", SqlDbType.NVarChar).Value = access
            cmdquery.Parameters.Add("@rights", SqlDbType.NVarChar).Value = rights
            If Form3.txtpassword.Text.Length > 1 Then
                If Form3.txtpassword.Text = Form3.txtconfirmpass.Text Then
                    If cmdquery.ExecuteNonQuery() > 0 Then
                        Form4.Show()
                        Form3.Dispose()
                        MsgBox("Register Successfully")
                    End If
                Else
                    MsgBox("Password Doesn't Match")
                End If
            Else
                MsgBox("Password must be 9 characters or more!")
            End If

        Catch ex As Exception
            'MsgBox(ex.Message)
            MsgBox("User already Exist!")
        End Try
        Return 0
    End Function

    Function testpopulatewc()
        Dim con2 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
        'Dim cmd2 As SqlCommand = New SqlCommand("SELECT WC.wc,WC.description FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID", con2)

        Dim cmd2 As SqlCommand = New SqlCommand("SELECT WC.wc, wc.description FROM wc INNER JOIN wcresourcegroup WCR ON WC.wc=WCR.wc INNER JOIN RGRPMBR000 RG ON RG.RGID=WCR.rgid INNER JOIN RESRC000 RS ON RS.RESID=RG.RESID", con2)

        Try
            con2.Open()
            Dim sqlreader1 As SqlDataReader = cmd2.ExecuteReader()
            While sqlreader1.Read()
                Dim wc As String = sqlreader1("wc").ToString()
                Dim desc As String = sqlreader1("description").ToString()
            End While
            sqlreader1.Close()

        Catch ex As Exception

        End Try
        Return 0
    End Function


    '    Function testdatafillup()
    '        Dim con1 As SqlConnection = New SqlConnection("Data Source=ERP-SVR;Initial Catalog=PI-SP_App;User ID=sa;Password=pi_dc_2011")
    '        Dim cmd As SqlCommand = New SqlCommand("  SELECT  rs.RESID AS resourceMachine,
    '		jr.job ,
    '		jr.suffix,
    '		jr.oper_num,
    '		js.start_date AS [date],
    '		jo.item AS jobitem ,
    '		itm.description ,itm.Uf_itemdesc_ext, -- DESCRIPTION
    '		jo.qty_released AS QTY,
    '		itm.u_m AS UM
    'FROM job  jo 
    'INNER JOIN jobroute jr 
    '		   ON jo.job=jr.job AND jo.suffix=jr.suffix
    'INNER JOIN job_sch js
    '		   ON jo.job=js.job AND jo.suffix=js.suffix
    'INNER JOIN item itm 
    '		   ON itm.item=jo.item
    'INNER JOIN wc 
    '		   ON wc.wc =jr.wc
    'INNER JOIN jrtresourcegroup jrs 
    '		   ON jr.job=jrs.job AND jr.suffix=jrs.suffix 
    '		                     AND jr.oper_num=jrs.oper_num
    'INNER JOIN RGRPMBR000 rg 
    '		   ON jrs.rgid=rg.RGID 
    'INNER JOIN RESRC000 rs 
    '		   ON rg.RESID=rs.RESID
    'WHERE jo.type='J' 
    '  AND jo.stat='R'
    '  AND JR.wc=@wc
    '  AND js.start_date >=@startdate
    '  AND js.end_date <=@enddate
    '  AND rs.RESID =@resid", con1)
    '        cmd.Parameters.AddWithValue("@wc", Form6.ComboBox1.Text)
    '        'cmd.Parameters.AddWithValue("@startdate", Form6.DateTimePicker1.Value)
    '        'cmd.Parameters.AddWithValue("@enddate", Form6.DateTimePicker2.Value)
    '        cmd.Parameters.Add("@startdate", SqlDbType.DateTime).Value = Form6.DateTimePicker1.Value
    '        cmd.Parameters.Add("@enddate", SqlDbType.DateTime).Value = Form6.DateTimePicker2.Value


    '        Dim a As New SqlDataAdapter(cmd)
    '        Dim dt As New DataTable
    '        a.Fill(dt)
    '        Form6.DataGridView2.DataSource = dt

    '        Return 0
    '    End Function

    Function resetptagline()
        Dim con As New SqlConnection("Data Source=ERP-SVR;Initial Catalog=Pallet_Tagging;User ID=sa;Password=pi_dc_2011")
        Try
            con.Open()
            Dim cmd As New SqlCommand("Update Ptag_line set [Select] = 0 where name=@name AND emp_num = @empnum AND job=@job AND suffix=@suffix AND Oper_num=@opernum AND [Select] = 1", con)

            cmd.Parameters.AddWithValue("name", Form2.txtoper_name.Text)
            cmd.Parameters.AddWithValue("empnum", Form2.txtoper_num.Text)
            cmd.Parameters.AddWithValue("job", Form2.txtjoborder.Text)
            cmd.Parameters.AddWithValue("suffix", Form2.txtjobsuffix.Text)
            cmd.Parameters.AddWithValue("opernum", Form2.txtjoboperation.Text)

            cmd.ExecuteNonQuery()


        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
        Return 0
    End Function
End Module
