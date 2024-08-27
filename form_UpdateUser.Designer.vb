<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_UpdateUser
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtposition = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtsection = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.txtconfirmpass = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtpassword = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtdept = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtname = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtempnum = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbsite = New System.Windows.Forms.ComboBox()
        Me.txtprevpass = New System.Windows.Forms.TextBox()
        Me.btnupdate = New System.Windows.Forms.Button()
        Me.btnexit = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmb_access = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmb_rights = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'txtposition
        '
        Me.txtposition.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtposition.Location = New System.Drawing.Point(290, 331)
        Me.txtposition.MaxLength = 50
        Me.txtposition.Name = "txtposition"
        Me.txtposition.Size = New System.Drawing.Size(378, 35)
        Me.txtposition.TabIndex = 35
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(31, 331)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(111, 31)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Position"
        '
        'txtsection
        '
        Me.txtsection.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtsection.Location = New System.Drawing.Point(290, 283)
        Me.txtsection.MaxLength = 50
        Me.txtsection.Name = "txtsection"
        Me.txtsection.Size = New System.Drawing.Size(378, 35)
        Me.txtsection.TabIndex = 33
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(31, 283)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(105, 31)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "Section"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(689, 391)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(102, 17)
        Me.CheckBox1.TabIndex = 31
        Me.CheckBox1.Text = "Show Password"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'txtconfirmpass
        '
        Me.txtconfirmpass.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtconfirmpass.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtconfirmpass.Location = New System.Drawing.Point(290, 434)
        Me.txtconfirmpass.Name = "txtconfirmpass"
        Me.txtconfirmpass.Size = New System.Drawing.Size(378, 35)
        Me.txtconfirmpass.TabIndex = 30
        Me.txtconfirmpass.UseSystemPasswordChar = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(30, 434)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(236, 31)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Confirm Password"
        '
        'txtpassword
        '
        Me.txtpassword.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtpassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpassword.Location = New System.Drawing.Point(290, 382)
        Me.txtpassword.Name = "txtpassword"
        Me.txtpassword.Size = New System.Drawing.Size(378, 35)
        Me.txtpassword.TabIndex = 28
        Me.txtpassword.UseSystemPasswordChar = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(30, 386)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(196, 31)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "New Password"
        '
        'txtdept
        '
        Me.txtdept.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdept.Location = New System.Drawing.Point(290, 230)
        Me.txtdept.MaxLength = 50
        Me.txtdept.Name = "txtdept"
        Me.txtdept.Size = New System.Drawing.Size(378, 35)
        Me.txtdept.TabIndex = 26
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(31, 234)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(156, 31)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Department"
        '
        'txtname
        '
        Me.txtname.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtname.Location = New System.Drawing.Point(290, 179)
        Me.txtname.MaxLength = 50
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(378, 35)
        Me.txtname.TabIndex = 24
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(31, 179)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 31)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Name"
        '
        'txtempnum
        '
        Me.txtempnum.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtempnum.Location = New System.Drawing.Point(290, 126)
        Me.txtempnum.MaxLength = 7
        Me.txtempnum.Name = "txtempnum"
        Me.txtempnum.ReadOnly = True
        Me.txtempnum.Size = New System.Drawing.Size(167, 35)
        Me.txtempnum.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(31, 126)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(192, 31)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Employee No.:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(31, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 31)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Site"
        '
        'cmbsite
        '
        Me.cmbsite.Enabled = False
        Me.cmbsite.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbsite.FormattingEnabled = True
        Me.cmbsite.Items.AddRange(New Object() {"PI-SP", "FP-SP", "PWPC"})
        Me.cmbsite.Location = New System.Drawing.Point(290, 72)
        Me.cmbsite.Name = "cmbsite"
        Me.cmbsite.Size = New System.Drawing.Size(144, 37)
        Me.cmbsite.TabIndex = 19
        '
        'txtprevpass
        '
        Me.txtprevpass.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtprevpass.Location = New System.Drawing.Point(499, 12)
        Me.txtprevpass.Name = "txtprevpass"
        Me.txtprevpass.Size = New System.Drawing.Size(378, 35)
        Me.txtprevpass.TabIndex = 36
        Me.txtprevpass.UseSystemPasswordChar = True
        Me.txtprevpass.Visible = False
        '
        'btnupdate
        '
        Me.btnupdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnupdate.Location = New System.Drawing.Point(141, 511)
        Me.btnupdate.Name = "btnupdate"
        Me.btnupdate.Size = New System.Drawing.Size(222, 90)
        Me.btnupdate.TabIndex = 38
        Me.btnupdate.Text = "Update"
        Me.btnupdate.UseVisualStyleBackColor = True
        '
        'btnexit
        '
        Me.btnexit.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnexit.Location = New System.Drawing.Point(490, 513)
        Me.btnexit.Name = "btnexit"
        Me.btnexit.Size = New System.Drawing.Size(202, 90)
        Me.btnexit.TabIndex = 37
        Me.btnexit.Text = "BACK"
        Me.btnexit.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(507, 84)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(118, 31)
        Me.Label10.TabIndex = 42
        Me.Label10.Text = "RIGHTS"
        '
        'cmb_access
        '
        Me.cmb_access.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_access.FormattingEnabled = True
        Me.cmb_access.Items.AddRange(New Object() {"TAGGING", "REPORT"})
        Me.cmb_access.Location = New System.Drawing.Point(631, 121)
        Me.cmb_access.Name = "cmb_access"
        Me.cmb_access.Size = New System.Drawing.Size(144, 37)
        Me.cmb_access.TabIndex = 41
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(499, 126)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(126, 31)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "ACCESS"
        '
        'cmb_rights
        '
        Me.cmb_rights.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_rights.FormattingEnabled = True
        Me.cmb_rights.Items.AddRange(New Object() {"EDIT", "VIEW", "DELETE", "ALL"})
        Me.cmb_rights.Location = New System.Drawing.Point(631, 78)
        Me.cmb_rights.Name = "cmb_rights"
        Me.cmb_rights.Size = New System.Drawing.Size(144, 37)
        Me.cmb_rights.TabIndex = 39
        '
        'form_UpdateUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(898, 639)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.cmb_access)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.cmb_rights)
        Me.Controls.Add(Me.btnupdate)
        Me.Controls.Add(Me.btnexit)
        Me.Controls.Add(Me.txtprevpass)
        Me.Controls.Add(Me.txtposition)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtsection)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.txtconfirmpass)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtpassword)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtdept)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtname)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtempnum)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbsite)
        Me.Name = "form_UpdateUser"
        Me.Text = "form_UpdateUser"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtposition As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtsection As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents txtconfirmpass As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtpassword As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtdept As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtname As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtempnum As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbsite As ComboBox
    Friend WithEvents txtprevpass As TextBox
    Friend WithEvents btnupdate As Button
    Friend WithEvents btnexit As Button
    Friend WithEvents Label10 As Label
    Friend WithEvents cmb_access As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents cmb_rights As ComboBox
End Class
