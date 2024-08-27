<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Transaction_Summary_Report
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Transaction_Summary_Report))
        Me.Label18 = New System.Windows.Forms.Label()
        Me.cmb_section = New System.Windows.Forms.ComboBox()
        Me.cmb_machine = New System.Windows.Forms.ComboBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.dtp_startdate = New System.Windows.Forms.DateTimePicker()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.dtp_enddate = New System.Windows.Forms.DateTimePicker()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.Label81 = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmb_report_type = New System.Windows.Forms.ComboBox()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(135, 143)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(79, 20)
        Me.Label18.TabIndex = 164
        Me.Label18.Text = "SECTION"
        '
        'cmb_section
        '
        Me.cmb_section.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_section.FormattingEnabled = True
        Me.cmb_section.Items.AddRange(New Object() {"OFFSET", "WEB", "DIGITAL PRESS", "DIE CUTTING", "LAMINATION", "GLUING", "FOLDING", "STITCHING", "PERFECT BINDING", "ERECTING", "MANUAL FINISHING", "BINDERY FINISHING", "STRIPPING", "MANUAL STRIPPING", "MACHINE STRIPPING", "SHEETING/SLITTING", "CUTTING", "3 KNIVES", "INSPECTION MACHINE", "QA"})
        Me.cmb_section.Location = New System.Drawing.Point(220, 140)
        Me.cmb_section.Name = "cmb_section"
        Me.cmb_section.Size = New System.Drawing.Size(329, 32)
        Me.cmb_section.TabIndex = 163
        '
        'cmb_machine
        '
        Me.cmb_machine.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_machine.FormattingEnabled = True
        Me.cmb_machine.Location = New System.Drawing.Point(219, 178)
        Me.cmb_machine.Name = "cmb_machine"
        Me.cmb_machine.Size = New System.Drawing.Size(330, 32)
        Me.cmb_machine.TabIndex = 159
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(70, 230)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(145, 20)
        Me.Label22.TabIndex = 158
        Me.Label22.Text = "SCHEDULE DATE"
        '
        'dtp_startdate
        '
        Me.dtp_startdate.CustomFormat = "M/dd/yyyy"
        Me.dtp_startdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp_startdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_startdate.Location = New System.Drawing.Point(221, 228)
        Me.dtp_startdate.Name = "dtp_startdate"
        Me.dtp_startdate.Size = New System.Drawing.Size(121, 26)
        Me.dtp_startdate.TabIndex = 160
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(35, 181)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(179, 20)
        Me.Label19.TabIndex = 157
        Me.Label19.Text = "RESOURCE/MACHINE"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(358, 226)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(19, 25)
        Me.Label23.TabIndex = 161
        Me.Label23.Text = "-"
        '
        'dtp_enddate
        '
        Me.dtp_enddate.CustomFormat = "M/dd/yyyy"
        Me.dtp_enddate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp_enddate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_enddate.Location = New System.Drawing.Point(393, 228)
        Me.dtp_enddate.Name = "dtp_enddate"
        Me.dtp_enddate.Size = New System.Drawing.Size(121, 26)
        Me.dtp_enddate.TabIndex = 162
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox5.ErrorImage = Nothing
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(0, 4)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(129, 111)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox5.TabIndex = 123
        Me.PictureBox5.TabStop = False
        '
        'Label81
        '
        Me.Label81.AutoSize = True
        Me.Label81.Font = New System.Drawing.Font("Calibri", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label81.ForeColor = System.Drawing.Color.White
        Me.Label81.Location = New System.Drawing.Point(135, 41)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(476, 39)
        Me.Label81.TabIndex = 124
        Me.Label81.Text = "TRANSACTION SUMMARY REPORT"
        Me.Label81.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.BackColor = System.Drawing.Color.RoyalBlue
        Me.Panel6.Controls.Add(Me.Label81)
        Me.Panel6.Controls.Add(Me.PictureBox5)
        Me.Panel6.Location = New System.Drawing.Point(0, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(818, 115)
        Me.Panel6.TabIndex = 165
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(218, 322)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(129, 39)
        Me.Button1.TabIndex = 166
        Me.Button1.Text = "PREVIEW"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(404, 322)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(129, 39)
        Me.Button2.TabIndex = 167
        Me.Button2.Text = "EXCEL"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(68, 273)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(146, 20)
        Me.Label1.TabIndex = 168
        Me.Label1.Text = "REPORT FORMAT"
        '
        'cmb_report_type
        '
        Me.cmb_report_type.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_report_type.FormattingEnabled = True
        Me.cmb_report_type.Items.AddRange(New Object() {"TIME EVALUATION REPORT", "OUTPUT VOLUME REPORT"})
        Me.cmb_report_type.Location = New System.Drawing.Point(221, 270)
        Me.cmb_report_type.Name = "cmb_report_type"
        Me.cmb_report_type.Size = New System.Drawing.Size(330, 28)
        Me.cmb_report_type.TabIndex = 169
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.IndianRed
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(657, 454)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(129, 39)
        Me.Button3.TabIndex = 170
        Me.Button3.Text = "Back"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Frm_Transaction_Summary_Report
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(815, 527)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.cmb_report_type)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.cmb_section)
        Me.Controls.Add(Me.cmb_machine)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.dtp_startdate)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.dtp_enddate)
        Me.Name = "Frm_Transaction_Summary_Report"
        Me.Text = "Frm_Transaction_Summary_Report"
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label18 As Label
    Friend WithEvents cmb_section As ComboBox
    Friend WithEvents cmb_machine As ComboBox
    Friend WithEvents Label22 As Label
    Friend WithEvents dtp_startdate As DateTimePicker
    Friend WithEvents Label19 As Label
    Friend WithEvents Label23 As Label
    Friend WithEvents dtp_enddate As DateTimePicker
    Friend WithEvents PictureBox5 As PictureBox
    Friend WithEvents Label81 As Label
    Friend WithEvents Panel6 As Panel
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents cmb_report_type As ComboBox
    Friend WithEvents Button3 As Button
End Class
