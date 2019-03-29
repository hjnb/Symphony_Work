<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 週間表
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblYmd = New System.Windows.Forms.Label()
        Me.btnUp = New System.Windows.Forms.Button()
        Me.btnDown = New System.Windows.Forms.Button()
        Me.rbn2F = New System.Windows.Forms.RadioButton()
        Me.rbn3F = New System.Windows.Forms.RadioButton()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.btnTouroku = New System.Windows.Forms.Button()
        Me.btnSakujo = New System.Windows.Forms.Button()
        Me.btnInnsatu = New System.Windows.Forms.Button()
        Me.btnTorikomi = New System.Windows.Forms.Button()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblYmd
        '
        Me.lblYmd.AutoSize = True
        Me.lblYmd.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblYmd.Location = New System.Drawing.Point(25, 13)
        Me.lblYmd.Name = "lblYmd"
        Me.lblYmd.Size = New System.Drawing.Size(126, 19)
        Me.lblYmd.TabIndex = 0
        Me.lblYmd.Text = "H30.12.31 (日)"
        '
        'btnUp
        '
        Me.btnUp.Location = New System.Drawing.Point(156, 3)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(17, 19)
        Me.btnUp.TabIndex = 1
        Me.btnUp.Text = "▲"
        Me.btnUp.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Location = New System.Drawing.Point(156, 21)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(17, 19)
        Me.btnDown.TabIndex = 2
        Me.btnDown.Text = "▼"
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'rbn2F
        '
        Me.rbn2F.AutoSize = True
        Me.rbn2F.Checked = True
        Me.rbn2F.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.rbn2F.Location = New System.Drawing.Point(202, 13)
        Me.rbn2F.Name = "rbn2F"
        Me.rbn2F.Size = New System.Drawing.Size(50, 20)
        Me.rbn2F.TabIndex = 3
        Me.rbn2F.TabStop = True
        Me.rbn2F.Text = "2階"
        Me.rbn2F.UseVisualStyleBackColor = True
        '
        'rbn3F
        '
        Me.rbn3F.AutoSize = True
        Me.rbn3F.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.rbn3F.Location = New System.Drawing.Point(263, 13)
        Me.rbn3F.Name = "rbn3F"
        Me.rbn3F.Size = New System.Drawing.Size(50, 20)
        Me.rbn3F.TabIndex = 4
        Me.rbn3F.Text = "3階"
        Me.rbn3F.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.DataGridView1.Location = New System.Drawing.Point(12, 43)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 21
        Me.DataGridView1.Size = New System.Drawing.Size(1174, 549)
        Me.DataGridView1.TabIndex = 5
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.ImeMode = System.Windows.Forms.ImeMode.Hiragana
        Me.DataGridView2.Location = New System.Drawing.Point(42, 591)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 21
        Me.DataGridView2.Size = New System.Drawing.Size(1144, 101)
        Me.DataGridView2.TabIndex = 6
        '
        'btnTouroku
        '
        Me.btnTouroku.Location = New System.Drawing.Point(398, 6)
        Me.btnTouroku.Name = "btnTouroku"
        Me.btnTouroku.Size = New System.Drawing.Size(83, 28)
        Me.btnTouroku.TabIndex = 8
        Me.btnTouroku.Text = "登録"
        Me.btnTouroku.UseVisualStyleBackColor = True
        Me.btnTouroku.Visible = False
        '
        'btnSakujo
        '
        Me.btnSakujo.Location = New System.Drawing.Point(499, 6)
        Me.btnSakujo.Name = "btnSakujo"
        Me.btnSakujo.Size = New System.Drawing.Size(83, 28)
        Me.btnSakujo.TabIndex = 9
        Me.btnSakujo.Text = "削除"
        Me.btnSakujo.UseVisualStyleBackColor = True
        Me.btnSakujo.Visible = False
        '
        'btnInnsatu
        '
        Me.btnInnsatu.Location = New System.Drawing.Point(600, 6)
        Me.btnInnsatu.Name = "btnInnsatu"
        Me.btnInnsatu.Size = New System.Drawing.Size(83, 28)
        Me.btnInnsatu.TabIndex = 10
        Me.btnInnsatu.Text = "印刷"
        Me.btnInnsatu.UseVisualStyleBackColor = True
        Me.btnInnsatu.Visible = False
        '
        'btnTorikomi
        '
        Me.btnTorikomi.Location = New System.Drawing.Point(701, 6)
        Me.btnTorikomi.Name = "btnTorikomi"
        Me.btnTorikomi.Size = New System.Drawing.Size(83, 28)
        Me.btnTorikomi.TabIndex = 11
        Me.btnTorikomi.Text = "取込"
        Me.btnTorikomi.UseVisualStyleBackColor = True
        Me.btnTorikomi.Visible = False
        '
        'Label52
        '
        Me.Label52.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label52.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label52.Location = New System.Drawing.Point(12, 72)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(1174, 2)
        Me.Label52.TabIndex = 398
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label1.Location = New System.Drawing.Point(12, 99)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(1174, 2)
        Me.Label1.TabIndex = 399
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label2.Location = New System.Drawing.Point(12, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(1174, 2)
        Me.Label2.TabIndex = 400
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label3.Location = New System.Drawing.Point(12, 142)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(1174, 2)
        Me.Label3.TabIndex = 401
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label4.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label4.Location = New System.Drawing.Point(12, 211)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(1174, 2)
        Me.Label4.TabIndex = 402
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label5.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label5.Location = New System.Drawing.Point(12, 282)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(1174, 2)
        Me.Label5.TabIndex = 403
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label6.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label6.Location = New System.Drawing.Point(12, 352)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(1174, 2)
        Me.Label6.TabIndex = 404
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label7.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label7.Location = New System.Drawing.Point(12, 379)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(1174, 2)
        Me.Label7.TabIndex = 405
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label8.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label8.Location = New System.Drawing.Point(12, 450)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(1174, 2)
        Me.Label8.TabIndex = 406
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label9.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label9.Location = New System.Drawing.Point(12, 519)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(1174, 2)
        Me.Label9.TabIndex = 407
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label10.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label10.Location = New System.Drawing.Point(42, 43)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(2, 549)
        Me.Label10.TabIndex = 408
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label11.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label11.Location = New System.Drawing.Point(205, 43)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(2, 549)
        Me.Label11.TabIndex = 409
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label12.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label12.Location = New System.Drawing.Point(368, 43)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(2, 549)
        Me.Label12.TabIndex = 410
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label13.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label13.Location = New System.Drawing.Point(531, 43)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(2, 549)
        Me.Label13.TabIndex = 411
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label14.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label14.Location = New System.Drawing.Point(694, 43)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(2, 549)
        Me.Label14.TabIndex = 412
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label15.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label15.Location = New System.Drawing.Point(858, 43)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(2, 549)
        Me.Label15.TabIndex = 413
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label16.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label16.Location = New System.Drawing.Point(1020, 43)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(2, 549)
        Me.Label16.TabIndex = 414
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(1203, 622)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowTemplate.Height = 21
        Me.DataGridView3.Size = New System.Drawing.Size(190, 182)
        Me.DataGridView3.TabIndex = 415
        '
        'DataGridView4
        '
        Me.DataGridView4.AllowUserToAddRows = False
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(1240, 317)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.RowTemplate.Height = 21
        Me.DataGridView4.Size = New System.Drawing.Size(152, 228)
        Me.DataGridView4.TabIndex = 416
        '
        '週間表
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(1419, 836)
        Me.Controls.Add(Me.DataGridView4)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label52)
        Me.Controls.Add(Me.btnTorikomi)
        Me.Controls.Add(Me.btnInnsatu)
        Me.Controls.Add(Me.btnSakujo)
        Me.Controls.Add(Me.btnTouroku)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.rbn3F)
        Me.Controls.Add(Me.rbn2F)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.btnUp)
        Me.Controls.Add(Me.lblYmd)
        Me.Controls.Add(Me.DataGridView2)
        Me.Name = "週間表"
        Me.Text = "週間表"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblYmd As System.Windows.Forms.Label
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents rbn2F As System.Windows.Forms.RadioButton
    Friend WithEvents rbn3F As System.Windows.Forms.RadioButton
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents btnTouroku As System.Windows.Forms.Button
    Friend WithEvents btnSakujo As System.Windows.Forms.Button
    Friend WithEvents btnInnsatu As System.Windows.Forms.Button
    Friend WithEvents btnTorikomi As System.Windows.Forms.Button
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
End Class
