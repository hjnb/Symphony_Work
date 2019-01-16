<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 同姓略名
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.abbreviationTextBox = New System.Windows.Forms.TextBox()
        Me.btnRegist = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.dgvNam = New System.Windows.Forms.DataGridView()
        Me.namList = New System.Windows.Forms.ListBox()
        Me.namLabel = New System.Windows.Forms.Label()
        CType(Me.dgvNam, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(118, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "略氏名"
        '
        'abbreviationTextBox
        '
        Me.abbreviationTextBox.Location = New System.Drawing.Point(165, 47)
        Me.abbreviationTextBox.Name = "abbreviationTextBox"
        Me.abbreviationTextBox.Size = New System.Drawing.Size(79, 19)
        Me.abbreviationTextBox.TabIndex = 1
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(253, 12)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(41, 23)
        Me.btnRegist.TabIndex = 2
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(253, 45)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(41, 23)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "削除"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'dgvNam
        '
        Me.dgvNam.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvNam.Location = New System.Drawing.Point(113, 74)
        Me.dgvNam.Name = "dgvNam"
        Me.dgvNam.RowTemplate.Height = 21
        Me.dgvNam.Size = New System.Drawing.Size(188, 126)
        Me.dgvNam.TabIndex = 4
        '
        'namList
        '
        Me.namList.FormattingEnabled = True
        Me.namList.ItemHeight = 12
        Me.namList.Location = New System.Drawing.Point(4, 6)
        Me.namList.Name = "namList"
        Me.namList.Size = New System.Drawing.Size(104, 196)
        Me.namList.TabIndex = 5
        '
        'namLabel
        '
        Me.namLabel.AutoSize = True
        Me.namLabel.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.namLabel.ForeColor = System.Drawing.Color.Blue
        Me.namLabel.Location = New System.Drawing.Point(121, 17)
        Me.namLabel.Name = "namLabel"
        Me.namLabel.Size = New System.Drawing.Size(0, 16)
        Me.namLabel.TabIndex = 6
        '
        '同姓略名
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(305, 208)
        Me.Controls.Add(Me.namLabel)
        Me.Controls.Add(Me.namList)
        Me.Controls.Add(Me.dgvNam)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.abbreviationTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Name = "同姓略名"
        Me.Text = "同姓略名"
        CType(Me.dgvNam, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents abbreviationTextBox As System.Windows.Forms.TextBox
    Friend WithEvents btnRegist As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents dgvNam As System.Windows.Forms.DataGridView
    Friend WithEvents namList As System.Windows.Forms.ListBox
    Friend WithEvents namLabel As System.Windows.Forms.Label
End Class
