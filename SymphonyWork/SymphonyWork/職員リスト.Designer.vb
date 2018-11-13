<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 職員リスト
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
        Me.btnPaint = New System.Windows.Forms.Button()
        Me.lstName = New System.Windows.Forms.ListBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnPaint
        '
        Me.btnPaint.Location = New System.Drawing.Point(21, 346)
        Me.btnPaint.Name = "btnPaint"
        Me.btnPaint.Size = New System.Drawing.Size(96, 27)
        Me.btnPaint.TabIndex = 15
        Me.btnPaint.Text = "PAINT"
        Me.btnPaint.UseVisualStyleBackColor = True
        '
        'lstName
        '
        Me.lstName.FormattingEnabled = True
        Me.lstName.ItemHeight = 12
        Me.lstName.Location = New System.Drawing.Point(21, 12)
        Me.lstName.Name = "lstName"
        Me.lstName.Size = New System.Drawing.Size(96, 328)
        Me.lstName.TabIndex = 14
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(107, 363)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 21
        Me.DataGridView1.Size = New System.Drawing.Size(10, 10)
        Me.DataGridView1.TabIndex = 16
        '
        '職員リスト
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(139, 392)
        Me.Controls.Add(Me.btnPaint)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.lstName)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "職員リスト"
        Me.Text = "職員リスト"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnPaint As System.Windows.Forms.Button
    Friend WithEvents lstName As System.Windows.Forms.ListBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
End Class
