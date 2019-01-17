<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CSV作成
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
        Me.ymBox = New ymdBox.ymdBox()
        Me.btnExecution = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.saveCsvFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'ymBox
        '
        Me.ymBox.boxType = 7
        Me.ymBox.DateText = ""
        Me.ymBox.EraLabelText = "H31"
        Me.ymBox.EraText = ""
        Me.ymBox.Location = New System.Drawing.Point(110, 21)
        Me.ymBox.MonthLabelText = "01"
        Me.ymBox.MonthText = ""
        Me.ymBox.Name = "ymBox"
        Me.ymBox.Size = New System.Drawing.Size(120, 46)
        Me.ymBox.TabIndex = 0
        '
        'btnExecution
        '
        Me.btnExecution.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnExecution.Location = New System.Drawing.Point(122, 106)
        Me.btnExecution.Name = "btnExecution"
        Me.btnExecution.Size = New System.Drawing.Size(97, 39)
        Me.btnExecution.TabIndex = 1
        Me.btnExecution.Text = "実行"
        Me.btnExecution.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 19)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "対象年月"
        '
        'CSV作成
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 396)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExecution)
        Me.Controls.Add(Me.ymBox)
        Me.Name = "CSV作成"
        Me.Text = "勤務割CSV作成"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ymBox As ymdBox.ymdBox
    Friend WithEvents btnExecution As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents saveCsvFileDialog As System.Windows.Forms.SaveFileDialog
End Class
