<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DB整理
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.deleteProgressBar = New System.Windows.Forms.ProgressBar()
        Me.btnDBDelete = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(44, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(347, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "処理効率向上のため、次のデータの５年以前を整理します"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(67, 90)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 14)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "勤務割表"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(67, 122)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 14)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "週間/食事担当表"
        '
        'deleteProgressBar
        '
        Me.deleteProgressBar.Location = New System.Drawing.Point(197, 167)
        Me.deleteProgressBar.Name = "deleteProgressBar"
        Me.deleteProgressBar.Size = New System.Drawing.Size(119, 16)
        Me.deleteProgressBar.TabIndex = 3
        '
        'btnDBDelete
        '
        Me.btnDBDelete.Location = New System.Drawing.Point(342, 162)
        Me.btnDBDelete.Name = "btnDBDelete"
        Me.btnDBDelete.Size = New System.Drawing.Size(66, 27)
        Me.btnDBDelete.TabIndex = 4
        Me.btnDBDelete.Text = "実行"
        Me.btnDBDelete.UseVisualStyleBackColor = True
        '
        'DB整理
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(448, 216)
        Me.Controls.Add(Me.btnDBDelete)
        Me.Controls.Add(Me.deleteProgressBar)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "DB整理"
        Me.Text = "DB整理"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents deleteProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents btnDBDelete As System.Windows.Forms.Button
End Class
