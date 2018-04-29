<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TopForm
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
        Me.btnWork = New System.Windows.Forms.Button()
        Me.btnWeek = New System.Windows.Forms.Button()
        Me.btnCreateCSV = New System.Windows.Forms.Button()
        Me.btnArrangementDB = New System.Windows.Forms.Button()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'btnWork
        '
        Me.btnWork.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnWork.Font = New System.Drawing.Font("ＭＳ ゴシック", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnWork.ForeColor = System.Drawing.Color.Blue
        Me.btnWork.Location = New System.Drawing.Point(168, 59)
        Me.btnWork.Name = "btnWork"
        Me.btnWork.Size = New System.Drawing.Size(301, 132)
        Me.btnWork.TabIndex = 0
        Me.btnWork.Text = "勤務割表"
        Me.btnWork.UseVisualStyleBackColor = False
        '
        'btnWeek
        '
        Me.btnWeek.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnWeek.Font = New System.Drawing.Font("ＭＳ ゴシック", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnWeek.ForeColor = System.Drawing.Color.Blue
        Me.btnWeek.Location = New System.Drawing.Point(168, 194)
        Me.btnWeek.Name = "btnWeek"
        Me.btnWeek.Size = New System.Drawing.Size(301, 132)
        Me.btnWeek.TabIndex = 1
        Me.btnWeek.Text = "週間/食事担当表"
        Me.btnWeek.UseVisualStyleBackColor = False
        '
        'btnCreateCSV
        '
        Me.btnCreateCSV.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnCreateCSV.Font = New System.Drawing.Font("ＭＳ ゴシック", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnCreateCSV.ForeColor = System.Drawing.Color.Blue
        Me.btnCreateCSV.Location = New System.Drawing.Point(168, 329)
        Me.btnCreateCSV.Name = "btnCreateCSV"
        Me.btnCreateCSV.Size = New System.Drawing.Size(301, 132)
        Me.btnCreateCSV.TabIndex = 2
        Me.btnCreateCSV.Text = "勤務割CSV作成"
        Me.btnCreateCSV.UseVisualStyleBackColor = False
        '
        'btnArrangementDB
        '
        Me.btnArrangementDB.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.btnArrangementDB.Font = New System.Drawing.Font("ＭＳ ゴシック", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnArrangementDB.ForeColor = System.Drawing.Color.Blue
        Me.btnArrangementDB.Location = New System.Drawing.Point(168, 465)
        Me.btnArrangementDB.Name = "btnArrangementDB"
        Me.btnArrangementDB.Size = New System.Drawing.Size(301, 132)
        Me.btnArrangementDB.TabIndex = 3
        Me.btnArrangementDB.Text = "DB整理"
        Me.btnArrangementDB.UseVisualStyleBackColor = False
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Checked = True
        Me.rbtnPreview.Location = New System.Drawing.Point(602, 171)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(67, 16)
        Me.rbtnPreview.TabIndex = 4
        Me.rbtnPreview.TabStop = True
        Me.rbtnPreview.Text = "プレビュー"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(688, 171)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(47, 16)
        Me.RadioButton1.TabIndex = 5
        Me.RadioButton1.Text = "印刷"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(936, 654)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.btnArrangementDB)
        Me.Controls.Add(Me.btnCreateCSV)
        Me.Controls.Add(Me.btnWeek)
        Me.Controls.Add(Me.btnWork)
        Me.Name = "TopForm"
        Me.Text = "Work 特養勤務割"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnWork As System.Windows.Forms.Button
    Friend WithEvents btnWeek As System.Windows.Forms.Button
    Friend WithEvents btnCreateCSV As System.Windows.Forms.Button
    Friend WithEvents btnArrangementDB As System.Windows.Forms.Button
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton

End Class
