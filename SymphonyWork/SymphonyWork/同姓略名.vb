Imports System.Data.OleDb

Public Class 同姓略名

    Private selectedYmstr As String

    Public Sub New(ymStr As String)
        InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterScreen
        selectedYmstr = ymStr
    End Sub

    '行ヘッダーのカレントセルを表す三角マークを非表示に設定する為のクラス。
    Public Class dgvRowHeaderCell

        'DataGridViewRowHeaderCell を継承
        Inherits DataGridViewRowHeaderCell

        'DataGridViewHeaderCell.Paint をオーバーライドして行ヘッダーを描画
        Protected Overrides Sub Paint(ByVal graphics As Graphics, ByVal clipBounds As Rectangle, _
           ByVal cellBounds As Rectangle, ByVal rowIndex As Integer, ByVal cellState As DataGridViewElementStates, _
           ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, _
           ByVal cellStyle As DataGridViewCellStyle, ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
           ByVal paintParts As DataGridViewPaintParts)
            '標準セルの描画からセル内容の背景だけ除いた物を描画(-5)
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, _
                     formattedValue, errorText, cellStyle, advancedBorderStyle, _
                     Not DataGridViewPaintParts.ContentBackground)
        End Sub

    End Class

    Private Sub 同姓略名_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '設定
        settingComponent()

        'リスト表示
        displayNamList()

        'dgv表示
        displayDgvNam()
    End Sub

    Private Sub displayNamList()
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & selectedYmstr & "' order by Seq2, Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        While Not rs.EOF
            namList.Items.Add(Util.checkDBNullValue(rs.Fields("Nam").Value))
            rs.MoveNext()
        End While
        rs.Close()
        cnn.Close()
    End Sub

    Private Sub displayDgvNam()
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT Nam, NNam FROM SNam order by Nam"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        da.Fill(ds, rs, "SNam")
        dgvNam.DataSource = ds.Tables("SNam")
        cnn.Close()

        '列設定等
        With dgvNam.Columns("Nam")
            .HeaderText = "氏名"
            .Width = 88
        End With
        With dgvNam.Columns("NNam")
            .HeaderText = "略氏名"
            .Width = 56
        End With

        '並べ替え禁止
        For Each c As DataGridViewColumn In dgvNam.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgvNam.CurrentCell = Nothing
    End Sub

    Private Sub settingComponent()
        'リストボックス
        namList.BackColor = Color.FromKnownColor(KnownColor.Control) '背景色

        'データグリッドビュー
        With dgvNam
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.Height = 15
            .ColumnHeadersHeight = 19
            .RowHeadersWidth = 25
            .ShowCellToolTips = False
            .RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .EnableHeadersVisualStyles = False
        End With
        Util.EnableDoubleBuffering(dgvNam)

        'テキストボックス



    End Sub

    Private Sub namList_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles namList.MouseClick
        namLabel.Text = namList.SelectedItem
    End Sub

    Private Sub dgvNam_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvNam.CellMouseClick
        If e.RowIndex >= 0 Then
            namLabel.Text = dgvNam("Nam", e.RowIndex).Value
            abbreviationTextBox.Text = dgvNam("NNam", e.RowIndex).Value
        End If
    End Sub

    Private Sub dgvNam_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvNam.CellPainting
        '行ヘッダーかどうか調べる
        If e.ColumnIndex < 0 AndAlso e.RowIndex >= 0 Then
            'セルを描画する
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する
            'e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)
            '行番号を描画する
            TextRenderer.DrawText(e.Graphics, _
                (e.RowIndex + 1).ToString(), _
                e.CellStyle.Font, _
                indexRect, _
                e.CellStyle.ForeColor, _
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            '描画が完了したことを知らせる
            e.Handled = True
        End If
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

    End Sub
End Class