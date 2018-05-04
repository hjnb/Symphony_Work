Imports System.Data.OleDb

Public Class workForm

    Private cn As OleDbConnection
    Private sqlCm As OleDbCommand
    Private adapter As OleDbDataAdapter
    Private dt As DataTable

    Private dayCharArray() As String = {"日", "月", "火", "水", "木", "金", "土"}

    Private Sub workForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        '当月のデータを表示
        'とりあえず今は仮でこれ
        displayWorkTable("2018/04", "2")
    End Sub

    Private Sub displayWorkTable(ymStr As String, floar As String)
        dgvWork.Columns.Clear()
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        dt = New DataTable()

        cn = New OleDbConnection(TopForm.DB_Work)
        sqlCm = cn.CreateCommand
        adapter = New OleDbDataAdapter(sqlCm)
        sqlCm.CommandText = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
        cn.Open()
        adapter.Fill(dt)
        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(adapter)
        adapter.SelectCommand.Connection = cn

        addHenkouRow(dt)
        addDayCharRow(dt, year, month)
        addTypeColumn(dt)
        addBlankRow(dt)

        settingDgv(dgvWork)
        dgvWork.DataSource = dt
        settingDgvColumns(dgvWork)
    End Sub

    Private Sub addHenkouRow(dt As DataTable)
        Dim row As DataRow
        For i As Integer = 0 To dt.Rows.Count * 2 - 1 Step 2
            row = dt.NewRow()
            For j As Integer = 1 To 31
                If (Not IsDBNull(dt.Rows(i)("Y" & j)) AndAlso Not IsDBNull(dt.Rows(i)("J" & j))) AndAlso dt.Rows(i)("Y" & j) <> dt.Rows(i)("J" & j) Then
                    row("Y" & j) = dt.Rows(i)("J" & j)
                End If
            Next
            dt.Rows.InsertAt(row, i + 1)
        Next
    End Sub

    Private Sub addTypeColumn(dt As DataTable)
        dt.Columns.Add("type", Type.GetType("System.String")).SetOrdinal(6)
        For i As Integer = 1 To dt.Rows.Count - 1
            If i Mod 2 = 0 Then
                dt.Rows(i).Item("type") = "変更"
            Else
                dt.Rows(i).Item("type") = "予定"
            End If
        Next
    End Sub

    Private Sub addBlankRow(dt As DataTable)
        Dim rowCount As Integer = dt.Rows.Count
        If rowCount = 51 Then
            Return
        End If

        For i As Integer = rowCount To 50
            Dim row As DataRow = dt.NewRow()
            dt.Rows.Add(row)
        Next
    End Sub

    Private Sub settingDgv(dgv As DataGridView)
        TopForm.EnableDoubleBuffering(dgv)

        With dgv
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .RowHeadersVisible = False '行ヘッダー非表示
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .RowTemplate.Height = 16
            .ColumnHeadersHeight = 19
        End With
    End Sub

    Private Sub settingDgvColumns(dgv As DataGridView)
        With dgv
            '並び替えができないようにする
            For Each c As DataGridViewColumn In dgv.Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            '非表示列
            .Columns("Ym").Visible = False
            .Columns("Seq").Visible = False
            .Columns("Seq2").Visible = False
            For i As Integer = 1 To 31
                .Columns("J" & i).Visible = False
            Next

            'ユニット
            With .Columns("Unt")
                .Width = 35
                .HeaderText = "ﾕﾆｯﾄ"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            'R
            With .Columns("Rdr")
                .Width = 20
                .HeaderText = "R"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            '氏名
            With .Columns("Nam")
                .Width = 90
                .HeaderText = "氏名"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With

            '予定
            With .Columns("type")
                .Frozen = True
                .ReadOnly = True
                .Width = 30
                .HeaderText = ""
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            For i As Integer = 1 To 31
                With .Columns("Y" & i)
                    .Width = 43
                    .HeaderText = i.ToString()
                    .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End With
            Next

        End With
    End Sub

    Private Sub setReadonlyCell(dgv As DataGridView)

    End Sub

    Private Sub addDayCharRow(dt As DataTable, year As Integer, month As Integer)
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek)
        Dim dayArray(31) As String
        Dim row As DataRow = dt.NewRow()

        For i As Integer = 1 To daysInMonth
            row("Y" & i) = dayCharArray((weekNumber + (i - 1)) Mod 7)
        Next

        dt.Rows.InsertAt(row, 0)
    End Sub

    Private Sub rbtnF_MouseClick(sender As Object, e As MouseEventArgs) Handles rbtn2F.MouseClick, rbtn3F.MouseClick
        If sender Is rbtn2F Then
            displayWorkTable("2018/04", "2")
        ElseIf sender Is rbtn3F Then
            displayWorkTable("2018/04", "3")
        End If
    End Sub

    Private Sub btnRowAdd_Click(sender As Object, e As EventArgs) Handles btnRowAdd.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 Then
            Return
        Else
            Dim row1 As DataRow = dt.NewRow()
            Dim row2 As DataRow = dt.NewRow()
            row1("Type") = "変更"
            row2("Type") = "予定"

            dt.Rows.InsertAt(row1, selectedRowIndex)
            dt.Rows.InsertAt(row2, selectedRowIndex)
        End If
    End Sub

    Private Sub btnRowDelete_Click(sender As Object, e As EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 Then
            Return
        Else
            dt.Rows.RemoveAt(selectedRowIndex)
            dt.Rows.RemoveAt(selectedRowIndex)
        End If
    End Sub

    Private Sub deleteMonthData(ymStr As String, floar As String)
        cn = New OleDbConnection(TopForm.DB_Work)
        sqlCm = cn.CreateCommand
        sqlCm.CommandText = "delete from KinD where Ym='" & ymStr & "' and (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9'))"
        cn.Open()
        sqlCm.ExecuteNonQuery()
        cn.Close()
        sqlCm.Dispose()
        cn.Dispose()
    End Sub

    Private Sub setAddState(dt As DataTable)
        For Each row As DataRow In dt.Rows
            If Not IsDBNull(row("Nam")) AndAlso row("Nam") <> "" Then
                row.SetAdded()
            End If
        Next
    End Sub

    Private Sub btnRegist_Click(sender As Object, e As EventArgs) Handles btnRegist.Click
        Dim floar As String = If(rbtn2F.Checked = True, "2", "3")
        dt.AcceptChanges()
        setAddState(dt)
        deleteMonthData("2018/04", floar)
        adapter.Update(dt)
        MsgBox("登録しました。")
    End Sub
End Class