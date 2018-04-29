Imports System.Data.OleDb

Public Class workForm

    Private cn As OleDbConnection
    Private sqlCm As OleDbCommand
    Private adapter As OleDbDataAdapter
    Private bs As BindingSource

    Private Sub workForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        displayWorkTable("2018/04", "2")
        settingDgv(dgvWork)
    End Sub

    Private Sub displayWorkTable(ymStr As String, floar As String)
        Dim reader As OleDbDataReader
        Dim yoteiDay(31) As String
        Dim henkouDay(31) As String
        bs = New BindingSource()

        cn = New OleDbConnection(TopForm.DB_Work)
        sqlCm = cn.CreateCommand
        adapter = New OleDbDataAdapter(sqlCm)
        sqlCm.CommandText = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
        cn.Open()
        reader = sqlCm.ExecuteReader()
        While reader.Read() = True
            For i = 1 To 31
                yoteiDay(i) = reader("Y" & i)
                henkouDay(i) = reader("J" & i)
            Next
            bs.Add(New workData(reader("Unt"), reader("Rdr"), reader("Nam"), "予定", yoteiDay))
            bs.Add(New workData("", "", "", "変更", henkouDay))
        End While
        reader.Close()
        cn.Close()
        sqlCm.Dispose()
        cn.Dispose()

        dgvWork.DataSource = bs
    End Sub

    Private Sub settingDgv(dgv As DataGridView)
        TopForm.EnableDoubleBuffering(dgv)

        '並び替えができないようにする
        For Each c As DataGridViewColumn In dgv.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
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
        End With
        setReadonlyCell(dgv)
    End Sub

    Private Sub setReadonlyCell(dgv As DataGridView)

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
        If selectedRowIndex = -1 Then
            Return
        Else
            bs.Insert(selectedRowIndex, New workData("変更"))
            bs.Insert(selectedRowIndex, New workData("予定"))
        End If
    End Sub

    Private Sub btnRowDelete_Click(sender As Object, e As EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 Then
            Return
        Else
            bs.RemoveAt(selectedRowIndex)
            bs.RemoveAt(selectedRowIndex)
        End If
    End Sub
End Class