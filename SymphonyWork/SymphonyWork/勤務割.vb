Public Class 勤務割

    Private workDt As DataTable
    Private editBeforeCellValue As String

    Private unitDictionary2F As Dictionary(Of String, String)
    Private unitDictionary3F As Dictionary(Of String, String)
    Private wordDictionary As Dictionary(Of String, String)
    Private dayCharArray() As String = {"日", "月", "火", "水", "木", "金", "土"}

    Private disableCellStyle As DataGridViewCellStyle
    Private namColumnCellStyle As DataGridViewCellStyle
    Private sundayColumnCellStyle As DataGridViewCellStyle
    Private sundayCharCellStyle As DataGridViewCellStyle
    Private workChangeCellStyle As DataGridViewCellStyle

    Private Const INPUT_ROW_COUNT As Integer = 50
    Private Const READONLY_ROW_COUNT As Integer = 32

    Private Sub 勤務割_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        createDictionary()
        createCellStyles()

        initDgvWork()
        rbtn2F.Checked = True
    End Sub

    Private Sub createDictionary()
        'ﾕﾆｯﾄ(2F)の連想配列作成
        unitDictionary2F = New Dictionary(Of String, String)
        unitDictionary2F.Add("※", "00")
        unitDictionary2F.Add("星", "21")
        unitDictionary2F.Add("森", "22")
        unitDictionary2F.Add("空", "23")

        'ﾕﾆｯﾄ(3F)の連想配列作成
        unitDictionary3F = New Dictionary(Of String, String)
        unitDictionary3F.Add("※", "00")
        unitDictionary3F.Add("月", "31")
        unitDictionary3F.Add("花", "32")
        unitDictionary3F.Add("海", "33")

        'Y1～Y31の列のセルの入力文字変換連想配列
        wordDictionary = New Dictionary(Of String, String)
        wordDictionary.Add("0", "")
        wordDictionary.Add("1", "早")
        wordDictionary.Add("2", "日早")
        wordDictionary.Add("3", "日")
        wordDictionary.Add("4", "日遅")
        wordDictionary.Add("5", "遅")
        wordDictionary.Add("6", "遅々")
        wordDictionary.Add("7", "夜")
        wordDictionary.Add("8", "深")
        wordDictionary.Add("10", "半")
        wordDictionary.Add("11", "半Ａ")
        wordDictionary.Add("12", "半Ｂ")
        wordDictionary.Add("13", "半夜")
        wordDictionary.Add("21", "半行")
        wordDictionary.Add("22", "研")
        wordDictionary.Add("31", "有")
        wordDictionary.Add("32", "公")
        wordDictionary.Add("33", "明")
        wordDictionary.Add("34", "希")
        wordDictionary.Add("35", "産")
        wordDictionary.Add("36", "特")
    End Sub

    Private Sub createCellStyles()
        '曜日の行、(予定or変更)の列のスタイル
        disableCellStyle = New DataGridViewCellStyle()
        disableCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        disableCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        disableCellStyle.SelectionForeColor = Color.Black
        disableCellStyle.Font = New Font("MS UI Gothic", 9, FontStyle.Bold)

        '氏名の列のスタイル
        namColumnCellStyle = New DataGridViewCellStyle()
        namColumnCellStyle.ForeColor = Color.Blue
        namColumnCellStyle.SelectionForeColor = Color.Blue
        namColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        '日曜日の列のスタイル
        sundayColumnCellStyle = New DataGridViewCellStyle()
        sundayColumnCellStyle.BackColor = Color.FromArgb(255, 200, 200)
        sundayColumnCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '日の文字のセルのスタイル
        sundayCharCellStyle = New DataGridViewCellStyle()
        sundayCharCellStyle.BackColor = Color.FromArgb(255, 200, 200)
        sundayCharCellStyle.SelectionBackColor = Color.FromArgb(255, 200, 200)
        sundayCharCellStyle.SelectionForeColor = Color.Black
        sundayCharCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sundayCharCellStyle.Font = New Font("MS UI Gothic", 9, FontStyle.Bold)

        'Y1～Y31列の変更の行のスタイル
        workChangeCellStyle = New DataGridViewCellStyle()
        workChangeCellStyle.ForeColor = Color.Red
        workChangeCellStyle.SelectionForeColor = Color.Red
        workChangeCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub initDgvWork()
        'dgv設定
        With dgvWork
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False '行削除禁止
            .RowHeadersVisible = False '行ヘッダー非表示
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .RowTemplate.Height = 15.5
            .ColumnHeadersHeight = 19
            .ShowCellToolTips = False
            .EnableHeadersVisualStyles = False
        End With

    End Sub

    Private Sub setEmptyCell()
        dgvWork.Columns.Clear()

        workDt = New DataTable()

        '列定義
        workDt.Columns.Add("Ym", Type.GetType("System.String"))
        workDt.Columns.Add("Seq", Type.GetType("System.String"))
        workDt.Columns.Add("Seq2", Type.GetType("System.String"))
        workDt.Columns.Add("Unt", Type.GetType("System.String"))
        workDt.Columns.Add("Rdr", Type.GetType("System.String"))
        workDt.Columns.Add("Nam", Type.GetType("System.String"))
        workDt.Columns.Add("Type", Type.GetType("System.String"))
        For i As Integer = 1 To 31
            workDt.Columns.Add("Y" & i, Type.GetType("System.String"))
        Next
        workDt.Columns.Add("月合計", Type.GetType("System.String"))
        workDt.Columns.Add("早出", Type.GetType("System.String"))
        workDt.Columns.Add("日早", Type.GetType("System.String"))
        workDt.Columns.Add("日勤", Type.GetType("System.String"))
        workDt.Columns.Add("日遅", Type.GetType("System.String"))
        workDt.Columns.Add("遅出", Type.GetType("System.String"))
        workDt.Columns.Add("遅々", Type.GetType("System.String"))
        workDt.Columns.Add("夜勤", Type.GetType("System.String"))
        workDt.Columns.Add("深夜", Type.GetType("System.String"))
        workDt.Columns.Add("半", Type.GetType("System.String"))
        workDt.Columns.Add("半A", Type.GetType("System.String"))
        workDt.Columns.Add("半B", Type.GetType("System.String"))
        workDt.Columns.Add("半夜", Type.GetType("System.String"))
        workDt.Columns.Add("半行", Type.GetType("System.String"))
        workDt.Columns.Add("研修", Type.GetType("System.String"))
        workDt.Columns.Add("公休", Type.GetType("System.String"))
        workDt.Columns.Add("明等", Type.GetType("System.String"))

        '空行追加
        For i = 0 To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
            workDt.Rows.Add(workDt.NewRow())
        Next

        '表示
        dgvWork.DataSource = workDt
    End Sub

    Private Sub settingDgvWorkColumnsAndRows(year As Integer, month As Integer)
        '空セル表示
        setEmptyCell()

        '曜日設定
        setDayCharRow(year, month)

        '列設定
        With dgvWork
            '並び替えができないようにする
            For Each c As DataGridViewColumn In .Columns
                c.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            '非表示列
            .Columns("Ym").Visible = False
            .Columns("Seq").Visible = False
            .Columns("Seq2").Visible = False

            '行固定
            .Rows(0).Frozen = True

            '列固定
            .Columns("type").Frozen = True

            'ユニット列
            With .Columns("Unt")
                .Width = 32
                .HeaderText = "ﾕﾆｯﾄ"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            'R列
            With .Columns("Rdr")
                .Width = 20
                .HeaderText = "R"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With

            '氏名列
            With .Columns("Nam")
                .Width = 90
                .HeaderText = "氏名"
                .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .DefaultCellStyle = namColumnCellStyle
            End With

            '予定or変更列
            With .Columns("type")
                .Width = 32
                .HeaderText = ""
                .DefaultCellStyle = disableCellStyle
            End With

            'Y1～Y31の列
            For i As Integer = 1 To 31
                With .Columns("Y" & i)
                    .Width = 43
                    .HeaderText = i.ToString()
                    .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End With
            Next

            'Y1～Y31の列の変更の行
            For i As Integer = 2 To INPUT_ROW_COUNT Step 2
                For j As Integer = 1 To 31
                    dgvWork("Y" & j, i).Style = workChangeCellStyle
                Next
            Next

            '日曜日の列
            For i As Integer = 1 To 31
                If Not IsDBNull(workDt.Rows(0).Item("Y" & i)) AndAlso workDt.Rows(0).Item("Y" & i) = "日" Then
                    dgvWork("Y" & i, 0).Style = sundayCharCellStyle
                    dgvWork.Columns("Y" & i).DefaultCellStyle = sundayColumnCellStyle
                End If
            Next

            '日曜日以外の曜日の行
            For Each cell As DataGridViewCell In dgvWork.Rows(0).Cells
                If IsDBNull(cell.Value) OrElse cell.Value <> "日" Then
                    cell.Style = disableCellStyle
                End If
            Next

            '小計記載行
            For i As Integer = 1 + INPUT_ROW_COUNT To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
                .Rows(i).DefaultCellStyle = disableCellStyle
            Next

            '小計記載列
            .Columns(38).DefaultCellStyle = disableCellStyle
            .Columns(38).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(38).Width = 60
            For i As Integer = 39 To 54
                .Columns(i).DefaultCellStyle = disableCellStyle
                .Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(i).Width = 35
            Next

            'ReadOnlyセルの設定
            setReadonlyCell()

        End With
    End Sub

    Private Sub setDayCharRow(year As Integer, month As Integer)
        Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)
        Dim firstDay As DateTime = New DateTime(year, month, 1)
        Dim weekNumber As Integer = CInt(firstDay.DayOfWeek)
        Dim row As DataRow = workDt.Rows(0)

        For i As Integer = 1 To daysInMonth
            row("Y" & i) = dayCharArray((weekNumber + (i - 1)) Mod 7)
        Next
    End Sub

    Private Sub setReadonlyCell()
        With dgvWork
            '曜日の行
            .Rows(0).ReadOnly = True

            '予定or変更の列
            .Columns("type").ReadOnly = True

            '変更の行のﾕﾆｯﾄ、R、Nam列のセル
            For i As Integer = 2 To dgvWork.Rows.Count - 1 Step 2
                dgvWork("Unt", i).ReadOnly = True
                dgvWork("Rdr", i).ReadOnly = True
                dgvWork("Nam", i).ReadOnly = True
            Next

            '小計記載行
            For i As Integer = 1 + INPUT_ROW_COUNT To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
                .Rows(i).ReadOnly = True
            Next

            '小計記載列
            For i As Integer = 38 To 54
                dgvWork.Columns(i).ReadOnly = True
            Next
        End With
    End Sub

    Private Sub displayWork(ymStr As String, floar As String)
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        settingDgvWorkColumnsAndRows(year, month)

        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            Dim warekiStr As String = Util.convADStrToWarekiStr(ymStr & "/01")
            Dim eraStr As String = warekiStr.Substring(0, 3)
            Dim monthStr As String = warekiStr.Substring(4, 2)
            Dim result As DialogResult = MessageBox.Show(eraStr & "年" & monthStr & "月分は登録されていません" & Environment.NewLine & "登録しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If result = Windows.Forms.DialogResult.Yes Then
                '前月データの名前を表示
                '
                '

                rs.Close()
                cnn.Close()
                Return
            Else
                rs.Close()
                cnn.Close()
                Return
            End If
        Else
            '表示処理
            Dim rowIndex As Integer = 1
            While Not rs.EOF
                '予定行の値設定
                dgvWork("Ym", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Ym").Value)
                dgvWork("Seq", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq").Value)
                dgvWork("Seq2", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq2").Value)
                dgvWork("Unt", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Unt").Value)
                dgvWork("Rdr", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Rdr").Value)
                dgvWork("Nam", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                dgvWork("Type", rowIndex).Value = "予定"
                For i As Integer = 1 To 31
                    dgvWork("Y" & i, rowIndex).Value = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                Next

                '変更行の値設定
                dgvWork("Type", (rowIndex + 1)).Value = "変更"
                For i As Integer = 1 To 31
                    dgvWork("Y" & i, (rowIndex + 1)).Value = If(Util.checkDBNullValue(rs.Fields("J" & i).Value) = Util.checkDBNullValue(rs.Fields("Y" & i).Value), "", Util.checkDBNullValue(rs.Fields("J" & i).Value))
                Next

                rowIndex += 2
                rs.MoveNext()
            End While
            rs.Close()
            cnn.Close()
        End If

    End Sub

    Private Sub ymBox_YmLabelTextChange(sender As Object, e As System.EventArgs) Handles ymBox.YmLabelTextChange
        Dim ym As String = ymBox.getADStr4Ym()
        Dim floar As String = If(rbtn2F.Checked, "2", "3")
        displayWork(ym, floar)
    End Sub

    Private Sub floarRadioButton_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtn2F.CheckedChanged, rbtn3F.CheckedChanged
        Dim rbtn As RadioButton = CType(sender, RadioButton)
        If rbtn.Checked = True Then
            rbtn.BackColor = Color.FromArgb(255, 255, 0)
            Dim floar As String = rbtn.Name.Substring(4, 1)
            displayWork(ymBox.getADStr4Ym(), floar)
        Else
            rbtn.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub

    Private Sub dgvWork_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvWork.CellBeginEdit
        editBeforeCellValue = If(IsDBNull(dgvWork(e.ColumnIndex, e.RowIndex).Value), "", dgvWork(e.ColumnIndex, e.RowIndex).Value)
    End Sub

    Private Sub dgvWork_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvWork.CellEndEdit
        If dgvWork.Columns(e.ColumnIndex).Name = "Unt" Then
            'ﾕﾆｯﾄ列の編集終了時、Seq2列のセルに対応した値を設定
            Dim inputStr As String = If(IsDBNull(dgvWork(e.ColumnIndex, e.RowIndex).Value), "", dgvWork(e.ColumnIndex, e.RowIndex).Value)
            Try
                If rbtn2F.Checked = True Then
                    dgvWork("Seq2", e.RowIndex).Value = unitDictionary2F(inputStr)
                Else
                    dgvWork("Seq2", e.RowIndex).Value = unitDictionary3F(inputStr)
                End If
            Catch ex As KeyNotFoundException
                dgvWork(e.ColumnIndex, e.RowIndex).Value = editBeforeCellValue
                MsgBox("正しいﾕﾆｯﾄ名を入力してください。")
            End Try
        ElseIf 7 <= e.ColumnIndex AndAlso e.ColumnIndex <= 37 Then
            'Y1～Y31列の編集終了時の処理、値の変換処理をする
            Try
                dgvWork(e.ColumnIndex, e.RowIndex).Value = wordDictionary(dgvWork(e.ColumnIndex, e.RowIndex).Value)
            Catch ex As KeyNotFoundException
                '何もしない
            End Try
        End If
    End Sub
End Class