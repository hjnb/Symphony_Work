Public Class 勤務割

    Private workDt As DataTable

    Private unitDictionary As Dictionary(Of String, String)
    Private wordDictionary As Dictionary(Of String, String)
    Private workTimeDictionary As Dictionary(Of String, Double)
    Private abbreviationDictionary As Dictionary(Of String, String)
    Private subtotalStrIndexDictionary As Dictionary(Of String, Integer)
    Private dayCharArray() As String = {"日", "月", "火", "水", "木", "金", "土"}

    Private disableCellStyle As DataGridViewCellStyle
    Private namColumnCellStyle As DataGridViewCellStyle
    Private sundayColumnCellStyle As DataGridViewCellStyle
    Private sundayCharCellStyle As DataGridViewCellStyle
    Private workChangeCellStyle As DataGridViewCellStyle
    Private subtotalPlanCellStyle As DataGridViewCellStyle
    Private subtotalChangeCellStyle As DataGridViewCellStyle

    Private Const INPUT_ROW_COUNT As Integer = 50
    Private Const READONLY_ROW_COUNT As Integer = 32

    Private abbreviationNamForm As 同姓略名

    Private Sub 勤務割_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.F12 Then
            btnRowAdd.Visible = Not btnRowAdd.Visible
            btnRowDelete.Visible = Not btnRowDelete.Visible
            btnRegist.Visible = Not btnRegist.Visible
            btnDelete.Visible = Not btnDelete.Visible
            btnPrint.Visible = Not btnPrint.Visible
            wordPanel.Visible = Not wordPanel.Visible
        End If

        If e.Alt AndAlso e.KeyCode = Keys.F11 Then
            If IsNothing(abbreviationNamForm) OrElse abbreviationNamForm.IsDisposed Then
                abbreviationNamForm = New 同姓略名(ymBox.getADStr4Ym())
                abbreviationNamForm.Show()
            End If
        End If
    End Sub

    Private Sub 勤務割_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.KeyPreview = True

        createDictionary()
        createCellStyles()

        initDgvWork()
        rbtn2F.Checked = True
    End Sub

    Private Sub createDictionary()
        'ﾕﾆｯﾄの連想配列作成
        unitDictionary = New Dictionary(Of String, String)
        unitDictionary.Add("※", "00")
        unitDictionary.Add("星", "21")
        unitDictionary.Add("森", "22")
        unitDictionary.Add("空", "23")
        unitDictionary.Add("月", "31")
        unitDictionary.Add("花", "32")
        unitDictionary.Add("海", "33")

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

        '略語
        abbreviationDictionary = New Dictionary(Of String, String)
        abbreviationDictionary.Add("早", "早出")
        abbreviationDictionary.Add("日早", "日早")
        abbreviationDictionary.Add("日", "日勤")
        abbreviationDictionary.Add("日遅", "日遅")
        abbreviationDictionary.Add("遅", "遅出")
        abbreviationDictionary.Add("遅々", "遅々")
        abbreviationDictionary.Add("夜", "夜勤")
        abbreviationDictionary.Add("深", "深夜")
        abbreviationDictionary.Add("半", "半")
        abbreviationDictionary.Add("半Ａ", "半Ａ")
        abbreviationDictionary.Add("半Ｂ", "半Ｂ")
        abbreviationDictionary.Add("半夜", "半夜")
        abbreviationDictionary.Add("半行", "半行")
        abbreviationDictionary.Add("研", "研修")
        abbreviationDictionary.Add("公", "公休")
        abbreviationDictionary.Add("明", "明等")

        '勤務時間
        workTimeDictionary = New Dictionary(Of String, Double)
        workTimeDictionary.Add("早", 7.5)
        workTimeDictionary.Add("日早", 7.5)
        workTimeDictionary.Add("日", 7.5)
        workTimeDictionary.Add("日遅", 7.5)
        workTimeDictionary.Add("遅", 7.5)
        workTimeDictionary.Add("遅々", 7.5)
        workTimeDictionary.Add("夜", 15.0)
        workTimeDictionary.Add("深", 7.5)
        workTimeDictionary.Add("半", 3.5)
        workTimeDictionary.Add("半Ａ", 3.5)
        workTimeDictionary.Add("半Ｂ", 3.5)
        workTimeDictionary.Add("半夜", 3.5)
        workTimeDictionary.Add("半行", 3.5)
        workTimeDictionary.Add("研", 7.5)
        workTimeDictionary.Add("有", 0.0)
        workTimeDictionary.Add("公", 7.5)
        workTimeDictionary.Add("明", 0.0)
        workTimeDictionary.Add("希", 0.0)
        workTimeDictionary.Add("産", 0.0)
        workTimeDictionary.Add("特", 0.0)
        workTimeDictionary.Add("A", 5.0)
        workTimeDictionary.Add("B", 5.5)
        workTimeDictionary.Add("C", 7.0)
        workTimeDictionary.Add("D", 3.5)
        workTimeDictionary.Add("E", 5.0)
        workTimeDictionary.Add("F", 6.0)
        workTimeDictionary.Add("G", 7.0)
        workTimeDictionary.Add("H", 4.0)
        workTimeDictionary.Add("I", 3.0)
        workTimeDictionary.Add("J", 5.5)
        workTimeDictionary.Add("K", 7.0)
        workTimeDictionary.Add("L", 2.5)
        workTimeDictionary.Add("M", 3.5)
        workTimeDictionary.Add("N", 2.0)
        workTimeDictionary.Add("P", 6.5)
        workTimeDictionary.Add("R", 2.5)
        workTimeDictionary.Add("S", 7.5)
        workTimeDictionary.Add("T", 4.5)

        '小計の行インデックス
        subtotalStrIndexDictionary = New Dictionary(Of String, Integer)
        subtotalStrIndexDictionary.Add("早出", 51)
        subtotalStrIndexDictionary.Add("日早", 53)
        subtotalStrIndexDictionary.Add("日勤", 55)
        subtotalStrIndexDictionary.Add("日遅", 57)
        subtotalStrIndexDictionary.Add("遅出", 59)
        subtotalStrIndexDictionary.Add("遅々", 61)
        subtotalStrIndexDictionary.Add("夜勤", 63)
        subtotalStrIndexDictionary.Add("深夜", 65)
        subtotalStrIndexDictionary.Add("半", 67)
        subtotalStrIndexDictionary.Add("半Ａ", 69)
        subtotalStrIndexDictionary.Add("半Ｂ", 71)
        subtotalStrIndexDictionary.Add("半夜", 73)
        subtotalStrIndexDictionary.Add("半行", 75)
        subtotalStrIndexDictionary.Add("研修", 77)
        subtotalStrIndexDictionary.Add("公休", 79)

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
        workChangeCellStyle.Font = New Font("MS UI Gothic", 8.5)

        '小計予定セルスタイル
        subtotalPlanCellStyle = New DataGridViewCellStyle()
        subtotalPlanCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalPlanCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalPlanCellStyle.ForeColor = Color.Blue
        subtotalPlanCellStyle.SelectionForeColor = Color.Blue
        subtotalPlanCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        '小計変更セルスタイル
        subtotalChangeCellStyle = New DataGridViewCellStyle()
        subtotalChangeCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalChangeCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        subtotalChangeCellStyle.ForeColor = Color.Red
        subtotalChangeCellStyle.SelectionForeColor = Color.Red
        subtotalChangeCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

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
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 8.5)
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
        workDt.Columns.Add("半Ａ", Type.GetType("System.String"))
        workDt.Columns.Add("半Ｂ", Type.GetType("System.String"))
        workDt.Columns.Add("半夜", Type.GetType("System.String"))
        workDt.Columns.Add("半行", Type.GetType("System.String"))
        workDt.Columns.Add("研修", Type.GetType("System.String"))
        workDt.Columns.Add("公休", Type.GetType("System.String"))
        workDt.Columns.Add("明等", Type.GetType("System.String"))

        '空行追加
        For i = 0 To 1 + INPUT_ROW_COUNT + READONLY_ROW_COUNT
            workDt.Rows.Add(workDt.NewRow())
        Next

        '小計項目名
        Dim itemArray As String() = {"早出", "日早", "日勤", "日遅", "遅出", "遅々", "夜勤", "深夜", "半", "半Ａ", "半Ｂ", "半夜", "半行", "研修", "公休"}
        Dim index As Integer = 0
        For i As Integer = 51 To 79 Step 2
            workDt.Rows(i).Item("Nam") = itemArray(index)
            index += 1
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
                    .Width = 50
                    .HeaderText = i.ToString()
                    .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End With
            Next

            'Y1～Y31,小計列の変更の行
            For i As Integer = 2 To INPUT_ROW_COUNT Step 2
                For j As Integer = 1 To 31
                    dgvWork("Y" & j, i).Style = workChangeCellStyle
                Next
                For k As Integer = 38 To 54
                    dgvWork(k, i).Style = subtotalChangeCellStyle
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
                If i Mod 2 = 1 Then
                    .Rows(i).DefaultCellStyle = subtotalPlanCellStyle
                Else
                    .Rows(i).DefaultCellStyle = subtotalChangeCellStyle
                End If
            Next

            '小計記載列
            .Columns(38).DefaultCellStyle = subtotalPlanCellStyle
            .Columns(38).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(38).Width = 60
            For i As Integer = 39 To 54
                .Columns(i).DefaultCellStyle = subtotalPlanCellStyle
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

    Private Sub setSeqValue()
        For i As Integer = 1 To INPUT_ROW_COUNT Step 2
            workDt.Rows(i).Item("Seq") = i + 1
        Next
    End Sub

    Private Sub displayWork(ymStr As String, floar As String, Optional deleteAfterFlg As Boolean = False)
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))
        settingDgvWorkColumnsAndRows(year, month)
        setSeqValue()

        If deleteAfterFlg Then
            Return
        End If

        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            Dim warekiStr As String = Util.convADStrToWarekiStr(ymStr & "/01")
            Dim eraStr As String = warekiStr.Substring(0, 3)
            Dim monthStr As String = warekiStr.Substring(4, 2)
            Dim result As DialogResult = MessageBox.Show(eraStr & "年" & monthStr & "月分は登録されていません" & Environment.NewLine & "登録しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If result = Windows.Forms.DialogResult.Yes Then
                rs.Close()

                '前月データの名前を表示
                Dim prevMonth As String
                Dim prevYear As String
                If month - 1 = 0 Then
                    prevYear = (year - 1).ToString()
                    prevMonth = "12"
                Else
                    prevYear = year.ToString()
                    prevMonth = If(month - 1 >= 10, (month - 1).ToString(), "0" & (month - 1).ToString())
                End If
                Dim prevYmStr As String = prevYear & "/" & prevMonth

                sql = "SELECT * FROM KinD WHERE YM='" & prevYmStr & "' AND ((Seq2='00' AND Unt='※') OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
                rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

                Dim rowIndex As Integer = 1
                While Not rs.EOF
                    '予定行の値設定
                    dgvWork("Seq", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq").Value)
                    dgvWork("Seq2", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Seq2").Value)
                    dgvWork("Unt", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Unt").Value)
                    dgvWork("Rdr", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Rdr").Value)
                    dgvWork("Nam", rowIndex).Value = Util.checkDBNullValue(rs.Fields("Nam").Value)
                    dgvWork("Type", rowIndex).Value = "予定"

                    '変更行の値設定
                    dgvWork("Type", (rowIndex + 1)).Value = "変更"

                    rowIndex += 2
                    rs.MoveNext()
                End While

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

    Private Sub dgvWork_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvWork.CellEndEdit
        Dim inputStr As String = If(IsDBNull(dgvWork(e.ColumnIndex, e.RowIndex).Value), "", dgvWork(e.ColumnIndex, e.RowIndex).Value)
        If dgvWork.Columns(e.ColumnIndex).Name = "Unt" Then
            'ﾕﾆｯﾄ列の編集終了時、Seq2列のセルに対応した値を設定
            Try
                dgvWork("Seq2", e.RowIndex).Value = unitDictionary(inputStr)
            Catch ex As KeyNotFoundException
                '何もしない
            End Try
        ElseIf 7 <= e.ColumnIndex AndAlso e.ColumnIndex <= 37 Then
            'Y1～Y31列の編集終了時の処理、値の変換処理をする
            Try
                dgvWork(e.ColumnIndex, e.RowIndex).Value = wordDictionary(inputStr)
            Catch ex As KeyNotFoundException
                '何もしない
            End Try
        End If
    End Sub

    Private Sub monthDataDelete(ymStr As String, floar As String, cnn As ADODB.Connection)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cnn
        cmd.CommandText = "delete from KinD where YM='" & ymStr & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9'))"
        cmd.Execute()
    End Sub

    Private Function existsWorkStr(row As DataGridViewRow) As Boolean
        For i As Integer = 1 To 31
            If Util.checkDBNullValue(row.Cells("Y" & i).Value) <> "" Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub subtotalClear()
        '小計列のクリア
        For i As Integer = 38 To 54
            For j As Integer = 1 To 50
                dgvWork(i, j).Value = ""
            Next
        Next

        '小計行のクリア
        For i As Integer = 51 To 80
            For j As Integer = 1 To 31
                dgvWork("Y" & j, i).Value = ""
            Next
        Next
    End Sub

    Private Sub btnRowAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnRowAdd.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 Then
            Return
        ElseIf Not IsDBNull(workDt.Rows(INPUT_ROW_COUNT - 1).Item("Nam")) AndAlso workDt.Rows(INPUT_ROW_COUNT - 1).Item("Nam") <> "" Then
            MsgBox("行挿入できません。")
            Return
        Else
            '変更の行を選択してる場合は予定の行を選択しているindexとする
            If selectedRowIndex Mod 2 = 0 Then
                selectedRowIndex -= 1
            End If

            Dim rowJ As DataRow = workDt.NewRow()
            Dim rowY As DataRow = workDt.NewRow()
            rowY("Seq") = selectedRowIndex + 1

            '行追加
            workDt.Rows.InsertAt(rowJ, selectedRowIndex) '変更行
            workDt.Rows.InsertAt(rowY, selectedRowIndex) '予定行

            '追加した変更行の設定
            dgvWork("Unt", selectedRowIndex + 1).ReadOnly = True
            dgvWork("Rdr", selectedRowIndex + 1).ReadOnly = True
            dgvWork("Nam", selectedRowIndex + 1).ReadOnly = True
            For i As Integer = 1 To 31 'Y1～Y31列のセルスタイル設定
                dgvWork("Y" & i, selectedRowIndex + 1).Style = workChangeCellStyle
            Next
            For i As Integer = 38 To 54 '小計部分のセルスタイル設定
                dgvWork(i, selectedRowIndex + 1).Style = subtotalChangeCellStyle
            Next

            '追加された行以降のSeqの値を更新
            For i As Integer = selectedRowIndex + 2 To INPUT_ROW_COUNT - 1 Step 2
                workDt.Rows(i).Item("Seq") += 2
            Next

            '下から２行削除
            workDt.Rows.RemoveAt(INPUT_ROW_COUNT + 2)
            workDt.Rows.RemoveAt(INPUT_ROW_COUNT + 1)
        End If
    End Sub

    Private Sub btnRowDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnRowDelete.Click
        Dim selectedRowIndex As Integer = If(IsNothing(dgvWork.CurrentRow), -1, dgvWork.CurrentRow.Index)
        If selectedRowIndex = -1 OrElse selectedRowIndex = 0 Then
            Return
        Else
            '変更の行を選択してる場合は予定の行を選択しているindexとする
            If selectedRowIndex Mod 2 = 0 Then
                selectedRowIndex -= 1
            End If

            '行削除
            workDt.Rows.RemoveAt(selectedRowIndex)
            workDt.Rows.RemoveAt(selectedRowIndex)

            '削除された行以降のSeqの値を更新
            For i As Integer = selectedRowIndex To INPUT_ROW_COUNT - 3 Step 2
                workDt.Rows(i).Item("Seq") -= 2
            Next

            '下に２行追加
            Dim row As DataRow = workDt.NewRow()
            row("Seq") = INPUT_ROW_COUNT
            workDt.Rows.InsertAt(workDt.NewRow(), INPUT_ROW_COUNT - 1)
            workDt.Rows.InsertAt(row, INPUT_ROW_COUNT - 1)

            '追加した変更行の設定
            dgvWork("Unt", INPUT_ROW_COUNT).ReadOnly = True
            dgvWork("Rdr", INPUT_ROW_COUNT).ReadOnly = True
            dgvWork("Nam", INPUT_ROW_COUNT).ReadOnly = True
            For i As Integer = 1 To 31 'Y1～Y31列のセルスタイル設定
                dgvWork("Y" & i, INPUT_ROW_COUNT).Style = workChangeCellStyle
            Next
            For i As Integer = 38 To 54 '小計部分のセルスタイル設定
                dgvWork(i, INPUT_ROW_COUNT).Style = subtotalChangeCellStyle
            Next
        End If
    End Sub

    Private Sub btnRegist_Click(sender As System.Object, e As System.EventArgs) Handles btnRegist.Click
        '
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        rs.Open("KinD", cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

        Dim ymStr As String = ymBox.getADStr4Ym()
        Dim floar As String = If(rbtn2F.Checked, "2", "3")
        Dim seq As Integer = 2
        Dim existsUnt As Boolean
        Dim existsNam As Boolean
        Dim existsWork As Boolean

        '登録チェック
        For i As Integer = 1 To 49 Step 2
            existsUnt = If(Util.checkDBNullValue(dgvWork("Unt", i).Value) <> "", True, False)
            existsNam = If(Util.checkDBNullValue(dgvWork("Nam", i).Value) <> "", True, False)
            existsWork = existsWorkStr(dgvWork.Rows(i))

            If (existsUnt AndAlso Not existsNam AndAlso existsWork) OrElse (Not existsUnt AndAlso Not existsNam AndAlso existsWork) Then
                MsgBox("氏名の無い行に入力しています。", MsgBoxStyle.Exclamation, "Work")
                rs.Close()
                cnn.Close()
                Return
            ElseIf (Not existsUnt AndAlso existsNam AndAlso existsWork) OrElse (Not existsUnt AndAlso existsNam AndAlso Not existsWork) Then
                MsgBox("ﾕﾆｯﾄが空白です。", MsgBoxStyle.Exclamation, "Work")
                rs.Close()
                cnn.Close()
                Return
            Else
                Continue For
            End If
        Next

        '既存データ削除
        monthDataDelete(ymStr, floar, cnn)

        '登録
        For i As Integer = 1 To 49 Step 2
            existsUnt = If(Util.checkDBNullValue(dgvWork("Unt", i).Value) <> "", True, False)
            existsNam = If(Util.checkDBNullValue(dgvWork("Nam", i).Value) <> "", True, False)
            existsWork = existsWorkStr(dgvWork.Rows(i))

            If (existsUnt AndAlso Not existsNam AndAlso Not existsWork) OrElse (Not existsUnt AndAlso Not existsNam AndAlso Not existsWork) Then
                Continue For
            Else
                If Not unitDictionary.ContainsKey(Util.checkDBNullValue(dgvWork("Unt", i).Value)) Then
                    Continue For
                End If
                With rs
                    .AddNew()
                    .Fields("Ym").Value = ymStr
                    .Fields("Seq").Value = seq
                    .Fields("Seq2").Value = Util.checkDBNullValue(dgvWork("Seq2", i).Value)
                    .Fields("Unt").Value = Util.checkDBNullValue(dgvWork("Unt", i).Value)
                    .Fields("Rdr").Value = Util.checkDBNullValue(dgvWork("Rdr", i).Value)
                    .Fields("Nam").Value = Util.checkDBNullValue(dgvWork("Nam", i).Value)
                    For j As Integer = 1 To 31
                        .Fields("Y" & j).Value = Util.checkDBNullValue(dgvWork("Y" & j, i).Value)
                        .Fields("J" & j).Value = If(Util.checkDBNullValue(dgvWork("Y" & j, i + 1).Value) = "", Util.checkDBNullValue(dgvWork("Y" & j, i).Value), Util.checkDBNullValue(dgvWork("Y" & j, i + 1).Value))
                    Next
                End With
                rs.Update()
                seq += 2
            End If
        Next
        rs.Close()
        cnn.Close()
        MsgBox("登録しました。", , "Work")
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Dim ymStr As String = ymBox.getADStr4Ym()
        Dim floar As String = If(rbtn2F.Checked, "2", "3")
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

        If rs.RecordCount <= 0 Then
            MsgBox("登録されていません", , "Work")
            rs.Close()
            cnn.Close()
        Else
            Dim result As DialogResult = MessageBox.Show("削除してよろしいですか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If result = Windows.Forms.DialogResult.Yes Then
                monthDataDelete(ymStr, floar, cnn)
                rs.Close()
                cnn.Close()

                '再表示
                displayWork(ymStr, floar, True)
                MsgBox("削除しました", , "Work")
            End If
        End If
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        'パスワードフォーム表示
        Dim passForm As Form = New passwordForm(TopForm.iniFilePath, 2)
        If passForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        Dim ymStr As String = ymBox.getADStr4Ym() '選択年月
        Dim floar As String = If(rbtn2F.Checked, "2", "3")
        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim rs As New ADODB.Recordset
        Dim sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('20' <= Seq2 AND Seq2 <= '39')) order by Seq2" '選択年月の全てのデータ(2階、3階共に)抽出
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
        If rs.RecordCount <= 0 Then
            MsgBox("該当がありません。", MsgBoxStyle.Exclamation, "Work")
            rs.Close()
            cnn.Close()
            Return
        Else
            rs.Close()
            subtotalClear()

            '予定の小計表示
            sql = "SELECT * FROM KinD WHERE YM='" & ymStr & "' AND ((Seq2='00' AND Unt='※') OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) order by Seq"
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)
            Dim rowIndex As Integer = 1
            Dim totalTime As Double
            While Not rs.EOF
                totalTime = 0.0
                For i As Integer = 1 To 31
                    Dim inputPlan As String = Util.checkDBNullValue(rs.Fields("Y" & i).Value)
                    If workTimeDictionary.ContainsKey(inputPlan) Then
                        totalTime = totalTime + workTimeDictionary(inputPlan)
                    End If
                    If Not abbreviationDictionary.ContainsKey(inputPlan) AndAlso inputPlan <> "" Then
                        inputPlan = "明"
                    End If
                    If abbreviationDictionary.ContainsKey(inputPlan) Then
                        Dim columnStr As String = abbreviationDictionary(inputPlan)
                        dgvWork(columnStr, rowIndex).Value = If(IsNumeric(dgvWork(columnStr, rowIndex).Value), CInt(dgvWork(columnStr, rowIndex).Value), 0) + 1
                        If columnStr <> "明等" Then
                            dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr)).Value = If(IsNumeric(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr)).Value), CInt(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr)).Value), 0) + 1
                        End If
                    End If
                Next
                If totalTime <> 0.0 Then
                    dgvWork("月合計", rowIndex).Value = totalTime.ToString("f1")
                End If
                rowIndex += 2
                rs.MoveNext()
            End While

            Dim changeRowResult As DialogResult = MessageBox.Show("縦/横計の変更分も表示しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            If changeRowResult = Windows.Forms.DialogResult.Yes Then
                '変更分の表示
                rs.MoveFirst()
                rowIndex = 1
                While Not rs.EOF
                    totalTime = 0.0
                    For i As Integer = 1 To 31
                        Dim inputChange As String = Util.checkDBNullValue(rs.Fields("J" & i).Value)
                        If workTimeDictionary.ContainsKey(inputChange) Then
                            totalTime = totalTime + workTimeDictionary(inputChange)
                        End If
                        If Not abbreviationDictionary.ContainsKey(inputChange) AndAlso inputChange <> "" Then
                            inputChange = "明"
                        End If
                        If abbreviationDictionary.ContainsKey(inputChange) Then
                            Dim columnStr As String = abbreviationDictionary(inputChange)
                            dgvWork(columnStr, rowIndex + 1).Value = If(IsNumeric(dgvWork(columnStr, rowIndex + 1).Value), CInt(dgvWork(columnStr, rowIndex + 1).Value), 0) + 1
                            If columnStr <> "明等" Then
                                dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr) + 1).Value = If(IsNumeric(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr) + 1).Value), CInt(dgvWork("Y" & i, subtotalStrIndexDictionary(columnStr) + 1).Value), 0) + 1
                            End If
                        End If
                    Next
                    If totalTime <> 0.0 Then
                        dgvWork("月合計", rowIndex + 1).Value = totalTime.ToString("f1")
                    End If
                    rowIndex += 2
                    rs.MoveNext()
                End While
            End If

            Dim workPrintResult As DialogResult = MessageBox.Show("勤務割表を印刷しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            Dim personalPrintResult As DialogResult = MessageBox.Show("個人別勤務割を印刷しますか？", "Work", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)

            '印刷
            If workPrintResult = Windows.Forms.DialogResult.Yes Then
                '勤務割表の印刷

            End If

            If personalPrintResult = Windows.Forms.DialogResult.Yes Then
                '個人別勤務割の印刷

            End If

        End If

    End Sub
End Class