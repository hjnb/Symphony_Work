Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class 週間表
    Private DGV1Table As DataTable
    Private DGV2Table As DataTable
    Private cellStyle1 As DataGridViewCellStyle
    Private cellStyle2 As DataGridViewCellStyle
    Private clearCellStyle As DataGridViewCellStyle
    Private whiteCellStyle As DataGridViewCellStyle
    Private pinkCellStyle As DataGridViewCellStyle
    Private Const HEISEI_Str As String = "H"
    Private Const NEXT_WAREKI As String = "X"
    Public cra3, cra4, cra5, cra6, cra8, cra9, cra10, cra11, cra12, cra13, cra14, cra15, cra16, cra17, cra18, cra19, cra20, cra21, cra22, cra23, cra24, cra25, cra26, cra27, cra28, cra29, cra30, cra31, cra32, cra33, cra34, cra35, cra36, cra37, cra38, cra39, cra40, cra41, cra42, crb3, crb4, crb5, crb6, crb8, crb9, crb10, crb11, crb12, crb13, crb14, crb15, crb16, crb17, crb18, crb19, crb20, crb21, crb22, crb23, crb24, crb25, crb26, crb27, crb28, crb29, crb30, crb31, crb32, crb33, crb34, crb35, crb36, crb37, crb38, crb39, crb40, crb41, crb42 As Integer

    Private Sub MadeStyle()
        '文字の大きさ指定
        cellStyle1 = New DataGridViewCellStyle()
        cellStyle1.Font = New Font("MS UI Gothic", 6)
        cellStyle1.ForeColor = Color.Blue
        cellStyle1.BackColor = Color.FromArgb(234, 234, 234)
        cellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter

        cellStyle2 = New DataGridViewCellStyle()
        cellStyle2.Font = New Font("MS UI Gothic", 8)
        cellStyle2.ForeColor = Color.Blue
        cellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter

        clearCellStyle = New DataGridViewCellStyle()
        clearCellStyle.BackColor = Color.FromArgb(234, 234, 234)
        clearCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        whiteCellStyle = New DataGridViewCellStyle()
        whiteCellStyle.Font = New Font("MS UI Gothic", 8)
        whiteCellStyle.BackColor = Color.FromArgb(255, 255, 255)
        whiteCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        pinkCellStyle = New DataGridViewCellStyle()
        pinkCellStyle.Font = New Font("MS UI Gothic", 8)
        pinkCellStyle.BackColor = Color.FromArgb(255, 192, 255)
        pinkCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub 週間表_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt = True Then
            If e.KeyCode = Keys.F12 Then
                btnTouroku.Visible = True
                btnSakujo.Visible = True
                btnInnsatu.Visible = True
                btnTorikomi.Visible = True
                Dim Staff As New 職員リスト()
                Staff.Owner = Me
                Staff.Show()
            End If
        End If
    End Sub

    Private Sub 週間表_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        MadeStyle()
        Me.WindowState = FormWindowState.Maximized
        'DataGridView1の設定
        Dim Cn As New OleDbConnection(TopForm.DB_Work)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        DGV1Table = New DataTable()
        Util.EnableDoubleBuffering(DataGridView1)

        With DataGridView1
            .RowTemplate.Height = 20
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 7)
        End With

        'DataGridView1列作成
        For i As Integer = 0 To 28
            DGV1Table.Columns.Add("a" & i, Type.GetType("System.String"))
        Next

        'DataGridView1行作成
        For i As Integer = 1 To 39
            DGV1Table.Rows.Add(DGV1Table.NewRow())
        Next

        'DataGridView1空を表示
        DataGridView1.DataSource = DGV1Table

        'DataGridView1列の設定
        For c As Integer = 0 To 28
            If c = 0 Then
                DataGridView1.Columns(c).Width = 30
            ElseIf c Mod 2 = 0 Then
                DataGridView1.Columns(c).Width = 55
                DataGridView1.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ElseIf c = 1 OrElse c = 5 OrElse c = 9 OrElse c = 13 OrElse c = 17 OrElse c = 21 OrElse c = 25 Then
                DataGridView1.Columns(c).Width = 28
            ElseIf c = 3 OrElse c = 7 OrElse c = 11 OrElse c = 15 OrElse c = 19 OrElse c = 23 OrElse c = 27 Then
                DataGridView1.Columns(c).Width = 17
            End If
        Next

        'DataGridView1の行の設定
        For r As Integer = 0 To 38
            DataGridView1.Rows(r).Height = 15
        Next

        'DataGridView2の設定
        Dim SQLCm2 As OleDbCommand = Cn.CreateCommand
        Dim Adapter2 As New OleDbDataAdapter(SQLCm2)
        DGV2Table = New DataTable()
        Util.EnableDoubleBuffering(DataGridView2)

        With DataGridView2
            '.RowTemplate.Height = 20
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 8)
        End With

        'DataGridView2の列作成
        For i As Integer = 0 To 6
            DGV2Table.Columns.Add("a" & i, Type.GetType("System.String"))
        Next

        'DataGridView2の行作成
        For i As Integer = 1 To 7
            DGV2Table.Rows.Add(DGV2Table.NewRow())
        Next

        'DataGridView2の空を表示
        DataGridView2.DataSource = DGV2Table

        'DataGridView2の列の設定
        For c As Integer = 0 To 6
            DataGridView2.Columns(c).Width = 155
        Next

        'DataGridView2の行の設定
        For r As Integer = 0 To 6
            DataGridView2.Rows(r).Height = 15
        Next

        KeyPreview = True

        '日付の設定
        Dim ymd As Date = Today
        Dim weekNumber As DayOfWeek = ymd.DayOfWeek
        For i As Integer = 0 To 6
            If weekNumber = i Then
                ymd = ymd.AddDays(-i)
                lblYmd.Text = ChangeWareki(ymd)
            End If
        Next
        lblYmd.Text = lblYmd.Text & "（日）"

        '各セルのスタイルの設定
        With DataGridView1
            For i = 0 To 1
                With .Rows(i)
                    .ReadOnly = True
                    .DefaultCellStyle = clearCellStyle
                End With
            Next
            With .Rows(6)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            With .Columns(0)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            For i As Integer = 1 To 28 Step 2
                .Columns(i).ReadOnly = True
                .Columns(i).DefaultCellStyle = cellStyle1
            Next
        End With

        For row As Integer = 11 To 23
            If row = 11 OrElse row = 16 OrElse row = 21 OrElse row = 22 OrElse row = 23 Then
                For col As Integer = 1 To 28
                    DataGridView1(col, row).Style = whiteCellStyle
                    DataGridView1(col, row).ReadOnly = False
                Next
            End If
        Next

        '各セルの固定値部分の設定
        Dim Youbi() As String = {"日", "月", "火", "水", "木", "金", "土"}
        For i As Integer = 0 To 6
            DataGridView1(i * 4 + 3, 0).Value = Youbi(i)
            DataGridView1(i * 4 + 2, 1).Value = "入浴"
            DataGridView1(i * 4 + 4, 1).Value = "着脱"
            DataGridView1(i * 4 + 2, 6).Value = "ＡＭ"
            DataGridView1(i * 4 + 4, 6).Value = "ＰＭ"
            DataGridView1(i * 4 + 3, 0).Style = cellStyle2
            DataGridView1(i * 4 + 2, 1).Style = cellStyle2
            DataGridView1(i * 4 + 4, 1).Style = cellStyle2
            DataGridView1(i * 4 + 2, 6).Style = cellStyle2
            DataGridView1(i * 4 + 4, 6).Style = cellStyle2
            For r As Integer = 1 To 3
                DataGridView1(i * 4 + 1, r * 5 + 2).Value = "早"
                DataGridView1(i * 4 + 1, r * 5 + 3).Value = "日早"
                DataGridView1(i * 4 + 1, r * 5 + 4).Value = "遅"
                DataGridView1(i * 4 + 1, r * 5 + 5).Value = "遅々"
                DataGridView1(i * 4 + 1, r * 5 + 19).Value = "朝"
                DataGridView1(i * 4 + 1, r * 5 + 21).Value = "昼"
                DataGridView1(i * 4 + 1, r * 5 + 23).Value = "夕"
                DataGridView1(i * 4 + 1, r * 5 + 2).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 3).Style = cellStyle1
                DataGridView1(i * 4 + 1, r * 5 + 4).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 5).Style = cellStyle1
                DataGridView1(i * 4 + 1, r * 5 + 19).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 21).Style = cellStyle2
                DataGridView1(i * 4 + 1, r * 5 + 23).Style = cellStyle2
            Next
        Next

        Dim Moji As String() = {"午前", "午後", "森", "ﾊﾟｰﾄ", "星", "ﾊﾟｰﾄ", "空", "ﾊﾟｰﾄ", "夜勤", "深夜", "森", "星", "空"}
        Dim Gyo As Integer() = {2, 4, 8, 11, 13, 16, 18, 21, 22, 23, 26, 31, 36}

        For n As Integer = 0 To 12
            DataGridView1(0, Gyo(n)).Style = cellStyle2
            DataGridView1(0, Gyo(n)).Value = Moji(n)
        Next

    End Sub

    Private Sub DataGridView_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting, DataGridView2.CellPainting
        '選択したセルに枠を付ける
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

            If (e.PaintParts And DataGridViewPaintParts.SelectionBackground) = DataGridViewPaintParts.SelectionBackground AndAlso (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                e.Graphics.DrawRectangle(New Pen(Color.Black, 2I), e.CellBounds.X + 1I, e.CellBounds.Y + 1I, e.CellBounds.Width - 3I, e.CellBounds.Height - 3I)
            End If

            Dim pParts As DataGridViewPaintParts = e.PaintParts And Not DataGridViewPaintParts.Background
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub

    Public Function ChangeWareki(ymd As Date) As String
        Dim wareki As String = ""
        Dim Result As String = ""
        If ymd <= "2019/04/30" Then
            wareki = "H"
            Dim YY As String = (Val(Strings.Left(ymd, 4)) - 1988)
            If YY.Length = 1 Then
                YY = "0" & (Val(Strings.Left(ymd, 4)) - 1988)
            End If
            Result = wareki & YY & "/" & Strings.Right(ymd, 5)
        ElseIf ymd > "2019/04/30" Then
            wareki = NEXT_WAREKI
            Dim YY As String = (Val(Strings.Left(ymd, 4)) - 2018)
            If YY.Length = 1 Then
                YY = "0" & (Val(Strings.Left(ymd, 4)) - 2018)
            End If
            Result = wareki & YY & "/" & Strings.Right(ymd, 5)
        End If
        Return Result
    End Function

    Public Function ChangeSeireki(ymd As String) As String
        Dim Seireki As Integer
        If Strings.Left(ymd, 1) = "H" Then
            Seireki = Val(Strings.Mid(ymd, 2, 2) + 1988)
        ElseIf Strings.Left(ymd, 1) = NEXT_WAREKI Then
            Seireki = Val(Strings.Mid(ymd, 2, 2) + 2018)
        End If

        Return Seireki
    End Function

    Private Sub rbn2F_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbn2F.CheckedChanged
        rbn2F.BackColor = Color.Yellow
        rbn3F.BackColor = SystemColors.Control

    End Sub

    Private Sub rbn3F_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbn3F.CheckedChanged
        rbn2F.BackColor = SystemColors.Control
        rbn3F.BackColor = Color.Yellow

        ChangeForm()
    End Sub

    Private Sub ChangeForm()
        '2階と3階で共通の部分
        DataGridView1.Columns.Clear()

        Dim Cn As New OleDbConnection(TopForm.DB_Work)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        DGV1Table = New DataTable()
        Util.EnableDoubleBuffering(DataGridView1)

        With DataGridView1
            .RowTemplate.Height = 20
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle.Font = New Font("MS UI Gothic", 7)
        End With

        'DataGridView1列作成
        For i As Integer = 0 To 28
            DGV1Table.Columns.Add("a" & i, Type.GetType("System.String"))
        Next


        If rbn2F.Checked = True Then    '2階の情報を表示
            Label52.Location = New Point(12, 74)
            Label1.Location = New Point(12, 103)
            Label2.Location = New Point(12, 134)
            Label3.Location = New Point(12, 149)
            Label4.Location = New Point(12, 223)
            Label5.Location = New Point(12, 298)
            Label6.Location = New Point(12, 373)
            Label7.Location = New Point(12, 403)
            Label8.Location = New Point(12, 478)
            Label9.Location = New Point(12, 554)

            For i As Integer = 10 To 16
                Controls("Label" & i).Size = New Size(2, 588)
            Next

            DataGridView1.Location = New Point(12, 43)
            DataGridView1.Size = New Size(1118, 588)
            DataGridView2.Location = New Point(42, 630)

            'DataGridView1行作成
            For i As Integer = 1 To 39
                DGV1Table.Rows.Add(DGV1Table.NewRow())
            Next
        ElseIf rbn3F.Checked = True Then    '3階の情報を表示
            Label52.Location = New Point(12, 72)
            Label1.Location = New Point(12, 100)
            Label2.Location = New Point(12, 128)
            Label3.Location = New Point(12, 142)
            Label4.Location = New Point(12, 212)
            Label5.Location = New Point(12, 282)
            Label6.Location = New Point(12, 365)
            Label7.Location = New Point(12, 393)
            Label8.Location = New Point(12, 491)
            Label9.Location = New Point(12, 589)
            For i As Integer = 10 To 16
                Controls("Label" & i).Size = New Size(2, 605)
            Next

            DataGridView1.Location = New Point(12, 43)
            DataGridView1.Size = New Size(1118, 605)
            DataGridView2.Location = New Point(42, 647)

            'DataGridView1行作成
            For i As Integer = 1 To 43
                DGV1Table.Rows.Add(DGV1Table.NewRow())
            Next
        End If

        '2階と3階で共通の部分

        'DataGridView1空を表示
        DataGridView1.DataSource = DGV1Table

        'DataGridView1列の設定
        For c As Integer = 0 To 28
            If c = 0 Then
                DataGridView1.Columns(c).Width = 30
            ElseIf c Mod 2 = 0 Then
                DataGridView1.Columns(c).Width = 55
                DataGridView1.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ElseIf c = 1 OrElse c = 5 OrElse c = 9 OrElse c = 13 OrElse c = 17 OrElse c = 21 OrElse c = 25 Then
                DataGridView1.Columns(c).Width = 28
            ElseIf c = 3 OrElse c = 7 OrElse c = 11 OrElse c = 15 OrElse c = 19 OrElse c = 23 OrElse c = 27 Then
                DataGridView1.Columns(c).Width = 17
            End If
        Next

        If rbn2F.Checked = True Then    '2階の情報を表示
            'DataGridView1行の設定
            For r As Integer = 0 To 38
                DataGridView1.Rows(r).Height = 15
            Next
        ElseIf rbn3F.Checked = True Then    '3階の情報を表示
            'DataGridView1行の設定
            For r As Integer = 0 To 42
                DataGridView1.Rows(r).Height = 14
            Next
        End If

        '2階と3階で共通の部分

        '各セルのスタイルの設定
        With DataGridView1
            For i = 0 To 1
                With .Rows(i)
                    .ReadOnly = True
                    .DefaultCellStyle = clearCellStyle
                End With
            Next
            With .Rows(6)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            With .Columns(0)
                .ReadOnly = True
                .DefaultCellStyle = clearCellStyle
            End With
            For i As Integer = 1 To 28 Step 2
                .Columns(i).ReadOnly = True
                .Columns(i).DefaultCellStyle = cellStyle1
            Next
        End With

        For row As Integer = 11 To 23
            If row = 11 OrElse row = 16 OrElse row = 21 OrElse row = 22 OrElse row = 23 Then
                For col As Integer = 1 To 28
                    DataGridView1(col, row).Style = whiteCellStyle
                    DataGridView1(col, row).ReadOnly = False
                Next
            End If
        Next

        If rbn2F.Checked = True Then    '2階の情報を表示
            '各セルの固定値部分の設定
            Dim Youbi() As String = {"日", "月", "火", "水", "木", "金", "土"}
            For i As Integer = 0 To 6
                DataGridView1(i * 4 + 3, 0).Value = Youbi(i)
                DataGridView1(i * 4 + 2, 1).Value = "入浴"
                DataGridView1(i * 4 + 4, 1).Value = "着脱"
                DataGridView1(i * 4 + 2, 6).Value = "ＡＭ"
                DataGridView1(i * 4 + 4, 6).Value = "ＰＭ"
                DataGridView1(i * 4 + 3, 0).Style = cellStyle2
                DataGridView1(i * 4 + 2, 1).Style = cellStyle2
                DataGridView1(i * 4 + 4, 1).Style = cellStyle2
                DataGridView1(i * 4 + 2, 6).Style = cellStyle2
                DataGridView1(i * 4 + 4, 6).Style = cellStyle2
                For r As Integer = 1 To 3
                    DataGridView1(i * 4 + 1, r * 5 + 2).Value = "早"
                    DataGridView1(i * 4 + 1, r * 5 + 3).Value = "日早"
                    DataGridView1(i * 4 + 1, r * 5 + 4).Value = "遅"
                    DataGridView1(i * 4 + 1, r * 5 + 5).Value = "遅々"
                    DataGridView1(i * 4 + 1, r * 5 + 19).Value = "朝"
                    DataGridView1(i * 4 + 1, r * 5 + 21).Value = "昼"
                    DataGridView1(i * 4 + 1, r * 5 + 23).Value = "夕"
                    DataGridView1(i * 4 + 1, r * 5 + 2).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 3).Style = cellStyle1
                    DataGridView1(i * 4 + 1, r * 5 + 4).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 5).Style = cellStyle1
                    DataGridView1(i * 4 + 1, r * 5 + 19).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 21).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 23).Style = cellStyle2
                Next
            Next
            'DataGridView1(0, 2).Value = "午前"
            'DataGridView1(0, 4).Value = "午後"
            'DataGridView1(0, 8).Value = "森"
            'DataGridView1(0, 11).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 13).Value = "星"
            'DataGridView1(0, 16).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 18).Value = "空"
            'DataGridView1(0, 21).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 22).Value = "夜勤"
            'DataGridView1(0, 23).Value = "深夜"
            'DataGridView1(0, 26).Value = "森"
            'DataGridView1(0, 31).Value = "星"
            'DataGridView1(0, 36).Value = "空"
            Dim Moji As String() = {"午前", "午後", "森", "ﾊﾟｰﾄ", "星", "ﾊﾟｰﾄ", "空", "ﾊﾟｰﾄ", "夜勤", "深夜", "森", "星", "空"}
            Dim Gyo As Integer() = {2, 4, 8, 11, 13, 16, 18, 21, 22, 23, 26, 31, 36}

            For n As Integer = 0 To 12
                DataGridView1(0, Gyo(n)).Style = cellStyle2
                DataGridView1(0, Gyo(n)).Value = Moji(n)
            Next

        ElseIf rbn3F.Checked = True Then    '3階の情報を表示
            
            For col As Integer = 1 To 28
                DataGridView1(col, 24).Style = whiteCellStyle
                DataGridView1(col, 24).ReadOnly = False
            Next
               

            '各セルの固定値部分の設定
            Dim Youbi() As String = {"日", "月", "火", "水", "木", "金", "土"}
            For i As Integer = 0 To 6
                DataGridView1(i * 4 + 3, 0).Value = Youbi(i)
                DataGridView1(i * 4 + 2, 1).Value = "入浴"
                DataGridView1(i * 4 + 4, 1).Value = "着脱"
                DataGridView1(i * 4 + 2, 6).Value = "ＡＭ"
                DataGridView1(i * 4 + 4, 6).Value = "ＰＭ"
                DataGridView1(i * 4 + 3, 0).Style = cellStyle2
                DataGridView1(i * 4 + 2, 1).Style = cellStyle2
                DataGridView1(i * 4 + 4, 1).Style = cellStyle2
                DataGridView1(i * 4 + 2, 6).Style = cellStyle2
                DataGridView1(i * 4 + 4, 6).Style = cellStyle2
                DataGridView1(i * 4 + 1, 24).Style = whiteCellStyle
                DataGridView1(i * 4 + 1, 24).ReadOnly = False
                DataGridView1(i * 4 + 3, 24).Style = whiteCellStyle
                DataGridView1(i * 4 + 3, 24).ReadOnly = False
                For r As Integer = 1 To 3
                    DataGridView1(i * 4 + 1, r * 5 + 2).Value = "早"
                    DataGridView1(i * 4 + 1, r * 5 + 3).Value = "日早"
                    DataGridView1(i * 4 + 1, r * 5 + 4).Value = "遅"
                    DataGridView1(i * 4 + 1, r * 5 + 5).Value = "遅々"
                    DataGridView1(i * 4 + 1, r * 7 + 18).Value = "朝"
                    DataGridView1(i * 4 + 1, r * 5 + 2).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 3).Style = cellStyle1
                    DataGridView1(i * 4 + 1, r * 5 + 4).Style = cellStyle2
                    DataGridView1(i * 4 + 1, r * 5 + 5).Style = cellStyle1
                    DataGridView1(i * 4 + 1, r * 7 + 18).Style = cellStyle2
                    If r = 1 OrElse r = 2 Then
                        DataGridView1(i * 4 + 1, r * 7 + 21).Value = "昼"
                        DataGridView1(i * 4 + 1, r * 7 + 23).Value = "夕"
                        DataGridView1(i * 4 + 1, r * 7 + 21).Style = cellStyle2
                        DataGridView1(i * 4 + 1, r * 7 + 23).Style = cellStyle2
                    ElseIf r = 3 Then
                        DataGridView1(i * 4 + 1, r * 7 + 19).Value = "昼"
                        DataGridView1(i * 4 + 1, r * 7 + 21).Value = "夕"
                        DataGridView1(i * 4 + 1, r * 7 + 19).Style = cellStyle2
                        DataGridView1(i * 4 + 1, r * 7 + 21).Style = cellStyle2
                    End If
                Next
            Next

            'DataGridView1(0, 2).Value = "午前"
            'DataGridView1(0, 4).Value = "午後"
            'DataGridView1(0, 9).Value = "月"
            'DataGridView1(0, 11).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 14).Value = "花"
            'DataGridView1(0, 16).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 19).Value = "海"
            'DataGridView1(0, 21).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 22).Value = "ﾊﾟｰﾄ"
            'DataGridView1(0, 23).Value = "夜勤"
            'DataGridView1(0, 24).Value = "深夜"
            'DataGridView1(0, 28).Value = "月"
            'DataGridView1(0, 35).Value = "花"
            'DataGridView1(0, 40).Value = "海"
            Dim Moji As String() = {"午前", "午後", "月", "ﾊﾟｰﾄ", "花", "ﾊﾟｰﾄ", "海", "ﾊﾟｰﾄ", "ﾊﾟｰﾄ", "夜勤", "深夜", "月", "花", "海"}
            Dim Gyo As Integer() = {2, 4, 9, 11, 14, 16, 19, 21, 22, 23, 24, 28, 35, 40}

            For n As Integer = 0 To 13
                DataGridView1(0, Gyo(n)).Style = cellStyle2
                DataGridView1(0, Gyo(n)).Value = Moji(n)
            Next

            
        End If

        '2階と3階で共通の部分
        For i As Integer = 0 To 6
            DataGridView1(4 * i + 2, 0).Value = Val(Strings.Mid(lblYmd.Text, 8, 2)) + i
        Next

        Dim Getumatu As Integer = Date.DaysInMonth(ChangeSeireki(Strings.Left(lblYmd.Text, 9)), Val(Strings.Mid(lblYmd.Text, 5, 2)))

        For i As Integer = 0 To 6
            If Val(DataGridView1(4 * i + 2, 0).Value) > Getumatu Then
                DataGridView1(4 * i + 2, 0).Value = Val(DataGridView1(4 * i + 2, 0).Value) - Getumatu
            End If
        Next

        DataIndication()

    End Sub

    Private Sub btnUp_Click(sender As System.Object, e As System.EventArgs) Handles btnUp.Click
        Dim ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        ymd = ymd.AddDays(7)
        lblYmd.Text = ChangeWareki(ymd) & "（日）"
    End Sub

    Private Sub btnDown_Click(sender As System.Object, e As System.EventArgs) Handles btnDown.Click
        Dim ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        ymd = ymd.AddDays(-7)
        lblYmd.Text = ChangeWareki(ymd) & "（日）"
    End Sub

    Private Sub lblYmd_TextChanged(sender As Object, e As System.EventArgs) Handles lblYmd.TextChanged
        If Strings.Left(lblYmd.Text, 9) = "H30.12.31" Then
            Return
        End If
        For i As Integer = 0 To 6
            DataGridView1(4 * i + 2, 0).Value = Val(Strings.Mid(lblYmd.Text, 8, 2)) + i
        Next

        Dim Getumatu As Integer = Date.DaysInMonth(ChangeSeireki(Strings.Left(lblYmd.Text, 9)), Val(Strings.Mid(lblYmd.Text, 5, 2)))

        For i As Integer = 0 To 6
            If Val(DataGridView1(4 * i + 2, 0).Value) > Getumatu Then
                DataGridView1(4 * i + 2, 0).Value = Val(DataGridView1(4 * i + 2, 0).Value) - Getumatu
            End If
        Next

        DataIndication()
    End Sub

    Private Sub DataClear()
        If rbn2F.Checked = True Then
            For column As Integer = 1 To 28
                If column Mod 2 = 0 Then    '偶数列
                    For row As Integer = 2 To 38    '入力可能な行をチェック
                        If row <> 6 Then    '6行目だけ入力する場所がないので外す
                            DataGridView1(column, row).Value = ""
                            DataGridView1(column, row).Style = whiteCellStyle
                        End If
                    Next
                Else
                    For row As Integer = 11 To 23
                        If row = 11 OrElse row = 16 OrElse row = 21 OrElse row = 22 OrElse row = 23 Then
                            DataGridView1(column, row).Value = ""
                            'DataGridView1(column, row).Style = clearCellStyle
                        End If
                    Next
                End If
            Next
        ElseIf rbn3F.Checked = True Then
            For column As Integer = 1 To 28
                If column Mod 2 = 0 Then    '偶数列
                    For row As Integer = 2 To 42    '入力可能な行をチェック
                        If row <> 6 Then    '6行目だけ入力する場所がないので外す
                            DataGridView1(column, row).Value = ""
                            DataGridView1(column, row).Style = whiteCellStyle
                        End If
                    Next
                Else
                    For row As Integer = 11 To 24
                        If row = 11 OrElse row = 16 OrElse row = 21 OrElse row = 22 OrElse row = 23 OrElse row = 24 Then
                            DataGridView1(column, row).Value = ""
                            'DataGridView1(column, row).Style = clearCellStyle
                        End If
                    Next
                End If
            Next
        End If
        

        For c2 As Integer = 0 To 6
            For r2 As Integer = 0 To 6
                DataGridView2(c2, r2).Value = ""
            Next
        Next
    End Sub

    Private Sub DataIndication()
        DataClear()

        If rbn2F.Checked = True Then    '2階
            Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
            Dim YmdAdd7 As Date = Ymd.AddDays(6)

            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from SHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            'Datagridview1への表示
            Dim ColumnsNo As Integer = 0
            While Not rs.EOF
                For RowNo As Integer = 2 To 38
                    If RowNo <= 5 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo - 1).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 3).Value
                        If rs.Fields(RowNo + 88).Value = 1 Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        Else
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 124).Value = 1 Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        Else
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    ElseIf 7 <= RowNo And RowNo <= 10 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 2).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 8).Value
                    ElseIf RowNo = 11 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 2).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 3).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 8).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 9).Value
                    ElseIf 12 <= RowNo And RowNo <= 15 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 9).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 15).Value
                    ElseIf RowNo = 16 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 9).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 10).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 15).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 16).Value
                    ElseIf 17 <= RowNo And RowNo <= 20 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 16).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 22).Value
                    ElseIf RowNo = 21 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 16).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 17).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 22).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 23).Value
                    ElseIf RowNo = 22 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 23).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 24).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 25).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 26).Value
                    ElseIf RowNo = 23 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 26).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 27).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 28).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 29).Value
                    ElseIf RowNo = 24 Or RowNo = 25 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 36).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 38).Value
                    ElseIf RowNo = 26 Or RowNo = 27 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 38).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 40).Value
                    ElseIf RowNo = 28 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 40).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 41).Value
                    ElseIf RowNo = 29 Or RowNo = 30 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 41).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 43).Value
                    ElseIf RowNo = 31 Or RowNo = 32 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 43).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 45).Value
                    ElseIf RowNo = 33 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 45).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 46).Value
                    ElseIf RowNo = 34 Or RowNo = 35 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 46).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 48).Value
                    ElseIf RowNo = 36 Or RowNo = 37 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 48).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 50).Value
                    ElseIf RowNo = 38 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 50).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 51).Value
                    End If
                    '色付け処理
                    If RowNo >= 7 Then
                        If rs.Fields(RowNo + 87).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 87).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 123).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 123).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    End If
                Next

                'Datagridview2への表示
                For rowno2 As Integer = 1 To 7
                    DGV2Table.Rows(rowno2 - 1).Item("a" & ColumnsNo) = rs.Fields(rowno2 + 52).Value
                Next

                rs.MoveNext()

                ColumnsNo = ColumnsNo + 1
            End While
            cnn.Close()

        ElseIf rbn3F.Checked = True Then    '3階
            Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
            Dim YmdAdd7 As Date = Ymd.AddDays(6)

            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from SHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            'Datagridview1への表示
            Dim ColumnsNo As Integer = 0
            While Not rs.EOF
                For RowNo As Integer = 2 To 42
                    If RowNo <= 5 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo - 1).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 3).Value
                        If rs.Fields(RowNo + 98).Value = 1 Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        Else
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 138).Value = 1 Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        Else
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    ElseIf 7 <= RowNo And RowNo <= 10 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 2).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 8).Value
                    ElseIf RowNo = 11 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 2).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 3).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 8).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 9).Value
                    ElseIf 12 <= RowNo And RowNo <= 15 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 9).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 15).Value
                    ElseIf RowNo = 16 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 9).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 10).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 15).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 16).Value
                    ElseIf 17 <= RowNo And RowNo <= 20 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 16).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 24).Value
                    ElseIf RowNo = 21 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 16).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 17).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 24).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 25).Value
                    ElseIf RowNo = 22 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 17).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 18).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 25).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 26).Value
                    ElseIf RowNo = 23 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 26).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 27).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 28).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 29).Value
                    ElseIf RowNo = 24 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 29).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 30).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 31).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 32).Value
                    ElseIf RowNo = 25 Or RowNo = 26 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 39).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 41).Value
                    ElseIf RowNo = 27 Or RowNo = 28 Or RowNo = 29 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 41).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 44).Value
                    ElseIf RowNo = 30 Or RowNo = 31 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 44).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 46).Value
                    ElseIf RowNo = 32 Or RowNo = 33 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 46).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 48).Value
                    ElseIf RowNo = 34 Or RowNo = 35 Or RowNo = 36 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 48).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 51).Value
                    ElseIf RowNo = 37 Or RowNo = 38 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 51).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 53).Value
                    ElseIf RowNo = 39 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 53).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 54).Value
                    ElseIf RowNo = 40 Or RowNo = 41 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 54).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 56).Value
                    ElseIf RowNo = 42 Then
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 56).Value
                        DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 57).Value
                    End If
                    '色付け処理
                    If RowNo >= 7 Then
                        If rs.Fields(RowNo + 97).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 97).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 2, RowNo).Style = whiteCellStyle
                        End If
                        If rs.Fields(RowNo + 137).Value = "1" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = pinkCellStyle
                        ElseIf rs.Fields(RowNo + 137).Value = "0" Then
                            DataGridView1(ColumnsNo * 4 + 4, RowNo).Style = whiteCellStyle
                        End If
                    End If
                Next

                'Datagridview2への表示
                For rowno2 As Integer = 1 To 7
                    DGV2Table.Rows(rowno2 - 1).Item("a" & ColumnsNo) = rs.Fields(rowno2 + 56).Value
                Next

                rs.MoveNext()

                ColumnsNo = ColumnsNo + 1
            End While
            cnn.Close()
        End If

    End Sub

    Private Sub btnTouroku_Click(sender As System.Object, e As System.EventArgs) Handles btnTouroku.Click
        If MsgBox("登録してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "登録確認") = MsgBoxResult.No Then
            Return
        End If

        Dim cnn As New ADODB.Connection
        cnn.Open(TopForm.DB_Work)
        Dim Honjitu As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        'データ削除
        Dim DelYmd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        Dim DelYmdAdd7 As Date = DelYmd.AddDays(6)
        Dim SQL As String = ""
        If rbn2F.Checked = True Then
            SQL = "DELETE FROM SHyo WHERE #" & DelYmd & "# <= Ymd and Ymd <= #" & DelYmdAdd7 & "#"
        ElseIf rbn3F.Checked = True Then
            SQL = "DELETE FROM SHyo3 WHERE #" & DelYmd & "# <= Ymd and Ymd <= #" & DelYmdAdd7 & "#"
        End If
        cnn.Execute(SQL)
        'データ登録
        If rbn2F.Checked = True Then    '2階の登録
            Dim ymd, nyuam1, nyuam2, nyupm1, nyupm2, tyaam1, tyaam2, tyapm1, tyapm2, moram1, moram2, moram3, moram4, morampk, moramp, morpm1, morpm2, morpm3, morpm4, morpmpk, morpmp, hosam1, hosam2, hosam3, hosam4, hosampk, hosamp, hospm1, hospm2, hospm3, hospm4, hospmpk, hospmp, soram1, soram2, soram3, soram4, sorampk, soramp, sorpm1, sorpm2, sorpm3, sorpm4, sorpmpk, sorpmp, yak1k, yak1, yak2k, yak2, sin1k, sin1, sin2k, sin2, gyo1, gyo2, gyo3, gyo4, gyo5, gyo6, gyo7, mora1, mora2, mora3, mora4, morh1, morh2, morh3, morh4, mory1, mory2, hosa1, hosa2, hosa3, hosa4, hosh1, hosh2, hosh3, hosh4, hosy1, hosy2, sora1, sora2, sora3, sora4, sorh1, sorh2, sorh3, sorh4, sory1, sory2 As String
            'Dim cra3, cra4, cra5, cra6, cra8, cra9, cra10, cra11, cra12, cra13, cra14, cra15, cra16, cra17, cra18, cra19, cra20, cra21, cra22, cra23, cra24, cra25, cra26, cra27, cra28, cra29, cra30, cra31, cra32, cra33, cra34, cra35, cra36, cra37, cra38, cra39, crb3, crb4, crb5, crb6, crb8, crb9, crb10, crb11, crb12, crb13, crb14, crb15, crb16, crb17, crb18, crb19, crb20, crb21, crb22, crb23, crb24, crb25, crb26, crb27, crb28, crb29, crb30, crb31, crb32, crb33, crb34, crb35, crb36, crb37, crb38, crb39 As Integer

            For dd As Integer = 0 To 6
                If dd = 0 Then
                    ymd = Honjitu
                Else
                    ymd = Honjitu.AddDays(dd)
                End If
                nyuam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 2).Value)
                nyuam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 3).Value)
                nyupm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 4).Value)
                nyupm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 5).Value)
                tyaam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 2).Value)
                tyaam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 3).Value)
                tyapm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 4).Value)
                tyapm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 5).Value)
                moram1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 7).Value)
                moram2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 8).Value)
                moram3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 9).Value)
                moram4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 10).Value)
                morampk = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 11).Value)
                moramp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 11).Value)
                morpm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 7).Value)
                morpm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 8).Value)
                morpm3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 9).Value)
                morpm4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 10).Value)
                morpmpk = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 11).Value)
                morpmp = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 11).Value)
                hosam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 12).Value)
                hosam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 13).Value)
                hosam3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 14).Value)
                hosam4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 15).Value)
                hosampk = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 16).Value)
                hosamp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 16).Value)
                hospm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 12).Value)
                hospm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 13).Value)
                hospm3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 14).Value)
                hospm4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 15).Value)
                hospmpk = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 16).Value)
                hospmp = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 16).Value)
                soram1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 17).Value)
                soram2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 18).Value)
                soram3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 19).Value)
                soram4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 20).Value)
                sorampk = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 21).Value)
                soramp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 21).Value)
                sorpm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 17).Value)
                sorpm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 18).Value)
                sorpm3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 19).Value)
                sorpm4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 20).Value)
                sorpmpk = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 21).Value)
                sorpmp = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 21).Value)
                yak1k = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 22).Value)
                yak1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 22).Value)
                yak2k = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 22).Value)
                yak2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 22).Value)
                sin1k = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 23).Value)
                sin1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 23).Value)
                sin2k = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 23).Value)
                sin2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 23).Value)
                gyo1 = Util.checkDBNullValue(DataGridView2(dd, 0).Value)
                gyo2 = Util.checkDBNullValue(DataGridView2(dd, 1).Value)
                gyo3 = Util.checkDBNullValue(DataGridView2(dd, 2).Value)
                gyo4 = Util.checkDBNullValue(DataGridView2(dd, 3).Value)
                gyo5 = Util.checkDBNullValue(DataGridView2(dd, 4).Value)
                gyo6 = Util.checkDBNullValue(DataGridView2(dd, 5).Value)
                gyo7 = Util.checkDBNullValue(DataGridView2(dd, 6).Value)
                mora1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 24).Value)
                mora2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 25).Value)
                mora3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 24).Value)
                mora4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 25).Value)
                morh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 26).Value)
                morh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 27).Value)
                morh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 26).Value)
                morh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 27).Value)
                mory1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 28).Value)
                mory2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 28).Value)
                hosa1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 29).Value)
                hosa2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 30).Value)
                hosa3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 29).Value)
                hosa4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 30).Value)
                hosh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 31).Value)
                hosh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 32).Value)
                hosh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 31).Value)
                hosh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 32).Value)
                hosy1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 33).Value)
                hosy2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 33).Value)
                sora1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 34).Value)
                sora2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 35).Value)
                sora3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 34).Value)
                sora4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 35).Value)
                sorh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 36).Value)
                sorh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 37).Value)
                sorh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 36).Value)
                sorh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 37).Value)
                sory1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 38).Value)
                sory2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 38).Value)

                For r As Integer = 2 To 38
                    If r <> 6 Then
                        Dim u As Type = GetType(週間表)
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("cra" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("cra" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("crb" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("crb" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    End If
                Next

                SQL = "INSERT INTO SHyo VALUES ('" & ymd & "', '" & nyuam1 & "', '" & nyuam2 & "', '" & nyupm1 & "', '" & nyupm2 & "', '" & tyaam1 & "', '" & tyaam2 & "', '" & tyapm1 & "', '" & tyapm2 & "', '" & moram1 & "', '" & moram2 & "', '" & moram3 & "', '" & moram4 & "', '" & morampk & "', '" & moramp & "', '" & morpm1 & "', '" & morpm2 & "', '" & morpm3 & "', '" & morpm4 & "', '" & morpmpk & "', '" & morpmp & "', '" & hosam1 & "', '" & hosam2 & "', '" & hosam3 & "', '" & hosam4 & "', '" & hosampk & "', '" & hosamp & "', '" & hospm1 & "', '" & hospm2 & "', '" & hospm3 & "', '" & hospm4 & "', '" & hospmpk & "', '" & hospmp & "', '" & soram1 & "', '" & soram2 & "', '" & soram3 & "', '" & soram4 & "', '" & sorampk & "', '" & soramp & "', '" & sorpm1 & "', '" & sorpm2 & "', '" & sorpm3 & "', '" & sorpm4 & "', '" & sorpmpk & "', '" & sorpmp & "', '" & yak1k & "', '" & yak1 & "', '" & yak2k & "', '" & yak2 & "', '" & sin1k & "', '" & sin1 & "', '" & sin2k & "', '" & sin2 & "', '" & gyo1 & "', '" & gyo2 & "', '" & gyo3 & "', '" & gyo4 & "', '" & gyo5 & "', '" & gyo6 & "', '" & gyo7 & "', '" & mora1 & "', '" & mora2 & "', '" & mora3 & "', '" & mora4 & "', '" & morh1 & "', '" & morh2 & "', '" & morh3 & "', '" & morh4 & "', '" & mory1 & "', '" & mory2 & "', '" & hosa1 & "', '" & hosa2 & "', '" & hosa3 & "', '" & hosa4 & "', '" & hosh1 & "', '" & hosh2 & "', '" & hosh3 & "', '" & hosh4 & "', '" & hosy1 & "', '" & hosy2 & "', '" & sora1 & "', '" & sora2 & "', '" & sora3 & "', '" & sora4 & "', '" & sorh1 & "', '" & sorh2 & "', '" & sorh3 & "', '" & sorh4 & "', '" & sory1 & "', '" & sory2 & "', '" & cra3 & "', '" & cra4 & "', '" & cra5 & "', '" & cra6 & "', '" & cra8 & "', '" & cra9 & "', '" & cra10 & "', '" & cra11 & "', '" & cra12 & "', '" & cra13 & "', '" & cra14 & "', '" & cra15 & "', '" & cra16 & "', '" & cra17 & "', '" & cra18 & "', '" & cra19 & "', '" & cra20 & "', '" & cra21 & "', '" & cra22 & "', '" & cra23 & "', '" & cra24 & "', '" & cra25 & "', '" & cra26 & "', '" & cra27 & "', '" & cra28 & "', '" & cra29 & "', '" & cra30 & "', '" & cra31 & "', '" & cra32 & "', '" & cra33 & "', '" & cra34 & "', '" & cra35 & "', '" & cra36 & "', '" & cra37 & "', '" & cra38 & "', '" & cra39 & "', '" & crb3 & "', '" & crb4 & "', '" & crb5 & "', '" & crb6 & "', '" & crb8 & "', '" & crb9 & "', '" & crb10 & "', '" & crb11 & "', '" & crb12 & "', '" & crb13 & "', '" & crb14 & "', '" & crb15 & "', '" & crb16 & "', '" & crb17 & "', '" & crb18 & "', '" & crb19 & "', '" & crb20 & "', '" & crb21 & "', '" & crb22 & "', '" & crb23 & "', '" & crb24 & "', '" & crb25 & "', '" & crb26 & "', '" & crb27 & "', '" & crb28 & "', '" & crb29 & "', '" & crb30 & "', '" & crb31 & "', '" & crb32 & "', '" & crb33 & "', '" & crb34 & "', '" & crb35 & "', '" & crb36 & "', '" & crb37 & "', '" & crb38 & "', '" & crb39 & "')"

                cnn.Execute(SQL)
            Next
        ElseIf rbn3F.Checked = True Then    '3階の登録
            Dim ymd, nyuam1, nyuam2, nyupm1, nyupm2, tyaam1, tyaam2, tyapm1, tyapm2, tukam1, tukam2, tukam3, tukam4, tukampk, tukamp, tukpm1, tukpm2, tukpm3, tukpm4, tukpmpk, tukpmp, hanam1, hanam2, hanam3, hanam4, hanampk, hanamp, hanpm1, hanpm2, hanpm3, hanpm4, hanpmpk, hanpmp, umiam1, umiam2, umiam3, umiam4, umiampk, umiamp, umiampk2, umiamp2, umipm1, umipm2, umipm3, umipm4, umipmpk, umipmp, umipmpk2, umipmp2, yak1k, yak1, yak2k, yak2, sin1k, sin1, sin2k, sin2, gyo1, gyo2, gyo3, gyo4, gyo5, gyo6, gyo7, tuka1, tuka2, tuka3, tuka4, tukh1, tukh2, tukh3, tukh4, tukh5, tukh6, tuky1, tuky2, tuky3, tuky4, hana1, hana2, hana3, hana4, hanh1, hanh2, hanh3, hanh4, hanh5, hanh6, hany1, hany2, hany3, hany4, umia1, umia2, umih1, umih2, umih3, umih4, umiy1, umiy2 As String
            Dim cra3, cra4, cra5, cra6, cra8, cra9, cra10, cra11, cra12, cra13, cra14, cra15, cra16, cra17, cra18, cra19, cra20, cra21, cra22, cra23, cra24, cra25, cra26, cra27, cra28, cra29, cra30, cra31, cra32, cra33, cra34, cra35, cra36, cra37, cra38, cra39, cra40, cra41, cra42, cra43, crb3, crb4, crb5, crb6, crb8, crb9, crb10, crb11, crb12, crb13, crb14, crb15, crb16, crb17, crb18, crb19, crb20, crb21, crb22, crb23, crb24, crb25, crb26, crb27, crb28, crb29, crb30, crb31, crb32, crb33, crb34, crb35, crb36, crb37, crb38, crb39, crb40, crb41, crb42, crb43 As Integer

            For dd As Integer = 0 To 6
                If dd = 0 Then
                    ymd = Honjitu
                Else
                    ymd = Honjitu.AddDays(dd)
                End If
                nyuam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 2).Value)
                nyuam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 3).Value)
                nyupm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 4).Value)
                nyupm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 5).Value)
                tyaam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 2).Value)
                tyaam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 3).Value)
                tyapm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 4).Value)
                tyapm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 5).Value)
                tukam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 7).Value)
                tukam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 8).Value)
                tukam3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 9).Value)
                tukam4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 10).Value)
                tukampk = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 11).Value)
                tukamp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 11).Value)
                tukpm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 7).Value)
                tukpm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 8).Value)
                tukpm3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 9).Value)
                tukpm4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 10).Value)
                tukpmpk = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 11).Value)
                tukpmp = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 11).Value)
                hanam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 12).Value)
                hanam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 13).Value)
                hanam3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 14).Value)
                hanam4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 15).Value)
                hanampk = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 16).Value)
                hanamp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 16).Value)
                hanpm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 12).Value)
                hanpm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 13).Value)
                hanpm3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 14).Value)
                hanpm4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 15).Value)
                hanpmpk = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 16).Value)
                hanpmp = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 16).Value)
                umiam1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 17).Value)
                umiam2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 18).Value)
                umiam3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 19).Value)
                umiam4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 20).Value)
                umiampk = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 21).Value)
                umiamp = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 21).Value)
                umiampk2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 22).Value)
                umiamp2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 22).Value)
                umipm1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 17).Value)
                umipm2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 18).Value)
                umipm3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 19).Value)
                umipm4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 20).Value)
                umipmpk = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 21).Value)
                umipmp = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 21).Value)
                umipmpk2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 22).Value)
                umipmp2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 22).Value)
                yak1k = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 23).Value)
                yak1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 23).Value)
                yak2k = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 23).Value)
                yak2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 23).Value)
                sin1k = Util.checkDBNullValue(DataGridView1(dd * 4 + 1, 24).Value)
                sin1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 24).Value)
                sin2k = Util.checkDBNullValue(DataGridView1(dd * 4 + 3, 24).Value)
                sin2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 24).Value)
                gyo1 = Util.checkDBNullValue(DataGridView2(dd, 0).Value)
                gyo2 = Util.checkDBNullValue(DataGridView2(dd, 1).Value)
                gyo3 = Util.checkDBNullValue(DataGridView2(dd, 2).Value)
                gyo4 = Util.checkDBNullValue(DataGridView2(dd, 3).Value)
                gyo5 = Util.checkDBNullValue(DataGridView2(dd, 4).Value)
                gyo6 = Util.checkDBNullValue(DataGridView2(dd, 5).Value)
                gyo7 = Util.checkDBNullValue(DataGridView2(dd, 6).Value)
                tuka1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 25).Value)
                tuka2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 26).Value)
                tuka3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 25).Value)
                tuka4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 26).Value)
                tukh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 27).Value)
                tukh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 28).Value)
                tukh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 29).Value)
                tukh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 27).Value)
                tukh5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 28).Value)
                tukh6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 29).Value)
                tuky1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 30).Value)
                tuky2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 31).Value)
                tuky3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 30).Value)
                tuky4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 31).Value)
                hana1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 32).Value)
                hana2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 33).Value)
                hana3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 32).Value)
                hana4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 33).Value)
                hanh1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 34).Value)
                hanh2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 35).Value)
                hanh3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 36).Value)
                hanh4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 34).Value)
                hanh5 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 35).Value)
                hanh6 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 36).Value)
                hany1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 37).Value)
                hany2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 38).Value)
                hany3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 37).Value)
                hany4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 38).Value)
                umia1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 39).Value)
                umia2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 39).Value)
                umih1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 40).Value)
                umih2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 41).Value)
                umih3 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 40).Value)
                umih4 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 41).Value)
                umiy1 = Util.checkDBNullValue(DataGridView1(dd * 4 + 2, 42).Value)
                umiy2 = Util.checkDBNullValue(DataGridView1(dd * 4 + 4, 42).Value)

                For r As Integer = 2 To 42
                    If r <> 6 Then
                        Dim u As Type = GetType(週間表)
                        If DataGridView1(dd * 4 + 2, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("cra" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("cra" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                        If DataGridView1(dd * 4 + 4, r).Style.BackColor = Color.FromArgb(255, 192, 255) Then
                            u.InvokeMember("crb" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {1})
                        Else
                            u.InvokeMember("crb" & r + 1, Reflection.BindingFlags.SetField, Nothing, Me, {0})
                        End If
                    End If
                Next

                SQL = "INSERT INTO SHyo3 VALUES ('" & ymd & "', '" & nyuam1 & "', '" & nyuam2 & "', '" & nyupm1 & "', '" & nyupm2 & "', '" & tyaam1 & "', '" & tyaam2 & "', '" & tyapm1 & "', '" & tyapm2 & "', '" & tukam1 & "', '" & tukam2 & "', '" & tukam3 & "', '" & tukam4 & "', '" & tukampk & "', '" & tukamp & "', '" & tukpm1 & "', '" & tukpm2 & "', '" & tukpm3 & "', '" & tukpm4 & "', '" & tukpmpk & "', '" & tukpmp & "', '" & hanam1 & "', '" & hanam2 & "', '" & hanam3 & "', '" & hanam4 & "', '" & hanampk & "', '" & hanamp & "', '" & hanpm1 & "', '" & hanpm2 & "', '" & hanpm3 & "', '" & hanpm4 & "', '" & hanpmpk & "', '" & hanpmp & "', '" & umiam1 & "', '" & umiam2 & "', '" & umiam3 & "', '" & umiam4 & "', '" & umiampk & "', '" & umiamp & "', '" & umiampk2 & "', '" & umiamp2 & "', '" & umipm1 & "', '" & umipm2 & "', '" & umipm3 & "', '" & umipm4 & "', '" & umipmpk & "', '" & umipmp & "', '" & umipmpk2 & "', '" & umipmp2 & "', '" & yak1k & "', '" & yak1 & "', '" & yak2k & "', '" & yak2 & "', '" & sin1k & "', '" & sin1 & "', '" & sin2k & "', '" & sin2 & "', '" & gyo1 & "', '" & gyo2 & "', '" & gyo3 & "', '" & gyo4 & "', '" & gyo5 & "', '" & gyo6 & "', '" & gyo7 & "', '" & tuka1 & "', '" & tuka2 & "', '" & tuka3 & "', '" & tuka4 & "', '" & tukh1 & "', '" & tukh2 & "', '" & tukh3 & "', '" & tukh4 & "', '" & tukh5 & "', '" & tukh6 & "', '" & tuky1 & "', '" & tuky2 & "', '" & tuky3 & "', '" & tuky4 & "', '" & hana1 & "', '" & hana2 & "', '" & hana3 & "', '" & hana4 & "', '" & hanh1 & "', '" & hanh2 & "', '" & hanh3 & "', '" & hanh4 & "', '" & hanh5 & "', '" & hanh6 & "', '" & hany1 & "', '" & hany2 & "', '" & hany3 & "', '" & hany4 & "', '" & umia1 & "', '" & umia2 & "', '" & umih1 & "', '" & umih2 & "', '" & umih3 & "', '" & umih4 & "', '" & umiy1 & "', '" & umiy2 & "', '" & cra3 & "', '" & cra4 & "', '" & cra5 & "', '" & cra6 & "', '" & cra8 & "', '" & cra9 & "', '" & cra10 & "', '" & cra11 & "', '" & cra12 & "', '" & cra13 & "', '" & cra14 & "', '" & cra15 & "', '" & cra16 & "', '" & cra17 & "', '" & cra18 & "', '" & cra19 & "', '" & cra20 & "', '" & cra21 & "', '" & cra22 & "', '" & cra23 & "', '" & cra24 & "', '" & cra25 & "', '" & cra26 & "', '" & cra27 & "', '" & cra28 & "', '" & cra29 & "', '" & cra30 & "', '" & cra31 & "', '" & cra32 & "', '" & cra33 & "', '" & cra34 & "', '" & cra35 & "', '" & cra36 & "', '" & cra37 & "', '" & cra38 & "', '" & cra39 & "', '" & cra40 & "', '" & cra41 & "', '" & cra42 & "', '" & cra43 & "', '" & crb3 & "', '" & crb4 & "', '" & crb5 & "', '" & crb6 & "', '" & crb8 & "', '" & crb9 & "', '" & crb10 & "', '" & crb11 & "', '" & crb12 & "', '" & crb13 & "', '" & crb14 & "', '" & crb15 & "', '" & crb16 & "', '" & crb17 & "', '" & crb18 & "', '" & crb19 & "', '" & crb20 & "', '" & crb21 & "', '" & crb22 & "', '" & crb23 & "', '" & crb24 & "', '" & crb25 & "', '" & crb26 & "', '" & crb27 & "', '" & crb28 & "', '" & crb29 & "', '" & crb30 & "', '" & crb31 & "', '" & crb32 & "', '" & crb33 & "', '" & crb34 & "', '" & crb35 & "', '" & crb36 & "', '" & crb37 & "', '" & crb38 & "', '" & crb39 & "', '" & crb40 & "', '" & crb41 & "', '" & crb42 & "', '" & crb43 & "')"
                cnn.Execute(SQL)
            Next
        End If
        cnn.Close()

        KinnmuwariTouroku()

    End Sub

    Private Sub KinnmuwariTouroku()
        If MsgBox("パートの勤務割に登録してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "ﾊﾟｰﾄ勤務割登録確認") = MsgBoxResult.No Then
            Return
        End If

        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim rsnextmonth As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        cnn.Open(TopForm.DB_Work)
        Dim SQL As String = ""
        Dim SQLnextmonth As String = ""
        Dim SQL2 As String = ""
        Dim updateSQL As String = ""

        Dim M As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        M = M.AddMonths(1)

        SQL2 = "SELECT * FROM SNam"
        rs2.Open(SQL2, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        If rbn2F.Checked = True Then
            Dim floar As String = 2
            Dim moriPwork1, moriPname1, moriPwork2, moriPname2, hosiPwork1, hosiPname1, hosiPwork2, hosiPname2, soraPwork1, soraPname1, soraPwork2, soraPname2 As String
            
            SQL = "SELECT * FROM KinD WHERE YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
            rs.Open(SQL, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If DataGridView1(2, 0).Value > 22 Then
                SQLnextmonth = "SELECT * FROM KinD WHERE YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(M, 5, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
                rsnextmonth.Open(SQLnextmonth, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            End If

            If rs.RecordCount <= 1 Then
                MsgBox("勤務割に該当月の登録データがありません")
            Else
                Dim listSQL As List(Of String) = New List(Of String)
                For dd As Integer = 0 To 6
                    moriPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 11).Value)
                    moriPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 11).Value)
                    moriPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 11).Value)
                    moriPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 11).Value)
                    hosiPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 16).Value)
                    hosiPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 16).Value)
                    hosiPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 16).Value)
                    hosiPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 16).Value)
                    soraPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 21).Value)
                    soraPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 21).Value)
                    soraPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 21).Value)
                    soraPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 21).Value)

                    Dim partname() As String = {moriPname1, moriPname2, hosiPname1, hosiPname2, soraPname1, soraPname2}
                    Dim partwork() As String = {moriPwork1, moriPwork2, hosiPwork1, hosiPwork2, soraPwork1, soraPwork2}

                    For i As Integer = 0 To 5
                        If partname(i) = "" Then

                        Else
                            If partnamecheck(rs, rs2, partname(i)) = True Then
                                Dim Hi As Integer = DataGridView1(4 * dd + 2, 0).Value
                                If Val(DataGridView1(2, 0).Value) - Val(DataGridView1(4 * dd + 2, 0).Value) > 20 Then   '月またぎがある週
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(M, 5, 2) & "')"
                                Else    '月またぎがない週
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "')"
                                End If
                                listSQL.Add(updateSQL)
                            Else
                                If i = 0 OrElse i = 1 Then
                                    MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：森：" & partname(i))
                                ElseIf i = 2 OrElse i = 3 Then
                                    MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：星：" & partname(i))
                                ElseIf i >= 4 Then
                                    MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：空：" & partname(i))
                                End If
                                Exit Sub
                            End If
                        End If

                    Next

                    'for each でのやり方
                    'For Each Name As String In partname
                    '    If partnamecheck(rs, rs2, Name) = True Then

                    '    Else

                    '    End If
                    'Next

                Next

                For Each i As String In listSQL
                    cnn.Execute(i)
                Next

            End If
            cnn.Close()

        ElseIf rbn3F.Checked = True Then
            Dim floar As String = 3
            Dim tukiPwork1, tukiPname1, tukiPwork2, tukiPname2, hanaPwork1, hanaPname1, hanaPwork2, hanaPname2, umiPwork1, umiPname1, umiPwork2, umiPname2, umiPwork3, umiPname3, umiPwork4, umiPname4 As String

            SQL = "SELECT * FROM KinD WHERE YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "' AND (Seq2='00' OR ('" & floar & "0' <= Seq2 AND Seq2 <= '" & floar & "9')) and Rdr = '' order by Seq"
            rs.Open(SQL, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.RecordCount <= 1 Then
                MsgBox("勤務割に該当月の登録データがありません")
            Else
                Dim listSQL As List(Of String) = New List(Of String)
                For dd As Integer = 0 To 6
                    tukiPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 11).Value)
                    tukiPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 11).Value)
                    tukiPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 11).Value)
                    tukiPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 11).Value)
                    hanaPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 16).Value)
                    hanaPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 16).Value)
                    hanaPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 16).Value)
                    hanaPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 16).Value)
                    umiPwork1 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 21).Value)
                    umiPname1 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 21).Value)
                    umiPwork2 = Util.checkDBNullValue(DataGridView1(4 * dd + 1, 22).Value)
                    umiPname2 = Util.checkDBNullValue(DataGridView1(4 * dd + 2, 22).Value)
                    umiPwork3 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 21).Value)
                    umiPname3 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 21).Value)
                    umiPwork4 = Util.checkDBNullValue(DataGridView1(4 * dd + 3, 22).Value)
                    umiPname4 = Util.checkDBNullValue(DataGridView1(4 * dd + 4, 22).Value)

                    Dim partname() As String = {tukiPname1, tukiPname2, hanaPname1, hanaPname2, umiPname1, umiPname2, umiPname3, umiPname4}
                    Dim partwork() As String = {tukiPwork1, tukiPwork2, hanaPwork1, hanaPwork2, umiPwork1, umiPwork2, umiPwork3, umiPwork4}
                    For i As Integer = 0 To 7
                        If partname(i) = "" Then

                        Else
                            If partnamecheck(rs, rs2, partname(i)) = True Then
                                Dim Hi As Integer = DataGridView1(4 * dd + 2, 0).Value
                                If Val(DataGridView1(2, 0).Value) - Val(DataGridView1(4 * dd + 2, 0).Value) > 20 Then   '月またぎがある週
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(M, 5, 2) & "')"
                                Else    '月またぎがない週
                                    updateSQL = "UPDATE KinD SET Y" & Hi & " = '" & partwork(i) & "', J" & Hi & " = '" & partwork(i) & "' WHERE (Nam LIKE '%" & partname(i) & "%') And (YM='" & ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 2) & "')"
                                End If
                                listSQL.Add(updateSQL)
                            Else
                                If i = 0 OrElse i = 1 Then
                                    MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：月：" & partname(i))
                                ElseIf i = 2 OrElse i = 3 Then
                                    MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：花：" & partname(i))
                                ElseIf i >= 4 Then
                                    MsgBox("ﾊﾟｰﾄの勤務割が登録されていません。" & vbCrLf & DataGridView1(4 * dd + 2, 0).Value & "日：海：" & partname(i))
                                End If
                                Exit Sub
                            End If
                        End If

                    Next

                Next

                For Each i As String In listSQL
                    cnn.Execute(i)
                Next

            End If
            cnn.Close()


        End If




    End Sub

    Private Function partnamecheck(rs As ADODB.Recordset, rs2 As ADODB.Recordset, Pname As String) As Boolean

        rs.MoveFirst()
        rs2.MoveFirst()

        If Pname <> "" Then
            '勤務割のほうに名前があるか
            While Not rs.EOF
                If System.Text.RegularExpressions.Regex.IsMatch(rs.Fields("Nam").Value, "^" & Pname) = True Then
                    Return True
                End If
                rs.MoveNext()
            End While
            
            '名前がない場合
            Return False

        End If

        Return False

    End Function

    Private Sub btnSakujo_Click(sender As System.Object, e As System.EventArgs) Handles btnSakujo.Click
        If MsgBox("削除してよろしいですか？", MsgBoxStyle.YesNo + vbExclamation, "削除確認") = MsgBoxResult.Yes Then
            Dim cnn As New ADODB.Connection
            cnn.Open(TopForm.DB_Work)

            Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
            Dim YmdAdd7 As Date = Ymd.AddDays(6)

            Dim SQL As String = ""
            If rbn2F.Checked = True Then
                SQL = "DELETE FROM SHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "#"
            ElseIf rbn3F.Checked = True Then
                SQL = "DELETE FROM SHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "#"
            End If

            cnn.Execute(SQL)
            cnn.Close()

            DataIndication()

        End If
    End Sub

    Private Sub btnTorikomi_Click(sender As System.Object, e As System.EventArgs) Handles btnTorikomi.Click
        Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        Dim YmdAdd7 As Date = Ymd.AddDays(6)

        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        If rbn2F.Checked = True Then
            If MsgBox("週間表3階の'入浴、夜勤等、特記事項'を取り込みますか？", MsgBoxStyle.YesNo + vbExclamation, "確認") = MsgBoxResult.Yes Then
                Dim sql As String = "select * from SHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
                cnn.Open(TopForm.DB_Work)
                rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

                'Datagridview1への表示
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    For RowNo As Integer = 2 To 38
                        If RowNo <= 5 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo - 1).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 3).Value
                            If rs.Fields(RowNo + 98).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 2, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                            End If
                            If rs.Fields(RowNo + 138).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 4, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                            End If
                        ElseIf RowNo = 22 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 27).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 28).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 29).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 30).Value
                        ElseIf RowNo = 23 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 30).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 31).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 32).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 33).Value
                        End If
                        ''色付け処理
                        'If rs.Fields(RowNo + 97).Value = "1" Then
                        '    DataGridView1(ColumnsNo * 4 + 2, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                        'End If
                        'If rs.Fields(RowNo + 137).Value = "1" Then
                        '    DataGridView1(ColumnsNo * 4 + 4, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                        'End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 7
                        DGV2Table.Rows(rowno2 - 1).Item("a" & ColumnsNo) = rs.Fields(rowno2 + 56).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

            End If
        ElseIf rbn3F.Checked = True Then
            If MsgBox("週間表2階の'入浴、夜勤等、特記事項'を取り込みますか？", MsgBoxStyle.YesNo + vbExclamation, "確認") = MsgBoxResult.Yes Then

                Dim sql As String = "select * from SHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
                cnn.Open(TopForm.DB_Work)
                rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

                'Datagridview1への表示
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    For RowNo As Integer = 2 To 42
                        If RowNo <= 5 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo - 1).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 3).Value
                            If rs.Fields(RowNo + 88).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 2, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                            End If
                            If rs.Fields(RowNo + 124).Value = 1 Then
                                DataGridView1(ColumnsNo * 4 + 4, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                            End If
                        ElseIf RowNo = 23 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 22).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 23).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 24).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 25).Value
                        ElseIf RowNo = 24 Then
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 1) = rs.Fields(RowNo + 25).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 2) = rs.Fields(RowNo + 26).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 3) = rs.Fields(RowNo + 27).Value
                            DGV1Table.Rows(RowNo).Item("a" & ColumnsNo * 4 + 4) = rs.Fields(RowNo + 28).Value
                        End If
                        ''色付け処理
                        'If rs.Fields(RowNo + 87).Value = "1" Then
                        '    DataGridView1(ColumnsNo * 4 + 2, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                        'ElseIf rs.Fields(RowNo + 123).Value = "1" Then
                        '    DataGridView1(ColumnsNo * 4 + 4, RowNo).Style.BackColor = Color.FromArgb(255, 192, 255)
                        'End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 7
                        DGV2Table.Rows(rowno2 - 1).Item("a" & ColumnsNo) = rs.Fields(rowno2 + 52).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

            End If
        End If
    End Sub

    Private Sub btnInnsatu_Click(sender As System.Object, e As System.EventArgs) Handles btnInnsatu.Click
        Dim Ymd As Date = ChangeSeireki(Strings.Left(lblYmd.Text, 9)) & "/" & Strings.Mid(lblYmd.Text, 5, 5)
        Dim YmdAdd7 As Date = Ymd.AddDays(6)

        If rbn2F.Checked = True Then        '2階の印刷
            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from SHyo WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.RecordCount > 0 Then
                Dim objExcel As Object
                Dim objWorkBooks As Object
                Dim objWorkBook As Object
                Dim oSheets As Object
                Dim oSheet As Object

                objExcel = CreateObject("Excel.Application")
                objWorkBooks = objExcel.Workbooks
                objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
                oSheets = objWorkBook.Worksheets
                oSheet = objWorkBook.Worksheets("週間食事表改")

                oSheet.Range("B2").Value = Strings.Mid(lblYmd.Text, 5, 2) & "月"

                Dim Cell() As String = {"D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE"}
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    oSheet.Range(Cell(ColumnsNo * 4 + 1) & "2").Value = DataGridView1((ColumnsNo * 4) + 2, 0).Value
                    For RowNo As Integer = 2 To 38
                        If RowNo <= 5 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo + 1).Value = rs.Fields(RowNo - 1).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo + 1).Value = rs.Fields(RowNo + 3).Value
                        ElseIf 7 <= RowNo And RowNo <= 10 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 2).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 8).Value
                        ElseIf RowNo = 11 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 2).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 3).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 8).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 9).Value
                        ElseIf 12 <= RowNo And RowNo <= 15 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 9).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 15).Value
                        ElseIf RowNo = 16 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 9).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 10).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 15).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 16).Value
                        ElseIf 17 <= RowNo And RowNo <= 20 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 16).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 22).Value
                        ElseIf RowNo = 21 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 16).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 17).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 22).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 23).Value
                        ElseIf RowNo = 22 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 23).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 24).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 25).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 26).Value
                        ElseIf RowNo = 23 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 26).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 27).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 28).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 29).Value
                        ElseIf RowNo = 24 Or RowNo = 25 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 36).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 38).Value
                        ElseIf RowNo = 26 Or RowNo = 27 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 38).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 40).Value
                        ElseIf RowNo = 28 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 40).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 41).Value
                        ElseIf RowNo = 29 Or RowNo = 30 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 41).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 43).Value
                        ElseIf RowNo = 31 Or RowNo = 32 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 43).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 45).Value
                        ElseIf RowNo = 33 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 45).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 46).Value
                        ElseIf RowNo = 34 Or RowNo = 35 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 46).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 48).Value
                        ElseIf RowNo = 36 Or RowNo = 37 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 48).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 50).Value
                        ElseIf RowNo = 38 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 50).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 51).Value
                        End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 7
                        oSheet.Range(Cell(ColumnsNo * 4) & rowno2 + 23).Value = rs.Fields(rowno2 + 52).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

                '保存
                objExcel.DisplayAlerts = False

                ' エクセル表示
                objExcel.Visible = True

                '印刷
                If TopForm.rbtnPreview.Checked = True Then
                    oSheet.PrintPreview(1)
                ElseIf TopForm.rbtnPrintout.Checked = True Then
                    oSheet.Printout(1)
                End If

                ' EXCEL解放
                objExcel.Quit()
                Marshal.ReleaseComObject(oSheet)
                Marshal.ReleaseComObject(objWorkBook)
                Marshal.ReleaseComObject(objExcel)
                oSheet = Nothing
                objWorkBook = Nothing
                objExcel = Nothing
            Else
                MsgBox("出力するデータがありません")
            End If

        ElseIf rbn3F.Checked = True Then        '3階の印刷
            Dim cnn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim sql As String = "select * from SHyo3 WHERE #" & Ymd & "# <= Ymd and Ymd <= #" & YmdAdd7 & "# order by Ymd"
            cnn.Open(TopForm.DB_Work)
            rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If rs.RecordCount > 0 Then
                Dim objExcel As Object
                Dim objWorkBooks As Object
                Dim objWorkBook As Object
                Dim oSheets As Object
                Dim oSheet As Object

                objExcel = CreateObject("Excel.Application")
                objWorkBooks = objExcel.Workbooks
                objWorkBook = objWorkBooks.Open(TopForm.excelFilePass)
                oSheets = objWorkBook.Worksheets
                oSheet = objWorkBook.Worksheets("週間食事表３改")

                oSheet.Range("B2").Value = Strings.Mid(lblYmd.Text, 5, 2) & "月"

                Dim Cell() As String = {"D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE"}
                Dim ColumnsNo As Integer = 0
                While Not rs.EOF
                    oSheet.Range(Cell(ColumnsNo * 4 + 1) & "2").Value = DataGridView1((ColumnsNo * 4) + 2, 0).Value
                    For RowNo As Integer = 2 To 42
                        If RowNo <= 5 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo + 1).Value = rs.Fields(RowNo - 1).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo + 1).Value = rs.Fields(RowNo + 3).Value
                        ElseIf 7 <= RowNo And RowNo <= 10 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 2).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 8).Value
                        ElseIf RowNo = 11 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 2).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 3).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 8).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 9).Value
                        ElseIf 12 <= RowNo And RowNo <= 15 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 9).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 15).Value
                        ElseIf RowNo = 16 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 9).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 10).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 15).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 16).Value
                        ElseIf 17 <= RowNo And RowNo <= 20 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 16).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 24).Value
                        ElseIf RowNo = 21 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 16).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 17).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 24).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 25).Value
                        ElseIf RowNo = 22 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 17).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 18).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 25).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 26).Value
                        ElseIf RowNo = 23 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 26).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 27).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 28).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 29).Value
                        ElseIf RowNo = 24 Then
                            oSheet.Range(Cell(ColumnsNo * 4) & RowNo).Value = rs.Fields(RowNo + 29).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo).Value = rs.Fields(RowNo + 30).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo).Value = rs.Fields(RowNo + 31).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 3) & RowNo).Value = rs.Fields(RowNo + 32).Value
                        ElseIf RowNo = 25 Or RowNo = 26 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 39).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 41).Value
                        ElseIf RowNo = 27 Or RowNo = 28 Or RowNo = 29 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 41).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 44).Value
                        ElseIf RowNo = 30 Or RowNo = 31 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 44).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 46).Value
                        ElseIf RowNo = 32 Or RowNo = 33 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 46).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 48).Value
                        ElseIf RowNo = 34 Or RowNo = 35 Or RowNo = 36 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 48).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 51).Value
                        ElseIf RowNo = 37 Or RowNo = 38 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 51).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 53).Value
                        ElseIf RowNo = 39 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 53).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 54).Value
                        ElseIf RowNo = 40 Or RowNo = 41 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 54).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 56).Value
                        ElseIf RowNo = 42 Then
                            oSheet.Range(Cell(ColumnsNo * 4 + 1) & RowNo + 7).Value = rs.Fields(RowNo + 56).Value
                            oSheet.Range(Cell(ColumnsNo * 4 + 2) & RowNo + 7).Value = rs.Fields(RowNo + 57).Value
                        End If
                    Next

                    'Datagridview2への表示
                    For rowno2 As Integer = 1 To 7
                        oSheet.Range(Cell(ColumnsNo * 4) & rowno2 + 24).Value = rs.Fields(rowno2 + 56).Value
                    Next

                    rs.MoveNext()

                    ColumnsNo = ColumnsNo + 1
                End While
                cnn.Close()

                '保存
                objExcel.DisplayAlerts = False

                ' エクセル表示
                objExcel.Visible = True

                '印刷
                If TopForm.rbtnPreview.Checked = True Then
                    oSheet.PrintPreview(1)
                ElseIf TopForm.rbtnPrintout.Checked = True Then
                    oSheet.Printout(1)
                End If

                ' EXCEL解放
                objExcel.Quit()
                Marshal.ReleaseComObject(oSheet)
                Marshal.ReleaseComObject(objWorkBook)
                Marshal.ReleaseComObject(objExcel)
                oSheet = Nothing
                objWorkBook = Nothing
                objExcel = Nothing
            Else
                MsgBox("出力するデータがありません")
            End If
        End If

    End Sub

End Class