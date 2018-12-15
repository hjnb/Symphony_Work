Public Class 勤務割

    Private workDt As DataTable

    Private unitDictionary2F As Dictionary(Of String, String)
    Private unitDictionary3F As Dictionary(Of String, String)
    Private wordDictionary As Dictionary(Of String, String)
    Private dayCharArray() As String = {"日", "月", "火", "水", "木", "金", "土"}

    Private Sub 勤務割_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        createDictionary()

        initDgvWork()
        displayWork(ymBox.getADStr4Ym, "2")

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

    Private Sub initDgvWork()
        workDt = New DataTable()

    End Sub

    Private Sub displayWork(ymStr As String, floar As String)
        dgvWork.Columns.Clear()
        Dim year As Integer = CInt(ymStr.Split("/")(0))
        Dim month As Integer = CInt(ymStr.Split("/")(1))

    End Sub








    Private Sub floarRadioButton_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtn2F.CheckedChanged, rbtn3F.CheckedChanged
        Dim rbtn As RadioButton = CType(sender, RadioButton)
        If rbtn.Checked = True Then
            rbtn.BackColor = Color.FromArgb(255, 255, 0)
        Else
            rbtn.BackColor = Color.FromKnownColor(KnownColor.Control)
        End If
    End Sub
End Class