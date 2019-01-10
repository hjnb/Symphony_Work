Imports System.Text

Public Class workDataGridView
    Inherits DataGridView

    Private unitDictionary As Dictionary(Of String, String)
    Private wordDictionary As Dictionary(Of String, String)

    Protected Overrides Sub InitLayout()
        MyBase.InitLayout()

        DoubleBuffered = True
    End Sub

    Protected Overrides Function ProcessDialogKey(keyData As System.Windows.Forms.Keys) As Boolean
        Dim inputStr As String = If(Not IsNothing(Me.EditingControl), CType(Me.EditingControl, DataGridViewTextBoxEditingControl).Text, "")
        Dim columnName As String = Me.Columns(CurrentCell.ColumnIndex).Name
        If keyData = Keys.Enter Then
            If columnName = "Unt" OrElse columnName = "Rdr" OrElse columnName = "Nam" Then
                EndEdit()
                Return False
            ElseIf 7 <= Me.CurrentCell.ColumnIndex AndAlso Me.CurrentCell.ColumnIndex <= 37 Then
                If inputStr = "" OrElse ("A" <= inputStr.Substring(0, 1) AndAlso inputStr.Substring(0, 1) <= "T") Then
                    Return Me.ProcessTabKey(keyData)
                Else
                    'Y1～Y31列の編集終了時の処理、値の変換処理をする
                    Try
                        CType(Me.EditingControl, DataGridViewTextBoxEditingControl).Text = wordDictionary(inputStr)
                    Catch ex As KeyNotFoundException
                        MsgBox("正しく入力して下さい。", MsgBoxStyle.Exclamation)
                        EndEdit()
                        Return False
                    End Try
                End If
                Return Me.ProcessTabKey(keyData)
            Else
                Return Me.ProcessTabKey(keyData)
            End If
        ElseIf keyData = Keys.Back Then
            If columnName = "Unt" OrElse columnName = "Rdr" OrElse columnName.Substring(0, 1) = "Y" Then
                CurrentCell.Value = ""
                BeginEdit(False)
            ElseIf columnName = "Nam" Then
                BeginEdit(True)
            End If
            Return MyBase.ProcessDialogKey(keyData)
        Else
            Return MyBase.ProcessDialogKey(keyData)
        End If
    End Function

    Protected Overrides Function ProcessDataGridViewKey(e As System.Windows.Forms.KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            Dim columnName As String = Me.Columns(CurrentCell.ColumnIndex).Name
            If columnName = "Unt" OrElse columnName = "Rdr" OrElse columnName = "Nam" Then
                BeginEdit(True)
                Return False
            Else
                Me.ProcessTabKey(e.KeyCode)
                BeginEdit(True)
                Return False
            End If
        End If

        Dim tb As DataGridViewTextBoxEditingControl = CType(Me.EditingControl, DataGridViewTextBoxEditingControl)
        If Not IsNothing(tb) AndAlso ((e.KeyCode = Keys.Left AndAlso tb.SelectionStart = 0) OrElse (e.KeyCode = Keys.Right AndAlso tb.SelectionStart = tb.TextLength)) Then
            Return False
        Else
            Return MyBase.ProcessDataGridViewKey(e)
        End If
    End Function

    Private Sub workDataGridView_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellEndEdit
        Dim inputStr As String = If(IsDBNull(Me(e.ColumnIndex, e.RowIndex).Value), "", Me(e.ColumnIndex, e.RowIndex).Value)
        If Me.Columns(e.ColumnIndex).Name = "Unt" Then
            'ﾕﾆｯﾄ列の編集終了時、Seq2列のセルに対応した値を設定
            Try
                Me("Seq2", e.RowIndex).Value = unitDictionary(inputStr)
            Catch ex As KeyNotFoundException
                '何もしない
            End Try
        ElseIf Me.Columns(e.ColumnIndex).Name = "Rdr" Then
            If inputStr <> "" AndAlso Not ("A" <= inputStr.Substring(0, 1) AndAlso inputStr.Substring(0, 1) <= "Z") Then
                Me(e.ColumnIndex, e.RowIndex).Value = ""
            End If
        End If
    End Sub

    Private Sub workDataGridView_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellEnter
        Dim columnName As String = Me.Columns(e.ColumnIndex).Name
        If columnName = "Unt" Then
            Me.ImeMode = Windows.Forms.ImeMode.Hiragana
        ElseIf columnName = "Rdr" Then
            Me.ImeMode = Windows.Forms.ImeMode.Off
        ElseIf columnName = "Nam" Then
            Me.ImeMode = Windows.Forms.ImeMode.Hiragana
        ElseIf columnName.Substring(0, 1) = "Y" Then
            Me.ImeMode = Windows.Forms.ImeMode.Off
        End If
    End Sub

    Private Sub ExDataGridView_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles Me.CellPainting
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

    Private Sub workDataGridView_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles Me.EditingControlShowing
        Dim tb As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)
        tb.CharacterCasing = CharacterCasing.Upper
        If Me.Columns(Me.CurrentCell.ColumnIndex).Name = "Rdr" Then
            tb.MaxLength = 1
        ElseIf Me.Columns(Me.CurrentCell.ColumnIndex).Name = "Unt" Then
            'イベントハンドラを削除、追加
            RemoveHandler tb.KeyPress, AddressOf dgvTextBox_KeyPress
            AddHandler tb.KeyPress, AddressOf dgvTextBox_KeyPress
        End If
    End Sub

    Private Sub dgvTextBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        Dim text As String = CType(sender, DataGridViewTextBoxEditingControl).Text
        Dim lengthByte As Integer = Encoding.GetEncoding("Shift_JIS").GetByteCount(text)
        Dim limitLengthByte As Integer = 2

        If lengthByte >= limitLengthByte Then '設定されているバイト数以上の時
            If e.KeyChar = ChrW(Keys.Back) Then
                'Backspaceは入力可能
                e.Handled = False
            Else
                '入力できなくする
                e.Handled = True
            End If
        End If
    End Sub

    Public Sub setUnitDictionary(unitDictionary As Dictionary(Of String, String))
        Me.unitDictionary = unitDictionary
    End Sub

    Public Sub setWordDictionary(wordDictionary As Dictionary(Of String, String))
        Me.wordDictionary = wordDictionary
    End Sub
End Class
