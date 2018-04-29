Imports System.Reflection

Public Class TopForm

    Public Shared DB_Work As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Owner\Desktop\sampleMdb\Work.mdb"
    Private workForm As workForm

    Private Sub TopForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        '画像の配置処理
        '
        '

    End Sub

    ''' <summary>
    ''' コントロールのDoubleBufferedプロパティをTrueにする
    ''' </summary>
    ''' <param name="control">対象のコントロール</param>
    Public Shared Sub EnableDoubleBuffering(control As Control)
        control.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, control, New Object() {True})
    End Sub

    Private Sub btnWork_Click(sender As Object, e As EventArgs) Handles btnWork.Click
        If IsNothing(workForm) OrElse workForm.IsDisposed Then
            workForm = New workForm()
            workForm.Owner = Me
            workForm.Show()
        End If
    End Sub
End Class
