Imports System.Reflection

Public Class workData

    'ユニット
    Public Property unt As String = ""
    'R
    Public Property rdr As String = ""
    '氏名
    Public Property nam As String = ""
    '予 or 変
    Public Property type As String = ""
    'day
    Public Property day1 As String = ""
    Public Property day2 As String = ""
    Public Property day3 As String = ""
    Public Property day4 As String = ""
    Public Property day5 As String = ""
    Public Property day6 As String = ""
    Public Property day7 As String = ""
    Public Property day8 As String = ""
    Public Property day9 As String = ""
    Public Property day10 As String = ""
    Public Property day11 As String = ""
    Public Property day12 As String = ""
    Public Property day13 As String = ""
    Public Property day14 As String = ""
    Public Property day15 As String = ""
    Public Property day16 As String = ""
    Public Property day17 As String = ""
    Public Property day18 As String = ""
    Public Property day19 As String = ""
    Public Property day20 As String = ""
    Public Property day21 As String = ""
    Public Property day22 As String = ""
    Public Property day23 As String = ""
    Public Property day24 As String = ""
    Public Property day25 As String = ""
    Public Property day26 As String = ""
    Public Property day27 As String = ""
    Public Property day28 As String = ""
    Public Property day29 As String = ""
    Public Property day30 As String = ""
    Public Property day31 As String = ""

    Public Sub New()

    End Sub

    Public Sub New(type As String)
        Me.type = type
    End Sub

    Public Sub New(unt As String, rdr As String, nam As String, type As String, dayArray() As String)
        Dim t As Type = GetType(workData)
        Me.unt = unt
        Me.rdr = rdr
        Me.nam = nam
        Me.type = type
        For i = 1 To 31
            t.InvokeMember("day" & i, BindingFlags.SetProperty, Nothing, Me, New String() {dayArray(i)})
        Next
    End Sub

End Class
