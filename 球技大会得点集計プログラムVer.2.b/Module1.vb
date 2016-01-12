Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

Public MustInherit Class PointData
    Public Property ClassName As String = ""
    Public Property Point As Integer = 0
    Public Property LostPoint As Integer = 0
    Public Property WinLose As Integer = 0
End Class

Public Class VolleyBall
    Inherits PointData
    Public Property WinSetCount As Integer = 0
    Public Property LoseSetCount As Integer = 0
    Public Property GetLoseRate As Integer = 0
End Class

Module Module1
    Private VolleyBall As Dictionary(Of String, VolleyBall)

    Sub Main()
        Console.WriteLine("～球技大会得点集計プログラムVer.2～")
        ShiraAuxiliarySys.DisableCloseButton()

        Console.WriteLine("処理種目別コードを入力し、Enterキーを押してください")
        Console.WriteLine("バレーボール→1")
        Console.WriteLine("バスケットボール→2")
        Console.WriteLine("卓球→3")
        Console.WriteLine("ドッチボール→4")
        Console.WriteLine("現在の集計結果の表示→5")
        Console.WriteLine("全データの初期化→6")
        Console.WriteLine("プログラムの終了→7")
        Dim strInPutKey As String = Console.ReadLine()
        Dim intInPutKey As Integer = ShiraAuxiliarySys.StrIntConv(strInPutKey)
    End Sub

End Module

Public Module ShiraAuxiliarySys
    Public Function StrIntConv(ByVal str As String) As Integer '"12345" => 12345
        str = StrConv(str, VbStrConv.Narrow)
        Dim outint As Integer = Integer.Parse(str)
        Return outint
    End Function
    Public Function StrStrConv(ByVal str As String) As String '"ａ１" => "A1"
        str = StrConv(str, VbStrConv.Narrow)
        str = str.ToUpper()
        Return str
    End Function

    '閉じるボタン無効化処理
    '(http://www.atmarkit.co.jp/fdotnet/dotnettips/896conclosebtn/conclosebtn.html)↓
    <DllImport("USER32.DLL")>
    Public Function GetSystemMenu(
      ByVal hWnd As IntPtr, ByVal bRevert As UInt32) As IntPtr
    End Function
    <DllImport("USER32.DLL")>
    Public Function RemoveMenu(
      ByVal hMenu As IntPtr, ByVal nPosition As UInt32,
      ByVal wFlags As UInt32) As UInt32
    End Function
    Const MF_BYCOMMAND As UInt32 = 0
    Const SC_CLOSE As UInt32 = &HF060
    Public Sub DisableCloseButton()
        Dim hWnd As IntPtr = Process.GetCurrentProcess.MainWindowHandle
        If hWnd <> IntPtr.Zero Then
            Dim hMenu As IntPtr = GetSystemMenu(hWnd, 0)
            RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
        End If
    End Sub
    '↑ここまで
End Module
