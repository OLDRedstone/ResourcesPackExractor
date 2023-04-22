Public Class Form1
    Dim SourceFilePath As String
    Dim FilePathStrings As String()
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If MsgBox($"请确保当前程序放置的位置{vbCrLf}""{Application.StartupPath}""{vbCrLf}是你所想要输出的位置。{vbCrLf}如果想要更改，可以随时取消。", MsgBoxStyle.OkCancel, "提示") <> MsgBoxResult.Ok Then
            MsgBox("已退出。",, "提示")
            End
        End If
        If OpenFileDialog1.ShowDialog() <> DialogResult.OK Then
            MsgBox("已退出。",, "提示")
            End
        End If
        SourceFilePath = OpenFileDialog1.FileName
        FilePathStrings = New IO.StreamReader(SourceFilePath).ReadToEnd.Replace("}, """, $"}}, #""").Split("#")
        OutputData()
        MsgBox("导出完毕。")
        End
    End Sub


    Sub OutputData()
        For Each I In FilePathStrings
            Dim FilePath = FindMidWord("""", """", I)
            Dim Hex = FindMidWord("{""hash"": """, """", I)
            If FilePath = "" Or Hex = "" Then
                MsgBox("导入错误，已退出。",, "提示")
                End
            End If
            My.Computer.FileSystem.CreateDirectory(
                $"{Application.StartupPath}Output\{IO.Path.GetDirectoryName(FilePath)}")
            IO.File.Copy(
                $"{IO.Directory.GetParent(SourceFilePath).Parent.FullName}\Objects\{Mid(Hex, 1, 2)}\{Hex}",
                $"{Application.StartupPath}Output\{FilePath}",
                True)
        Next
    End Sub

    Function FindMidWord(startStr As String, endStr As String, text As String) As String
        Dim startIndex As Integer = text.IndexOf(startStr)
        Dim endIndex As Integer = text.IndexOf(endStr, startIndex + startStr.Length)
        If startIndex = -1 OrElse endIndex = -1 Then
            Return ""
        End If
        Dim contentStartIndex As Integer = startIndex + startStr.Length
        Dim contentLength As Integer = endIndex - contentStartIndex
        If contentLength < 0 Then
            Return ""
        End If
        Return text.Substring(contentStartIndex, contentLength)
    End Function
End Class
