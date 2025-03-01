Imports AxWMPLib
Imports Microsoft.Win32
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices

Public Class Form1
    Private WithEvents Timer1 As New System.Windows.Forms.Timer With {.Interval = 10} '定时器
    Private WithEvents MonMouce As New System.Windows.Forms.Timer With {.Interval = 10, .Enabled = True} '定时器监控鼠标
    Private WithEvents Exiting As New System.Windows.Forms.Timer With {.Interval = 20, .Enabled = False} '执行渐淡退出
    Private WithEvents BackGroundWorker As New System.ComponentModel.BackgroundWorker '后台操作修改启动界面
    Private WithEvents BackGroundWorker2 As New System.ComponentModel.BackgroundWorker '后台操作运行启动列表的程序
    Private TempVideoPath0 As String = Path.GetTempFileName() '0字节临时文件路径
    Private TempVideoPath As String = TempVideoPath0 & ".mp4" '临时文件真实路径
    Public Shared CommandLine As String = "" '进程的命令行，准备传递
    Public Shared HadStarted As Boolean = False
    Public Shared BoardReady As Boolean = False '指示白板是否有后台进程
    Public Shared Event Touching(x As Integer, y As Integer) '定义触控点击事件

    Structure TOUCHINPUT
        Dim x As Integer
        Dim y As Integer
    End Structure
    <DllImport("user32.dll")>
    Public Shared Function RegisterTouchWindow(hwnd As IntPtr, flags As UInteger) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Shared Function GetTouchInputInfo(hTouchInput As IntPtr, cInputs As Integer, <Out> pInputs As TOUCHINPUT(), cbSize As Integer) As Boolean
    End Function
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_TOUCH As Integer = &H240
        If m.Msg = WM_TOUCH Then
            Dim inputCount As Integer = m.WParam.ToInt32()
            Dim inputs(inputCount - 1) As TOUCHINPUT
            If GetTouchInputInfo(m.LParam, inputCount, inputs, Marshal.SizeOf(GetType(TOUCHINPUT))) Then
                For Each touch In inputs
                    RaiseEvent Touching(touch.x \ 100, touch.y \ 10) '触发事件并将触控坐标转换为像素
                Next
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        '应用程序全屏
        Me.TopMost = True
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Left = 0
        Me.Top = 0
        Me.Height = My.Computer.Screen.Bounds.Height
        Me.Width = My.Computer.Screen.Bounds.Width
        '播放器布局
        Player.Left = 0
        Player.Top = 0
        Player.Width = Me.ClientSize.Width
        Player.Height = Me.ClientSize.Height
    End Sub

    ' 窗体加载时提取资源文件
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '预设播放器属性
        Player.uiMode = "none" '隐藏播放器下方选项
        Player.settings.volume = 10 '音量10%
        Player.stretchToFit = True '自动缩放
        Player.enableContextMenu = False '禁止右键菜单

        '设置播放路径并播放
        If File.Exists("Start.mp4") Then
            Player.URL = "Start.mp4"
        Else
            WriteResourceToFile(My.Resources.YuanShenStart_MP4, TempVideoPath, True) ' 资源中读取视频，覆盖写入视频文件
            Player.URL = TempVideoPath
        End If
        Me.Text = Player.URL
        Player.Ctlcontrols.play()
        '进行其他操作
        BoardReady = System.Diagnostics.Process.GetProcessesByName("EasiNote").Length > 0 '扫描白板进程
        CommandLine = Interaction.Command() '更新进程的命令行
        If BackGroundWorker.IsBusy = False Then BackGroundWorker.RunWorkerAsync() '启动异步线程
        If BackGroundWorker2.IsBusy = False Then BackGroundWorker2.RunWorkerAsync() '启动异步线程
        Timer1.Enabled = True '定时器开始操作，计算剩余视频时间
    End Sub

    Private Sub BackGroundWorker_DoWork() Handles BackGroundWorker.DoWork
        Dim path1 = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Seewo\EasiNote5\Resources\Banner\"
        Try
            If File.Exists(path1 & "Banner.png") Then
                If Not File.Exists(path1 & "Banner_Backup.png") Then
                    Dim file0 As New FileInfo(path1 & "Banner.png")
                    file0.MoveTo(path1 & "Banner_Backup.png")
                End If
                If File.Exists("Background.png") Then
                    Dim file2 As New FileInfo("Background.png")
                    file2.CopyTo(path1 & "Banner.png", True)
                Else
                    WriteResourceToFile(My.Resources.SeewoStart_PNG, path1 & "Banner.png", True)
                End If
            Else
                If Directory.Exists(path1) Then
                    If File.Exists("Background.png") Then
                        Dim file2 As New FileInfo("Background.png")
                        file2.CopyTo(path1 & "Banner.png", True)
                    Else
                        WriteResourceToFile(My.Resources.SeewoStart_PNG, path1 & "Banner.png", True)
                    End If
                End If
            End If
        Catch
        End Try

        Dim possiblePaths As New List(Of String) From {
                "SOFTWARE\Seewo\EasiNote",
                "SOFTWARE\WOW6432Node\Seewo\EasiNote",
                "SOFTWARE\Seewo\EasiNote5"  ' 兼容旧版本
                }
        For Each path In possiblePaths
            Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(path)
                If key IsNot Nothing Then
                    Try
                        Dim Dir = key.GetValue("VersionPath") & "\Main\Assets\"
                        If File.Exists(Dir & "SplashScreen.png") Then
                            Dim file0 As New FileInfo(Dir & "SplashScreen.png")
                            If Not File.Exists(Dir & "SplashScreen_Backup.png") Then
                                file0.MoveTo(Dir & "SplashScreen_Backup.png") '创建备份
                            End If
                            file0 = Nothing
                        End If
                        If File.Exists("Background.png") Then
                            Dim file2 As New FileInfo("Background.png")
                            file2.CopyTo(Dir & "SplashScreen.png", True)
                        Else
                            WriteResourceToFile(My.Resources.SeewoStart_PNG, Dir & "SplashScreen.png", True)
                        End If
                    Catch
                    End Try
                End If
            End Using
        Next
    End Sub

    Private Sub BackGroundWorker2_DoWork() Handles BackGroundWorker2.DoWork
        On Error Resume Next
        If File.Exists("StartList.txt") Then
            Dim Paths As String() = File.ReadAllLines("StartList.txt")
            Err.Clear()
            For Each path0 In Paths
                Interaction.Shell(path0)
                Err.Clear()
            Next
        End If
        Err.Clear()
    End Sub

    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        'If e.KeyChar = Chr(27) Then'按下ESC
        'End If
        Timer1.Enabled = False
        StartYuanShenBoard()
        Exiting.Enabled = True
    End Sub

    Private Sub Touching_Fun() Handles MyClass.Touching
        Timer1.Enabled = False
        StartYuanShenBoard()
        Exiting.Enabled = True
    End Sub

    Private Sub MonMouce_Tick() Handles MonMouce.Tick
        Dim Mouce As Boolean = False
        If (Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left Then
            Mouce = True
        ElseIf (Control.MouseButtons And MouseButtons.Right) = MouseButtons.Right Then
            Mouce = True
        ElseIf (Control.MouseButtons And MouseButtons.Middle) = MouseButtons.Middle Then
            Mouce = True
        ElseIf (Control.MouseButtons And MouseButtons.XButton1) = MouseButtons.XButton1 Then
            Mouce = True
        ElseIf (Control.MouseButtons And MouseButtons.XButton2) = MouseButtons.XButton2 Then
            Mouce = True
        End If
        If Mouce And (MyBase.Focused Or Player.Focused) Then
            StartYuanShenBoard()
            Exiting.Enabled = True
        End If
    End Sub

    Private Sub Timer1_Tick() Handles Timer1.Tick
        If Player.currentMedia.duration <> 0 Then '视频有长度
            Dim Remain = GetRemainingSeconds(Player)
            If Remain <= 1 Then '剩余时长小于等于1秒，执行渐淡退出
                Exiting.Enabled = True
            ElseIf Remain = 0 Then '视频结束
                Timer1.Enabled = False
                System.Windows.Forms.Application.Exit()
            End If
            If BoardReady Then '白板启动过，有后台进程，因此启动速度会快很多
                If Player.currentMedia.duration >= 4.5 Then '视频总时长大于等于4.5秒
                    If Remain <= 4 Then
                        StartYuanShenBoard()
                    End If
                Else
                    StartYuanShenBoard()
                End If
            Else
                If Player.currentMedia.duration >= 6.5 Then '视频总时长大于等于6.5秒
                    If Remain <= 6 Then
                        StartYuanShenBoard()
                    End If
                Else
                    StartYuanShenBoard()
                End If
            End If
        End If
    End Sub

    Private Sub Exiting_Tick() Handles Exiting.Tick '执行渐淡退出
        If Me.Opacity >= 0.02 Then
            Me.Opacity -= 0.02
        Else
            Me.Visible = False
            Me.Opacity = 1
            Exiting.Enabled = False
            Player.close()
            System.Windows.Forms.Application.Exit()
        End If
    End Sub

    Public Sub StartYuanShenBoard()
        If Not HadStarted Then
            Dim possiblePaths As New List(Of String) From {
                "SOFTWARE\Seewo\EasiNote",
                "SOFTWARE\WOW6432Node\Seewo\EasiNote",
                "SOFTWARE\Seewo\EasiNote5"  ' 兼容旧版本
                }
            For Each path In possiblePaths
                Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(path)
                    If key IsNot Nothing Then
                        Try
                            Dim installDir = key.GetValue("ExePath")
                            Dim startInfo As New ProcessStartInfo()
                            startInfo.FileName = installDir
                            startInfo.Arguments = CommandLine
                            If Process.Start(startInfo).Id <> 0 Then
                                HadStarted = True
                                Exit Sub
                            End If
                        Catch
                        End Try
                    End If
                End Using
            Next
            Dim startInfo1 As New ProcessStartInfo()
            startInfo1.FileName = "C:\Program Files (x86)\Seewo\EasiNote5\swenlauncher\swenlauncher.exe"
            startInfo1.Arguments = CommandLine
            If Process.Start(startInfo1).Id <> 0 Then
                HadStarted = True
                Exit Sub
            End If
            Try
                If File.Exists(Environment.CurrentDirectory & "\path.txt") Then
                    Dim filep = File.ReadAllLines(Environment.CurrentDirectory & "\path.txt")(0)
                    Dim startInfo2 As New ProcessStartInfo()
                    startInfo2.FileName = filep
                    startInfo2.Arguments = CommandLine
                    If Process.Start(startInfo2).Id <> 0 Then
                        HadStarted = True
                        Exit Sub
                    End If
                End If
            Catch
            End Try
            MsgBox("希沃白板启动失败：找不到应用程序。请采用传统的启动方式！")
        End If
    End Sub

    ' 关闭窗体时清理临时文件
    Private Sub Form1_FormClosed() Handles MyBase.FormClosed
        If File.Exists(TempVideoPath) Then
            File.Delete(TempVideoPath)
        End If
        If File.Exists(TempVideoPath0) Then
            File.Delete(TempVideoPath0)
        End If
    End Sub

    ' 获取视频剩余秒数
    Private Function GetRemainingSeconds(Player As AxWindowsMediaPlayer) As Double
        If Player.playState = WMPLib.WMPPlayState.wmppsPlaying Then
            Dim totalSeconds As Double = Player.currentMedia.duration
            Dim currentPosition As Double = Player.Ctlcontrols.currentPosition
            Return Math.Max(totalSeconds - currentPosition, 0) ' 确保非负
        Else
            Return 0
        End If
    End Function

    Public Sub WriteResourceToFile(resource As Object, targetPath As String, Optional overwrite As Boolean = False)
        If File.Exists(targetPath) AndAlso Not overwrite Then Exit Sub ' 存在且不覆盖则退出
        Using fileStream As New FileStream(targetPath, FileMode.Create)
            If TypeOf resource Is String Then
                Using writer As New StreamWriter(fileStream)
                    writer.Write(resource)
                End Using
            Else
                Dim bytes As Byte() = DirectCast(resource, Byte())
                fileStream.Write(bytes, 0, bytes.Length)
            End If
        End Using
    End Sub
End Class