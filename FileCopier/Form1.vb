Imports Scripting

Public Class Form1
    Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Integer
    Public name As String
    Public Size As Single
    Public now_s As Boolean
    Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Integer)
    Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Integer, lpBytesPerSector As Integer, lpNumberOfFreeClusters As Integer, lpTotalNumberOfClusters As Integer) As Integer
    Public adress As String
    Public n As String
    Public l As Integer
    'Dim aa As System.IO.StreamWriter = New System.IO.StreamWriter(System.Windows.Forms.Application.StartupPath & "\log\" & Now & ".log")
    Public Shared Sub delay(ByVal Interval)
        Dim __time As DateTime = DateTime.Now
        Dim __Span As Int64 = Interval * 10000   '因为时间是以100纳秒为单位。   
        While (DateTime.Now.Ticks - __time.Ticks < __Span)
            Application.DoEvents()
        End While
    End Sub
    Private Sub cout(ByVal message As String)
        Dim s As String
        s = "#" & now & "#" & message
        ListBox1.Items.Add(s)
        ListBox1.SelectedItem = ListBox1.Items(ListBox1.Items.Count - 1)
        'aa.WriteLine(1, s)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox("注意！由于copy命令缺陷，复制文件夹时不会复制母文件夹！（可在文件夹内再建立一个相同名称的子文件夹）", MsgBoxStyle.Information, "注意")
        n = 0
        l = 0
        cout("请选择功能")
        now_s = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim se As String
        FolderBrowserDialog1.Description = "请选择要复制文件的目录"
        se = Me.FolderBrowserDialog1.ShowDialog()
        If se = DialogResult.OK Then
            adress = FolderBrowserDialog1.SelectedPath
            cout("你选择的文件夹路径为：" & adress)
            n = 1
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            adress = OpenFileDialog1.FileName
            cout("你选择的文件为：" & adress)
            n = 2
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Button1.Text = "停止") Then
            Timer1.Stop()
            now_s = False
            Button1.Text = "开始"
            cout("操作正在停止中")
            Exit Sub
        End If
        cout("操作已开始，请不要关闭程序")
        Dim Fso As New FileSystemObject
        Dim FileSys As New FileSystemObject
        Dim FolderObj As Folder
        Dim ws
        ws = CreateObject("WScript.Shell")
        FileSys = CreateObject("scripting.filesystemobject")
        Try
            Size = FileLen(adress) / 1024 / 1024
            cout("需复制的文件大小为" & Size & "MB")
            Timer1.Start()
        Catch ex As Exception
            Try
                FolderObj = FileSys.GetFolder(adress)
                Size = CStr(FolderObj.Size) / 1024 / 1024
                cout("需复制的文件夹大小为" & Size & "MB")
                Timer1.Start()
            Catch ex_2 As Exception
                MsgBox(ex_2.Message, MsgBoxStyle.Exclamation, "错误")
            End Try
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'cout("操作已开始，请不要关闭程序")

        If (Not now_s) Then
            Button1.Text = "停止"
            'Button1.Enabled = False
            Try
                now_s = True
                Dim type As Integer
                Dim kind As Integer
                Dim Fso As New FileSystemObject
                Dim drvDisk As Drive, strResult As String
                Dim FileSys As New FileSystemObject
                Dim cmdd As String
                Dim new_folder As String
                Dim ws
                ws = CreateObject("WScript.Shell")
                FileSys = CreateObject("scripting.filesystemobject")
                For type = 65 To 90
                    kind = GetDriveType(Chr(type) & ":\")
                    If kind = 2 Then
                        cout("已检测到可移动磁盘, 盘符为" & Chr(type))
                        drvDisk = Fso.GetDrive(Chr(type) & ":")
                        strResult = "磁盘卷标:" & drvDisk.VolumeName & vbCrLf
                        cout(strResult)
                        strResult = "磁盘序列号:" & drvDisk.SerialNumber & vbCrLf
                        cout(strResult)
                        strResult = "磁盘类型:" & drvDisk.DriveType & vbCrLf
                        cout(strResult)
                        strResult = "文件系统:" & drvDisk.FileSystem & vbCrLf
                        cout(strResult)
                        strResult = "磁盘容量(GB): " & FormatNumber(((drvDisk.TotalSize / 1024) / 1024) / 1024, 2, , , Microsoft.VisualBasic.TriState.True) & vbCrLf
                        cout(strResult)
                        strResult = "可用空间(GB): " & FormatNumber(((drvDisk.FreeSpace / 1024) / 1024) / 1024, 2, , , Microsoft.VisualBasic.TriState.True) & vbCrLf
                        cout(strResult)
                        strResult = "已用空间(GB): " & FormatNumber(((((drvDisk.TotalSize - drvDisk.FreeSpace) / 1024) / 1024) / 1024), 2, , , Microsoft.VisualBasic.TriState.True)
                        cout(strResult)
                        If drvDisk.FreeSpace / 1024 / 1024 > Size Then
                            If (n = 1) Then
                                new_folder = adress.Substring(adress.LastIndexOf("\") + 1, adress.Length - adress.LastIndexOf("\") - 1)
                                cmdd = "cmd /c xcopy " & Chr(34) & adress & Chr(34) & " " & Chr(34) & Chr(type) & ":\" & new_folder & "\" & Chr(34) & " /Y /E /V /F /K /H"
                            End If
                            If (n = 2) Then
                                cmdd = "cmd /c xcopy " & Chr(34) & adress & Chr(34) & " " & Chr(34) & Chr(type) & ":\" & Chr(34) & " /Y /V /F /K /H"
                            End If
                            cout("可用空间充足，即将开始复制，请注意弹出的CMD窗口")
                            cout("正在执行：" & cmdd)
                            cout("等待用户输入")
                            ws = CreateObject("wscript.shell")
                            ws.Run(cmdd, 1, True)
                            cout("复制成功！")
                        Else
                            cout("可用空间不足，缺少" & -(Int(drvDisk.FreeSpace / 1024 / 1024) - Size) & "MB，请删除一些文件后再试")
                        End If
                        cout("操作完成，请拔出U盘")
                        Do
                            kind = GetDriveType(Chr(type) & ":\")
                            If kind <> 2 Then
                                Exit Do
                            End If
                            delay(2)
                        Loop
                        'Button1.Enabled = True
                        cout("U盘已拔出，继续扫描U盘")
                    End If
                Next type
                now_s = False
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "错误")
                now_s = False
            End Try
        End If

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Dispose()
        Dim ws
        ws = CreateObject("WScript.Shell")
        Try
            ws.Run("taskkill /F /IM 文件自动拷贝.exe /T", 1, True)
        Catch ex As Exception

        End Try
    End Sub

End Class
