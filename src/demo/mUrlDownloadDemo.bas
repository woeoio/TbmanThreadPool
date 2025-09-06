Module mUrlDownloadDemo

    ' WinHttp API 声明
    Private Declare Function WinHttpOpen Lib "winhttp" Alias "WinHttpOpen" ( _
        ByVal pwszUserAgent As LongPtr, _
        ByVal dwAccessType As Long, _
        ByVal pwszProxyName As LongPtr, _
        ByVal pwszProxyBypass As LongPtr, _
        ByVal dwFlags As Long _
    ) As LongPtr

    Private Declare Function WinHttpConnect Lib "winhttp" Alias "WinHttpConnect" ( _
        ByVal hSession As LongPtr, _
        ByVal pwszServerName As LongPtr, _
        ByVal nServerPort As Long, _
        ByVal dwReserved As Long _
    ) As LongPtr

    Private Declare Function WinHttpOpenRequest Lib "winhttp" Alias "WinHttpOpenRequest" ( _
        ByVal hConnect As LongPtr, _
        ByVal pwszVerb As LongPtr, _
        ByVal pwszObjectName As LongPtr, _
        ByVal pwszVersion As LongPtr, _
        ByVal pwszReferrer As LongPtr, _
        ByVal ppwszAcceptTypes As LongPtr, _
        ByVal dwFlags As Long _
    ) As LongPtr

    Private Declare Function WinHttpSendRequest Lib "winhttp" Alias "WinHttpSendRequest" ( _
        ByVal hRequest As LongPtr, _
        ByVal lpszHeaders As LongPtr, _
        ByVal dwHeadersLength As Long, _
        ByVal lpOptional As LongPtr, _
        ByVal dwOptionalLength As Long, _
        ByVal dwTotalLength As Long, _
        ByVal dwContext As LongPtr _
    ) As Long

    Private Declare Function WinHttpReceiveResponse Lib "winhttp" Alias "WinHttpReceiveResponse" ( _
        ByVal hRequest As LongPtr, _
        ByVal lpReserved As LongPtr _
    ) As Long

    Private Declare Function WinHttpReadData Lib "winhttp" Alias "WinHttpReadData" ( _
        ByVal hRequest As LongPtr, _
        ByVal lpBuffer As LongPtr, _
        ByVal dwNumberOfBytesToRead As Long, _
        ByRef lpdwNumberOfBytesRead As Long _
    ) As Long

    Private Declare Function WinHttpCloseHandle Lib "winhttp" Alias "WinHttpCloseHandle" ( _
        ByVal hInternet As LongPtr _
    ) As Long

    Private Type DownloadInfo
        Url As String
        OutputFile As String
    End Type

    Private Function DownloadProc(ByVal param As LongPtr) As Long
        Dim task As cThread
        Set task = mTask.ObjectFromPtr(param)
        
        ' 获取下载信息
        Dim info As DownloadInfo
        CopyMemory info, ByVal task.Tag, LenB(task.Tag)
        
        ' 检查任务是否被取消
        If task.CancelRequested Then
            DownloadProc = 1
            Exit Function
        End If
        
        ' 解析URL
        Dim url As tURL
        If Not ParseURL(info.Url, url) Then
            DownloadProc = 2
            Exit Function
        End If
        
        ' 初始化WinHttp
        Dim hSession As LongPtr
        hSession = WinHttpOpen(StrPtr("TwinBasic Downloader"), 0, 0, 0, 0)
        If hSession = 0 Then
            DownloadProc = 3
            Exit Function
        End If
        
        ' 连接服务器
        Dim hConnect As LongPtr
        hConnect = WinHttpConnect(hSession, StrPtr(url.Host), url.Port, 0)
        If hConnect = 0 Then
            WinHttpCloseHandle hSession
            DownloadProc = 4
            Exit Function
        End If
        
        ' 创建请求
        Dim hRequest As LongPtr
        hRequest = WinHttpOpenRequest(hConnect, StrPtr("GET"), StrPtr(url.Path), 0, 0, 0, 0)
        If hRequest = 0 Then
            WinHttpCloseHandle hConnect
            WinHttpCloseHandle hSession
            DownloadProc = 5
            Exit Function
        End If
        
        ' 发送请求
        If WinHttpSendRequest(hRequest, 0, 0, 0, 0, 0, 0) = 0 Then
            WinHttpCloseHandle hRequest
            WinHttpCloseHandle hConnect
            WinHttpCloseHandle hSession
            DownloadProc = 6
            Exit Function
        End If
        
        ' 等待响应
        If WinHttpReceiveResponse(hRequest, 0) = 0 Then
            WinHttpCloseHandle hRequest
            WinHttpCloseHandle hConnect
            WinHttpCloseHandle hSession
            DownloadProc = 7
            Exit Function
        End If
        
        ' 读取数据
        Dim buffer(1 To 8192) As Byte
        Dim bytesRead As Long
        Dim fNum As Integer
        
        fNum = FreeFile
        Open info.OutputFile For Binary As #fNum
        
        Do
            ' 检查任务是否被取消
            If task.CancelRequested Then
                Close #fNum
                Kill info.OutputFile  ' 删除未完成的文件
                WinHttpCloseHandle hRequest
                WinHttpCloseHandle hConnect
                WinHttpCloseHandle hSession
                DownloadProc = 8
                Exit Function
            End If
            
            If WinHttpReadData(hRequest, VarPtr(buffer(1)), UBound(buffer), bytesRead) = 0 Then
                Exit Do
            End If
            
            If bytesRead > 0 Then
                Put #fNum, , buffer
            End If
        Loop While bytesRead > 0
        
        Close #fNum
        
        ' 清理
        WinHttpCloseHandle hRequest
        WinHttpCloseHandle hConnect
        WinHttpCloseHandle hSession
        
        DownloadProc = 0
    End Function

    Private Type tURL
        Protocol As String
        Host As String
        Port As Long
        Path As String
    End Type

    Private Function ParseURL(ByVal strURL As String, ByRef outURL As tURL) As Boolean
        On Error GoTo ErrorHandler
        
        ' 设置默认值
        outURL.Protocol = "https"
        outURL.Port = 443
        
        ' 解析协议
        Dim pos As Long
        pos = InStr(1, strURL, "://")
        If pos > 0 Then
            outURL.Protocol = LCase$(Left$(strURL, pos - 1))
            strURL = Mid$(strURL, pos + 3)
            
            ' 根据协议设置默认端口
            If outURL.Protocol = "http" Then
                outURL.Port = 80
            End If
        End If
        
        ' 解析主机和端口
        pos = InStr(1, strURL, "/")
        Dim hostPort As String
        If pos > 0 Then
            hostPort = Left$(strURL, pos - 1)
            outURL.Path = Mid$(strURL, pos)
        Else
            hostPort = strURL
            outURL.Path = "/"
        End If
        
        ' 检查是否有端口
        pos = InStr(1, hostPort, ":")
        If pos > 0 Then
            outURL.Host = Left$(hostPort, pos - 1)
            outURL.Port = CLng(Mid$(hostPort, pos + 1))
        Else
            outURL.Host = hostPort
        End If
        
        ParseURL = True
        Exit Function
        
    ErrorHandler:
        ParseURL = False
    End Function

    Public Sub DownloadDemo()
        ' 创建输出目录
        Dim outputDir As String
        outputDir = "D:\Downloads\"
        MkDir outputDir
        
        ' 创建线程池
        Dim pool As New cThreadPool
        pool.Create 4  ' 使用4个线程
        pool.EnableLogging outputDir & "download_pool.log"
        
        ' 示例URL列表
        Dim urls() As String
        ReDim urls(1 To 100)
        Dim i As Long
        For i = 1 To 100
            urls(i) = "https://example.com/api/data" & i & ".json"
        Next
        
        ' 添加下载任务
        Dim info As DownloadInfo
        For i = 1 To UBound(urls)
            info.Url = urls(i)
            info.OutputFile = outputDir & "data" & i & ".json"
            
            ' 创建任务并设置下载信息
            Dim task As cThread
            Set task = pool.AddTask(AddressOf DownloadProc, info)
            
            ' 设置超时和重试策略
            task.SetTimeout 30000  ' 30秒超时
            task.SetRetryPolicy 3, 5000  ' 最多重试3次，间隔5秒
            
            Debug.Print "添加下载任务 " & i & ": " & info.Url
        Next
        
        ' 演示动态添加任务
        ' 等待1秒后添加新任务
        Sleep 1000
        
        ' 动态添加3个新任务
        For i = 1 To 3
            info.Url = "https://example.com/api/extra" & i & ".json"
            info.OutputFile = outputDir & "extra" & i & ".json"
            
            Set task = pool.AddTask(AddressOf DownloadProc, info)
            task.SetTimeout 30000
            task.Priority = Priority_High  ' 设置高优先级
            
            Debug.Print "动态添加下载任务: " & info.Url
        Next
        
        ' 等待所有任务完成
        pool.WaitForAll
        
        ' 输出统计信息
        Debug.Print "下载完成!"
        Debug.Print pool.GetPoolStats
        
        ' 清理线程池
        pool.Shutdown
    End Sub

End Module
