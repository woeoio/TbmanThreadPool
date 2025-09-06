# 示例代码

## 基础示例

### 1. 简单任务执行

```vb
' 简单的任务过程
Public Function SimpleTask(ByVal param As LongPtr) As Long
    Debug.Print "执行任务：" & param
    Sleep 1000  ' 模拟耗时操作
    SimpleTask = 0  ' 返回成功
End Function

' 使用示例
Public Sub SimpleTaskDemo()
    ' 创建线程池
    Dim pool As New cThreadPool
    pool.Create 4  ' 4个线程
    
    ' 添加任务
    Dim task As cThread
    Set task = pool.AddTask(AddressOf SimpleTask, "Hello")
    
    ' 等待完成
    task.WaitForCompletion
    
    ' 检查结果
    If task.Result = 0 Then
        Debug.Print "任务成功完成"
    Else
        Debug.Print "任务失败"
    End If
End Sub
```

### 2. 批量任务处理

```vb
Public Sub BatchTaskDemo()
    ' 创建线程池
    Dim pool As New cThreadPool
    pool.Create 4
    
    ' 创建多个任务
    Dim i As Long
    For i = 1 To 10
        pool.AddTask AddressOf SimpleTask, "任务" & i
    Next i
    
    ' 等待所有任务完成
    If pool.WaitForAll(10000) Then  ' 10秒超时
        Debug.Print "所有任务完成"
    Else
        Debug.Print "部分任务未完成"
    End If
End Sub
```

## 中级示例

### 1. 带回调的任务

```vb
' 任务完成回调
Public Function TaskCompleted(ByVal taskPtr As LongPtr, _
                            ByVal result As Long, _
                            ByVal wParam As LongPtr, _
                            ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(taskPtr)
    
    Debug.Print "任务完成：" & task.Tag
    Debug.Print "执行时间：" & task.ExecutionTime & "ms"
    
    TaskCompleted = 0
End Function

' 错误回调
Public Function TaskError(ByVal taskPtr As LongPtr, _
                         ByVal errorCode As Long, _
                         ByVal wParam As LongPtr, _
                         ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(taskPtr)
    
    Debug.Print "任务错误：" & task.ErrorDescription
    
    TaskError = 0
End Function

' 使用示例
Public Sub CallbackDemo()
    Dim pool As New cThreadPool
    pool.Create 4
    
    Dim task As cThread
    Set task = pool.AddTask(AddressOf SimpleTask, "带回调的任务")
    
    ' 设置回调
    task.SetOnComplete AddressOf TaskCompleted
    task.SetOnError AddressOf TaskError
    
    ' 等待完成
    task.WaitForCompletion
End Sub
```

### 2. 文件下载示例

```vb
' 下载信息结构
Private Type DownloadInfo
    Url As String
    OutputFile As String
End Type

' 下载过程
Public Function DownloadProc(ByVal param As LongPtr) As Long
    Dim info As DownloadInfo
    info = param
    
    ' 创建HTTP请求
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    On Error GoTo ErrorHandler
    
    ' 发送请求
    http.Open "GET", info.Url, False
    http.Send
    
    ' 检查响应
    If http.Status = 200 Then
        ' 保存文件
        Dim fNum As Integer
        fNum = FreeFile
        Open info.OutputFile For Binary As #fNum
            Put #fNum, , http.ResponseBody
        Close #fNum
        DownloadProc = 0
    Else
        DownloadProc = http.Status
    End If
    Exit Function
    
ErrorHandler:
    DownloadProc = Err.Number
End Function

' 下载演示
Public Sub DownloadDemo()
    Dim pool As New cThreadPool
    pool.Create 4
    
    ' 准备下载信息
    Dim info As DownloadInfo
    info.Url = "https://example.com/file.txt"
    info.OutputFile = "C:\download\file.txt"
    
    ' 创建下载任务
    Dim task As cThread
    Set task = pool.AddTask(AddressOf DownloadProc, info)
    
    ' 设置超时和重试
    task.SetTimeout 30000        ' 30秒超时
    task.SetRetryPolicy 3, 1000  ' 最多重试3次
    
    ' 等待完成
    If task.WaitForCompletion Then
        If task.Result = 0 Then
            Debug.Print "下载成功"
        Else
            Debug.Print "下载失败：" & task.Result
        End If
    Else
        Debug.Print "下载超时"
    End If
End Sub
```

## 高级示例

### 1. 任务依赖链

```vb
' 依赖任务示例
Public Sub DependencyDemo()
    Dim pool As New cThreadPool
    pool.Create 4
    
    ' 创建任务链
    Dim task1 As cThread, task2 As cThread, task3 As cThread
    
    ' 第一个任务：准备数据
    Set task1 = pool.AddTask(AddressOf PrepareData)
    task1.Tag = "准备数据"
    
    ' 第二个任务：处理数据
    Set task2 = pool.AddTask(AddressOf ProcessData)
    task2.Tag = "处理数据"
    task2.AddDependency task1  ' 依赖任务1
    
    ' 第三个任务：保存结果
    Set task3 = pool.AddTask(AddressOf SaveResult)
    task3.Tag = "保存结果"
    task3.AddDependency task2  ' 依赖任务2
    
    ' 等待整个链完成
    pool.WaitForAll
End Sub
```

### 2. 自动扩展线程池

```vb
' 压力测试任务
Public Function StressTask(ByVal param As LongPtr) As Long
    ' 模拟不同负载的任务
    Dim duration As Long
    duration = Int(Rnd * 5000) + 1000  ' 1-6秒随机持续时间
    Sleep duration
    StressTask = 0
End Function

' 自动扩展示例
Public Sub AutoScaleDemo()
    Dim pool As New cThreadPool
    pool.Create 4  ' 初始4个线程
    
    ' 配置自动扩展
    pool.SetAutoScale True, _
        minThreads:=2, _
        maxThreads:=8, _
        loadFactor:=0.75
    
    ' 启用性能日志
    pool.EnablePerformanceLog "C:\poolstress.log"
    
    ' 添加大量任务
    Dim i As Long
    For i = 1 To 50
        pool.AddTask AddressOf StressTask, "压力测试" & i
        Sleep 100  ' 间隔添加任务
    Next i
    
    ' 等待所有任务完成
    pool.WaitForAll
    
    ' 输出统计信息
    Debug.Print "平均执行时间：" & pool.AverageTaskTime & "ms"
    Debug.Print "最大线程数：" & pool.PeakThreadCount
End Sub
```

### 3. 复杂错误处理

```vb
' 错误处理示例
Public Sub ErrorHandlingDemo()
    Dim pool As New cThreadPool
    pool.Create 4
    
    ' 设置全局错误处理
    pool.SetErrorHandler AddressOf GlobalErrorHandler
    
    ' 添加可能失败的任务
    Dim task As cThread
    Set task = pool.AddTask(AddressOf RiskyTask)
    
    ' 设置重试策略
    task.SetRetryPolicy 3, 1000
    
    ' 设置自定义错误处理
    task.SetOnError AddressOf CustomErrorHandler
    
    ' 等待完成
    task.WaitForCompletion
    
    ' 检查错误状态
    If task.LastError <> 0 Then
        Debug.Print "任务失败："
        Debug.Print "错误代码：" & task.LastError
        Debug.Print "错误描述：" & task.ErrorDescription
        Debug.Print "重试次数：" & task.RetryCount
    End If
End Sub

' 全局错误处理
Public Function GlobalErrorHandler(ByVal pool As cThreadPool, _
                                 ByVal errorCode As Long) As Long
    ' 记录错误
    Debug.Print "线程池错误：" & errorCode
    GlobalErrorHandler = True  ' 继续执行
End Function

' 自定义错误处理
Public Function CustomErrorHandler(ByVal taskPtr As LongPtr, _
                                 ByVal errorCode As Long, _
                                 ByVal wParam As LongPtr, _
                                 ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(taskPtr)
    
    Select Case errorCode
        Case 1001  ' 自定义错误码
            ' 特殊处理
            CustomErrorHandler = 1  ' 请求重试
        Case Else
            ' 其他错误
            CustomErrorHandler = 0  ' 不再重试
    End Select
End Function
```

### 4. UI更新示例

```vb
' UI更新示例
Public Sub UiUpdateDemo()
    Dim pool As New cThreadPool
    pool.Create 4
    
    ' 创建进度更新任务
    Dim task As cThread
    Set task = pool.AddTask(AddressOf LongProcess)
    
    ' 设置进度回调
    task.SetProgressCallback AddressOf UpdateProgress
    
    ' 等待完成
    task.WaitForCompletion
End Sub

' 进度更新回调
Public Function UpdateProgress(ByVal taskPtr As LongPtr, _
                             ByVal progress As Long, _
                             ByVal wParam As LongPtr, _
                             ByVal lParam As LongPtr) As Long
    ' 注意：这是在工作线程中调用的
    ' 需要使用PostMessage来更新UI
    PostMessage FormHandle, WM_UPDATE_PROGRESS, progress, 0
    UpdateProgress = 0
End Function

' 窗体代码
Private Sub Form_Message(ByVal msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long)
    If msg = WM_UPDATE_PROGRESS Then
        ' 更新进度条
        ProgressBar1.Value = wParam
    End If
End Sub
```
