Module mWorkStateExample
    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Document: https://doc.twinbasic.vb6.pro/en/tbman/threadPool/

    ' 工作状态示例模块
    ' 演示如何在任务执行过程中正确设置线程工作状态

    ' 示例任务数据结构
    Private Type TaskWorkData
        TaskId As Long
        WorkDuration As Long ' 工作持续时间（毫秒）
        Data As String
    End Type

    ' 示例任务过程 - 正确使用工作状态（自动管理版本）
    Private Function ExampleWorkProc(ByVal param As LongPtr) As Long
        ' 获取线程对象
        Dim task As cThread
        Set task = mThread.ReturnFromPtr(param)
        
        ' *** 重要提示：工作状态现在自动管理！***
        ' - 任务开始时会自动设置为 ThreadWork_Busy
        ' - 正常完成时会自动设置为 ThreadWork_Completed
        ' - 发生错误时会自动设置为 ThreadWork_Error
        
        ' 获取任务数据
        Dim workData As TaskWorkData
        CopyMemory workData, ByVal task.Tag, LenB(task.Tag)
        
        ' 检查任务是否被取消
        If task.CancelRequested Then
            ExampleWorkProc = 1
            Exit Function
        End If
        
        ' 模拟工作过程
        Dim startTime As Currency
        Dim currentTime As Currency
        QueryPerformanceCounter startTime
        
        Do
            ' 模拟一些工作
            Sleep 100  ' 模拟工作负载
            
            ' 检查是否被取消
            If task.CancelRequested Then
                ExampleWorkProc = 2
                Exit Function
            End If
            
            ' 检查工作时间
            QueryPerformanceCounter currentTime
            Dim elapsedMs As Long
            elapsedMs = (currentTime - startTime) * 1000 / task.PerformanceFrequency
            
        Loop While elapsedMs < workData.WorkDuration
        
        ' 任务完成（状态会自动设置为 ThreadWork_Completed）
        ExampleWorkProc = 0
    End Function
    
    ' 错误处理示例任务过程（自动管理版本）
    Private Function ExampleErrorProc(ByVal param As LongPtr) As Long
        ' 获取线程对象
        Dim task As cThread
        Set task = mThread.ReturnFromPtr(param)
        
        ' *** 重要提示：工作状态现在自动管理！***
        ' - 发生错误时会自动设置为 ThreadWork_Error
        
        ' 模拟一个会发生错误的任务
        On Error GoTo ErrorHandler
        
        ' 模拟一些工作
        Sleep 500
        
        ' 模拟错误
        Err.Raise vbObjectError + 1000, "ExampleErrorProc", "模拟的任务错误"
        
        ' 正常完成（这行不会执行到）
        ExampleErrorProc = 0
        Exit Function
        
    ErrorHandler:
        ' 错误状态会自动设置，只需要返回错误代码
        ExampleErrorProc = -1
    End Function
    
    ' 演示如何使用工作状态监控
    Public Sub DemoWorkStateMonitoring()
        ' 创建线程池
        Dim pool As New cThreadPool
        pool.MaxThreads = 3
        
        ' 添加正常工作的任务
        Dim i As Long
        For i = 1 To 5
            Dim workData As TaskWorkData
            workData.TaskId = i
            workData.WorkDuration = 2000 + (i * 500) ' 不同的工作时间
            workData.Data = "Task " & i
            
            pool.AddTask AddressOf ExampleWorkProc, VarPtr(workData), , True
        Next
        
        ' 添加一个会出错的任务
        pool.AddTask AddressOf ExampleErrorProc, 0, , True
        
        ' 监控工作状态
        Debug.Print "开始监控工作状态..."
        
        Do While Not pool.IsAllTasksCompleted
            Debug.Print String(50, "-")
            Debug.Print "总任务数: " & pool.TotalTasks
            Debug.Print "运行中任务: " & pool.RunningTasks & " (实际工作的线程)"
            Debug.Print "等待中任务: " & pool.QueuedTasks
            Debug.Print "完成进度: " & Format(pool.CompletionPercentage, "0.0") & "%"
            
            ' 显示每个线程的详细状态
            Dim j As Long
            Dim thread As cThread
            j = 1
            For Each thread In pool.Tasks
                Debug.Print "  线程 " & j & ": " & _
                           "运行=" & thread.IsRunning & _
                           ", 忙碌=" & thread.IsBusy & _
                           ", 空闲=" & thread.IsIdle & _
                           ", 状态=" & thread.GetWorkStateText()
                j = j + 1
            Next
            
            Sleep 1000 ' 每秒更新一次
        Loop
        
        Debug.Print "所有任务已完成!"
        
        ' 清理
        Set pool = Nothing
    End Sub

End Module
