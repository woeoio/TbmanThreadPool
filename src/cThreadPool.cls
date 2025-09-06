Class cThreadPool

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Document: https://doc.twinbasic.vb6.pro/en/tbman/threadPool/

    '@Module
    '@Folder("Threading")
    '@Description "线程池管理器"
    
    ' API 声明
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    
    Private m_Tasks As Collection            ' 所有任务的集合
    Private m_PendingTasks As Collection     ' 待执行的任务队列
    Private m_MaxThreads As Long             ' 最大线程数
    Private m_IsTerminating As Boolean       ' 是否正在终止
    Private m_MinThreads As Long             ' 最小线程数
    Private m_ActiveThreads As Long          ' 活动线程数
    Private m_LastError As Long              ' 最后错误码
    Private m_ErrorDescription As String     ' 最后错误描述
    Private m_CompletedTasks As Long         ' 已完成任务数
    Private m_TotalExecutionTime As Double   ' 总执行时间
    Private m_State As ePoolState            ' 线程池状态
    Private m_LogFile As String              ' 日志文件路径
    Private m_LoadFactor As Double           ' 负载因子
    Private m_AutoScale As Boolean           ' 是否启用自动扩展
    Private m_UINotifier As cThreadUINotifier      ' UI通知器
    
    Public Enum ePoolState
        PoolState_Running = 0
        PoolState_Stopping = 1
        PoolState_Stopped = 2
    End Enum
    Public Enum ePoolWaitBehavior
        PoolWait_ReturnOnTimeout = 0
        PoolWait_TerminateAllOnTimeout = 1
        PoolWait_CancelAllOnTimeout = 2
    End Enum
    
    '@Description "创建新线程池"
    Public Function Create(Optional ByVal MaxThreads As Long = 0) As cThreadPool
        If MaxThreads < 1 Then MaxThreads = 1
        Initialize MaxThreads
        Set Create = Me
    End Function

    '@Description "初始化线程池"
    Friend Sub Initialize(Optional ByVal MaxThreads As Long = 0)
        Set m_Tasks = New Collection
        Set m_PendingTasks = New Collection
        m_MaxThreads = MaxThreads
        m_MinThreads = 1
        m_ActiveThreads = 0
        m_CompletedTasks = 0
        m_TotalExecutionTime = 0
        m_State = PoolState_Running
        m_LoadFactor = 0.75
        m_AutoScale = False
        m_LogFile = ""
    End Sub

    '@Description "添加新任务到线程池"
    Public Function AddTask(ByVal TaskProcedure As LongPtr, _
                          Optional ByVal Tag As Variant, _
                          Optional ByVal StartImmediately As Boolean = True, _
                          Optional ByVal Priority As eTaskPriority = Priority_Normal) As cThread
        
        ' 首先进行自动清理
        AutoCleanupCompleted
        
        ' 创建新任务
        Dim newTask As New cThread
        newTask.Create TaskProcedure, Tag, False ' 先不启动
        newTask.Priority = Priority
        
        ' 如果启用了日志，为新任务也启用日志
        If m_LogFile <> "" Then
            newTask.EnableLogging m_LogFile
        End If
        
        ' 如果有UI通知器，为新任务设置
        If Not (m_UINotifier Is Nothing) Then
            newTask.SetUINotifier m_UINotifier
        End If
        
        ' 添加到任务集合
        m_Tasks.Add newTask
        
        WriteLog "New task added with priority " & Priority
        
        If StartImmediately Then
            ' 获取当前活动线程数（实时计算）
            Dim currentActiveThreads As Long
            currentActiveThreads = Me.ActiveThreads
            
            ' 如果有空闲线程，直接启动
            If m_MaxThreads = 0 Or currentActiveThreads < m_MaxThreads Then
                newTask.Start
                If newTask.IsRunning Then
                    WriteLog "Task started immediately"
                End If
            Else
                ' 否则加入待执行队列
                AddToPendingQueue newTask
                WriteLog "Task added to pending queue"
            End If
        Else
            ' 加入待执行队列
            AddToPendingQueue newTask
            WriteLog "Task added to pending queue (start deferred)"
        End If
        
        ' 调整线程池大小
        AdjustPoolSize
        
        ' 处理待执行的任务
        ProcessPendingTasks
        
        Set AddTask = newTask
    End Function

    '@Description "移除并终止任务"
    Public Sub RemoveTask(ByVal Index As Variant)
        Dim task As cThread
    
        ' 支持按索引或对象引用移除
        If VarType(Index) = vbObject Then
            Set task = Index
            For i As Long = 1 To m_Tasks.Count
                If m_Tasks(i) Is task Then
                    If task.IsRunning Then
                        task.Terminate
                        WriteLog "Task terminated and removed from pool"
                    End If
                    m_Tasks.Remove i
                    Exit For
                End If
            Next
        Else
            Set task = m_Tasks(Index)
            If task.IsRunning Then
                task.Terminate
                WriteLog "Task terminated and removed from pool"
            End If
            m_Tasks.Remove Index
        End If
        
        ' 处理待执行的任务（移除任务后可能有空闲线程）
        ProcessPendingTasks
    End Sub

    '@Description "暂停所有运行中的任务"
    Public Sub PauseAll()
        Dim task As cThread
        For Each task In m_Tasks
            If task.IsRunning And Not task.IsPaused Then
                task.Pause
            End If
        Next
    End Sub

    '@Description "恢复所有暂停的任务"
    Public Sub ResumeAll()
        Dim task As cThread
        For Each task In m_Tasks
            If task.IsRunning And task.IsPaused Then
                task.Resume
            End If
        Next
    End Sub

    '@Description "请求取消所有任务"
    Public Sub RequestCancelAll()
        Dim task As cThread
        For Each task In m_Tasks
            If task.IsRunning Then
                task.RequestCancel
            End If
        Next
    End Sub

    '@Description "终止所有任务"
    Public Sub TerminateAll()
        If m_IsTerminating Then Exit Sub
        m_IsTerminating = True
        
        WriteLog "Terminating all tasks..."
        
        ' 终止所有运行中的任务
        Dim task As cThread
        For Each task In m_Tasks
            If task.IsRunning Then
                task.Terminate
            End If
        Next
        
        ' 清空集合
        Set m_Tasks = New Collection
        Set m_PendingTasks = New Collection
        m_State = PoolState_Stopped
        m_ActiveThreads = 0  ' 重置计数器
        
        WriteLog "All tasks terminated and pool stopped"
        m_IsTerminating = False
    End Sub
    
    Private Sub Class_Terminate()
        TerminateAll
    End Sub

    '@Description "设置UI通知器"
    Public Sub SetUINotifier(ByRef notifier As cThreadUINotifier)
        Set m_UINotifier = notifier
        
        ' 为所有现有任务设置UI通知器
        Dim task As cThread
        For Each task In m_Tasks
            task.SetUINotifier notifier
        Next
    End Sub
    
    '@Description "获取UI通知器"
    Public Property Get UINotifier() As cThreadUINotifier
        Set UINotifier = m_UINotifier
    End Property

    '@Description "等待所有任务完成"
    Public Sub WaitForAll(Optional ByVal TimeoutMs As Long = INFINITE)
        Dim handles() As LongPtr
        Dim count As Long
        Dim task As cThread
        Dim startTime As Currency
        
        QueryPerformanceCounter startTime
        WriteLog "WaitForAll started, timeout=" & TimeoutMs & "ms"
    
        ' 收集所有运行中任务的句柄
        count = 0
        For Each task In m_Tasks
            If task.IsRunning Then
                ReDim Preserve handles(0 To count)
                handles(count) = task.SafeThreadHandle
                If handles(count) <> 0 Then
                    count = count + 1
                    WriteLog "Added running task to wait list, handle=" & handles(count - 1)
                End If
            End If
        Next
        
        WriteLog "Total running tasks to wait for: " & count
    
        If count > 0 Then
            ' 使用WaitForMultipleObjects等待所有线程
            Dim result As Long
            result = WaitForMultipleObjects(count, handles(0), True, TimeoutMs)
            
            Dim endTime As Currency
            QueryPerformanceCounter endTime
            Dim freq As Currency
            QueryPerformanceFrequency freq
            Dim elapsed As Double
            elapsed = (endTime - startTime) / freq * 1000
            
            WriteLog "WaitForMultipleObjects result=" & result & ", elapsed=" & Format$(elapsed, "0.0") & "ms"
        
            Select Case result
                Case WAIT_OBJECT_0 To WAIT_OBJECT_0 + count - 1
                    WriteLog "All tasks completed successfully"
                Case WAIT_TIMEOUT
                    WriteLog "WaitForAll timed out after " & TimeoutMs & "ms"
                    ' 检查哪些任务还在运行
                    CheckStillRunningTasks
                Case WAIT_FAILED
                    WriteLog "WaitForMultipleObjects failed with error: " & Err.LastDllError
                    Err.Raise vbObjectError + 2001, "cThreadPool.WaitForAll", "WaitForMultipleObjects failed with error: " & Err.LastDllError
            End Select
        
            ' 更新活动线程计数
            UpdateActiveThreadCount
        Else
            WriteLog "No running tasks to wait for"
        End If
    End Sub
    
    Private Sub CheckStillRunningTasks()
        Dim task As cThread
        Dim runningCount As Long
        runningCount = 0
        
        For Each task In m_Tasks
            If task.IsRunning Then
                runningCount = runningCount + 1
                WriteLog "Task still running: ThreadID=" & task.ThreadID & ", Handle=" & task.ThreadHandle
            End If
        Next
        
        WriteLog "Total tasks still running: " & runningCount
    End Sub
    
    Private Sub UpdateActiveThreadCount()
        ' 由于ActiveThreads属性现在实时计算，这个方法主要用于日志记录
        Dim actualRunning As Long
        actualRunning = Me.ActiveThreads  ' 调用实时计算的属性
        WriteLog "Active thread count updated to: " & actualRunning
    End Sub

    '@Description "获取任务数量"
    Public Property Get Count() As Long
        ' 在获取计数前进行自动清理
        AutoCleanupCompleted
        Count = m_Tasks.Count
    End Property

    '@Description "获取活动线程数"
    Public Property Get ActiveThreads() As Long
        ' 使用精确的运行任务计数方法
        Dim count As Long
        count = GetActualRunningTaskCount()
        
        ' 同时更新缓存的计数器
        m_ActiveThreads = count
        ActiveThreads = count
    End Property

    '@Description "获取最大线程数限制"
    Public Property Get MaxThreads() As Long
        MaxThreads = m_MaxThreads
    End Property

    '@Description "设置最大线程数限制"
    Public Property Let MaxThreads(ByVal Value As Long)
        m_MaxThreads = Value
    End Property

    '@Description "按索引获取任务"
    Public Property Get Task(ByVal Index As Long) As cThread
        Set Task = m_Tasks(Index)
    End Property

    '@Description "清理已完成的任务"
    Public Sub CleanupCompleted()
        Dim i As Long
        i = 1
        Do While i <= m_Tasks.Count
            If Not m_Tasks(i).IsRunning Then
                UpdateTaskStats m_Tasks(i)
                m_Tasks.Remove i
            Else
                i = i + 1
            End If
        Loop
    End Sub
    
    '@Description "自动清理已完成的任务（内部调用）"
    Private Sub AutoCleanupCompleted()
        ' 自动清理，但不会在每次调用时都执行
        Static lastCleanupTime As Double
        Dim currentTime As Double
        currentTime = Timer
        
        ' 每秒最多清理一次，避免频繁操作
        If currentTime - lastCleanupTime > 1 Then
            CleanupCompleted
            lastCleanupTime = currentTime
            WriteLog "Auto cleanup completed, remaining tasks: " & m_Tasks.Count
        End If
    End Sub
    
    '@Description "优雅关闭线程池"
    Public Sub Shutdown(Optional ByVal TimeoutMs As Long = 5000)
        m_State = PoolState_Stopping
        WriteLog "Starting graceful shutdown..."
        
        ' 首先尝试取消所有任务
        RequestCancelAll
        
        ' 等待指定时间
        WaitForAll TimeoutMs
        
        ' 检查是否还有活动线程
        Dim currentActiveThreads As Long
        currentActiveThreads = Me.ActiveThreads
        
        ' 如果还有任务在运行，强制终止
        If currentActiveThreads > 0 Then
            WriteLog "Force terminating " & currentActiveThreads & " remaining tasks"
            TerminateAll
        End If
        
        m_State = PoolState_Stopped
        WriteLog "Thread pool shutdown completed"
    End Sub
    
    '@Description "获取线程池状态"
    Public Property Get State() As ePoolState
        State = m_State
    End Property
    
    '@Description "获取最后一次错误信息"
    Public Property Get LastError() As String
        LastError = m_ErrorDescription
    End Property
    
    '@Description "获取已完成的任务数量"
    Public Property Get CompletedTasks() As Long
        CompletedTasks = m_CompletedTasks
    End Property
    
    '@Description "获取平均任务执行时间(毫秒)"
    Public Property Get AverageExecutionTime() As Double
        If m_CompletedTasks > 0 Then
            AverageExecutionTime = m_TotalExecutionTime / m_CompletedTasks
        End If
    End Property
    
    '@Description "获取运行中的任务数量（精确检查线程状态）"
    Public Property Get RunningTasks() As Long
        RunningTasks = GetActualRunningTaskCount()
    End Property
    
    '@Description "精确计算运行中的任务数量（私有方法）"
    Private Function GetActualRunningTaskCount() As Long
        ' 先进行自动清理以移除已完成的任务
        AutoCleanupCompleted
        
        Dim count As Long
        Dim task As cThread
        count = 0
        
        For Each task In m_Tasks
            ' 触发智能状态更新
            task.UpdateWorkState
            
            ' 检查任务是否真正在忙碌状态工作
            If task.IsRunning And task.IsBusy Then
                ' 可以添加额外的验证，比如检查线程句柄是否有效
                Dim threadHandle As LongPtr
                threadHandle = task.SafeThreadHandle
                If threadHandle <> 0 Then
                    count = count + 1
                End If
            End If
        Next
        
        WriteLog "Actual working task count: " & count
        GetActualRunningTaskCount = count
    End Function
    
    '@Description "获取未完成的任务数量（运行中+等待中）"
    Public Property Get PendingTasks() As Long
        ' 使用精确的运行中任务数量
        PendingTasks = Me.RunningTasks + m_PendingTasks.Count
    End Property
    
    '@Description "获取等待执行的任务数量"
    Public Property Get QueuedTasks() As Long
        QueuedTasks = m_PendingTasks.Count
    End Property
    
    '@Description "获取总任务数量（运行中+等待中+已在集合中但未清理的已完成任务）"
    Public Property Get TotalTasks() As Long
        ' 自动清理以获得准确计数
        AutoCleanupCompleted
        TotalTasks = m_Tasks.Count + m_PendingTasks.Count
    End Property
    
    '@Description "检查是否所有任务都已完成"
    Public Property Get IsAllTasksCompleted() As Boolean
        ' 使用精确的运行中任务数量检查
        IsAllTasksCompleted = (Me.RunningTasks = 0 And m_PendingTasks.Count = 0)
    End Property
    
    '@Description "获取任务完成百分比"
    Public Property Get CompletionPercentage() As Double
        ' 使用精确的运行中任务数量
        Dim currentPending As Long
        currentPending = Me.PendingTasks
        
        Dim totalEverAdded As Long
        totalEverAdded = m_CompletedTasks + currentPending
        
        If totalEverAdded > 0 Then
            CompletionPercentage = (m_CompletedTasks * 100.0) / totalEverAdded
        Else
            CompletionPercentage = 100.0  ' 没有任务时视为100%完成
        End If
    End Property
    
    Private Sub UpdateTaskStats(ByVal task As cThread)
        m_CompletedTasks = m_CompletedTasks + 1
        m_TotalExecutionTime = m_TotalExecutionTime + task.ExecutionTime
    End Sub
    
    Private Sub SetError(ByVal ErrorNumber As Long, ByVal Description As String)
        m_LastError = ErrorNumber
        m_ErrorDescription = Description
        WriteLog "Error: " & Description
    End Sub
    
    '@Description "启用日志记录"
    Public Sub EnableLogging(ByVal LogFilePath As String)
        m_LogFile = LogFilePath
        WriteLog "Thread pool logging enabled"
        
        ' 为所有任务启用日志
        Dim task As cThread
        For Each task In m_Tasks
            task.EnableLogging LogFilePath
        Next
    End Sub
    
    '@Description "写入日志"
    Private Sub WriteLog(ByVal Message As String)
        If m_LogFile = "" Then Exit Sub
        
        Dim fNum As Integer
        fNum = FreeFile
        
        On Error Resume Next
        Open m_LogFile For Append As #fNum
        Print #fNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [Pool] " & Message
        Close #fNum
    End Sub
    
    '@Description "设置线程池自动扩展"
    Public Sub SetAutoScale(ByVal Enable As Boolean, Optional ByVal MinThreads As Long = 1, Optional ByVal LoadFactor As Double = 0.75)
        m_AutoScale = Enable
        m_MinThreads = MinThreads
        m_LoadFactor = LoadFactor
        WriteLog "AutoScale " & IIf(Enable, "enabled", "disabled") & " (MinThreads=" & MinThreads & ", LoadFactor=" & LoadFactor & ")"
    End Sub
    
    '@Description "调整线程池大小"
    Private Sub AdjustPoolSize()
        If Not m_AutoScale Then Exit Sub
        
        ' 获取当前活动线程数（实时计算）
        Dim currentActiveThreads As Long
        currentActiveThreads = Me.ActiveThreads
        
        ' 计算当前负载
        Dim currentLoad As Double
        If m_MaxThreads > 0 Then
            currentLoad = currentActiveThreads / m_MaxThreads
        Else
            currentLoad = 0
        End If
        
        ' 如果负载过高，增加线程数
        If currentLoad > m_LoadFactor Then
            Dim newMaxThreads As Long
            newMaxThreads = m_MaxThreads * 2
            WriteLog "Expanding pool size from " & m_MaxThreads & " to " & newMaxThreads & " threads (load=" & Format$(currentLoad * 100, "0.0") & "%)"
            m_MaxThreads = newMaxThreads
        ' 如果负载过低，减少线程数（但不低于最小线程数）
        ElseIf currentLoad < (m_LoadFactor / 2) And m_MaxThreads > m_MinThreads Then
            newMaxThreads = Max(m_MinThreads, m_MaxThreads \ 2)
            WriteLog "Shrinking pool size from " & m_MaxThreads & " to " & newMaxThreads & " threads (load=" & Format$(currentLoad * 100, "0.0") & "%)"
            m_MaxThreads = newMaxThreads
        End If
    End Sub
    
    '@Description "按优先级添加任务"
    Private Sub AddToPendingQueue(ByVal Task As cThread)
        Dim i As Long
        Dim currentTask As cThread
        
        ' 按优先级插入队列
        For i = 1 To m_PendingTasks.Count
            Set currentTask = m_PendingTasks(i)
            If Task.Priority > currentTask.Priority Then
                m_PendingTasks.Add Task, , i
                Exit Sub
            End If
        Next
        
        ' 如果没有找到合适的位置，添加到末尾
        m_PendingTasks.Add Task
    End Sub
    
    '@Description "处理待执行的任务"
    Private Sub ProcessPendingTasks()
        If m_PendingTasks.Count = 0 Then Exit Sub
        
        ' 获取当前活动线程数（实时计算）
        Dim currentActiveThreads As Long
        currentActiveThreads = Me.ActiveThreads
        
        ' 检查是否可以执行更多任务
        If m_MaxThreads > 0 And currentActiveThreads >= m_MaxThreads Then Exit Sub
        
        ' 获取下一个要执行的任务
        Dim nextTask As cThread
        Set nextTask = m_PendingTasks(1)
        m_PendingTasks.Remove 1
        
        ' 启动任务
        nextTask.Start
        If nextTask.IsRunning Then
            WriteLog "Pending task started, TaskId: " & nextTask.TaskId
        End If
        
        ' 调整线程池大小
        AdjustPoolSize
        
        ' 递归处理更多待执行任务
        ProcessPendingTasks
    End Sub
    
    '@Description "获取线程池统计信息"
    Public Function GetPoolStats() As String
        ' 先进行自动清理以获得准确统计
        AutoCleanupCompleted
        
        ' 获取详细统计
        Dim runningTasks As Long
        Dim queuedTasks As Long
        Dim totalTasks As Long
        Dim pendingTasks As Long
        Dim isCompleted As Boolean
        
        runningTasks = Me.RunningTasks
        queuedTasks = Me.QueuedTasks
        totalTasks = Me.TotalTasks
        pendingTasks = Me.PendingTasks
        isCompleted = Me.IsAllTasksCompleted
        
        Dim stats As String
        stats = "Thread Pool Statistics:" & vbCrLf & _
               "Running Tasks: " & runningTasks & "/" & m_MaxThreads & vbCrLf & _
               "Queued Tasks: " & queuedTasks & vbCrLf & _
               "Total Active Tasks: " & totalTasks & vbCrLf & _
               "Pending Tasks: " & pendingTasks & vbCrLf & _
               "Completed Tasks: " & Me.CompletedTasks & vbCrLf & _
               "Tasks in Collection: " & m_Tasks.Count & vbCrLf & _
               "Completion %: " & Format$(Me.CompletionPercentage, "0.0") & "%" & vbCrLf & _
               "All Completed: " & IIf(isCompleted, "Yes", "No") & vbCrLf & _
               "Average Execution Time: " & Format$(AverageExecutionTime, "0.00") & "ms" & vbCrLf & _
               "Pool State: " & IIf(m_State = PoolState_Running, "Running", IIf(m_State = PoolState_Stopping, "Stopping", "Stopped")) & vbCrLf & _
               "Auto-Scale: " & IIf(m_AutoScale, "Enabled", "Disabled")
        GetPoolStats = stats
    End Function
    
    Private Function Max(ByVal a As Long, ByVal b As Long) As Long
        If a > b Then
            Max = a
        Else
            Max = b
        End If
    End Function

End Class