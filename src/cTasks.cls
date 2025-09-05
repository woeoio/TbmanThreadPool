Class cTasks

    '@Module
    '@Folder("Threading")
    '@Description "线程池管理器"
    Private m_Tasks As Collection
    Private m_MaxThreads As Long
    Private m_ActiveThreads As Long
    Private m_LastError As Long
    Private m_ErrorDescription As String
    Private m_CompletedTasks As Long
    Private m_TotalExecutionTime As Double
    Private m_State As ePoolState
    
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
    Public Function Create(Optional ByVal MaxThreads As Long = 0) As cTasks
        If MaxThreads < 1 Then MaxThreads = 1
        Initialize MaxThreads
        Set Create = Me
    End Function

    '@Description "初始化线程池"
    Friend Sub Initialize(Optional ByVal MaxThreads As Long = 0)
        Set m_Tasks = New Collection
        m_MaxThreads = MaxThreads
        m_ActiveThreads = 0
    End Sub

    '@Description "添加新任务到线程池"
    Public Function AddTask(ByVal TaskProcedure As LongPtr, Optional ByVal Tag As Variant, Optional ByVal StartImmediately As Boolean = True) As cTask
        ' 检查是否达到最大线程数
        If m_MaxThreads > 0 And m_ActiveThreads >= m_MaxThreads Then
            Err.Raise vbObjectError + 2000, "cTasks.AddTask", "Maximum number of threads (" & m_MaxThreads & ") reached"
        End If
    
        Dim newTask As New cTask
        newTask.Create(TaskProcedure, Tag, StartImmediately)
    
        m_Tasks.Add newTask
        If newTask.IsRunning Then
            m_ActiveThreads = m_ActiveThreads + 1
        End If
    
        Set AddTask = newTask
    End Function

    '@Description "移除并终止任务"
    Public Sub RemoveTask(ByVal Index As Variant)
        Dim task As cTask
    
        ' 支持按索引或对象引用移除
        If VarType(Index) = vbObject Then
            Set task = Index
            For i As Long = 1 To m_Tasks.Count
                If m_Tasks(i) Is task Then
                    If task.IsRunning Then
                        task.Terminate
                        m_ActiveThreads = m_ActiveThreads - 1
                    End If
                    m_Tasks.Remove i
                    Exit For
                End If
            Next
        Else
            Set task = m_Tasks(Index)
            If task.IsRunning Then
                task.Terminate
                m_ActiveThreads = m_ActiveThreads - 1
            End If
            m_Tasks.Remove Index
        End If
    End Sub

    '@Description "暂停所有运行中的任务"
    Public Sub PauseAll()
        Dim task As cTask
        For Each task In m_Tasks
            If task.IsRunning And Not task.IsPaused Then
                task.Pause
            End If
        Next
    End Sub

    '@Description "恢复所有暂停的任务"
    Public Sub ResumeAll()
        Dim task As cTask
        For Each task In m_Tasks
            If task.IsRunning And task.IsPaused Then
                task.Resume
            End If
        Next
    End Sub

    '@Description "请求取消所有任务"
    Public Sub RequestCancelAll()
        Dim task As cTask
        For Each task In m_Tasks
            If task.IsRunning Then
                task.RequestCancel
            End If
        Next
    End Sub

    '@Description "终止所有任务"
    Public Sub TerminateAll()
        Dim task As cTask
        For Each task In m_Tasks
            If task.IsRunning Then
                task.Terminate
                m_ActiveThreads = m_ActiveThreads - 1
            End If
        Next
    End Sub

    '@Description "等待所有任务完成"
    Public Sub WaitForAll(Optional ByVal TimeoutMs As Long = INFINITE)
        Dim handles() As LongPtr
        Dim count As Long
        Dim task As cTask
    
        ' 收集所有运行中任务的句柄
        count = 0
        For Each task In m_Tasks
            If task.IsRunning Then
                ReDim Preserve handles(0 To count)
                handles(count) = task.SafeThreadHandle
                count = count + 1
            End If
        Next
    
        If count > 0 Then
            ' 使用WaitForMultipleObjects等待所有线程
            Dim result As Long
            result = WaitForMultipleObjects(count, handles(0), True, TimeoutMs)
        
            If result = WAIT_FAILED Then
                Err.Raise vbObjectError + 2001, "cTasks.WaitForAll", "WaitForMultipleObjects failed with error: " & Err.LastDllError
            End If
        
            ' 更新活动线程计数
            m_ActiveThreads = m_ActiveThreads - count
        End If
    End Sub

    '@Description "获取任务数量"
    Public Property Get Count() As Long
        Count = m_Tasks.Count
    End Property

    '@Description "获取活动线程数"
    Public Property Get ActiveThreads() As Long
        ActiveThreads = m_ActiveThreads
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
    Public Property Get Task(ByVal Index As Long) As cTask
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
    
    '@Description "优雅关闭线程池"
    Public Sub Shutdown(Optional ByVal TimeoutMs As Long = 5000)
        m_State = PoolState_Stopping
        
        ' 首先尝试取消所有任务
        RequestCancelAll
        
        ' 等待指定时间
        WaitForAll TimeoutMs
        
        ' 如果还有任务在运行，强制终止
        If m_ActiveThreads > 0 Then
            TerminateAll
        End If
        
        m_State = PoolState_Stopped
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
    
    Private Sub UpdateTaskStats(ByVal task As cTask)
        m_CompletedTasks = m_CompletedTasks + 1
        m_TotalExecutionTime = m_TotalExecutionTime + task.ExecutionTime
    End Sub
    
    Private Sub SetError(ByVal ErrorNumber As Long, ByVal Description As String)
        m_LastError = ErrorNumber
        m_ErrorDescription = Description
    End Sub

End Class