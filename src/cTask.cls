Class cTask

    '@Module
    '@Folder("Threading")
    '@Description "代表一个可管理的线程任务"

    ' API 声明
    Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( _
        ByVal lpMutexAttributes As LongPtr, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String _
    ) As LongPtr

    Private Declare Function ReleaseMutex Lib "kernel32" ( _
        ByVal hMutex As LongPtr _
    ) As Long

    Private Declare Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As LongPtr, _
        ByVal dwMilliseconds As Long _
    ) As Long

    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr _
    ) As Long

    Private Declare Function GetExitCodeThread Lib "kernel32" ( _
        ByVal hThread As LongPtr, _
        ByRef lpExitCode As Long _
    ) As Long

    ' 常量
    Private Const STILL_ACTIVE As Long = &H103
    Private Const INFINITE As Long = &HFFFFFFFF
    Private Const WAIT_OBJECT_0 As Long = 0
    Private Const WAIT_TIMEOUT As Long = 258
    Private Const WAIT_FAILED As Long = &HFFFFFFFF

    ' 成员变量
    Private m_ThreadHandle As LongPtr
    Private m_ThreadID As Long
    Private m_IsRunning As Boolean
    Private m_IsPaused As Boolean
    Private m_TaskProcedure As LongPtr
    Private m_CancelRequested As Boolean
    Private m_Tag As Variant
    Private m_LastError As Long
    Private m_ErrorDescription As String
    Private m_StateLock As LongPtr  ' 用于状态同步的互斥体
    Private m_StartTime As Currency  ' 用于记录任务开始时间
    Private m_EndTime As Currency    ' 用于记录任务结束时间
    Private m_Priority As eTaskPriority  ' 任务优先级
    Private m_MaxRetries As Long     ' 最大重试次数
    Private m_RetryCount As Long     ' 当前重试次数
    Private m_RetryDelay As Long     ' 重试间隔(毫秒)
    Private m_TaskTimeout As Long    ' 任务超时时间(毫秒)
    Private m_Dependencies As Collection  ' 任务依赖集合
    Private m_Result As eTaskResult      ' 任务执行结果
    Private m_OnComplete As LongPtr      ' 完成回调函数
    Private m_OnError As LongPtr         ' 错误回调函数
    Private m_LogFile As String          ' 日志文件路径

    ' 在类模块顶部添加（公共枚举）
    Public Enum eTaskWaitBehavior
        ' 默认行为：超时后直接返回
        TaskWait_ReturnOnTimeout = 0
    
        ' 超时后终止线程（强制结束）
        TaskWait_TerminateOnTimeout = 1
    
        ' 超时后请求取消（优雅结束）
        TaskWait_CancelOnTimeout = 2
    
        ' 超时后暂停线程（可恢复）
        TaskWait_PauseOnTimeout = 3
    End Enum
    
    Public Enum eTaskPriority
        Priority_Lowest = 0
        Priority_Low = 1
        Priority_Normal = 2
        Priority_High = 3
        Priority_Highest = 4
    End Enum
    
    ' 任务执行结果枚举
    Public Enum eTaskResult
        TaskResult_Success = 0
        TaskResult_Failed = 1
        TaskResult_Cancelled = 2
        TaskResult_TimedOut = 3
    End Enum
    ' ======================
    ' 初始化与清理
    ' ======================

    Private Sub Class_Initialize()
        m_StateLock = CreateMutex(0, False, 0)
    End Sub

    Private Sub Class_Terminate()
        ' 确保线程终止并清理资源
        If m_IsRunning Then
            Terminate
        End If
        If m_StateLock <> 0 Then
            CloseHandle m_StateLock
            m_StateLock = 0
        End If
    End Sub

    ' ======================
    ' 公共方法
    ' ======================

    '@Description "创建新任务实例"
    Public Function Create(ByVal TaskProcedure As LongPtr, _
                          Optional ByVal Tag As Variant, _
                          Optional ByVal StartImmediately As Boolean = True _
                          ) As cTask
       Initialize TaskProcedure, Tag
        If StartImmediately Then
            Start
        End If
        Set Create = Me
    End Function

    '@Description "初始化任务"
    Friend Sub Initialize(ByVal TaskProcedure As LongPtr, Optional ByVal Tag As Variant)
        LockState
        m_TaskProcedure = TaskProcedure
        m_Tag = Tag
        m_IsRunning = False
        m_IsPaused = False
        m_CancelRequested = False
        m_Priority = Priority_Normal
        m_MaxRetries = 0
        m_RetryCount = 0
        m_RetryDelay = 1000
        m_TaskTimeout = INFINITE
        m_Result = TaskResult_Success
        Set m_Dependencies = New Collection
        m_LogFile = ""
        UnlockState
    End Sub

    '@Description "启动线程"
    Public Sub Start()
        LockState
        On Error GoTo ErrorHandler
    
        If m_IsRunning Then
            UnlockState
            Exit Sub
        End If
        
        QueryPerformanceCounter m_StartTime
    
        m_CancelRequested = False
        m_ThreadHandle = CreateThread( _
            ByVal 0&, _
            ByVal 0&, _
            m_TaskProcedure, _
            ByVal ObjPtr(Me), _
            ByVal 0&, _
            m_ThreadID _
        )
    
        If m_ThreadHandle = 0 Then
            m_LastError = Err.LastDllError
            m_ErrorDescription = "Failed to create thread"
            Err.Raise vbObjectError + 1000, "cTask.Start", "Thread creation failed with error: " & m_LastError
        Else
            m_IsRunning = True
        End If
    
        UnlockState
        Exit Sub
    
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, Err.Source, Err.Description
    End Sub

    ' ======================
    ' 线程控制方法
    ' ======================

    '@Description "暂停线程"
    Public Sub Pause()
        LockState
        On Error GoTo ErrorHandler
    
        If Not m_IsRunning Or m_IsPaused Then
            UnlockState
            Exit Sub
        End If
    
        If SuspendThread(m_ThreadHandle) = -1 Then
            m_LastError = Err.LastDllError
            m_ErrorDescription = "Failed to suspend thread"
            Err.Raise vbObjectError + 1001, "cTask.Pause", "Thread suspension failed with error: " & m_LastError
        Else
            m_IsPaused = True
        End If
    
        UnlockState
        Exit Sub
    
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, Err.Source, Err.Description
    End Sub

    '@Description "恢复线程"
    Public Sub [Resume]()
        LockState
        On Error GoTo ErrorHandler
    
        If Not m_IsRunning Or Not m_IsPaused Then
            UnlockState
            Exit Sub
        End If
    
        If ResumeThread(m_ThreadHandle) = -1 Then
            m_LastError = Err.LastDllError
            m_ErrorDescription = "Failed to resume thread"
            Err.Raise vbObjectError + 1002, "cTask.Resume", "Thread resumption failed with error: " & m_LastError
        Else
            m_IsPaused = False
        End If
    
        UnlockState
        Exit Sub
    
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, Err.Source, Err.Description
    End Sub

    '@Description "请求取消线程"
    Public Sub RequestCancel()
        LockState
        m_CancelRequested = True
        UnlockState
    End Sub

    '@Description "终止线程"
    Public Sub Terminate()
        LockState
        On Error GoTo ErrorHandler
    
        If Not m_IsRunning Then
            UnlockState
            Exit Sub
        End If
    
        If TerminateThread(m_ThreadHandle, 1) = 0 Then
            m_LastError = Err.LastDllError
            m_ErrorDescription = "Failed to terminate thread"
            Err.Raise vbObjectError + 1003, "cTask.Terminate", "Thread termination failed with error: " & m_LastError
        Else
            m_IsRunning = False
            m_IsPaused = False
            CloseHandle m_ThreadHandle
            m_ThreadHandle = 0
        End If
    
        UnlockState
        Exit Sub
    
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, Err.Source, Err.Description
    End Sub

    ' ======================
    ' 状态检查方法
    ' ======================

    '@Description "检查线程是否存活"
    Public Function IsThreadAlive() As Boolean
        If m_ThreadHandle = 0 Then
            IsThreadAlive = False
            Exit Function
        End If
    
        Dim exitCode As Long
        LockState
        If GetExitCodeThread(m_ThreadHandle, exitCode) <> 0 Then
            IsThreadAlive = (exitCode = STILL_ACTIVE)
        Else
            IsThreadAlive = False
        End If
        UnlockState
    End Function

    '@Description "等待线程结束"
    Public Sub WaitForCompletion(Optional ByVal TimeoutMs As Long = INFINITE, _
                                Optional ByVal BehaviorOnTimeout As eTaskWaitBehavior = TaskWait_ReturnOnTimeout)
        If Not m_IsRunning Then Exit Sub
    
        Dim result As Long
        LockState
        result = WaitForSingleObject(m_ThreadHandle, TimeoutMs)
    
        Select Case result
            Case WAIT_OBJECT_0
                ' 正常结束
                m_IsRunning = False
                m_IsPaused = False
                CloseHandle m_ThreadHandle
                m_ThreadHandle = 0
            
            Case WAIT_TIMEOUT
                ' 超时处理
                Select Case BehaviorOnTimeout
                    Case TaskWait_TerminateOnTimeout
                        Terminate
                    Case TaskWait_CancelOnTimeout
                        RequestCancel
                    Case TaskWait_PauseOnTimeout
                        Pause
                End Select
            
            Case WAIT_FAILED
                m_LastError = Err.LastDllError
                m_ErrorDescription = "Wait failed"
                UnlockState
                Err.Raise vbObjectError + 1004, "cTask.WaitForCompletion", "Wait failed with error: " & m_LastError
        End Select
        UnlockState
    End Sub

    ' ======================
    ' 属性访问器
    ' ======================

    '@Description "获取线程是否正在运行"
    Public Property Get IsRunning() As Boolean
        LockState
        IsRunning = m_IsRunning
        UnlockState
    End Property

    '@Description "获取线程是否已暂停"
    Public Property Get IsPaused() As Boolean
        LockState
        IsPaused = m_IsPaused
        UnlockState
    End Property

    '@Description "获取是否已请求取消"
    Public Property Get CancelRequested() As Boolean
        LockState
        CancelRequested = m_CancelRequested
        UnlockState
    End Property

    '@Description "获取线程TAG"
    Public Property Get Tag() As Variant
        LockState
        Tag = m_Tag
        UnlockState
    End Property
    
    '@Description "获取线程ID"
    Public Property Get ThreadID() As LongPtr
        LockState
        ThreadID = m_ThreadID
        UnlockState
    End Property
    
    '@Description "获取线程句柄"
    Public Property Get ThreadHandle() As LongPtr
        LockState
        ThreadHandle = m_ThreadHandle
        UnlockState
    End Property

    '@Description "安全获取线程句柄"
    Public Property Get SafeThreadHandle() As LongPtr
        LockState
        If m_IsRunning Then
            SafeThreadHandle = m_ThreadHandle
        Else
            SafeThreadHandle = 0
        End If
        UnlockState
    End Property

    ' ======================
    ' 私有方法
    ' ======================

    Private Sub LockState()
        WaitForSingleObject m_StateLock, INFINITE
    End Sub

    Private Sub UnlockState()
        ReleaseMutex m_StateLock
    End Sub
    
    '@Description "获取任务执行时间(毫秒)"
    Public Property Get ExecutionTime() As Double
        Dim freq As Currency
        QueryPerformanceFrequency freq
        
        If m_EndTime = 0 Then
            Dim currentTime As Currency
            QueryPerformanceCounter currentTime
            ExecutionTime = (currentTime - m_StartTime) * 1000 / freq
        Else
            ExecutionTime = (m_EndTime - m_StartTime) * 1000 / freq
        End If
    End Property

    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    
    ' 日志记录相关API
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    
    '@Description "设置任务优先级"
    Public Property Let Priority(ByVal Value As eTaskPriority)
        LockState
        m_Priority = Value
        UnlockState
    End Property
    
    '@Description "获取任务优先级"
    Public Property Get Priority() As eTaskPriority
        LockState
        Priority = m_Priority
        UnlockState
    End Property
    
    '@Description "设置任务超时时间"
    Public Sub SetTimeout(ByVal TimeoutMs As Long)
        LockState
        m_TaskTimeout = TimeoutMs
        UnlockState
    End Sub
    
    '@Description "设置重试参数"
    Public Sub SetRetryPolicy(ByVal MaxRetries As Long, Optional ByVal RetryDelayMs As Long = 1000)
        LockState
        m_MaxRetries = MaxRetries
        m_RetryDelay = RetryDelayMs
        UnlockState
    End Sub
    
    '@Description "添加任务依赖"
    Public Sub AddDependency(ByVal Task As cTask)
        LockState
        m_Dependencies.Add Task
        UnlockState
    End Sub
    
    '@Description "检查所有依赖任务是否完成"
    Private Function AreDependenciesMet() As Boolean
        Dim task As cTask
        For Each task In m_Dependencies
            If task.IsRunning Then
                AreDependenciesMet = False
                Exit Function
            End If
        Next
        AreDependenciesMet = True
    End Function
    
    '@Description "设置完成回调"
    Public Sub SetOnComplete(ByVal CallbackProc As LongPtr)
        LockState
        m_OnComplete = CallbackProc
        UnlockState
    End Sub
    
    '@Description "设置错误回调"
    Public Sub SetOnError(ByVal CallbackProc As LongPtr)
        LockState
        m_OnError = CallbackProc
        UnlockState
    End Sub
    
    '@Description "启用日志记录"
    Public Sub EnableLogging(ByVal LogFilePath As String)
        LockState
        m_LogFile = LogFilePath
        WriteLog "Task logging enabled"
        UnlockState
    End Sub
    
    '@Description "写入日志"
    Private Sub WriteLog(ByVal Message As String)
        If m_LogFile = "" Then Exit Sub
        
        Dim fNum As Integer
        fNum = FreeFile
        
        On Error Resume Next
        Open m_LogFile For Append As #fNum
        Print #fNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [PID:" & GetCurrentProcessId & "|TID:" & GetCurrentThreadId & "] " & Message
        Close #fNum
    End Sub
    
    '@Description "获取任务结果"
    Public Property Get Result() As eTaskResult
        LockState
        Result = m_Result
        UnlockState
    End Property
    
    '@Description "获取当前重试次数"
    Public Property Get RetryCount() As Long
        LockState
        RetryCount = m_RetryCount
        UnlockState
    End Property
    
    '@Description "执行任务（包含重试逻辑）"
    Private Function ExecuteWithRetry() As Boolean
        Dim success As Boolean
        m_RetryCount = 0
        
        Do
            WriteLog "Executing task (Attempt " & (m_RetryCount + 1) & " of " & (m_MaxRetries + 1) & ")"
            
            success = ExecuteTask()
            
            If success Then
                WriteLog "Task executed successfully"
                ExecuteWithRetry = True
                Exit Function
            End If
            
            m_RetryCount = m_RetryCount + 1
            
            If m_RetryCount <= m_MaxRetries Then
                WriteLog "Task failed, retrying in " & m_RetryDelay & "ms"
                Sleep m_RetryDelay
            End If
        Loop While m_RetryCount <= m_MaxRetries And Not m_CancelRequested
        
        If Not success Then
            WriteLog "Task failed after " & m_RetryCount & " retries"
        End If
        
        ExecuteWithRetry = success
    End Function
    
    '@Description "任务实际执行逻辑"
    Private Function ExecuteTask() As Boolean
        On Error GoTo ErrorHandler
        
        ' 检查依赖
        If Not AreDependenciesMet Then
            WriteLog "Dependencies not met"
            m_Result = TaskResult_Failed
            Exit Function
        End If
        
        ' 设置超时定时器（如果指定了超时）
        Dim timeoutThread As LongPtr
        If m_TaskTimeout <> INFINITE Then
            timeoutThread = CreateThread(0, 0, AddressOf TimeoutProc, ObjPtr(Me), 0, 0)
        End If
        
        ' 执行任务
        Dim taskResult As Long
        taskResult = CallWindowProc(m_TaskProcedure, ObjPtr(Me), 0, 0, 0)
        
        ' 清理超时线程
        If timeoutThread <> 0 Then
            TerminateThread timeoutThread, 0
            CloseHandle timeoutThread
        End If
        
        ' 调用完成回调
        If m_OnComplete <> 0 Then
            CallWindowProc m_OnComplete, ObjPtr(Me), taskResult, 0, 0
        End If
        
        ExecuteTask = True
        Exit Function
        
    ErrorHandler:
        WriteLog "Error in ExecuteTask: " & Err.Description
        m_LastError = Err.Number
        m_ErrorDescription = Err.Description
        m_Result = TaskResult_Failed
        
        ' 调用错误回调
        If m_OnError <> 0 Then
            CallWindowProc m_OnError, ObjPtr(Me), Err.Number, 0, 0
        End If
        
        ExecuteTask = False
    End Function
    
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As LongPtr, _
        ByVal hWnd As LongPtr, _
        ByVal Msg As Long, _
        ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr _
    ) As Long

End Class