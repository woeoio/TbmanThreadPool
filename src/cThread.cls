Class cThread

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Document: https://doc.twinbasic.vb6.pro/en/tbman/threadPool/

    '@Module
    '@Folder("Threading")
    '@Description "代表一个可管理的线程任务"

    ' API 声明
    Private Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( _
        ByVal lpMutexAttributes As LongPtr, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String _
    ) As LongPtr

    Private Declare PtrSafe Function ReleaseMutex Lib "kernel32" ( _
        ByVal hMutex As LongPtr _
    ) As Long

    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As LongPtr, _
        ByVal dwMilliseconds As Long _
    ) As Long

    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr _
    ) As Long

    Private Declare PtrSafe Function GetExitCodeThread Lib "kernel32" ( _
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
    Private m_TaskData As Collection     ' 任务输入参数集合
    Private m_TaskResult As Collection   ' 任务结果集合
    Private m_UINotifier As cThreadUINotifier  ' UI通知器
    Private m_TaskId As String           ' 任务唯一标识符

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
        ' 创建一个命名互斥锁，以对象指针作为唯一标识
        m_StateLock = CreateMutex(0, False, "TB_Task_" & ObjPtr(Me))
        If m_StateLock = 0 Then
            m_LastError = Err.LastDllError
            m_ErrorDescription = "Failed to create mutex"
            WriteLog "Failed to create mutex: " & m_LastError
        End If
        Set m_TaskData = New Collection
        Set m_TaskResult = New Collection
        
        ' 生成唯一的任务ID
        m_TaskId = "Task_" & Format$(Now, "yyyymmddhhnnss") & "_" & ObjPtr(Me)
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        
        ' 先尝试释放锁，以免其他线程还在等待
        If m_StateLock <> 0 Then
            ReleaseMutex m_StateLock
        End If
        
        ' 确保线程终止并清理资源
        If m_IsRunning Then
            Terminate
        End If
        
        ' 最后关闭互斥锁句柄
        If m_StateLock <> 0 Then
            CloseHandle m_StateLock
            m_StateLock = 0
        End If
        
        Set m_TaskData = Nothing
        Set m_TaskResult = Nothing
    End Sub

    ' ======================
    ' 公共方法
    ' ======================

    '@Description "创建新任务实例"
    Public Function Create(ByVal TaskProcedure As LongPtr, _
                          Optional ByVal Tag As Variant, _
                          Optional ByVal StartImmediately As Boolean = True _
                          ) As cThread
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
    Public Function Start() As Boolean
        Dim success As Boolean
        
        If Not TryWithLock("Start", success) Then
            WriteLog "Failed to acquire lock in Start"
            Start = False
            Exit Function
        End If
        
        If m_IsRunning Then
            Start = True ' 已经在运行，视为成功
            Exit Function
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
            WriteLog "Thread creation failed with error: " & m_LastError
            Start = False
            Exit Function
        End If
        
        m_IsRunning = True
        Start = True
    End Function

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
            Err.Raise vbObjectError + 1001, "cThread.Pause", "Thread suspension failed with error: " & m_LastError
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
            Err.Raise vbObjectError + 1002, "cThread.Resume", "Thread resumption failed with error: " & m_LastError
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
            Err.Raise vbObjectError + 1003, "cThread.Terminate", "Thread termination failed with error: " & m_LastError
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
                Err.Raise vbObjectError + 1004, "cThread.WaitForCompletion", "Wait failed with error: " & m_LastError
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
        ' Tag 通常在初始化后不变，不需要锁
        Tag = m_Tag
    End Property
    
    '@Description "获取线程ID"
    Public Property Get ThreadID() As LongPtr
        ' ThreadID 在线程创建后不变，不需要锁
        ThreadID = m_ThreadID
    End Property
    
    '@Description "获取线程句柄"
    Public Property Get ThreadHandle() As LongPtr
        ' ThreadHandle 在线程创建后不变，不需要锁
        ThreadHandle = m_ThreadHandle
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

    Private Function LockState(Optional ByVal TimeoutMs As Long = 5000) As Boolean
        On Error Resume Next
        
        ' 检查互斥锁是否有效
        If m_StateLock = 0 Then
            ' 尝试重新创建互斥锁
            m_StateLock = CreateMutex(0, False, "TB_Task_" & ObjPtr(Me))
            If m_StateLock = 0 Then
                WriteLog "Failed to recreate mutex: " & Err.LastDllError
                LockState = False
                Exit Function
            End If
        End If
        
        ' 尝试获取锁
        Dim result As Long
        result = WaitForSingleObject(m_StateLock, TimeoutMs)
        
        Select Case result
            Case WAIT_OBJECT_0
                ' 成功获取锁
                LockState = True
            
            Case WAIT_TIMEOUT
                ' 获取锁超时
                WriteLog "Lock acquisition timed out after " & TimeoutMs & "ms"
                LockState = False
            
            Case WAIT_FAILED
                ' 获取锁失败
                m_LastError = Err.LastDllError
                m_ErrorDescription = "Failed to acquire lock"
                WriteLog "Lock acquisition failed with error: " & m_LastError
                
                ' 如果是无效句柄，尝试重新创建
                If m_LastError = 6 Then  ' ERROR_INVALID_HANDLE
                    CloseHandle m_StateLock
                    m_StateLock = CreateMutex(0, False, "TB_Task_" & ObjPtr(Me))
                    If m_StateLock <> 0 Then
                        ' 重试一次
                        result = WaitForSingleObject(m_StateLock, TimeoutMs)
                        LockState = (result = WAIT_OBJECT_0)
                    End If
                End If
                
                If Not LockState Then
                    WriteLog "Lock retry failed"
                End If
        End Select
    End Function

    Private Sub UnlockState()
        On Error Resume Next
        If m_StateLock <> 0 Then
            ReleaseMutex m_StateLock
            If Err.LastDllError = 6 Then  ' ERROR_INVALID_HANDLE
                m_StateLock = 0  ' 标记为无效，下次会重新创建
            End If
        End If
    End Sub
    
    '@Description "尝试执行受保护的操作"
    Private Function TryWithLock(ByVal Action As String, ByRef Result As Boolean) As Boolean
        Result = False
        
        ' 尝试获取锁
        If Not LockState() Then
            WriteLog "Failed to acquire lock for action: " & Action
            Exit Function
        End If
        
        On Error GoTo ErrorHandler
        Result = True ' 将在实际操作中被重写
        TryWithLock = True
        
    ExitFunction:
        UnlockState
        Exit Function
        
    ErrorHandler:
        WriteLog "Error in " & Action & ": " & Err.Description
        Result = False
        Resume ExitFunction
    End Function
    
    '@Description "获取任务执行时间(毫秒)"
    Public Property Get ExecutionTime() As Double
        ' 时间读取是原子操作，不需要锁
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

    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    
    ' 日志记录相关API
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    
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
    Public Sub AddDependency(ByVal Task As cThread)
        LockState
        m_Dependencies.Add Task
        UnlockState
    End Sub
    
    '@Description "检查所有依赖任务是否完成"
    Private Function AreDependenciesMet() As Boolean
        Dim task As cThread
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
    
    '@Description "设置UI通知器"
    Public Sub SetUINotifier(ByRef notifier As cThreadUINotifier)
        LockState
        Set m_UINotifier = notifier
        UnlockState
    End Sub
    
    '@Description "向UI线程发送通知"
    Public Sub NotifyToUI(ByVal NotificationCode As Long, Optional ByVal Data As Variant)
        ' 这个方法可以在任何线程中安全调用
        If m_UINotifier Is Nothing Then Exit Sub
        
        m_UINotifier.PostNotification m_TaskId, NotificationCode, Data
    End Sub
    
    '@Description "向UI线程发送进度通知"
    Public Sub NotifyProgress(ByVal CurrentValue As Long, ByVal MaxValue As Long, Optional ByVal Message As String = "")
        If m_UINotifier Is Nothing Then Exit Sub
        
        m_UINotifier.PostProgress m_TaskId, CurrentValue, MaxValue, Message
    End Sub
    
    '@Description "向UI线程发送状态通知"
    Public Sub NotifyStatus(ByVal Message As String)
        If m_UINotifier Is Nothing Then Exit Sub
        
        m_UINotifier.PostNotification m_TaskId, 2, Message  ' 2 = Notify_Status
    End Sub
    
    '@Description "向UI线程发送错误通知"
    Public Sub NotifyError(ByVal ErrorCode As Long, ByVal ErrorMessage As String)
        If m_UINotifier Is Nothing Then Exit Sub
        
        m_UINotifier.PostError m_TaskId, ErrorCode, ErrorMessage
    End Sub
    
    '@Description "向UI线程发送完成通知"
    Public Sub NotifyCompleted(Optional ByVal Result As Variant)
        If m_UINotifier Is Nothing Then Exit Sub
        
        m_UINotifier.PostCompleted m_TaskId, Result
    End Sub
    
    '@Description "获取任务ID"
    Public Property Get TaskId() As String
        TaskId = m_TaskId
    End Property
    
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
    
    '@Description "设置任务输入参数"
    Public Sub SetTaskData(ByVal Key As String, ByVal Value As Variant)
        LockState
        On Error Resume Next
        m_TaskData.Remove Key
        On Error GoTo ErrorHandler
        m_TaskData.Add Value, Key
        UnlockState
        Exit Sub
        
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, "cThread.SetTaskData", "Failed to set task parameter: " & Err.Description
    End Sub
    
    '@Description "获取任务输入参数"
    Public Function GetTaskData(ByVal Key As String) As Variant
        LockState
        On Error GoTo ErrorHandler
        If IsObject(m_TaskData(Key)) Then
            Set GetTaskData = m_TaskData(Key)
        Else
            GetTaskData = m_TaskData(Key)
        End If
        UnlockState
        Exit Function
        
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, "cThread.GetTaskData", "Failed to get task parameter: " & Err.Description
    End Function
    
    '@Description "设置任务结果"
    Public Sub SetTaskResult(ByVal Key As String, ByVal Value As Variant)
        LockState
        On Error Resume Next
        m_TaskResult.Remove Key
        On Error GoTo ErrorHandler
        m_TaskResult.Add Value, Key
        UnlockState
        Exit Sub
        
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, "cThread.SetTaskResult", "Failed to set task result: " & Err.Description
    End Sub
    
    '@Description "获取任务结果"
    Public Function GetTaskResult(ByVal Key As String) As Variant
        LockState
        On Error GoTo ErrorHandler
        If IsObject(m_TaskResult(Key)) Then
            Set GetTaskResult = m_TaskResult(Key)
        Else
            GetTaskResult = m_TaskResult(Key)
        End If
        UnlockState
        Exit Function
        
    ErrorHandler:
        UnlockState
        Err.Raise Err.Number, "cThread.GetTaskResult", "Failed to get task result: " & Err.Description
    End Function
    
    '@Description "清除任务数据"
    Public Sub ClearTaskData()
        LockState
        Set m_TaskData = New Collection
        UnlockState
    End Sub
    
    '@Description "清除任务结果"
    Public Sub ClearTaskResults()
        LockState
        Set m_TaskResult = New Collection
        UnlockState
    End Sub
    
    '@Description "获取所有任务数据键"
    Public Function GetTaskDataKeys() As Collection
        Dim result As New Collection
        Dim item As Variant
        Dim i As Long
        
        LockState
        For i = 1 To m_TaskData.Count
            result.Add m_TaskData(i)
        Next i
        UnlockState
        
        Set GetTaskDataKeys = result
    End Function
    
    '@Description "获取所有任务结果键"
    Public Function GetTaskResultKeys() As Collection
        Dim result As New Collection
        Dim item As Variant
        Dim i As Long
        
        LockState
        For i = 1 To m_TaskResult.Count
            result.Add m_TaskResult(i)
        Next i
        UnlockState
        
        Set GetTaskResultKeys = result
    End Function
    
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
    
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As LongPtr, _
        ByVal hWnd As LongPtr, _
        ByVal Msg As Long, _
        ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr _
    ) As Long

    ' 超时处理专用的Friend方法
    Friend Sub HandleTimeout()
        If m_IsRunning Then
            WriteLog "Task timed out after " & m_TaskTimeout & "ms"
            m_Result = TaskResult_TimedOut
            RequestCancel
        End If
    End Sub
    
    '@Description "超时处理线程过程"
    Private Function TimeoutProc(ByVal param As LongPtr) As Long
        On Error GoTo ErrorHandler
        
        Dim task As cThread
        Set task = mThread.ReturnFromPtr(param)
        
        ' 使用公共接口获取超时时间
        Dim timeoutValue As Long
        timeoutValue = task.GetTimeout
        
        ' 等待指定的超时时间
        Sleep timeoutValue
        
        ' 通过Friend方法处理超时
        task.HandleTimeout
        
        TimeoutProc = 0
        Exit Function
        
    ErrorHandler:
        TimeoutProc = 1
    End Function
    
    '@Description "获取超时时间"
    Public Function GetTimeout() As Long
        ' 超时时间通常在设置后不变，不需要锁
        GetTimeout = m_TaskTimeout
    End Function
    
    '@Description "任务是否超时"
    Public Property Get IsTimedOut() As Boolean
        LockState
        IsTimedOut = (m_Result = TaskResult_TimedOut)
        UnlockState
    End Property

End Class