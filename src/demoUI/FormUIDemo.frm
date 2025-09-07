[Description("")]
[FormDesignerId("AEF87A46-FA0F-41E0-BE51-80166E1DC4C5")]
[PredeclaredId]
Class FormUIDemo

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Document: https://doc.twinbasic.vb6.pro/en/tbman/threadPool/
    '@Description "UI通知器演示窗体"

    ' UI通知器 - 使用WithEvents来接收事件
    Private WithEvents m_UINotifier As cThreadUINotifier
    Private m_ThreadPool As cThreadPool
    Private m_TaskCounter As Long
    
    Private m_TaskidMapProcessBarIndex As New Collection

    ' 窗体加载事件
    Private Sub Form_Load()
        ' 创建UI通知器（自动初始化）
        Set m_UINotifier = New cThreadUINotifier
        
        ' 创建线程池
        Set m_ThreadPool = New cThreadPool
        m_ThreadPool.Create 3  ' 最多3个并发线程
        m_ThreadPool.SetUINotifier m_UINotifier
        
        ' 初始化UI
        InitializeUI
        
        AddLogMessage "UI Notifier Demo initialized successfully"
        AddLogMessage "Thread pool created with max 3 concurrent threads"
        
        ' 显示线程池初始状态
        AddLogMessage "Initial pool state: " & vbCrLf & m_ThreadPool.GetPoolStats()
    End Sub

    ' 窗体卸载事件
    Private Sub Form_Unload(Cancel As Integer)
        AddLogMessage "Form unloading - cleaning up resources..."
        
        ' 显示卸载前的线程池状态
        If Not (m_ThreadPool Is Nothing) Then
            AddLogMessage "Final cleanup - Active threads: " & m_ThreadPool.ActiveThreads
        End If
        
        ' 线程池会在对象销毁时自动终止所有线程并清理资源
        Set m_ThreadPool = Nothing
        Set m_UINotifier = Nothing
        
        AddLogMessage "Resources cleaned up successfully"
    End Sub

    ' 初始化UI控件
    Private Sub InitializeUI()
        ' 这里应该设置各种UI控件的初始状态
        ' 例如：进度条、列表框、按钮等
        
        ' 示例（需要根据实际的控件名称调整）:
        Dim i As Long
        For i = 0 To 2
            ProgressBar1(i).Min = 0
            ProgressBar1(i).Max = 100
            ProgressBar1(i).Value = 0
            Label1(i).Caption = "就绪"
        Next
        LabelStatus.Caption = "Ready"
        ListBoxLog.Clear
        ButtonStart.Caption = "Start Tasks"
        ButtonCancel.Caption = "Cancel All"
        ButtonCancel.Enabled = False
    End Sub

    ' =====================================================
    ' UI通知器事件处理
    ' =====================================================

    ' 处理线程进度通知
    Private Sub m_UINotifier_ThreadProgress(ByVal TaskId As String, ByVal CurrentValue As Long, ByVal MaxValue As Long, ByVal Message As String)
        ' 更新进度条
        Dim i As Long = m_TaskidMapProcessBarIndex.Item(TaskId)
        ProgressBar1(i).Max = MaxValue
        ProgressBar1(i).Value = CurrentValue
        Label1(i).Caption = Message
        
        ' 更新状态标签
        Dim percent As Long
        If MaxValue > 0 Then
            percent = (CurrentValue * 100) \ MaxValue
        End If
        
        LabelStatus.Caption = "Doing..."
        ' LabelStatus.Caption = "Task " & GetTaskDisplayName(TaskId) & ": " & percent & "% - " & Message
        
        ' 记录到日志
        AddLogMessage "[PROGRESS] " & GetTaskDisplayName(TaskId) & ": " & CurrentValue & "/" & MaxValue & " - " & Message
        
        ' 强制刷新UI
        DoEvents
    End Sub

    ' 处理线程完成通知
    Private Sub m_UINotifier_ThreadCompleted(ByVal TaskId As String, ByVal Result As Variant)
        Dim i As Long = m_TaskidMapProcessBarIndex.Item(TaskId)
        Label1(i).Caption = "已完成"
        
        AddLogMessage "[COMPLETED] " & GetTaskDisplayName(TaskId) & " finished with result: " & CStr(Result)
        
        ' 更新UI状态
        UpdateTaskStatus TaskId, "Completed"
        
        
        ' 注释：线程池现在会自动清理已完成的任务，不需要手动调用 CleanupCompleted
        ' If Not (m_ThreadPool Is Nothing) Then
        '     m_ThreadPool.CleanupCompleted
        ' End If
        
        ' 检查是否所有任务都完成了
        CheckAllTasksCompleted
    End Sub

    ' 处理线程错误通知
    Private Sub m_UINotifier_ThreadError(ByVal TaskId As String, ByVal ErrorCode As Long, ByVal ErrorMessage As String)
        AddLogMessage "[ERROR] " & GetTaskDisplayName(TaskId) & " error " & ErrorCode & ": " & ErrorMessage
        
        ' 更新UI状态
        UpdateTaskStatus TaskId, "Error: " & ErrorMessage
        
        ' 更新对应的进度条和标签
        On Error Resume Next
        Dim i As Long = m_TaskidMapProcessBarIndex.Item(TaskId)
        If Err.Number = 0 Then
            Label1(i).Caption = "错误"
        End If
        On Error GoTo 0
        
        ' 注释：线程池现在会自动清理已完成的任务，不需要手动调用 CleanupCompleted
        ' If Not (m_ThreadPool Is Nothing) Then
        '     m_ThreadPool.CleanupCompleted
        ' End If
        
        ' 检查是否所有任务都完成了
        CheckAllTasksCompleted
        
        ' 可以选择显示错误对话框
        ' MsgBox "Task error: " & ErrorMessage, vbExclamation
    End Sub

    ' 处理一般线程通知
    Private Sub m_UINotifier_ThreadNotification(ByVal TaskId As String, ByVal NotificationCode As Long, ByVal Data As Variant)
        Dim message As String
        
        Select Case NotificationCode
            Case 2  ' Notify_Status
                message = "[STATUS] " & GetTaskDisplayName(TaskId) & ": " & CStr(Data)
                UpdateTaskStatus TaskId, CStr(Data)
                
            Case 3  ' Notify_Warning
                message = "[WARNING] " & GetTaskDisplayName(TaskId) & ": " & CStr(Data)
                
            Case 101  ' 自定义通知（数据验证警告）
                message = "[DATA] " & GetTaskDisplayName(TaskId) & ": " & CStr(Data)
                
            Case Else
                message = "[NOTIFY " & NotificationCode & "] " & GetTaskDisplayName(TaskId) & ": " & CStr(Data)
        End Select
        
        AddLogMessage message
    End Sub

    ' =====================================================
    ' 按钮事件处理
    ' =====================================================

    ' 开始任务按钮
    Private Sub ButtonStart_Click()
        ' 清理之前的日志（可选）
        ' ListBoxLog.Clear
        
        ' 重置任务计数器
        m_TaskCounter = 0
        
        AddLogMessage "Starting demo tasks..."
        
        ' 显示线程池初始状态
        If Not (m_ThreadPool Is Nothing) Then
            AddLogMessage "Pool status before starting tasks: " & vbCrLf & m_ThreadPool.GetPoolStats()
        End If
        
        ' 启动不同类型的任务
        m_TaskidMapProcessBarIndex.Clear()
        StartFileProcessTask
        StartDownloadTask
        StartDatabaseTask
        
        ' 显示任务启动后的线程池状态
        If Not (m_ThreadPool Is Nothing) Then
            AddLogMessage "Pool status after starting tasks - Active: " & m_ThreadPool.ActiveThreads & ", Total: " & m_ThreadPool.Count
        End If
        
        ' 更新UI状态
        ButtonStart.Enabled = False
        ButtonCancel.Enabled = True
        
        AddLogMessage "All tasks started. Monitor progress below."
    End Sub

    ' 取消所有任务按钮
    Private Sub ButtonCancel_Click()
        If Not (m_ThreadPool Is Nothing) Then
            AddLogMessage "Cancelling all tasks..."
            
            ' 显示取消前的状态
            AddLogMessage "Before cancel - Active: " & m_ThreadPool.ActiveThreads & ", Total: " & m_ThreadPool.Count
            
            m_ThreadPool.TerminateAll
            AddLogMessage "All tasks cancelled by user"
            
            ' 显示取消后的状态
            AddLogMessage "After cancel - Active: " & m_ThreadPool.ActiveThreads & ", Total: " & m_ThreadPool.Count
            
            ' 清空任务映射集合
            m_TaskidMapProcessBarIndex.Clear()
            
            ' 更新UI状态
            ButtonStart.Enabled = True
            ButtonCancel.Enabled = False
           
            InitializeUI()
        End If
    End Sub

    ' =====================================================
    ' 任务启动方法
    ' =====================================================

    Private Sub StartFileProcessTask()
        Dim task As cThread
        Set task = m_ThreadPool.AddTask(AddressOf mUINotifierExample.FileProcessTask, "FileTask")
        AddLogMessage "Started file processing task: " & task.TaskId
        m_TaskidMapProcessBarIndex.Add(0, task.TaskId)
    End Sub

    Private Sub StartDownloadTask()
        Dim task As cThread
        Set task = m_ThreadPool.AddTask(AddressOf mUINotifierExample.DownloadTask, "DownloadTask")
        AddLogMessage "Started download task: " & task.TaskId
        m_TaskidMapProcessBarIndex.Add(1, task.TaskId)
    End Sub

    Private Sub StartDatabaseTask()
        Dim task As cThread
        Set task = m_ThreadPool.AddTask(AddressOf mUINotifierExample.DatabaseBatchTask, "DatabaseTask")
        AddLogMessage "Started database task: " & task.TaskId
        m_TaskidMapProcessBarIndex.Add(2, task.TaskId)
    End Sub

    ' =====================================================
    ' 辅助方法
    ' =====================================================

    ' 添加日志消息
    Private Sub AddLogMessage(ByVal message As String)
        Dim timeStamp As String
        timeStamp = Format$(Now, "hh:nn:ss")
        
        Dim logEntry As String
        logEntry = timeStamp & " - " & message
        
        ' 添加到日志列表框
        ListBoxLog.AddItem logEntry
        ListBoxLog.TopIndex = ListBoxLog.ListCount - 1  ' 自动滚动到底部
        
        ' 同时输出到调试窗口
        Debug.Print logEntry
    End Sub

    ' 获取任务显示名称
    Private Function GetTaskDisplayName(ByVal TaskId As String) As String
        ' 从TaskId中提取有意义的显示名称
        If InStr(TaskId, "Task_") > 0 Then
            ' 可以根据Tag或其他信息来确定任务类型
            GetTaskDisplayName = "Task[" & Right$(TaskId, 6) & "]"
        Else
            GetTaskDisplayName = TaskId
        End If
    End Function

    ' 更新任务状态
    Private Sub UpdateTaskStatus(ByVal TaskId As String, ByVal status As String)
        ' 这里可以更新任务状态显示
        ' 例如在列表视图或树形控件中显示各个任务的状态
        Debug.Print "Task " & GetTaskDisplayName(TaskId) & " status: " & status
    End Sub

    ' 检查所有任务是否完成
    Private Sub CheckAllTasksCompleted()
        ' 使用线程池的内置完成检查
        If m_ThreadPool.IsAllTasksCompleted Then
            AddLogMessage "All tasks completed! (Timer confirmation)"
                            
            ' 重新启用开始按钮
            ButtonStart.Enabled = True
            ButtonCancel.Enabled = False
            LabelStatus.Caption = "All tasks completed"

            Exit Sub
        End If
    End Sub



End Class
