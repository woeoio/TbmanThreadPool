# ThreadPool 特性集成指南

> 重要提示：ThreadPool 已内置实现了以下核心特性：
> - ⚡ 超时处理 (`SetTimeout`, `HandleTimeout`)
> - 🔄 重试机制 (`SetRetryPolicy`, `MaxRetries`, `RetryDelay`)
> - 📊 性能监控 (`ExecutionTime`, 任务状态跟踪)
> - ⏸️ 暂停/恢复 (`Pause`, `Resume`)
> - ❌ 取消支持 (`RequestCancel`, `CancelRequested`)
> - 📝 日志记录 (`EnableLogging`, `WriteLog`)
> 
> 本文档展示如何将这些内置特性与新增的 TaskData/Result 机制结合使用，以构建更强大的任务处理方案。

## 1. 超时处理与TaskData/Result集成

### 利用内置超时特性

### 扩展内置超时功能
```vb
' ThreadPool 已内置超时处理，这里展示如何扩展它
With New cThread
    ' 使用内置超时机制
    .SetTimeout 30000                        ' 设置总体超时（内置功能）
    
    ' 使用 TaskData/Result 扩展超时处理能力
    .SetTaskData "subOperationTimeout", 5000 ' 子操作超时
    .SetTaskData "timeoutAction", "retry"    ' 自定义超时行为
    .SetTaskData "timeoutRetries", 3         ' 超时后重试次数
End With
```

### 超时处理与结果记录
```vb
Public Function TimeoutAwareTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(param)
    
    ' 获取超时配置
    Dim operationTimeout As Long
    operationTimeout = task.GetTaskData("operationTimeout")
    
    ' 执行操作并记录超时情况
    task.SetTaskResult "startTime", Now
    
    Dim success As Boolean
    success = PerformOperation(operationTimeout)
    
    If Not success Then
        task.SetTaskResult "timedOut", True
        task.SetTaskResult "timeoutDuration", operationTimeout
        task.SetTaskResult "failurePoint", "operation"
    End If
    
    TimeoutAwareTaskProc = IIf(success, 0, 1)
End Function
```

## 2. 重试机制与状态管理

### 增强内置重试功能
```vb
Public Sub ConfigureTaskWithRetry(task As cThread)
    ' 使用内置重试机制
    task.SetRetryPolicy 3, 1000  ' 设置最大重试次数和延迟（内置功能）
    
    ' 使用 TaskData/Result 扩展重试功能
    task.SetTaskData "retryBackoff", 2    ' 自定义指数退避系数
    task.SetTaskData "retryConditions", CreateRetryConditions() ' 自定义重试条件
    
    ' 重试状态追踪
    task.SetTaskResult "retryCount", 0
    task.SetTaskResult "retryHistory", CreateObject("Scripting.Dictionary")
End Sub

Public Function RetryAwareTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(param)
    
    Dim retryCount As Long
    retryCount = task.GetTaskResult("retryCount")
    
    ' 执行操作
    Dim success As Boolean
    success = PerformOperation()
    
    ' 更新重试历史
    Dim history As Object
    Set history = task.GetTaskResult("retryHistory")
    history.Add Now, "Attempt " & (retryCount + 1) & ": " & IIf(success, "Success", "Failed")
    
    ' 更新重试计数
    task.SetTaskResult "retryCount", retryCount + 1
    
    RetryAwareTaskProc = IIf(success, 0, 1)
End Function
```

## 3. 进度报告与任务控制

### 结合内置控制功能
> ThreadPool 已提供了 `Pause`、`Resume`、`RequestCancel` 等控制功能，
> 下面展示如何利用 TaskData/Result 增强这些功能。

```vb
Public Function ProgressAwareTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(param)
    
    ' 初始化进度信息
    task.SetTaskResult "progress", 0
    task.SetTaskResult "status", "initializing"
    task.SetTaskResult "startTime", Now
    task.SetTaskResult "estimatedCompletion", Empty
    
    ' 处理过程
    Dim totalSteps As Long
    totalSteps = task.GetTaskData("totalSteps")
    
    For i = 1 To totalSteps
        If task.CancelRequested Then
            ' 记录取消状态
            task.SetTaskResult "cancelTime", Now
            task.SetTaskResult "cancelReason", "user_requested"
            task.SetTaskResult "lastCompletedStep", i - 1
            Exit For
        End If
        
        ' 更新进度
        UpdateTaskProgress task, i, totalSteps
        
        ' 执行步骤...
    Next i
    
    ProgressAwareTaskProc = 0
End Function

Private Sub UpdateTaskProgress(task As cThread, currentStep As Long, totalSteps As Long)
    Dim progress As Double
    progress = currentStep / totalSteps * 100
    
    ' 更新进度信息
    task.SetTaskResult "progress", progress
    task.SetTaskResult "currentStep", currentStep
    task.SetTaskResult "status", "processing"
    
    ' 计算预计完成时间
    If currentStep > 1 Then
        Dim startTime As Date
        startTime = task.GetTaskResult("startTime")
        Dim timePerStep As Double
        timePerStep = DateDiff("s", startTime, Now) / (currentStep - 1)
        Dim remainingSteps As Long
        remainingSteps = totalSteps - currentStep
        Dim estimatedCompletion As Date
        estimatedCompletion = DateAdd("s", timePerStep * remainingSteps, Now)
        
        task.SetTaskResult "estimatedCompletion", estimatedCompletion
    End If
End Sub
```

## 4. 日志与调试集成

### 扩展内置日志功能
> ThreadPool 已内置了基础的日志功能（通过 `EnableLogging` 和 `WriteLog`），
> 下面展示如何使用 TaskData/Result 构建更强大的日志和调试系统。
```vb
Private Sub LogTaskEvent(task As cThread, eventType As String, message As String)
    ' 确保日志容器存在
    If task.GetTaskResult("eventLog") = Empty Then
        task.SetTaskResult "eventLog", CreateObject("Scripting.Dictionary")
    End If
    
    Dim log As Object
    Set log = task.GetTaskResult("eventLog")
    
    ' 创建日志条目
    Dim entry As New Dictionary
    With entry
        .Add "timestamp", Now
        .Add "type", eventType
        .Add "message", message
        .Add "taskId", task.ThreadID
        .Add "progress", task.GetTaskResult("progress")
        .Add "status", task.GetTaskResult("status")
    End With
    
    ' 添加到日志
    log.Add log.Count + 1, entry
    
    ' 如果需要，写入文件
    If task.GetTaskData("logToFile") Then
        WriteLogToFile task, entry
    End If
End Sub
```

### 调试辅助功能
```vb
Public Function GetTaskDebugInfo(task As cThread) As Dictionary
    Dim debug As New Dictionary
    
    ' 基本信息
    debug.Add "threadId", task.ThreadID
    debug.Add "status", task.GetTaskResult("status")
    debug.Add "runtime", DateDiff("s", task.GetTaskResult("startTime"), Now)
    
    ' 性能指标
    debug.Add "progress", task.GetTaskResult("progress")
    debug.Add "memoryUsage", GetTaskMemoryUsage(task)
    
    ' 错误信息
    If task.GetTaskResult("lastError") <> Empty Then
        debug.Add "lastError", task.GetTaskResult("lastError")
        debug.Add "errorTimestamp", task.GetTaskResult("errorTime")
    End If
    
    ' 重试信息
    debug.Add "retryCount", task.GetTaskResult("retryCount")
    debug.Add "retryHistory", task.GetTaskResult("retryHistory")
    
    Set GetTaskDebugInfo = debug
End Function
```

## 5. 组合内置特性与扩展功能

### 完整示例：整合所有特性
> 本示例展示如何将 ThreadPool 的内置特性（超时、重试、日志等）
> 与 TaskData/Result 机制结合使用，构建功能完整的任务处理系统。
```vb
Public Function ComplexTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(param)
    
    ' 初始化
    InitializeTaskTracking task
    
    Do While ShouldRetry(task)
        ' 更新重试状态
        UpdateRetryStatus task
        
        ' 执行带超时的操作
        Dim success As Boolean
        success = ExecuteWithTimeout(task)
        
        If success Then
            ' 记录成功
            LogTaskSuccess task
            ComplexTaskProc = 0
            Exit Function
        Else
            ' 记录失败
            LogTaskFailure task
            
            ' 检查是否应该重试
            If Not ShouldRetry(task) Then
                ComplexTaskProc = 1
                Exit Function
            End If
            
            ' 等待重试
            WaitForRetry task
        End If
    Loop
    
    ComplexTaskProc = 1
End Function

Private Sub InitializeTaskTracking(task As cThread)
    ' 初始化所有跟踪状态
    With task
        .SetTaskResult "attempts", 0
        .SetTaskResult "startTime", Now
        .SetTaskResult "status", "initialized"
        .SetTaskResult "lastError", Empty
        .SetTaskResult "progress", 0
    End With
End Sub
```

## 总结

通过合理集成各个特性，可以构建出健壮、可监控、高性能的任务处理系统。关键点包括：

1. 使用 TaskData/Result 存储配置和状态
2. 实现全面的进度和状态跟踪
3. 提供详细的日志和调试信息
4. 优雅处理超时和重试
5. 支持取消和清理操作

这些特性的组合使用可以显著提高任务处理的可靠性和可维护性。
