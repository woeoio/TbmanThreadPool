# ThreadPool 故障排除指南

## 1. 常见问题与解决方案

### 任务不启动
#### 症状
- 任务创建后没有执行
- `IsRunning` 返回 False
- 没有任何错误提示

#### 可能的原因
1. 任务过程地址无效
2. 资源限制
3. 线程创建失败

#### 解决方案
```vb
' 检查任务过程地址
Public Function ValidateTaskProc(taskProc As LongPtr) As Boolean
    If taskProc = 0 Then
        Debug.Print "无效的任务过程地址"
        ValidateTaskProc = False
        Exit Function
    End If
    
    ' 其他验证...
    ValidateTaskProc = True
End Function

' 资源检查
Public Function CheckSystemResources() As Boolean
    ' 检查系统资源
    Dim availableMemory As Long
    availableMemory = GetAvailableMemory()
    
    If availableMemory < MINIMUM_REQUIRED_MEMORY Then
        Debug.Print "系统内存不足"
        CheckSystemResources = False
        Exit Function
    End If
    
    CheckSystemResources = True
End Function
```

### 任务卡死
#### 症状
- 任务长时间处于运行状态
- 没有进度更新
- 无法取消

#### 诊断步骤
1. 检查任务状态
2. 分析任务日志
3. 检查资源使用情况

```vb
Public Sub DiagnoseHungTask(task As cThread)
    ' 收集诊断信息
    Debug.Print "任务状态诊断:"
    Debug.Print "运行时间: " & task.ExecutionTime & "ms"
    Debug.Print "当前状态: " & task.GetTaskResult("status")
    Debug.Print "最后进度更新: " & task.GetTaskResult("lastProgressUpdate")
    Debug.Print "资源使用情况:"
    
    ' 获取详细任务信息
    Dim taskInfo As Dictionary
    Set taskInfo = GetTaskDebugInfo(task)
    
    For Each key In taskInfo.Keys
        Debug.Print key & ": " & taskInfo(key)
    Next
End Sub
```

### 内存泄漏
#### 症状
- 内存使用持续增长
- 性能逐渐下降
- 任务执行变慢

#### 排查方法
1. 监控对象引用
2. 跟踪资源分配
3. 检查循环引用

```vb
' 内存使用跟踪
Public Sub TrackMemoryUsage(task As cThread)
    Static lastCheck As Date
    Static memoryReadings As Collection
    
    If memoryReadings Is Nothing Then
        Set memoryReadings = New Collection
    End If
    
    ' 每分钟记录一次内存使用
    If DateDiff("s", lastCheck, Now) >= 60 Then
        Dim usage As Dictionary
        Set usage = New Dictionary
        
        usage.Add "timestamp", Now
        usage.Add "memoryUsage", GetProcessMemoryUsage()
        usage.Add "taskObjects", CountTaskObjects()
        
        memoryReadings.Add usage
        lastCheck = Now
        
        ' 分析内存趋势
        AnalyzeMemoryTrend memoryReadings
    End If
End Sub
```

### 死锁问题
#### 症状
- 多个任务互相等待
- 系统响应变慢
- 任务执行卡住

#### 解决方案
1. 实现死锁检测
2. 使用锁超时
3. 避免嵌套锁

```vb
' 死锁检测
Public Sub DetectDeadlock(tasks As Collection)
    Dim lockGraph As Dictionary
    Set lockGraph = BuildLockGraph(tasks)
    
    If HasCycle(lockGraph) Then
        ' 发现潜在死锁
        LogDeadlockSituation lockGraph
        
        ' 尝试解决
        ResolveDeadlock tasks
    End If
End Sub

' 带超时的锁获取
Public Function TryLockWithTimeout(ByVal timeoutMs As Long) As Boolean
    Dim result As Long
    result = WaitForSingleObject(m_StateLock, timeoutMs)
    
    Select Case result
        Case WAIT_OBJECT_0
            TryLockWithTimeout = True
        Case WAIT_TIMEOUT
            ' 记录锁超时
            LogLockTimeout
            TryLockWithTimeout = False
        Case Else
            ' 处理其他错误
            HandleLockError result
            TryLockWithTimeout = False
    End Select
End Function
```

## 2. 性能问题诊断

### 性能监控
```vb
' 性能监控包装器
Public Function MonitorTaskPerformance(task As cThread) As Dictionary
    Dim metrics As New Dictionary
    
    ' 收集基本指标
    metrics.Add "executionTime", task.ExecutionTime
    metrics.Add "cpuUsage", GetTaskCPUUsage(task.ThreadID)
    metrics.Add "memoryUsage", GetTaskMemoryUsage(task.ThreadID)
    
    ' 分析性能瓶颈
    AnalyzePerformanceBottlenecks metrics
    
    Set MonitorTaskPerformance = metrics
End Function
```

### 性能优化建议
1. 使用性能分析工具
2. 优化数据结构使用
3. 减少锁竞争
4. 优化资源使用

## 3. 日志分析

### 日志级别
```vb
Public Enum LogLevel
    LogLevel_Error = 1
    LogLevel_Warning = 2
    LogLevel_Info = 3
    LogLevel_Debug = 4
End Enum

' 增强的日志记录
Public Sub LogTaskEvent(task As cThread, level As LogLevel, message As String)
    If level <= GetCurrentLogLevel() Then
        Dim logEntry As String
        logEntry = BuildLogEntry(task, level, message)
        WriteToLog logEntry
        
        ' 对于错误级别，额外记录诊断信息
        If level = LogLevel_Error Then
            LogDiagnosticInfo task
        End If
    End If
End Sub
```

### 日志分析工具
```vb
' 日志分析器
Public Function AnalyzeTaskLogs(task As cThread) As Dictionary
    Dim analysis As New Dictionary
    
    ' 获取任务日志
    Dim logs As Object
    Set logs = task.GetTaskResult("eventLog")
    
    ' 分析错误模式
    analysis.Add "errorPatterns", AnalyzeErrorPatterns(logs)
    
    ' 分析性能模式
    analysis.Add "performanceMetrics", AnalyzePerformanceMetrics(logs)
    
    ' 生成建议
    analysis.Add "recommendations", GenerateRecommendations(analysis)
    
    Set AnalyzeTaskLogs = analysis
End Function
```

## 4. 调试技巧

### 调试辅助函数
```vb
' 任务状态快照
Public Function CaptureTaskSnapshot(task As cThread) As Dictionary
    Dim snapshot As New Dictionary
    
    ' 基本信息
    snapshot.Add "threadId", task.ThreadID
    snapshot.Add "status", task.GetTaskResult("status")
    snapshot.Add "runtime", task.ExecutionTime
    
    ' 任务数据
    Dim taskData As New Dictionary
    For Each key In task.GetTaskDataKeys
        taskData.Add key, task.GetTaskData(key)
    Next
    snapshot.Add "taskData", taskData
    
    ' 任务结果
    Dim taskResults As New Dictionary
    For Each key In task.GetTaskResultKeys
        taskResults.Add key, task.GetTaskResult(key)
    Next
    snapshot.Add "taskResults", taskResults
    
    Set CaptureTaskSnapshot = snapshot
End Function

' 调试点
Public Sub DebugCheckpoint(task As cThread, checkpoint As String)
    Static checkpoints As Dictionary
    If checkpoints Is Nothing Then Set checkpoints = New Dictionary
    
    ' 记录检查点时间
    If Not checkpoints.Exists(checkpoint) Then
        checkpoints.Add checkpoint, Now
    End If
    
    ' 分析执行时间
    Dim timeSpent As Double
    timeSpent = DateDiff("s", checkpoints(checkpoint), Now)
    
    ' 记录检查点信息
    LogTaskEvent task, LogLevel_Debug, _
                "Checkpoint: " & checkpoint & _
                ", Time spent: " & timeSpent & "s"
End Sub
```

## 5. 最佳实践

### 预防措施
1. 实现健康检查
2. 设置合理的超时
3. 使用断言验证
4. 实现优雅降级

```vb
' 任务健康检查
Public Function CheckTaskHealth(task As cThread) As Boolean
    ' 验证基本状态
    If Not ValidateTaskState(task) Then
        LogTaskEvent task, LogLevel_Warning, "任务状态异常"
        Exit Function
    End If
    
    ' 检查资源使用
    If Not CheckResourceUsage(task) Then
        LogTaskEvent task, LogLevel_Warning, "资源使用异常"
        Exit Function
    End If
    
    ' 验证数据一致性
    If Not ValidateTaskData(task) Then
        LogTaskEvent task, LogLevel_Warning, "数据一致性检查失败"
        Exit Function
    End If
    
    CheckTaskHealth = True
End Function
```

### 应急处理
1. 实现紧急停止机制
2. 保存关键数据
3. 记录详细日志
4. 通知相关人员

```vb
' 紧急处理
Public Sub HandleEmergency(task As cThread, emergency As String)
    ' 记录紧急情况
    LogTaskEvent task, LogLevel_Error, "紧急情况: " & emergency
    
    ' 保存任务状态
    SaveTaskState task
    
    ' 通知监控系统
    NotifyMonitoringSystem task, emergency
    
    ' 执行紧急清理
    EmergencyCleanup task
End Sub
```

## 6. 故障恢复

### 自动恢复策略
```vb
Public Sub AttemptTaskRecovery(task As cThread)
    ' 分析故障原因
    Dim failureReason As String
    failureReason = AnalyzeFailure(task)
    
    ' 选择恢复策略
    Select Case failureReason
        Case "timeout"
            HandleTimeoutRecovery task
        Case "resource_exhaustion"
            HandleResourceRecovery task
        Case "deadlock"
            HandleDeadlockRecovery task
        Case Else
            HandleGenericRecovery task
    End Select
End Sub
```

### 数据恢复
```vb
' 任务状态恢复
Public Function RestoreTaskState(task As cThread) As Boolean
    ' 尝试从备份恢复
    If RestoreFromBackup(task) Then
        LogTaskEvent task, LogLevel_Info, "已从备份恢复"
        RestoreTaskState = True
        Exit Function
    End If
    
    ' 尝试重建状态
    If RebuildTaskState(task) Then
        LogTaskEvent task, LogLevel_Info, "已重建任务状态"
        RestoreTaskState = True
        Exit Function
    End If
    
    RestoreTaskState = False
End Function
```
