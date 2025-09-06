# ThreadPool 性能优化指南

## 1. 数据传递优化

### 大数据处理
- 对于大型数据集，使用引用而不是值传递
- 考虑使用文件或数据库作为中间存储
- 使用数据分块处理

```vb
' 推荐：使用引用传递大数据
Dim largeData As Dictionary
Set largeData = New Dictionary
' ... 填充数据 ...
task.SetTaskData "data", largeData  ' 引用传递

' 不推荐：直接传递大字符串或数组
task.SetTaskData "data", largeString  ' 值传递可能导致性能问题
```

### 内存管理
- 及时清理不再需要的数据
- 使用数据流而不是一次性加载
- 实现自定义的内存池机制

```vb
' 使用后清理大型对象
Public Sub CleanupTask(task As cThread)
    Dim obj As Object
    Set obj = task.GetTaskResult("largeObject")
    If Not obj Is Nothing Then
        ' 清理对象资源
        obj.Cleanup  ' 假设对象有清理方法
        Set obj = Nothing
    End If
    task.ClearTaskResults
End Sub
```

## 2. 并发优化

### 任务分割
- 将大任务分割成小任务
- 使用适当的任务粒度
- 实现任务队列优先级

```vb
' 任务分割示例
Public Function SplitTask(data As Variant, chunkSize As Long) As Collection
    Dim tasks As New Collection
    Dim i As Long
    
    For i = 1 To Length(data) Step chunkSize
        Dim task As New cThread
        With task
            .SetTaskData "startIndex", i
            .SetTaskData "endIndex", Min(i + chunkSize - 1, Length(data))
            .SetTaskData "data", GetDataChunk(data, i, chunkSize)
            .Priority = CalculatePriority(i)
        End With
        tasks.Add task
    Next i
    
    Set SplitTask = tasks
End Function
```

### 资源控制
- 实现资源池
- 控制并发任务数
- 监控资源使用情况

```vb
' 资源池实现
Private Type ResourcePool
    Available As Collection
    InUse As Collection
    MaxSize As Long
End Type

Private Function AcquireResource(pool As ResourcePool) As Variant
    If pool.Available.Count > 0 Then
        Set AcquireResource = pool.Available(1)
        pool.Available.Remove 1
        pool.InUse.Add AcquireResource
    End If
End Function
```

## 3. 锁优化

### 细粒度锁定
- 缩小临界区范围
- 使用多个专用锁
- 避免锁的嵌套

```vb
' 改进前：粗粒度锁定
Public Sub CoarseGrainedLock()
    LockState
    ' 执行多个操作...
    UnlockState
End Sub

' 改进后：细粒度锁定
Public Sub FineGrainedLock()
    ' 只锁定必要的操作
    LockState
    Dim value As Variant
    value = m_TaskData("key")
    UnlockState
    
    ' 执行不需要锁定的操作...
    
    LockState
    m_TaskResult.Add result, "key"
    UnlockState
End Sub
```

### 读写分离
- 实现读写锁
- 优化频繁读取操作
- 批量更新机制

```vb
Private Type ReadWriteLock
    ReadCount As Long
    WriteCount As Long
    ReadMutex As LongPtr
    WriteMutex As LongPtr
End Type

' 读锁实现
Private Sub AcquireReadLock(lock As ReadWriteLock)
    WaitForSingleObject lock.ReadMutex, INFINITE
    lock.ReadCount = lock.ReadCount + 1
    If lock.ReadCount = 1 Then
        WaitForSingleObject lock.WriteMutex, INFINITE
    End If
    ReleaseMutex lock.ReadMutex
End Sub
```

## 4. 监控与诊断

### 性能指标收集
- 记录关键操作时间
- 跟踪资源使用情况
- 实现性能计数器

```vb
' 性能监控包装器
Private Function MeasurePerformance(task As cThread, operation As String) As Double
    Dim startTime As Currency
    QueryPerformanceCounter startTime
    
    ' 执行操作...
    
    Dim endTime As Currency
    QueryPerformanceCounter endTime
    
    Dim freq As Currency
    QueryPerformanceFrequency freq
    
    MeasurePerformance = (endTime - startTime) / freq * 1000
    
    ' 记录性能数据
    task.SetTaskResult "perf_" & operation, MeasurePerformance
End Function
```

### 诊断信息
- 实现详细日志记录
- 添加性能跟踪点
- 导出诊断数据

```vb
' 诊断日志记录
Private Sub LogDiagnostics(task As cThread, category As String, message As String)
    Dim logEntry As String
    logEntry = Format$(Now, "yyyy-mm-dd hh:nn:ss.000") & " | " & _
               category & " | " & _
               "Task[" & task.ThreadID & "] | " & _
               message
               
    ' 添加详细信息
    If task.GetTaskResult("diagnostics") = Empty Then
        task.SetTaskResult "diagnostics", CreateObject("Scripting.Dictionary")
    End If
    
    Dim diag As Object
    Set diag = task.GetTaskResult("diagnostics")
    diag.Add Now, logEntry
End Sub
```

## 5. 最佳实践建议

### 代码优化
1. 预分配集合大小
2. 使用适当的数据结构
3. 避免不必要的对象创建
4. 实现对象池模式

### 并发控制
1. 避免长时间持有锁
2. 使用非阻塞算法
3. 实现任务超时机制
4. 合理设置任务优先级

### 资源管理
1. 及时释放资源
2. 实现资源限制
3. 监控资源泄漏
4. 定期清理缓存

### 错误处理
1. 实现优雅降级
2. 添加重试机制
3. 记录详细错误信息
4. 实现自动恢复机制
