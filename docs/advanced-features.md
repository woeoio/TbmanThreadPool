# 高级特性

## 任务优先级管理

### 优先级机制

线程池支持以下优先级级别：

```vb
Public Enum TaskPriority
    Priority_Low = 0
    Priority_Normal = 1
    Priority_High = 2
    Priority_Critical = 3
End Enum
```

### 使用优先级

1. 创建任务时指定：
```vb
' 创建高优先级任务
Dim task As cThread
Set task = pool.AddTask(AddressOf MyProc, param, Priority_High)
```

2. 动态修改优先级：
```vb
' 修改已有任务的优先级
task.Priority = Priority_Critical
```

### 优先级调度策略

- 高优先级任务优先执行
- 同优先级任务按FIFO顺序执行
- 新加入的高优先级任务可以抢占等待队列中的低优先级任务

## 错误处理机制

### 错误状态

任务可能的错误状态：
```vb
Public Enum TaskResult
    TaskResult_Success = 0
    TaskResult_Failed = 1
    TaskResult_Cancelled = 2
    TaskResult_TimedOut = 3
    TaskResult_Exception = 4
End Enum
```

### 错误追踪

每个任务都维护详细的错误信息：
```vb
' 获取错误信息
Debug.Print "错误代码：" & task.LastError
Debug.Print "错误描述：" & task.ErrorDescription
```

### 错误回调

```vb
' 错误处理回调
Public Function OnError(ByVal taskPtr As LongPtr, _
                       ByVal errorCode As Long, _
                       ByVal wParam As LongPtr, _
                       ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(taskPtr)
    
    ' 记录错误
    Debug.Print "任务失败：" & task.ErrorDescription
    
    ' 决定是否重试
    If errorCode = SomeRetryableError Then
        OnError = 1  ' 返回1表示需要重试
    Else
        OnError = 0  ' 返回0表示不再重试
    End If
End Function
```

## 性能监控

### 统计信息

线程池提供详细的性能统计：

```vb
' 获取线程池统计信息
Debug.Print "总线程数：" & pool.ThreadCount
Debug.Print "活动线程数：" & pool.ActiveThreadCount
Debug.Print "等待任务数：" & pool.PendingTaskCount
Debug.Print "已完成任务数：" & pool.CompletedTaskCount
Debug.Print "平均执行时间：" & pool.AverageTaskTime & "ms"
```

### 性能日志

```vb
' 启用性能日志
pool.EnablePerformanceLog "C:\threadpool.log"

' 日志内容示例：
' 2023-09-05 10:30:15 任务开始 ID=1
' 2023-09-05 10:30:16 任务完成 ID=1 执行时间=1023ms
' 2023-09-05 10:30:16 线程池状态：活动=3/4 等待=5
```

## 自动负载均衡

### 自动扩展配置

```vb
' 配置自动扩展
pool.SetAutoScale True, _
    minThreads:=2, _       ' 最小线程数
    maxThreads:=8, _       ' 最大线程数
    loadFactor:=0.75       ' 负载因子
```

### 扩展触发条件

- 当前负载 = 活动线程数 / 总线程数
- 如果当前负载 > loadFactor，尝试增加线程
- 如果当前负载 < loadFactor/2，尝试减少线程
- 始终保持最小线程数

### 负载监控

```vb
' 获取当前负载情况
Dim load As Double
load = pool.GetCurrentLoad  ' 返回0-1之间的值

' 根据负载调整策略
If load > 0.9 Then
    ' 负载过高，可以临时提高线程数
    pool.SetMaxThreads pool.ThreadCount + 2
End If
```

## 任务依赖管理

### 设置任务依赖

```vb
' 创建依赖任务
Dim task1 As cThread, task2 As cThread
Set task1 = pool.AddTask(AddressOf Proc1)
Set task2 = pool.AddTask(AddressOf Proc2)

' 设置依赖关系
task2.AddDependency task1  ' task2 依赖于 task1
```

### 依赖执行规则

- 依赖任务必须等待前置任务完成
- 如果前置任务失败，依赖任务不会执行
- 支持多个依赖（AND关系）

### 依赖链示例

```vb
' 创建任务链
Dim task1 As cThread, task2 As cThread, task3 As cThread
Set task1 = pool.AddTask(AddressOf Step1)
Set task2 = pool.AddTask(AddressOf Step2)
Set task3 = pool.AddTask(AddressOf Step3)

' 设置依赖链
task2.AddDependency task1
task3.AddDependency task2

' 启动任务链
task1.Start
```

## 资源管理

### 线程池资源控制

```vb
' 设置线程池资源限制
pool.SetResourceLimits _
    maxMemory:=1024 * 1024 * 100, _  ' 100MB内存限制
    maxCpu:=75                        ' 最高75%CPU使用率
```

### 资源监控

```vb
' 监控资源使用
Dim memUsage As Long
memUsage = pool.GetMemoryUsage  ' 获取当前内存使用量

' 根据资源使用情况调整策略
If memUsage > pool.MaxMemory * 0.9 Then
    ' 内存使用接近限制，采取措施
    pool.PauseNewTasks  ' 暂停接受新任务
End If
```

## 任务取消和超时

### 取消机制

```vb
' 支持取消的任务过程
Public Function CancellableProc(ByVal param As Variant) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(param)
    
    Do
        If task.CancelRequested Then
            ' 清理资源
            CleanupResources
            CancellableProc = TaskResult_Cancelled
            Exit Function
        End If
        
        ' 执行任务步骤...
        
    Loop
End Function

' 请求取消任务
task.RequestCancel
```

### 超时处理

```vb
' 设置任务超时
task.SetTimeout 5000  ' 5秒超时

' 超时回调
Public Function OnTimeout(ByVal taskPtr As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(taskPtr)
    
    ' 执行超时清理
    CleanupResources
    
    OnTimeout = TaskResult_TimedOut
End Function
```

## 日志和调试

### 启用调试日志

```vb
' 配置日志
pool.SetLogLevel LogLevel_Debug
pool.SetLogFile "C:\threadpool_debug.log"
```

### 日志回调

```vb
' 自定义日志处理
Public Function LogHandler(ByVal level As Long, _
                         ByVal message As String) As Long
    ' 处理日志消息
    Select Case level
        Case LogLevel_Error
            LogToEventLog message
        Case LogLevel_Debug
            Debug.Print message
    End Select
End Function

' 设置日志回调
pool.SetLogCallback AddressOf LogHandler
```
