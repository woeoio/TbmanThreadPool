# 线程工作状态使用指南

## 概述

线程池现在支持**自动的**工作状态跟踪，能够区分线程的运行状态和实际工作状态。系统会自动管理状态转换，无需手动调用状态设置方法。

## 工作状态枚举

```vb
Public Enum eThreadWorkState
    ThreadWork_Idle = 0      ' 空闲 - 线程运行中但未执行任务
    ThreadWork_Busy = 1      ' 忙碌 - 线程正在执行任务
    ThreadWork_Paused = 2    ' 暂停 - 线程被暂停
    ThreadWork_Stopped = 3   ' 停止 - 线程已停止
    ThreadWork_Error = 4     ' 错误 - 线程执行出错
    ThreadWork_Completed = 5 ' 完成 - 任务已完成
End Enum
```

## 核心区别

### 运行状态 vs 工作状态

- **IsRunning**: 表示线程是否在运行（操作系统级别）
- **IsBusy**: 表示线程是否正在执行任务
- **IsIdle**: 表示线程运行中但处于空闲状态

## 🎉 自动状态管理

### 新特性：无需手动管理状态！

从现在开始，工作状态完全自动管理：

- ✅ **任务开始时**：自动设置为 `ThreadWork_Busy`
- ✅ **任务正常完成时**：自动设置为 `ThreadWork_Completed`  
- ✅ **任务发生错误时**：自动设置为 `ThreadWork_Error`

## 使用方法

### 1. 简化的任务函数（推荐）

```vb
Public Function MyTask(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    ' 无需手动调用 task.SetBusy - 系统自动处理！
    
    ' 执行任务逻辑
    For i = 1 To 100
        ' 检查取消请求
        If task.CancelRequested Then
            MyTask = -1  ' 返回错误代码，状态自动设置为 Error
            Exit Function
        End If
        
        ' 执行实际工作
        Sleep 10
        
        ' 更新进度
        task.NotifyProgress i, 100, "Processing item " & i
    Next
    
    ' 无需手动调用 task.SetCompleted - 系统自动处理！
    MyTask = 0  ' 返回成功代码，状态自动设置为 Completed
End Function
```

### 2. 错误处理（自动管理）

```vb
Public Function TaskWithErrorHandling(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    ' 状态自动设置为 Busy
    
    On Error GoTo ErrorHandler
    
    ' 执行可能出错的操作
    ' ... 任务逻辑 ...
    
    ' 正常完成 - 状态自动设置为 Completed
    TaskWithErrorHandling = 0
    Exit Function
    
ErrorHandler:
    ' 发生错误 - 状态自动设置为 Error
    TaskWithErrorHandling = -1
End Function
```

### 3. 线程池状态监控

```vb
Public Sub MonitorThreadPool(pool As cThreadPool)
    Do While Not pool.IsAllTasksCompleted
        Dim runningCount As Long, busyCount As Long, idleCount As Long
        
        ' 统计不同状态的线程数
        Dim thread As cThread
        For Each thread In pool.Tasks
            If thread.IsRunning Then
                runningCount = runningCount + 1
                If thread.IsBusy Then
                    busyCount = busyCount + 1
                ElseIf thread.IsIdle Then
                    idleCount = idleCount + 1
                End If
            End If
        Next
        
        Debug.Print "Running: " & runningCount & _
                   ", Busy: " & busyCount & _
                   ", Idle: " & idleCount & _
                   ", Queued: " & pool.QueuedTasks
        
        Sleep 1000
    Loop
End Sub
```

### 4. 详细状态显示

```vb
Public Sub ShowDetailedThreadStatus(pool As cThreadPool)
    Dim i As Long
    Dim thread As cThread
    
    i = 1
    For Each thread In pool.Tasks
        Debug.Print "Thread " & i & ":" & _
                   " Running=" & thread.IsRunning & _
                   ", Busy=" & thread.IsBusy & _
                   ", Idle=" & thread.IsIdle & _
                   ", State=" & thread.GetWorkStateText()
        i = i + 1
    Next
End Sub
```

## 线程池属性

### 精确的任务计数

- **RunningTasks**: 实际正在工作的线程数（IsBusy = True）
- **QueuedTasks**: 等待执行的任务数
- **CompletedTasks**: 已完成的任务数
- **TotalTasks**: 总任务数

### 使用示例

```vb
Private Sub UpdateProgressDisplay()
    LabelStatus.Caption = "Progress: " & _
        Format$(m_ThreadPool.CompletionPercentage, "0.0") & "% (" & _
        m_ThreadPool.RunningTasks & " working, " & _
        GetIdleThreadCount() & " idle, " & _
        m_ThreadPool.QueuedTasks & " queued)"
End Sub

Private Function GetIdleThreadCount() As Long
    Dim count As Long
    Dim thread As cThread
    
    For Each thread In m_ThreadPool.Tasks
        If thread.IsRunning And thread.IsIdle Then
            count = count + 1
        End If
    Next
    
    GetIdleThreadCount = count
End Function
```

## 最佳实践

### 1. 🎉 简化的开发模式

**新版本（自动管理）**：
- ✅ 只需专注业务逻辑
- ✅ 正常返回 0 表示成功
- ✅ 返回非零值表示错误
- ✅ 状态完全自动管理

**旧版本（手动管理）**：
- ❌ 需要记住调用 `task.SetBusy`
- ❌ 需要记住调用 `task.SetCompleted` 
- ❌ 容易忘记状态设置导致计数错误

### 2. 监控建议

- 使用 `RunningTasks` 获取精确的工作线程数
- 结合 `IsBusy` 和 `IsIdle` 属性进行详细状态检查
- 使用 `GetWorkStateText()` 方法获取易读的状态描述

### 3. 向后兼容性

手动状态设置方法仍然可用，但不再必需：
- `SetBusy()` - 可选，系统会自动调用
- `SetCompleted()` - 可选，系统会自动调用
- `SetError()` - 可选，系统会自动调用

### 3. 调试技巧

```vb
' 记录状态变化
Debug.Print "Task state changed: " & thread.GetWorkStateText()

' 监控状态转换
Private Sub LogStateTransition(thread As cThread, action As String)
    WriteLog action & " - Thread state: " & thread.GetWorkStateText() & _
            ", Running: " & thread.IsRunning & _
            ", Busy: " & thread.IsBusy
End Sub
```

## 注意事项

1. **自动化管理**: 工作状态现在完全自动管理，无需手动调用状态设置方法
2. **线程安全**: 所有工作状态操作都是线程安全的
3. **性能**: 状态检查操作是轻量级的，可以频繁调用
4. **兼容性**: 现有的 `IsRunning` 属性保持不变，新功能是附加的
5. **返回值**: 任务函数应返回 0 表示成功，非零值表示错误

## 故障排除

### 问题：任务完成但RunningTasks没有减少
**解决**: ✅ 已解决！现在自动管理状态

### 问题：线程显示为运行但不工作  
**解决**: ✅ 已解决！现在可以精确区分运行和工作状态

### 问题：错误任务没有正确清理
**解决**: ✅ 已解决！错误状态自动设置
