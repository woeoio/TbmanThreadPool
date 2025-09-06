# ThreadPool 详细用例解析

## 概述

本文档基于 `Form1.frm` 中的测试用例，详细解析线程池的使用方法和最佳实践。该用例展示了如何正确地创建线程池、添加任务、处理取消请求以及优化性能。

## 完整用例代码

```vb
[Description("")]
[FormDesignerId("71B7D6D9-7FFB-4151-A8CD-3D6FAF10C90E")]
[PredeclaredId]
Class Form1
    ' 线程过程示例
    Function ThreadProc(ByVal param As LongPtr) As Long
        Dim t As cThread '从指针还原线程对象，可以使用对象成员
        Set t = mThread.ReturnFromPtr(param)
        ' 一次性获取所有需要的信息，避免循环体内获取，减少锁竞争
        Dim threadHandle As LongPtr = t.ThreadHandle
        Dim threadID As Long = t.ThreadID
        Dim tag As Variant = t.Tag
        ' 执行工作循环（不再访问需要锁的属性）
        Dim i As Long
        For i = 1 To 5
            Debug.Print i, threadHandle, threadID, tag
            ' 只在需要时检查取消状态
            If t.CancelRequested Then Exit For
            ' 模拟一些实际工作
            Dim j As Long
            For j = 1 To 1000000: Next j  ' 简单的CPU工作
        Next
        ThreadProc = 0
    End Function
    
    Private Sub Command1_Click()
        Dim pool As New cThreadPool
        Dim Task As New cThread
        Dim i As Long
        '单个线程示例
        ' Task.Create(AddressOf ThreadProc, "dw")
        '线程池示例，创建5个工作线程
        pool.Create(5)
        ' 添加任务，支持无数个任务，线程池会自动分配工作线程去完成
        For i = 1 To 5
            Set Task = pool.AddTask(AddressOf ThreadProc, "wl-" & i)
            Task.EnableLogging(App.Path & "\trace.log")
        Next
        ' 线程是异步的，如果需要同步流程，可WaitForAll等待所有任务
        pool.WaitForAll 5000  ' 等待5秒（可选）
        ' 确保任务完全清理
        Set pool = Nothing
        Set Task = Nothing
        Debug.Print "all done"
    End Sub
End Class
```

## 代码解析

### 1. 线程过程函数 (ThreadProc)

#### 1.1 线程对象恢复
```vb
Dim t As cThread
Set t = mThread.ReturnFromPtr(param)
```
- 从传入的指针参数恢复线程对象
- 这允许在线程过程中访问线程的属性和方法

#### 1.2 性能优化：一次性获取线程信息
```vb
' 一次性获取所有需要的信息，避免循环体内获取，减少锁竞争
Dim threadHandle As LongPtr = t.ThreadHandle
Dim threadID As Long = t.ThreadID
Dim tag As Variant = t.Tag
```
**重要优化点：**
- 在循环开始前一次性获取所有需要的线程信息
- 避免在循环体内重复访问这些属性
- 减少锁竞争，提高性能
- 特别是在高频访问场景下，这种优化效果显著

#### 1.3 工作循环
```vb
Dim i As Long
For i = 1 To 5
    Debug.Print i, threadHandle, threadID, tag
    ' 只在需要时检查取消状态
    If t.CancelRequested Then Exit For
    ' 模拟一些实际工作
    Dim j As Long
    For j = 1 To 1000000: Next j  ' 简单的CPU工作
Next
```
**关键特性：**
- 每次循环都检查取消状态
- 及时响应取消请求
- 模拟实际的CPU密集型工作

### 2. 主程序逻辑 (Command1_Click)

#### 2.1 线程池初始化
```vb
Dim pool As New cThreadPool
pool.Create(5)
```
- 创建包含5个工作线程的线程池
- 工作线程数量应根据系统CPU核心数和任务特性调整

#### 2.2 任务添加
```vb
For i = 1 To 5
    Set Task = pool.AddTask(AddressOf ThreadProc, "wl-" & i)
    Task.EnableLogging(App.Path & "\trace.log")
Next
```
- 添加5个任务到线程池
- 每个任务都有唯一的标识 ("wl-1", "wl-2", 等)
- 启用日志记录到 trace.log 文件

#### 2.3 同步等待
```vb
pool.WaitForAll 5000  ' 等待5秒（可选）
```
- 等待所有任务完成，最多等待5秒
- 超时后继续执行，不会无限等待

#### 2.4 资源清理
```vb
Set pool = Nothing
Set Task = Nothing
```
- 确保正确释放线程池和任务对象
- 防止内存泄漏

## 最佳实践总结

### 1. 性能优化
- **一次性获取属性**：在循环开始前获取所有需要的线程属性
- **减少锁竞争**：避免在循环体内频繁访问需要同步的属性
- **合理的任务粒度**：既不要太细碎，也不要太大块

### 2. 取消处理
- **及时检查**：在循环中定期检查 `CancelRequested`
- **优雅退出**：使用 `Exit For` 而不是强制终止
- **状态报告**：通过返回值报告任务完成状态

### 3. 资源管理
- **正确初始化**：确保线程池正确创建
- **适当清理**：使用 `Set ... = Nothing` 释放对象
- **超时控制**：使用 `WaitForAll` 的超时参数避免无限等待

### 4. 调试和监控
- **启用日志**：使用 `EnableLogging` 记录执行过程
- **输出状态**：使用 `Debug.Print` 跟踪执行进度
- **错误处理**：在实际应用中添加适当的错误处理

## 实际应用场景

这个用例适用于以下场景：

1. **批量数据处理**：处理大量相似的数据项
2. **并行计算**：执行可以并行化的计算任务
3. **网络请求**：并发处理多个网络请求
4. **文件操作**：并行处理多个文件

## 扩展建议

### 1. 添加错误处理
```vb
Function ThreadProc(ByVal param As LongPtr) As Long
    On Error GoTo ErrorHandler
    ' ... 原有代码 ...
    ThreadProc = 0  ' 成功
    Exit Function
ErrorHandler:
    Debug.Print "线程错误: " & Err.Description
    ThreadProc = -1  ' 错误
End Function
```

### 2. 进度报告
```vb
For i = 1 To 5
    ' 计算进度百分比
    Dim progress As Long = (i * 100) \ 5
    ' 可以通过自定义属性或回调报告进度
    If t.CancelRequested Then Exit For
    ' ... 工作代码 ...
Next
```

### 3. 动态任务调整
```vb
' 根据系统性能动态调整工作线程数
Dim optimalThreadCount As Long = GetOptimalThreadCount()
pool.Create(optimalThreadCount)
```

这个用例展示了ThreadPool的核心功能和最佳实践，为实际项目开发提供了良好的参考模板。
