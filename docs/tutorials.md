# 使用教程

## 基础使用

### 1. 创建线程池

最基本的线程池创建方式：

```vb
' 创建一个包含4个线程的池
Dim pool As New cThreadPool
pool.Create 4
```

### 2. 添加简单任务

```vb
' 定义任务过程
Public Function MyProc(ByVal param As Variant) As Long
    ' 执行一些操作
    Debug.Print "处理任务：" & param
    Sleep 1000  ' 模拟耗时操作
    MyProc = 0  ' 返回成功
End Function

' 添加任务到线程池
Dim task As cThread
Set task = pool.AddTask(AddressOf MyProc, "任务1")

' 等待任务完成
task.WaitForCompletion
```

### 3. 使用任务回调

```vb
' 完成回调
Public Function TaskCompleted(ByVal taskPtr As LongPtr, _
                            ByVal result As Long, _
                            ByVal wParam As LongPtr, _
                            ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(taskPtr)
    Debug.Print "任务完成，结果：" & result
    TaskCompleted = 0
End Function

' 错误回调
Public Function TaskError(ByVal taskPtr As LongPtr, _
                         ByVal errorCode As Long, _
                         ByVal wParam As LongPtr, _
                         ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(taskPtr)
    Debug.Print "任务错误：" & errorCode
    TaskError = 0
End Function

' 设置回调
Dim task As cThread
Set task = pool.AddTask(AddressOf MyProc)
task.SetOnComplete AddressOf TaskCompleted
task.SetOnError AddressOf TaskError
```

## 高级使用

### 1. 任务优先级

```vb
' 添加高优先级任务
Dim highTask As cThread
Set highTask = pool.AddTask(AddressOf MyProc, "高优先级", Priority_High)

' 添加普通优先级任务
Dim normalTask As cThread
Set normalTask = pool.AddTask(AddressOf MyProc, "普通优先级", Priority_Normal)
```

### 2. 任务超时控制

```vb
' 设置任务超时
Dim task As cThread
Set task = pool.AddTask(AddressOf MyProc)
task.SetTimeout 5000  ' 5秒超时

' 等待任务完成
If Not task.WaitForCompletion Then
    Debug.Print "任务超时或失败"
End If
```

### 3. 任务重试机制

```vb
' 设置重试策略
Dim task As cThread
Set task = pool.AddTask(AddressOf MyProc)
task.SetRetryPolicy 3, 1000  ' 最多重试3次，间隔1秒
```

### 4. 批量任务处理

```vb
' 添加多个任务
Dim tasks() As cThread
ReDim tasks(1 To 10)

' 创建10个任务
Dim i As Long
For i = 1 To 10
    Set tasks(i) = pool.AddTask(AddressOf MyProc, "任务" & i)
Next i

' 等待所有任务完成
pool.WaitForAll 10000  ' 最多等待10秒
```

### 5. 自动扩展线程池

```vb
' 创建可自动扩展的线程池
Dim pool As New cThreadPool
pool.Create 4
pool.SetAutoScale True, 2, 0.75  ' 启用自动扩展，最小2线程，负载因子0.75
```

## 最佳实践

### 1. 错误处理

```vb
' 任务过程中的错误处理
Public Function MyProc(ByVal param As Variant) As Long
    On Error GoTo ErrorHandler
    
    ' 执行任务...
    
    MyProc = 0  ' 成功
    Exit Function
    
ErrorHandler:
    MyProc = Err.Number  ' 返回错误码
End Function
```

### 2. 取消支持

```vb
' 支持取消的任务过程
Public Function CancellableProc(ByVal param As Variant) As Long
    Dim task As cThread
    Set task = mThread.ObjectFromPtr(param)
    
    Do
        ' 检查是否请求取消
        If task.CancelRequested Then
            CancellableProc = TaskResult_Cancelled
            Exit Function
        End If
        
        ' 执行一些操作...
        Sleep 100
    Loop
    
    CancellableProc = 0
End Function
```

### 3. 资源清理

```vb
' 优雅关闭线程池
Private Sub Class_Terminate()
    If Not pool Is Nothing Then
        pool.Shutdown True  ' 等待任务完成
        Set pool = Nothing
    End If
End Sub
```

### 4. 性能优化

1. 合理设置线程数
   - 通常设置为CPU核心数的1-2倍
   - 考虑任务类型（CPU密集或IO密集）

2. 避免过小的任务
   - 任务执行时间应显著大于线程调度开销
   - 考虑合并小任务

3. 使用任务批处理
   - 批量添加任务而不是逐个添加
   - 使用WaitForAll而不是逐个等待

4. 合理使用超时
   - 为所有任务设置合理的超时时间
   - 避免无限等待

## 注意事项

1. 线程安全
   - 任务过程应该是线程安全的
   - 共享资源需要适当的同步
   
2. 资源管理
   - 正确释放资源
   - 避免资源泄露
   
3. 错误处理
   - 使用错误回调
   - 记录详细错误信息
   
4. UI更新
   - 不要在任务线程直接更新UI
   - 使用PostMessage或类似机制
