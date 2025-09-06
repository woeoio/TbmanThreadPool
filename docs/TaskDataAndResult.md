# 任务数据与结果管理

## 概述
ThreadPool 库提供了灵活的任务数据传递和结果存储机制。通过 `TaskData` 和 `TaskResult` 集合，我们可以在任务执行前、执行中和执行后方便地传递和存储各种类型的数据。

## 主要特点

1. **类型安全**
   - 支持任意类型的数据，包括基本类型和对象
   - 自动处理对象和非对象类型的赋值

2. **命名参数**
   - 使用字符串键访问数据
   - 比传统的 `Tag` 属性更清晰和结构化
   - 支持多个参数和结果的同时存储

3. **线程安全**
   - 所有数据访问都通过互斥锁保护
   - 确保在多线程环境下的数据一致性

4. **灵活性**
   - 支持动态添加和修改数据
   - 可在任务执行过程中随时更新结果
   - 支持复杂数据结构的传递

## API 参考

### TaskData 相关方法

```vb
' 设置任务输入参数
Public Sub SetTaskData(ByVal Key As String, ByVal Value As Variant)

' 获取任务输入参数
Public Function GetTaskData(ByVal Key As String) As Variant

' 清除所有任务数据
Public Sub ClearTaskData()

' 获取所有任务数据键
Public Function GetTaskDataKeys() As Collection
```

### TaskResult 相关方法

```vb
' 设置任务结果
Public Sub SetTaskResult(ByVal Key As String, ByVal Value As Variant)

' 获取任务结果
Public Function GetTaskResult(ByVal Key As String) As Variant

' 清除所有任务结果
Public Sub ClearTaskResults()

' 获取所有任务结果键
Public Function GetTaskResultKeys() As Collection
```

## 使用示例

### 基本使用

```vb
' 创建任务并设置输入参数
Dim task As New cThread
With task
    .SetTaskData "url", "http://example.com"
    .SetTaskData "timeout", 5000
    .SetTaskData "headers", CreateObject("Scripting.Dictionary")
End With

' 启动任务
task.Create AddressOf DownloadTaskProc
```

### 在任务过程中使用

```vb
Public Function DownloadTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mTask.ObjectFromPtr(param)
    
    ' 获取输入参数
    Dim url As String
    url = task.GetTaskData("url")
    Dim timeout As Long
    timeout = task.GetTaskData("timeout")
    
    ' 执行下载...
    
    ' 设置结果
    task.SetTaskResult "statusCode", 200
    task.SetTaskResult "content", content
    task.SetTaskResult "responseTime", Timer
    task.SetTaskResult "headers", responseHeaders
    
    DownloadTaskProc = 0
End Function
```

### 在回调中处理结果

```vb
Public Function TaskCompletedCallback(ByVal taskPtr As LongPtr, _
                                    ByVal result As Long, _
                                    ByVal wParam As LongPtr, _
                                    ByVal lParam As LongPtr) As Long
    Dim task As cThread
    Set task = mTask.ObjectFromPtr(taskPtr)
    
    ' 获取结果
    Dim statusCode As Long
    statusCode = task.GetTaskResult("statusCode")
    Dim content As String
    content = task.GetTaskResult("content")
    
    ' 处理结果...
    Debug.Print "下载完成: 状态码=" & statusCode & _
                ", 内容长度=" & Len(content)
    
    TaskCompletedCallback = 0
End Function
```

## 最佳实践

1. **参数命名规范**
   - 使用清晰、描述性的键名
   - 建议使用驼峰命名法
   - 相关参数可使用共同前缀

2. **错误处理**
   - 总是检查 GetTaskData/GetTaskResult 的返回值
   - 使用 On Error 处理可能的异常

```vb
On Error Resume Next
Dim value As Variant
value = task.GetTaskData("someKey")
If Err.Number <> 0 Then
    ' 处理错误...
    Exit Function
End If
On Error GoTo 0
```

3. **资源管理**
   - 及时清理不再需要的数据（使用 ClearTaskData/ClearTaskResults）
   - 对于大型对象，在使用完后设置为 Nothing

4. **并发考虑**
   - 避免过长时间的数据访问操作
   - 考虑在关键部分使用额外的同步机制

## 应用场景

1. **Web 请求处理**
   - 传递 URL、headers、请求参数
   - 存储响应状态、内容、元数据

2. **文件操作**
   - 传递文件路径、操作选项
   - 存储处理结果、统计信息

3. **数据处理**
   - 传递数据集、处理参数
   - 存储处理结果、统计数据

4. **多步骤任务**
   - 在不同步骤间传递数据
   - 累积中间结果

## 注意事项

1. 避免存储过大的数据对象，可能影响性能
2. 注意及时清理不再需要的数据
3. 在高并发场景下注意锁的使用
4. 建议为关键数据提供默认值
5. 定期检查和维护数据一致性

## 总结

TaskData 和 TaskResult 机制提供了一个强大而灵活的方式来处理任务相关的数据传递和结果存储。通过合理使用这个机制，可以显著提高代码的可维护性和可扩展性。
