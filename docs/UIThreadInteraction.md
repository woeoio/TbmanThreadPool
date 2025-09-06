# UI线程交互指南

本文档介绍如何使用线程池库的UI线程交互功能，实现子线程与主UI线程的安全通信。

## 概述

新的UI通知系统解决了多线程环境中子线程无法直接操作UI控件的问题。通过消息机制，子线程可以安全地向UI线程发送通知，而UI线程可以通过事件处理来更新界面。

## 核心组件

### 1. cThreadUINotifier 类
负责管理UI线程通知的核心类，提供以下功能：
- 创建隐藏的消息接收窗口
- 处理跨线程消息传递
- 触发UI线程事件
- 支持多种通知类型

### 2. cThread 类的新方法
扩展了 cThread 类，新增了以下UI通知方法：
- `NotifyToUI()` - 发送自定义通知
- `NotifyProgress()` - 发送进度通知
- `NotifyStatus()` - 发送状态通知
- `NotifyError()` - 发送错误通知
- `NotifyCompleted()` - 发送完成通知

### 3. cThreadPool 类的UI支持
线程池类新增了UI通知器支持：
- `SetUINotifier()` - 设置UI通知器
- 自动为新任务配置UI通知器

## 使用步骤

### 第一步：在Form中声明UI通知器

```vb
' 在Form类中声明（使用WithEvents接收事件）
Private WithEvents m_UINotifier As cThreadUINotifier
Private m_ThreadPool As cThreadPool
```

### 第二步：初始化UI通知器

```vb
Private Sub Form_Load()
    ' 创建UI通知器（自动初始化）
    Set m_UINotifier = New cThreadUINotifier
    
    ' 创建线程池并设置UI通知器
    Set m_ThreadPool = New cThreadPool
    m_ThreadPool.Create 3  ' 最大3个并发线程
    m_ThreadPool.SetUINotifier m_UINotifier
End Sub
```

### 第三步：处理UI通知事件

```vb
' 处理进度通知
Private Sub m_UINotifier_ThreadProgress(ByVal TaskId As String, _
                                       ByVal CurrentValue As Long, _
                                       ByVal MaxValue As Long, _
                                       ByVal Message As String)
    ' 更新进度条
    ProgressBar1.Max = MaxValue
    ProgressBar1.Value = CurrentValue
    
    ' 更新状态标签
    Dim percent As Long
    percent = (CurrentValue * 100) \ MaxValue
    LabelStatus.Caption = "Task " & TaskId & ": " & percent & "% - " & Message
End Sub

' 处理完成通知
Private Sub m_UINotifier_ThreadCompleted(ByVal TaskId As String, _
                                        ByVal Result As Variant)
    ListBox1.AddItem "Task " & TaskId & " completed: " & CStr(Result)
End Sub

' 处理错误通知
Private Sub m_UINotifier_ThreadError(ByVal TaskId As String, _
                                   ByVal ErrorCode As Long, _
                                   ByVal ErrorMessage As String)
    ListBox1.AddItem "Task " & TaskId & " error: " & ErrorMessage
    MsgBox "Task error: " & ErrorMessage, vbExclamation
End Sub

' 处理一般通知
Private Sub m_UINotifier_ThreadNotification(ByVal TaskId As String, _
                                           ByVal NotificationCode As Long, _
                                           ByVal Data As Variant)
    Select Case NotificationCode
        Case 2  ' Status notification
            LabelStatus.Caption = CStr(Data)
        Case 101  ' Custom notification
            ListBox1.AddItem "Custom: " & CStr(Data)
        Case Else
            ListBox1.AddItem "Notification " & NotificationCode & ": " & CStr(Data)
    End Select
End Sub
```

### 第四步：在任务过程中发送通知

```vb
Public Function MyTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    ' 发送状态通知
    task.NotifyStatus "Task started"
    
    Dim totalItems As Long
    totalItems = 100
    
    Dim i As Long
    For i = 1 To totalItems
        ' 检查取消请求
        If task.CancelRequested Then
            task.NotifyStatus "Task cancelled"
            MyTaskProc = -1
            Exit Function
        End If
        
        ' 执行实际工作
        ' ... 处理逻辑 ...
        
        ' 发送进度通知
        task.NotifyProgress i, totalItems, "Processing item " & i
        
        ' 可选：发送自定义通知
        If (i Mod 10) = 0 Then
            task.NotifyToUI 101, "Checkpoint reached: " & i
        End If
    Next
    
    ' 发送完成通知
    task.NotifyCompleted "Processed " & totalItems & " items"
    
    MyTaskProc = 0
End Function
```

### 第五步：清理资源

```vb
Private Sub Form_Unload(Cancel As Integer)
    ' 线程池会在对象销毁时自动终止所有线程并清理资源
    Set m_ThreadPool = Nothing
    Set m_UINotifier = Nothing
End Sub
```

## 通知类型

### 预定义通知代码

| 代码 | 名称 | 用途 |
|------|------|------|
| 1 | Notify_Progress | 进度更新 |
| 2 | Notify_Status | 状态信息 |
| 3 | Notify_Warning | 警告信息 |
| 4 | Notify_Error | 错误信息 |
| 5 | Notify_Completed | 完成信息 |
| 100+ | Notify_Custom | 用户自定义 |

### 数据类型支持

UI通知器支持以下数据类型：
- Long/Integer - 数值类型
- String - 字符串类型  
- Double/Single - 浮点数类型
- Boolean - 布尔类型

## 最佳实践

### 1. 合理控制通知频率
```vb
' 不要每次循环都发送通知
For i = 1 To 10000
    ' 只在特定条件下发送通知
    If (i Mod 100) = 0 Then
        task.NotifyProgress i, 10000, "Processing..."
    End If
Next
```

### 2. 使用有意义的消息
```vb
' 好的做法
task.NotifyProgress 50, 100, "Downloading file 50 of 100"

' 避免的做法
task.NotifyProgress 50, 100, ""
```

### 3. 适当的错误处理
```vb
Public Function SafeTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    On Error GoTo ErrorHandler
    
    ' 任务逻辑...
    
    task.NotifyCompleted "Success"
    SafeTaskProc = 0
    Exit Function
    
ErrorHandler:
    task.NotifyError Err.Number, Err.Description
    SafeTaskProc = -1
End Function
```

### 4. 任务ID的使用
每个任务都有唯一的TaskId，可以用来：
- 区分不同任务的通知
- 在UI中显示任务状态
- 取消特定任务

## 注意事项

1. **线程安全**: NotifyToUI方法是线程安全的，可以在任何线程中调用
2. **非阻塞**: PostMessage是非阻塞调用，不会影响任务执行性能
3. **内存管理**: UI通知器会自动管理通知数据的内存，无需手动清理
4. **错误恢复**: 如果UI线程忙碌，通知可能会排队，但不会丢失

## 完整示例

参考以下文件获取完整的使用示例：
- `FormUIDemo.frm` - 完整的演示窗体
- `mUINotifierExample.bas` - 任务过程示例
- `cThreadUINotifier.cls` - UI通知器类实现
- `cThreadUINotificationData.cls` - UI通知数据封装类

通过这个UI通知系统，您可以轻松实现线程池与UI界面的安全交互，创建响应迅速且用户友好的多线程应用程序。
