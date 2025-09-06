Module mCallbackDemo

    ' 定义示例数据类型
    Private Type TaskData
        TaskId As Long
        Data As String
    End Type

    ' 完成回调函数
    ' 参数说明：
    ' taskPtr: 任务对象的指针
    ' result: 任务返回值
    ' wParam: 预留参数
    ' lParam: 预留参数
    Public Function TaskCompletedCallback(ByVal taskPtr As LongPtr, _
                                        ByVal result As Long, _
                                        ByVal wParam As LongPtr, _
                                        ByVal lParam As LongPtr) As Long
        ' 获取任务对象
        Dim task As cThread
        Set task = mThread.ObjectFromPtr(taskPtr)
        
        ' 获取任务相关数据
        Dim taskData As TaskData
        CopyMemory taskData, ByVal task.Tag, LenB(task.Tag)
        
        ' 处理完成事件
        Debug.Print "任务完成: ID=" & taskData.TaskId & _
                   ", 结果=" & result & _
                   ", 执行时间=" & Format$(task.ExecutionTime, "0.00") & "ms"
        
        TaskCompletedCallback = 0
    End Function
    
    ' 错误回调函数
    ' 参数说明：
    ' taskPtr: 任务对象的指针
    ' errorCode: 错误代码
    ' wParam: 预留参数
    ' lParam: 预留参数
    Public Function TaskErrorCallback(ByVal taskPtr As LongPtr, _
                                    ByVal errorCode As Long, _
                                    ByVal wParam As LongPtr, _
                                    ByVal lParam As LongPtr) As Long
        ' 获取任务对象
        Dim task As cThread
        Set task = mThread.ObjectFromPtr(taskPtr)
        
        ' 获取任务相关数据
        Dim taskData As TaskData
        CopyMemory taskData, ByVal task.Tag, LenB(task.Tag)
        
        ' 处理错误事件
        Debug.Print "任务错误: ID=" & taskData.TaskId & _
                   ", 错误代码=" & errorCode & _
                   ", 错误描述=" & task.LastError
        
        TaskErrorCallback = 0
    End Function
    
    ' 示例任务过程
    Public Function SampleTaskProc(ByVal param As LongPtr) As Long
        Dim task As cThread
        Set task = mThread.ObjectFromPtr(param)
        
        ' 模拟一些工作
        Dim i As Long
        For i = 1 To 10
            ' 检查是否请求取消
            If task.CancelRequested Then
                SampleTaskProc = 1
                Exit Function
            End If
            
            ' 模拟工作负载
            Sleep 100
        Next
        
        ' 模拟随机错误
        If Rnd > 0.7 Then
            Err.Raise vbObjectError + 1001, "SampleTaskProc", "随机错误"
        End If
        
        SampleTaskProc = 0
    End Function
    
    ' 演示如何使用回调的示例
    Public Sub CallbackDemo()
        Randomize Timer
        
        ' 创建线程池
        Dim pool As New cThreadPool
        pool.Create 4
        
        ' 准备任务数据
        Dim taskData As TaskData
        taskData.TaskId = 1
        taskData.Data = "示例数据"
        
        ' 创建任务
        Dim task As cThread
        Set task = pool.AddTask(AddressOf SampleTaskProc, taskData)
        
        ' 设置回调
        task.SetOnComplete AddressOf TaskCompletedCallback
        task.SetOnError AddressOf TaskErrorCallback
        
        ' 设置超时和重试
        task.SetTimeout 5000  ' 5秒超时
        task.SetRetryPolicy 3, 1000  ' 最多重试3次，间隔1秒
        
        ' 等待任务完成
        pool.WaitForAll
        
        ' 关闭线程池
        pool.Shutdown
    End Sub

End Module
