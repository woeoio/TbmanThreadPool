[Description("")]
[FormDesignerId("71B7D6D9-7FFB-4151-A8CD-3D6FAF10C90E")]
[PredeclaredId]
Class Form1
    ' 线程过程示例
    Function ThreadProc(ByVal param As LongPtr) As Long
        Dim t As cThread
        Set t = mTask.ObjectFromPtr(param)
        
        ' 一次性获取所有需要的信息，减少锁竞争
        Dim threadHandle As LongPtr
        Dim threadID As Long
        Dim tag As Variant
        
        ' 获取信息（这会触发锁）
        threadHandle = t.ThreadHandle
        threadID = t.ThreadID
        tag = t.Tag
        
        ' 执行工作循环（不再访问需要锁的属性）
        Dim i As Long
        For i = 1 To 5
            Debug.Print i, threadHandle, threadID, tag
            
            ' 只在需要时检查取消状态
            If i Mod 2 = 0 Then  ' 减少检查频率
                If t.CancelRequested Then Exit For
            End If
            
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
        pool.WaitForAll 5000  ' 等待5秒
        ' 确保任务完全清理
        Set pool = Nothing
        Set Task = Nothing
        Debug.Print "all done"
    End Sub
End Class