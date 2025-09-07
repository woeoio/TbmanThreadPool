[Description("")]
[FormDesignerId("71B7D6D9-7FFB-4151-A8CD-3D6FAF10C90E")]
[PredeclaredId]
Class Form1

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Document: https://doc.twinbasic.vb6.pro/en/tbman/threadPool/

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