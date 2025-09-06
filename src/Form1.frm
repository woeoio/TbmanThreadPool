[Description("")]
[FormDesignerId("976A6C7A-CB62-451C-A1D8-BAF38064F787")]
[PredeclaredId]
Class Form1

    Dim pool As New cTasks
    Dim Task As New cTask
    
    Sub New()
    End Sub
    
    ' 线程过程示例
    Function ThreadProc(ByVal param As LongPtr) As Long
        Dim t As cTask
        Set t = mTask.ObjectFromPtr(param)
    
        Dim i As Long
        For i = 1 To 10
            Debug.Print i, t.ThreadHandle, t.Tag
            If t.CancelRequested Then Exit For
            ' 模拟工作
            ' Sleep 50
        Next
    
        ThreadProc = 0
    End Function
    
    Private Sub Form_Load()
        
        Task.Create(AddressOf ThreadProc, "dw")
        
        pool.Create(10)

        ' ' 添加任务

        pool.AddTask(AddressOf ThreadProc)

        ' ' 等待所有任务完成（超时5秒后取消所有任务）
        pool.WaitForAll 5000

        ' ' 不需要用户自己定义任何枚举！
    End Sub
    
End Class