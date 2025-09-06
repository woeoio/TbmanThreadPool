Module mUINotifierExample

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Description "UI通知器使用示例"

    ' 示例：在Form中使用UI通知器
    ' 
    ' 1. 在Form中声明UI通知器：
    ' Private WithEvents m_UINotifier As cThreadUINotifier
    ' 
    ' 2. 在Form_Load中初始化：
    ' Set m_UINotifier = New cThreadUINotifier
    ' 
    ' 3. 为线程池设置UI通知器：
    ' Dim pool As New cThreadPool
    ' pool.Create 3
    ' pool.SetUINotifier m_UINotifier
    ' 
    ' 4. 在Form中处理通知事件：
    ' Private Sub m_UINotifier_ThreadProgress(ByVal TaskId As String, ByVal CurrentValue As Long, ByVal MaxValue As Long, ByVal Message As String)
    '     ProgressBar1.Max = MaxValue
    '     ProgressBar1.Value = CurrentValue
    '     Label1.Caption = "Task " & TaskId & ": " & Message & " (" & CurrentValue & "/" & MaxValue & ")"
    ' End Sub
    ' 
    ' Private Sub m_UINotifier_ThreadCompleted(ByVal TaskId As String, ByVal Result As Variant)
    '     ListBox1.AddItem "Task " & TaskId & " completed with result: " & CStr(Result)
    ' End Sub
    ' 
    ' Private Sub m_UINotifier_ThreadError(ByVal TaskId As String, ByVal ErrorCode As Long, ByVal ErrorMessage As String)
    '     ListBox1.AddItem "Task " & TaskId & " error " & ErrorCode & ": " & ErrorMessage
    ' End Sub
    
    ' 示例任务过程 - 模拟文件处理
    Public Function FileProcessTask(ByVal param As LongPtr) As Long
        Dim task As cThread
        Set task = mThread.ReturnFromPtr(param)
        
        Dim fileCount As Long
        fileCount = 100  ' 假设要处理100个文件
        
        Dim i As Long
        For i = 1 To fileCount
            ' 检查是否请求取消
            If task.CancelRequested Then
                task.NotifyStatus "Task cancelled by user"
                FileProcessTask = -1
                Exit Function
            End If
            
            ' 模拟文件处理
            Sleep 50  ' 模拟处理时间
            
            ' 发送进度通知
            task.NotifyProgress i, fileCount, "Processing file " & i
            
            ' 模拟错误情况
            If i = 50 Then
                task.NotifyStatus "Halfway through processing..."
            End If
            
            ' 模拟偶发错误
            If i = 75 And Rnd() < 0.3 Then
                task.NotifyError 1001, "Simulated processing error at file " & i
                ' 继续处理其他文件
            End If
        Next
        
        ' 发送完成通知
        task.NotifyCompleted "Processed " & fileCount & " files successfully"
        
        FileProcessTask = 0
    End Function
    
    ' 示例任务过程 - 模拟网络下载
    Public Function DownloadTask(ByVal param As LongPtr) As Long
        Dim task As cThread
        Set task = mThread.ReturnFromPtr(param)
        
        Dim totalBytes As Long
        Dim downloadedBytes As Long
        totalBytes = 1024000  ' 1MB
        
        task.NotifyStatus "Starting download..."
        
        ' 模拟分块下载
        Dim chunkSize As Long
        chunkSize = 8192  ' 8KB chunks
        
        Do While downloadedBytes < totalBytes
            ' 检查是否请求取消
            If task.CancelRequested Then
                task.NotifyStatus "Download cancelled"
                DownloadTask = -1
                Exit Function
            End If
            
            ' 模拟下载一个数据块
            Sleep 10
            downloadedBytes = downloadedBytes + chunkSize
            If downloadedBytes > totalBytes Then downloadedBytes = totalBytes
            
            ' 发送进度通知
            Dim percent As Long
            percent = (downloadedBytes * 100) \ totalBytes
            task.NotifyProgress downloadedBytes, totalBytes, "Downloaded " & percent & "%"
            
            ' 模拟网络错误
            If percent = 50 And Rnd() < 0.2 Then
                task.NotifyError 2001, "Network connection timeout, retrying..."
                Sleep 100  ' 模拟重试延迟
            End If
        Loop
        
        task.NotifyCompleted "Download completed: " & totalBytes & " bytes"
        DownloadTask = 0
    End Function
    
    ' 示例任务过程 - 模拟数据库批处理
    Public Function DatabaseBatchTask(ByVal param As LongPtr) As Long
        Dim task As cThread
        Set task = mThread.ReturnFromPtr(param)
        
        Dim recordCount As Long
        recordCount = 500
        
        task.NotifyStatus "Connecting to database..."
        Sleep 200  ' 模拟连接时间
        
        task.NotifyStatus "Starting batch processing..."
        
        Dim i As Long
        For i = 1 To recordCount
            If task.CancelRequested Then
                task.NotifyStatus "Batch processing cancelled"
                DatabaseBatchTask = -1
                Exit Function
            End If
            
            ' 模拟记录处理
            Sleep 5
            
            ' 每10%报告一次进度
            If (i Mod (recordCount \ 10)) = 0 Then
                task.NotifyProgress i, recordCount, "Processed " & i & " records"
            End If
            
            ' 模拟数据验证错误
            If (i Mod 100) = 0 And Rnd() < 0.1 Then
                task.NotifyToUI 101, "Data validation warning at record " & i  ' 自定义通知
            End If
        Next
        
        task.NotifyStatus "Committing transaction..."
        Sleep 100
        
        task.NotifyCompleted "Batch processing completed: " & recordCount & " records processed"
        DatabaseBatchTask = 0
    End Function
    
    ' 完整的使用示例函数
    Public Sub RunUINotifierExample()
        ' 注意：这个示例需要在有UI的环境中运行
        ' 通常在Form的事件中调用
        
        Dim notifier As New cThreadUINotifier
        
        ' 创建线程池并设置UI通知器
        Dim pool As New cThreadPool
        pool.Create 3
        pool.SetUINotifier notifier
        
        ' 添加不同类型的任务
        Dim task1 As cThread
        Dim task2 As cThread
        Dim task3 As cThread
        
        Set task1 = pool.AddTask(AddressOf FileProcessTask, "Files")
        Set task2 = pool.AddTask(AddressOf DownloadTask, "Download")
        Set task3 = pool.AddTask(AddressOf DatabaseBatchTask, "Database")
        
        Debug.Print "Started 3 tasks with UI notification:"
        Debug.Print "Task 1 ID: " & task1.TaskId
        Debug.Print "Task 2 ID: " & task2.TaskId
        Debug.Print "Task 3 ID: " & task3.TaskId
        Debug.Print "Monitor the Form events to see UI notifications"
        
        ' 等待所有任务完成（可选）
        ' pool.WaitForAll 30000  ' 等待最多30秒
    End Sub

End Module
