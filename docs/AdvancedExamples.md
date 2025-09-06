# ThreadPool 高级应用示例

## 1. 批量文件处理

### 图片处理任务
```vb
Public Function ImageProcessTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    ' 获取处理参数
    Dim inputPath As String: inputPath = task.GetTaskData("inputPath")
    Dim outputPath As String: outputPath = task.GetTaskData("outputPath")
    Dim quality As Long: quality = task.GetTaskData("quality")
    Dim maxSize As Long: maxSize = task.GetTaskData("maxSize")
    
    ' 处理过程中记录进度
    task.SetTaskResult "status", "processing"
    task.SetTaskResult "progress", 0
    
    ' 执行图片处理...
    
    ' 记录结果
    task.SetTaskResult "outputSize", FileLen(outputPath)
    task.SetTaskResult "processingTime", Timer - startTime
    task.SetTaskResult "status", "completed"
    
    ImageProcessTaskProc = 0
End Function
```

### 数据导出任务
```vb
' 创建导出任务
With New cThread
    .SetTaskData "format", "CSV"
    .SetTaskData "query", "SELECT * FROM Sales WHERE Year = 2025"
    .SetTaskData "parameters", CreateObject("Scripting.Dictionary")
    .SetTaskData "outputPath", "D:\exports\sales_2025.csv"
    .Create AddressOf ExportTaskProc
End With
```

## 2. 异步API调用

### Web API 请求链
```vb
Public Function ChainedAPITaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    ' 第一个API调用
    task.SetTaskResult "step1Status", "running"
    ' ... 执行调用 ...
    task.SetTaskResult "step1Result", response1
    
    ' 基于第一个调用的结果执行第二个调用
    Dim response1 As Dictionary
    Set response1 = task.GetTaskResult("step1Result")
    
    task.SetTaskResult "step2Status", "running"
    ' ... 执行第二个调用 ...
    task.SetTaskResult "step2Result", response2
    
    ChainedAPITaskProc = 0
End Function
```

## 3. 实时数据处理

### 股票数据监控
```vb
Public Function StockMonitorTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    Dim symbols As Variant
    symbols = task.GetTaskData("symbols")  ' 股票代码数组
    Dim interval As Long
    interval = task.GetTaskData("interval") ' 更新间隔
    
    Do While Not task.CancelRequested
        ' 获取最新数据
        task.SetTaskResult "lastUpdate", Now
        task.SetTaskResult "prices", GetCurrentPrices(symbols)
        
        ' 检查预警条件
        CheckAlertConditions task
        
        Sleep interval
    Loop
    
    StockMonitorTaskProc = 0
End Function

Private Sub CheckAlertConditions(task As cThread)
    Dim alerts As Dictionary
    Set alerts = task.GetTaskData("alertConditions")
    
    ' 处理预警...
    task.SetTaskResult "triggeredAlerts", triggeredAlerts
End Sub
```

## 4. 多任务协作

### 主从任务模式
```vb
' 主任务
Public Function MasterTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    ' 创建子任务列表
    Dim subTasks As New Collection
    For i = 1 To 5
        Dim subTask As New cThread
        With subTask
            .SetTaskData "masterTask", task
            .SetTaskData "workerId", i
            .Create AddressOf WorkerTaskProc
        End With
        subTasks.Add subTask
    Next i
    
    ' 等待所有子任务完成
    task.SetTaskResult "subTasks", subTasks
    WaitForSubTasks subTasks
    
    ' 合并结果
    MergeResults task, subTasks
    
    MasterTaskProc = 0
End Function

' 工作任务
Public Function WorkerTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    Dim workerId As Long
    workerId = task.GetTaskData("workerId")
    
    ' 执行分配的工作...
    task.SetTaskResult "workerResult", result
    
    WorkerTaskProc = 0
End Function
```

## 5. 进度报告与状态更新

### 长时间运行任务
```vb
Public Function LongRunningTaskProc(ByVal param As LongPtr) As Long
    Dim task As cThread
    Set task = mThread.ReturnFromPtr(param)
    
    Dim totalSteps As Long
    totalSteps = task.GetTaskData("totalSteps")
    
    For i = 1 To totalSteps
        If task.CancelRequested Then Exit For
        
        ' 更新进度
        task.SetTaskResult "currentStep", i
        task.SetTaskResult "progress", i / totalSteps * 100
        task.SetTaskResult "status", "Processing step " & i
        
        ' 执行步骤...
        
        ' 更新步骤结果
        task.SetTaskResult "step" & i & "Result", stepResult
    Next i
    
    task.SetTaskResult "status", "Completed"
    LongRunningTaskProc = 0
End Function
```
